import asyncio
import base64
import hmac
import json
import logging
import os
import re
import subprocess
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import fitz  # PyMuPDF
import requests
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import JSONResponse, Response
from PIL import Image
from pydantic import BaseModel, Field

logger = logging.getLogger("doc_converter")
logging.basicConfig(level=logging.INFO)

API_KEY = os.environ.get("WSM_API_KEY", "")

SF_ID_PATTERN = re.compile(r"^[a-zA-Z0-9]{15,18}$")
INSTANCE_URL_PATTERN = re.compile(r"^https://[a-zA-Z0-9.-]+\.my\.salesforce\.com$")
API_VERSION_PATTERN = re.compile(r"^\d{2,3}\.\d$")

MAX_DOCS = 50
MAX_BODY_BYTES = 200_000_000  # 200MB for base64-encoded files
MAX_FILE_SIZE = 100_000_000  # 100MB per decoded file
MAX_CONCURRENT_JOBS = 3
ABORT_THRESHOLD_SECONDS = 280
LIBREOFFICE_TIMEOUT = 120

OFFICE_EXTENSIONS = {
    "docx", "doc", "xlsx", "xls", "pptx", "ppt",
    "odt", "ods", "odp", "rtf", "csv",
}
IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "tiff", "tif", "bmp", "gif", "webp"}
PDF_EXTENSION = "pdf"

_job_semaphore = threading.Semaphore(MAX_CONCURRENT_JOBS)

app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)


class ConvertRequest(BaseModel):
    file_base64: str
    filename: str


class ConvertAndMergeRequest(BaseModel):
    documents: list[ConvertRequest] = Field(..., min_length=1, max_length=MAX_DOCS)
    output_filename: str = "Merged Document.pdf"
    return_individual: bool = False


class ConvertSfRequest(BaseModel):
    content_document_ids: list[str] = Field(..., min_length=1, max_length=MAX_DOCS)
    instance_url: str
    access_token: str
    parent_record_id: str
    output_filename: str = "Merged Document.pdf"
    merge: bool = True
    upload_individual: bool = False
    api_version: str = "63.0"


def _verify_api_key(request: Request):
    if not API_KEY:
        raise HTTPException(status_code=500, detail="WSM_API_KEY not configured")
    provided = request.headers.get("X-API-Key", "")
    if not hmac.compare_digest(provided, API_KEY):
        raise HTTPException(status_code=401, detail="Invalid API key")


def _get_extension(filename: str) -> str:
    return Path(filename).suffix.lstrip(".").lower()


def _convert_office_to_pdf(input_path: Path, output_dir: Path) -> Path:
    env = os.environ.copy()
    user_profile = tempfile.mkdtemp(prefix="lo_profile_")
    env["HOME"] = user_profile

    try:
        result = subprocess.run(
            [
                "soffice",
                "--headless",
                "--nofirststartwizard",
                f"-env:UserInstallation=file://{user_profile}",
                "--convert-to", "pdf",
                "--outdir", str(output_dir),
                str(input_path),
            ],
            check=True,
            timeout=LIBREOFFICE_TIMEOUT,
            capture_output=True,
            text=True,
            env=env,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"LibreOffice conversion timed out after {LIBREOFFICE_TIMEOUT}s")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"LibreOffice conversion failed: {e.stderr}")
    finally:
        import shutil
        shutil.rmtree(user_profile, ignore_errors=True)

    pdf_path = output_dir / (input_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError(f"LibreOffice produced no output. stderr: {result.stderr}")
    return pdf_path


def _convert_image_to_pdf(input_path: Path, output_dir: Path) -> Path:
    img = Image.open(input_path)
    img_width, img_height = img.size
    img.close()

    doc = fitz.open()
    page = doc.new_page(width=img_width, height=img_height)
    page.insert_image(fitz.Rect(0, 0, img_width, img_height), filename=str(input_path))

    pdf_path = output_dir / (input_path.stem + ".pdf")
    doc.save(str(pdf_path))
    doc.close()
    return pdf_path


def _convert_file_to_pdf(file_bytes: bytes, filename: str, work_dir: Path) -> bytes:
    ext = _get_extension(filename)

    if ext == PDF_EXTENSION:
        if file_bytes[:4] != b"%PDF":
            raise ValueError(f"{filename} has .pdf extension but is not a valid PDF")
        return file_bytes

    safe_name = re.sub(r'[^\w.\-]', '_', filename)
    input_path = work_dir / safe_name
    input_path.write_bytes(file_bytes)

    output_dir = work_dir / "output"
    output_dir.mkdir(exist_ok=True)

    if ext in OFFICE_EXTENSIONS:
        pdf_path = _convert_office_to_pdf(input_path, output_dir)
    elif ext in IMAGE_EXTENSIONS:
        pdf_path = _convert_image_to_pdf(input_path, output_dir)
    else:
        raise ValueError(f"Unsupported file type: .{ext}")

    return pdf_path.read_bytes()


def _merge_pdfs(pdf_list: list[bytes], output_filename: str) -> bytes:
    merged = fitz.open()
    page_counts = []
    try:
        for i, pdf_bytes in enumerate(pdf_list):
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            page_counts.append(len(doc))
            merged.insert_pdf(doc)
            doc.close()

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            merged_path = tmp.name
        merged.save(merged_path)
    finally:
        merged.close()

    try:
        merged_bytes = Path(merged_path).read_bytes()
    finally:
        os.unlink(merged_path)

    return merged_bytes, page_counts


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/convert")
async def convert(request: Request):
    _verify_api_key(request)

    body_bytes = await request.body()
    if len(body_bytes) > MAX_BODY_BYTES:
        raise HTTPException(status_code=400, detail="Request body too large")

    try:
        data = ConvertRequest.model_validate_json(body_bytes)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    ext = _get_extension(data.filename)
    if ext not in OFFICE_EXTENSIONS and ext not in IMAGE_EXTENSIONS and ext != PDF_EXTENSION:
        raise HTTPException(status_code=400, detail=f"Unsupported file type: .{ext}")

    try:
        file_bytes = base64.b64decode(data.file_base64)
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid base64 encoding")

    if len(file_bytes) > MAX_FILE_SIZE:
        raise HTTPException(status_code=400, detail=f"File too large ({len(file_bytes)} bytes, max {MAX_FILE_SIZE})")

    acquired = _job_semaphore.acquire(blocking=False)
    if not acquired:
        raise HTTPException(status_code=429, detail="Service is busy. Try again later.")

    try:
        pdf_bytes = await asyncio.to_thread(_do_convert, file_bytes, data.filename)
    finally:
        _job_semaphore.release()

    pdf_filename = Path(data.filename).stem + ".pdf"
    return Response(
        content=pdf_bytes,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{pdf_filename}"'},
    )


def _do_convert(file_bytes: bytes, filename: str) -> bytes:
    with tempfile.TemporaryDirectory(prefix="conv_") as work_dir:
        return _convert_file_to_pdf(file_bytes, filename, Path(work_dir))


@app.post("/convert-and-merge")
async def convert_and_merge(request: Request):
    _verify_api_key(request)

    body_bytes = await request.body()
    if len(body_bytes) > MAX_BODY_BYTES:
        raise HTTPException(status_code=400, detail="Request body too large")

    try:
        data = ConvertAndMergeRequest.model_validate_json(body_bytes)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    sanitized = os.path.basename(data.output_filename).strip()
    if not sanitized or len(sanitized) > 255 or any(ord(c) < 32 for c in sanitized):
        raise HTTPException(status_code=400, detail="Invalid output filename")
    if not sanitized.endswith(".pdf"):
        sanitized += ".pdf"
    data.output_filename = sanitized

    for doc in data.documents:
        ext = _get_extension(doc.filename)
        if ext not in OFFICE_EXTENSIONS and ext not in IMAGE_EXTENSIONS and ext != PDF_EXTENSION:
            raise HTTPException(status_code=400, detail=f"Unsupported file type for {doc.filename}: .{ext}")

    acquired = _job_semaphore.acquire(blocking=False)
    if not acquired:
        raise HTTPException(status_code=429, detail="Service is busy. Try again later.")

    try:
        result = await asyncio.to_thread(_do_convert_and_merge, data)
    finally:
        _job_semaphore.release()

    return result


def _do_convert_and_merge(data: ConvertAndMergeRequest) -> dict:
    converted_pdfs = []
    individual_results = []

    with tempfile.TemporaryDirectory(prefix="batch_") as work_dir:
        work_path = Path(work_dir)

        for i, doc in enumerate(data.documents):
            try:
                file_bytes = base64.b64decode(doc.file_base64)
            except Exception:
                raise HTTPException(status_code=400, detail=f"Invalid base64 for document {i + 1}: {doc.filename}")

            if len(file_bytes) > MAX_FILE_SIZE:
                raise HTTPException(status_code=400, detail=f"File too large: {doc.filename}")

            doc_dir = work_path / f"doc_{i}"
            doc_dir.mkdir()

            pdf_bytes = _convert_file_to_pdf(file_bytes, doc.filename, doc_dir)
            converted_pdfs.append(pdf_bytes)

            if data.return_individual:
                individual_results.append({
                    "filename": Path(doc.filename).stem + ".pdf",
                    "pdf_base64": base64.b64encode(pdf_bytes).decode("utf-8"),
                })

    merged_bytes, page_counts = _merge_pdfs(converted_pdfs, data.output_filename)

    result = {
        "success": True,
        "merged_pdf_base64": base64.b64encode(merged_bytes).decode("utf-8"),
        "output_filename": data.output_filename,
        "total_pages": sum(page_counts),
        "page_counts": page_counts,
        "file_size_bytes": len(merged_bytes),
        "documents_converted": len(data.documents),
    }

    if data.return_individual:
        result["individual_pdfs"] = individual_results

    return result


@app.post("/convert-sf")
async def convert_sf(request: Request):
    _verify_api_key(request)

    body_bytes = await request.body()
    if len(body_bytes) > MAX_BODY_BYTES:
        raise HTTPException(status_code=400, detail="Request body too large")

    try:
        data = ConvertSfRequest.model_validate_json(body_bytes)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

    for doc_id in data.content_document_ids:
        if not SF_ID_PATTERN.match(doc_id):
            raise HTTPException(status_code=400, detail=f"Invalid ContentDocumentId: {doc_id}")

    if not INSTANCE_URL_PATTERN.match(data.instance_url):
        raise HTTPException(status_code=400, detail="Invalid instance_url")

    if not SF_ID_PATTERN.match(data.parent_record_id):
        raise HTTPException(status_code=400, detail=f"Invalid parent_record_id: {data.parent_record_id}")

    if not API_VERSION_PATTERN.match(data.api_version):
        raise HTTPException(status_code=400, detail="Invalid api_version format")

    sanitized = os.path.basename(data.output_filename).strip()
    if not sanitized or len(sanitized) > 255 or any(ord(c) < 32 for c in sanitized):
        raise HTTPException(status_code=400, detail="Invalid output filename")
    if not sanitized.endswith(".pdf"):
        sanitized += ".pdf"
    data.output_filename = sanitized

    acquired = _job_semaphore.acquire(blocking=False)
    if not acquired:
        raise HTTPException(status_code=429, detail="Service is busy. Try again later.")

    try:
        return await asyncio.to_thread(_do_convert_sf, data)
    finally:
        _job_semaphore.release()


def _do_convert_sf(data: ConvertSfRequest):
    start = time.time()
    org_domain = data.instance_url.split("//")[1]
    auth_headers = {
        "Authorization": f"Bearer {data.access_token}",
        "Content-Type": "application/json",
    }
    base_url = f"{data.instance_url}/services/data/v{data.api_version}"

    logger.info("Convert-SF started: org=%s docs=%d merge=%s", org_domain, len(data.content_document_ids), data.merge)

    try:
        id_list = "','".join(data.content_document_ids)
        soql = (
            f"SELECT Id, ContentDocumentId, Title, FileExtension, ContentSize "
            f"FROM ContentVersion "
            f"WHERE ContentDocumentId IN ('{id_list}') AND IsLatest = true"
        )
        query_resp = requests.get(
            f"{base_url}/query",
            headers=auth_headers,
            params={"q": soql},
            timeout=30,
        )
        if query_resp.status_code != 200:
            return _error_response(f"Failed to query ContentVersions (HTTP {query_resp.status_code}): {query_resp.text[:200]}", 502)

        records = query_resp.json().get("records", [])
        if not records:
            return _error_response(f"No ContentVersions found for the provided document IDs", 404)

        doc_id_to_record = {r["ContentDocumentId"]: r for r in records}
        ordered_records = [doc_id_to_record[did] for did in data.content_document_ids if did in doc_id_to_record]

        if len(ordered_records) != len(data.content_document_ids):
            found = {r["ContentDocumentId"] for r in ordered_records}
            missing = [did for did in data.content_document_ids if did not in found]
            return _error_response(f"Documents not found: {', '.join(missing)}", 404)

        download_headers = {"Authorization": f"Bearer {data.access_token}"}

        def download_one(index, cv_id):
            for attempt in range(3):
                if time.time() - start > ABORT_THRESHOLD_SECONDS:
                    raise Exception("Approaching timeout limit, aborting")
                resp = requests.get(
                    f"{data.instance_url}/services/data/v{data.api_version}/sobjects/ContentVersion/{cv_id}/VersionData",
                    headers=download_headers,
                    timeout=60,
                )
                if resp.status_code == 200:
                    return index, resp.content
                if resp.status_code in (429, 503) and attempt < 2:
                    time.sleep(2 ** attempt)
                    continue
                raise Exception(f"Download failed for document {index + 1} (HTTP {resp.status_code})")
            raise Exception(f"Download failed for document {index + 1} after retries")

        file_data_list = [None] * len(ordered_records)
        with ThreadPoolExecutor(max_workers=3) as executor:
            futures = {
                executor.submit(download_one, i, r["Id"]): i
                for i, r in enumerate(ordered_records)
            }
            for future in as_completed(futures):
                try:
                    idx, content = future.result()
                    file_data_list[idx] = content
                except Exception:
                    for f in futures:
                        f.cancel()
                    raise

        converted_pdfs = []
        conversion_results = []
        with tempfile.TemporaryDirectory(prefix="sf_conv_") as work_dir:
            work_path = Path(work_dir)

            for i, record in enumerate(ordered_records):
                file_bytes = file_data_list[i]
                file_data_list[i] = None
                ext = (record.get("FileExtension") or "").lower()
                title = record.get("Title", f"document_{i}")
                filename = f"{title}.{ext}" if ext else title

                doc_dir = work_path / f"doc_{i}"
                doc_dir.mkdir()

                try:
                    pdf_bytes = _convert_file_to_pdf(file_bytes, filename, doc_dir)
                    converted_pdfs.append(pdf_bytes)
                    conversion_results.append({
                        "title": title,
                        "original_type": ext,
                        "status": "converted",
                        "content_document_id": record["ContentDocumentId"],
                    })
                except Exception as e:
                    return _error_response(
                        f"Failed to convert document {i + 1} ({title}.{ext}): {str(e)}", 400
                    )

                del file_bytes

        uploaded_ids = []

        if data.upload_individual and len(converted_pdfs) > 1:
            for i, pdf_bytes in enumerate(converted_pdfs):
                if time.time() - start > ABORT_THRESHOLD_SECONDS:
                    return _error_response("Operation timed out before upload", 504)

                individual_name = f"{ordered_records[i]['Title']}.pdf"
                cv_id = _upload_to_salesforce(
                    pdf_bytes, individual_name, data.parent_record_id,
                    base_url, data.access_token, data.output_filename,
                )
                if cv_id:
                    uploaded_ids.append({"title": individual_name, "content_version_id": cv_id})

        if data.merge:
            merged_bytes, page_counts = _merge_pdfs(converted_pdfs, data.output_filename)

            if time.time() - start > ABORT_THRESHOLD_SECONDS:
                return _error_response("Operation timed out before upload", 504)

            merged_cv_id = _upload_to_salesforce(
                merged_bytes, data.output_filename, data.parent_record_id,
                base_url, data.access_token, data.output_filename,
            )

            duration = round(time.time() - start, 2)
            logger.info(
                "Convert-SF complete: org=%s docs=%d pages=%d duration=%.2fs",
                org_domain, len(ordered_records), sum(page_counts), duration,
            )

            result = {
                "success": True,
                "merged_content_version_id": merged_cv_id,
                "total_pages": sum(page_counts),
                "page_counts": page_counts,
                "file_size_bytes": len(merged_bytes),
                "documents_converted": len(converted_pdfs),
                "conversion_results": conversion_results,
            }
            if uploaded_ids:
                result["individual_uploads"] = uploaded_ids
            return result

        else:
            if not uploaded_ids:
                for i, pdf_bytes in enumerate(converted_pdfs):
                    if time.time() - start > ABORT_THRESHOLD_SECONDS:
                        return _error_response("Operation timed out before upload", 504)

                    individual_name = f"{ordered_records[i]['Title']}.pdf"
                    cv_id = _upload_to_salesforce(
                        pdf_bytes, individual_name, data.parent_record_id,
                        base_url, data.access_token, data.output_filename,
                    )
                    if cv_id:
                        uploaded_ids.append({"title": individual_name, "content_version_id": cv_id})

            duration = round(time.time() - start, 2)
            logger.info(
                "Convert-SF complete (no merge): org=%s docs=%d duration=%.2fs",
                org_domain, len(ordered_records), duration,
            )

            return {
                "success": True,
                "documents_converted": len(converted_pdfs),
                "conversion_results": conversion_results,
                "individual_uploads": uploaded_ids,
            }

    except HTTPException:
        raise
    except Exception as e:
        logger.error("Convert-SF failed: org=%s error=%s", org_domain, str(e))
        return _error_response(str(e), 500)


def _upload_to_salesforce(
    pdf_bytes: bytes, filename: str, parent_id: str,
    base_url: str, access_token: str, fallback_name: str,
) -> str | None:
    upload_url = f"{base_url}/sobjects/ContentVersion"
    title = filename.removesuffix(".pdf") if filename.endswith(".pdf") else filename
    entity = {
        "Title": title,
        "PathOnClient": filename if filename.endswith(".pdf") else filename + ".pdf",
        "FirstPublishLocationId": parent_id,
    }
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name

    try:
        with open(tmp_path, "rb") as f:
            upload_resp = requests.post(
                upload_url,
                headers={"Authorization": f"Bearer {access_token}"},
                files={
                    "entity_content": (None, json.dumps(entity), "application/json"),
                    "VersionData": (filename, f, "application/pdf"),
                },
                timeout=120,
            )
    finally:
        os.unlink(tmp_path)

    if upload_resp.status_code not in (200, 201):
        logger.error("Upload failed for %s: HTTP %d", filename, upload_resp.status_code)
        return None

    return upload_resp.json().get("id")


def _error_response(message: str, status_code: int) -> JSONResponse:
    return JSONResponse(
        status_code=status_code,
        content={"success": False, "error": message},
    )
