"""Exercise an installed PDF2MD server using only the Python standard library."""

from __future__ import annotations

import argparse
import json
import mimetypes
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from pathlib import Path


def multipart_body(
    fields: dict[str, str], file_path: Path, file_field: str = "file"
) -> tuple[bytes, str]:
    boundary = "----PDF2MDTest" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.extend([
            f"--{boundary}\r\n".encode(),
            f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode(),
            str(value).encode("utf-8"),
            b"\r\n",
        ])
    content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    chunks.extend([
        f"--{boundary}\r\n".encode(),
        f'Content-Disposition: form-data; name="{file_field}"; filename="{file_path.name}"\r\n'.encode("utf-8"),
        f"Content-Type: {content_type}\r\n\r\n".encode(),
        file_path.read_bytes(),
        b"\r\n",
        f"--{boundary}--\r\n".encode(),
    ])
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def get_json(url: str) -> dict:
    with urllib.request.urlopen(url, timeout=15) as response:
        return json.loads(response.read().decode("utf-8"))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", required=True)
    parser.add_argument("--pdf", required=True, type=Path)
    args = parser.parse_args()
    base_url = args.url.rstrip("/")
    start_request = urllib.request.Request(
        base_url + "/upload/start",
        data=json.dumps({
            "ocrMode": "auto",
            "ocrDpi": "300",
            "optImages": False,
            "optHeaders": True,
        }).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(start_request, timeout=30) as response:
        job_id = json.loads(response.read().decode("utf-8"))["job_id"]

    body, content_type = multipart_body({}, args.pdf)
    upload_request = urllib.request.Request(
        base_url + "/upload/file/" + job_id,
        data=body,
        headers={"Content-Type": content_type},
        method="POST",
    )
    with urllib.request.urlopen(upload_request, timeout=30) as response:
        uploaded = json.loads(response.read().decode("utf-8"))
    if uploaded.get("status") != "uploaded":
        raise RuntimeError(json.dumps(uploaded, ensure_ascii=False))

    finish_request = urllib.request.Request(
        base_url + "/upload/finish/" + job_id,
        data=b"",
        method="POST",
    )
    with urllib.request.urlopen(finish_request, timeout=30) as response:
        accepted = json.loads(response.read().decode("utf-8"))
    if accepted.get("status") != "running":
        raise RuntimeError(json.dumps(accepted, ensure_ascii=False))
    deadline = time.time() + 90
    status = {}
    while time.time() < deadline:
        status = get_json(base_url + "/status/" + job_id)
        if status.get("status") == "done":
            break
        time.sleep(0.5)
    if status.get("status") != "done":
        raise RuntimeError("La conversión empaquetada no terminó a tiempo")
    if status.get("progress", {}).get("done") != 1:
        raise RuntimeError(json.dumps(status, ensure_ascii=False))
    output_name = status["files"][0]["output"]
    output_url = base_url + "/download/" + job_id + "/" + urllib.parse.quote(output_name)
    with urllib.request.urlopen(output_url, timeout=20) as response:
        markdown = response.read().decode("utf-8")
    if "COLOMBIA" not in markdown.upper():
        raise RuntimeError("El OCR empaquetado no reconoció el texto de prueba")
    print(json.dumps({
        "job_id": job_id,
        "status": status["status"],
        "ocr_ready": True,
        "recognized": "COLOMBIA",
        "progress": status.get("progress"),
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
