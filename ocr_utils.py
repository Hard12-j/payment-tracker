"""OCR via the OCR.space API (no OpenCV / system libraries needed)."""
import io

import requests
from PIL import Image

from settings import get_setting

API_URL = "https://api.ocr.space/parse/image"


def _shrink(image_bytes: bytes, max_bytes: int = 900_000) -> bytes:
    """Free tier accepts files up to ~1 MB, so re-encode as JPEG and shrink if needed."""
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    quality = 90
    while True:
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=quality)
        if buf.tell() <= max_bytes or (quality <= 40 and img.width < 700):
            return buf.getvalue()
        if quality > 50:
            quality -= 15
        else:
            img = img.resize((int(img.width * 0.8), int(img.height * 0.8)))


def extract_text(image_bytes: bytes) -> str:
    key = get_setting("OCR_SPACE_API_KEY")
    if not key:
        raise RuntimeError("OCR_SPACE_API_KEY is not set")
    data = _shrink(image_bytes)
    resp = requests.post(
        API_URL,
        files={"file": ("screenshot.jpg", data, "image/jpeg")},
        data={"apikey": key, "language": "eng", "OCREngine": "2",
              "scale": "true", "isOverlayRequired": "false"},
        timeout=60,
    )
    resp.raise_for_status()
    j = resp.json()
    if j.get("IsErroredOnProcessing"):
        raise RuntimeError(str(j.get("ErrorMessage")))
    return "\n".join(r.get("ParsedText", "") for r in j.get("ParsedResults", []))