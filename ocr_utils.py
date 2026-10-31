"""OCR wrapper. RapidOCR is pip-only (no separate Tesseract install needed)."""
import io

import numpy as np
from PIL import Image

_engine = None


def _get_engine():
    global _engine
    if _engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _engine = RapidOCR()
    return _engine


def extract_text(image_bytes: bytes) -> str:
    """Returns all recognised text, one line per OCR box, top-to-bottom."""
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    # upscale small screenshots: OCR accuracy drops on tiny text
    if img.width < 900:
        scale = 900 / img.width
        img = img.resize((900, int(img.height * scale)))
    result, _ = _get_engine()(np.array(img))
    if not result:
        return ""
    return "\n".join(item[1] for item in result)
