"""File conversion engine.

Supports: pdf, docx, pptx, png, jpeg, heic (and .hecc, treated as heic).

Document conversions (anything involving docx/pptx/pdf as office docs)
use LibreOffice headless when available. Image work uses Pillow/pillow-heif,
PDF rendering uses PyMuPDF.
"""

from __future__ import annotations

import io
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIF_OK = True
except Exception:
    HEIF_OK = False

try:
    import pymupdf  # PyMuPDF
    PDF_OK = True
except Exception:
    PDF_OK = False

try:
    from pptx import Presentation
    from pptx.util import Inches
    PPTX_OK = True
except Exception:
    PPTX_OK = False

try:
    from docx import Document
    from docx.shared import Inches as DocxInches
    DOCX_OK = True
except Exception:
    DOCX_OK = False


FORMATS = ("pdf", "docx", "pptx", "png", "jpeg", "heic")

IMAGE_FORMATS = ("png", "jpeg", "heic")
OFFICE_FORMATS = ("docx", "pptx")

EXT_TO_FORMAT = {
    ".pdf": "pdf",
    ".docx": "docx",
    ".pptx": "pptx",
    ".png": "png",
    ".jpg": "jpeg",
    ".jpeg": "jpeg",
    ".heic": "heic",
    ".heif": "heic",
    ".hecc": "heic",  # common typo; treat as heic
}

FORMAT_TO_EXT = {
    "pdf": ".pdf",
    "docx": ".docx",
    "pptx": ".pptx",
    "png": ".png",
    "jpeg": ".jpg",
    "heic": ".heic",
}


class ConversionError(Exception):
    pass


def detect_format(path: str | Path) -> str | None:
    return EXT_TO_FORMAT.get(Path(path).suffix.lower())


def soffice_available() -> bool:
    return shutil.which("soffice") is not None or shutil.which("libreoffice") is not None


def _unique_path(directory: Path, stem: str, ext: str) -> Path:
    candidate = directory / f"{stem}{ext}"
    i = 2
    while candidate.exists():
        candidate = directory / f"{stem} ({i}){ext}"
        i += 1
    return candidate


# ---------------------------------------------------------------- images

def _open_image(path: Path) -> Image.Image:
    if path.suffix.lower() in (".heic", ".heif", ".hecc") and not HEIF_OK:
        raise ConversionError("HEIC support (pillow-heif) is not installed.")
    try:
        img = Image.open(path)
        img.load()
    except Exception as e:
        raise ConversionError(f"Could not read image: {e}")
    return img


def _prep_image(img: Image.Image, fmt: str) -> Image.Image:
    """Normalize an image for saving to ``fmt``."""
    if fmt == "jpeg":
        if img.mode in ("RGBA", "LA", "PA"):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            img = bg
        elif img.mode != "RGB":
            img = img.convert("RGB")
    elif fmt == "png":
        if img.mode not in ("RGB", "RGBA", "L", "LA", "P"):
            img = img.convert("RGB")
    return img


def image_to_image(src: Path, dst: Path, fmt: str) -> Path:
    img = _open_image(src)
    img = _prep_image(img, fmt)
    save_kwargs: dict = {}
    if fmt == "jpeg":
        save_kwargs = {"format": "JPEG", "quality": 95}
    elif fmt == "png":
        save_kwargs = {"format": "PNG"}
    elif fmt == "heic":
        if not HEIF_OK:
            raise ConversionError("HEIC support (pillow-heif) is not installed.")
        save_kwargs = {"format": "HEIF", "quality": 90}
    try:
        img.save(dst, **save_kwargs)
    except Exception as e:
        raise ConversionError(f"Could not write {fmt.upper()}: {e}")
    return dst


def images_to_pdf(images: list[Image.Image], dst: Path) -> Path:
    pages = [_prep_image(im, "jpeg").convert("RGB") for im in images]
    try:
        pages[0].save(dst, "PDF", save_all=True, append_images=pages[1:])
    except Exception as e:
        raise ConversionError(f"Could not write PDF: {e}")
    return dst


# ---------------------------------------------------------------- pdf rendering

def pdf_to_images(src: Path, out_dir: Path, stem: str, fmt: str, dpi: int = 200) -> list[Path]:
    if not PDF_OK:
        raise ConversionError("PyMuPDF is not installed.")
    try:
        doc = pymupdf.open(src)
    except Exception as e:
        raise ConversionError(f"Could not read PDF: {e}")
    outputs: list[Path] = []
    zoom = dpi / 72
    mat = pymupdf.Matrix(zoom, zoom)
    for i, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=mat)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        dst = _unique_path(out_dir, f"{stem}_p{i}", FORMAT_TO_EXT[fmt])
        image_to_image_via_pil(img, dst, fmt)
        outputs.append(dst)
    doc.close()
    return outputs


def image_to_image_via_pil(img: Image.Image, dst: Path, fmt: str) -> Path:
    img = _prep_image(img, fmt)
    kwargs: dict = {}
    if fmt == "jpeg":
        kwargs = {"format": "JPEG", "quality": 95}
    elif fmt == "png":
        kwargs = {"format": "PNG"}
    elif fmt == "heic":
        if not HEIF_OK:
            raise ConversionError("HEIC support (pillow-heif) is not installed.")
        kwargs = {"format": "HEIF", "quality": 90}
    try:
        img.save(dst, **kwargs)
    except Exception as e:
        raise ConversionError(f"Could not write {fmt.upper()}: {e}")
    return dst


# ---------------------------------------------------------------- office via LibreOffice

# Map target extension -> LibreOffice --convert-to filter spec.
# (Bare "docx"/"pptx" have no export filter; the full filter names are needed.)
SOFFICE_FILTERS = {
    "pdf": "pdf",
    "docx": "docx:MS Word 2007 XML",
    "pptx": "pptx:Impress MS PowerPoint 2007 XML",
}


def _soffice_convert(src: Path, out_dir: Path, to_ext: str) -> Path:
    """Convert an office document using LibreOffice headless. Returns output path."""
    binary = shutil.which("soffice") or shutil.which("libreoffice")
    if not binary:
        raise ConversionError(
            "LibreOffice is required for this conversion but was not found. "
            "Install LibreOffice and try again."
        )
    convert_to = SOFFICE_FILTERS.get(to_ext, to_ext)
    with tempfile.TemporaryDirectory() as tmp:
        cmd = [binary, "--headless", "--convert-to", convert_to,
               "--outdir", tmp, str(src)]
        try:
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        except subprocess.TimeoutExpired:
            raise ConversionError("LibreOffice conversion timed out.")
        expected = Path(tmp) / (src.stem + "." + to_ext)
        if not expected.exists():
            # LibreOffice sometimes picks a slightly different name; take first match
            matches = list(Path(tmp).glob(f"{src.stem}.*"))
            if not matches:
                err = (proc.stderr or proc.stdout or "").strip()[-500:]
                raise ConversionError(f"LibreOffice conversion failed. {err}")
            expected = matches[0]
        dst = _unique_path(out_dir, src.stem, "." + to_ext)
        shutil.move(str(expected), dst)
        return dst


# ---------------------------------------------------------------- office construction from images

def _image_to_png_bytes(img: Image.Image) -> io.BytesIO:
    buf = io.BytesIO()
    _prep_image(img, "png").save(buf, format="PNG")
    buf.seek(0)
    return buf


def images_to_pptx(images: list[Image.Image], dst: Path) -> Path:
    if not PPTX_OK:
        raise ConversionError("python-pptx is not installed.")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    for img in images:
        slide = prs.slides.add_slide(blank)
        buf = _image_to_png_bytes(img)
        # fit image inside slide, keep aspect ratio
        iw, ih = img.size
        slide_ratio = prs.slide_width / prs.slide_height
        img_ratio = iw / ih if ih else 1
        if img_ratio > slide_ratio:
            w = prs.slide_width
            h = int(prs.slide_width / img_ratio)
        else:
            h = prs.slide_height
            w = int(prs.slide_height * img_ratio)
        left = int((prs.slide_width - w) / 2)
        top = int((prs.slide_height - h) / 2)
        slide.shapes.add_picture(buf, left, top, width=w, height=h)
    try:
        prs.save(dst)
    except Exception as e:
        raise ConversionError(f"Could not write PPTX: {e}")
    return dst


def images_to_docx(images: list[Image.Image], dst: Path) -> Path:
    if not DOCX_OK:
        raise ConversionError("python-docx is not installed.")
    doc = Document()
    for i, img in enumerate(images):
        if i > 0:
            doc.add_page_break()
        buf = _image_to_png_bytes(img)
        # scale to fit page width (6.5") keeping aspect ratio
        iw, ih = img.size
        width = DocxInches(6.5)
        height = DocxInches(6.5 * ih / iw) if iw else DocxInches(6.5)
        doc.add_picture(buf, width=width, height=height)
    try:
        doc.save(dst)
    except Exception as e:
        raise ConversionError(f"Could not write DOCX: {e}")
    return dst


# ---------------------------------------------------------------- main entry point

def convert_file(src: str | Path, target: str, out_dir: str | Path) -> list[Path]:
    """Convert ``src`` to ``target`` format, writing into ``out_dir``.

    Returns the list of files created (usually one; several for
    multi-page PDF/image outputs).
    """
    src = Path(src)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    if target not in FORMATS:
        raise ConversionError(f"Unknown target format: {target}")
    source = detect_format(src)
    if source is None:
        raise ConversionError(f"Unsupported file type: {src.suffix or '(no extension)'}")
    if source == target:
        # same type -> still produce a new file (a copy)
        dst = _unique_path(out_dir, f"{src.stem}_copy", FORMAT_TO_EXT[target])
        shutil.copy2(src, dst)
        return [dst]

    stem = src.stem
    ext = FORMAT_TO_EXT[target]

    # ---- image -> *
    if source in IMAGE_FORMATS:
        img = _open_image(src)
        if target in IMAGE_FORMATS:
            return [image_to_image(src, _unique_path(out_dir, stem, ext), target)]
        if target == "pdf":
            return [images_to_pdf([img], _unique_path(out_dir, stem, ext))]
        if target == "pptx":
            return [images_to_pptx([img], _unique_path(out_dir, stem, ext))]
        if target == "docx":
            return [images_to_docx([img], _unique_path(out_dir, stem, ext))]

    # ---- pdf -> *
    if source == "pdf":
        if target in IMAGE_FORMATS:
            return pdf_to_images(src, out_dir, stem, target)
        if target == "pptx":
            pages = _pdf_page_images(src)
            return [images_to_pptx(pages, _unique_path(out_dir, stem, ext))]
        if target == "docx":
            pages = _pdf_page_images(src)
            return [images_to_docx(pages, _unique_path(out_dir, stem, ext))]

    # ---- office (docx/pptx) -> *
    if source in OFFICE_FORMATS:
        if target == "pdf":
            return [_soffice_convert(src, out_dir, "pdf")]
        if target in OFFICE_FORMATS:
            if source == "pptx" and target == "docx":
                # LibreOffice cannot export Impress directly to DOCX;
                # go via PDF and embed one slide image per page.
                with tempfile.TemporaryDirectory() as tmp:
                    pdf_path = _soffice_convert(src, Path(tmp), "pdf")
                    pages = _pdf_page_images(pdf_path)
                    return [images_to_docx(pages, _unique_path(out_dir, stem, ext))]
            return [_soffice_convert(src, out_dir, target)]
        if target in IMAGE_FORMATS:
            # office -> pdf -> render pages
            with tempfile.TemporaryDirectory() as tmp:
                pdf_path = _soffice_convert(src, Path(tmp), "pdf")
                return pdf_to_images(pdf_path, out_dir, stem, target)

    raise ConversionError(f"Conversion from {source} to {target} is not supported.")


def _pdf_page_images(src: Path, dpi: int = 200) -> list[Image.Image]:
    if not PDF_OK:
        raise ConversionError("PyMuPDF is not installed.")
    try:
        doc = pymupdf.open(src)
    except Exception as e:
        raise ConversionError(f"Could not read PDF: {e}")
    zoom = dpi / 72
    mat = pymupdf.Matrix(zoom, zoom)
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        images.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
    doc.close()
    if not images:
        raise ConversionError("PDF has no pages.")
    return images
