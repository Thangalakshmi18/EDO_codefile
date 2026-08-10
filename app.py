"""
pdf_diagram_padding_check.py
=============================

Standalone tool to test the PDF diagram-extraction logic from
`imageedo_final.py` (functions: is_valid_rect, _get_drawing_region,
extract_pdf_page_diagrams) and visually check how the crop `padding`
value affects the extracted images.

It launches a small local web page where you can:
  - enter/confirm the PDF file path
  - adjust the padding (in points, same unit as the original PAD var)
  - click "Extract" and see every extracted diagram image inline

Run:
    pip install flask pymupdf --break-system-packages   # if not installed
    python pdf_diagram_padding_check.py --pdf "/path/to/your.pdf"

Then open the printed URL (default http://127.0.0.1:5050) in a browser.
"""

import os
import io
import base64
import logging
import argparse

import fitz  # PyMuPDF
from flask import Flask, request, render_template_string

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# ----------------------------------------------------------------
# Defaults (mirrors the constants in imageedo_final.py)
# ----------------------------------------------------------------
DEFAULT_PDF_PATH = "C:\\Users\\Ranjith\\OneDrive\\Desktop\\edo\\181995 Rev 4.pdf"       # filled in from --pdf at startup, or typed in the form
DEFAULT_PADDING_X = 40
DEFAULT_PADDING_Y = 40

MIN_PADDING_X = 0
MIN_PADDING_Y = 0


# ==================================================================
# ---- Ported unchanged from imageedo_final.py --------------------
# ==================================================================

def is_valid_rect(rect):
    """True if a PyMuPDF rect is non-empty, finite, and has real size."""
    if rect is None:
        return False
    if rect.is_empty or rect.is_infinite:
        return False
    if rect.width <= 1 or rect.height <= 1:
        return False
    return True


def _get_drawing_region(page, proximity=150, paragraph_char_threshold=80,
                         paragraph_height_threshold=40, band_vertical_tolerance=60):
    """
    Finds the bounding box of just the actual line-art drawing on the
    page - excludes large text paragraphs but includes small
    annotation/callout/dimension labels sitting next to the drawing.

    Returns a fitz.Rect, or None if no drawing content was found.
    """
    page_w, page_h = page.rect.width, page.rect.height
    content_bboxes = []

    # Embedded raster images (if any)
    for img in page.get_images(full=True):
        try:
            xref = img[0]
            for bbox in page.get_image_rects(xref):
                if is_valid_rect(bbox):
                    content_bboxes.append(fitz.Rect(bbox))
        except Exception:
            pass

    # Vector line-art (the CAD drawing itself)
    try:
        raw_rects = []
        for d in page.get_drawings():
            r = fitz.Rect(d["rect"])
            if is_valid_rect(r) and r.width > 3 and r.height > 3:
                raw_rects.append(r)
    except Exception:
        raw_rects = []

    # Drop page-border / zone-grid / table-gridline strokes
    raw_rects = [
        r for r in raw_rects
        if not (r.width > 0.85 * page_w and r.height < 0.02 * page_h)
        and not (r.height > 0.85 * page_h and r.width < 0.02 * page_w)
    ]

    used = [False] * len(raw_rects)
    clusters = []
    for i, r in enumerate(raw_rects):
        if used[i]:
            continue
        cluster = fitz.Rect(r)
        changed = True
        while changed:
            changed = False
            for j, r2 in enumerate(raw_rects):
                if used[j]:
                    continue
                expanded = fitz.Rect(cluster.x0 - 40, cluster.y0 - 40, cluster.x1 + 40, cluster.y1 + 40)
                if expanded.intersects(r2):
                    cluster |= r2
                    used[j] = True
                    changed = True
        used[i] = True
        clusters.append(cluster)

    clusters = [
        c for c in clusters
        if c.width > 30 and c.height > 30
        and not (c.width > 0.9 * page_w and c.height > 0.9 * page_h)
    ]
    content_bboxes.extend(clusters)

    if not content_bboxes:
        return None

    main_bbox = max(content_bboxes, key=lambda r: r.width * r.height)
    changed = True
    while changed:
        changed = False
        for r in content_bboxes:
            if r is main_bbox or main_bbox.contains(r):
                continue
            vertical_overlap = min(main_bbox.y1, r.y1) - max(main_bbox.y0, r.y0)
            shares_band = vertical_overlap > -band_vertical_tolerance
            expanded = fitz.Rect(main_bbox.x0 - proximity, main_bbox.y0 - proximity,
                                  main_bbox.x1 + proximity, main_bbox.y1 + proximity)
            if shares_band or expanded.intersects(r):
                merged = fitz.Rect(main_bbox)
                merged |= r
                if merged != main_bbox:
                    main_bbox = merged
                    changed = True

    for block in page.get_text("blocks"):
        bx0, by0, bx1, by1, text = block[0], block[1], block[2], block[3], block[4]
        rect = fitz.Rect(bx0, by0, bx1, by1)
        is_paragraph = (
            len(text.strip()) > paragraph_char_threshold
            or rect.height > paragraph_height_threshold
        )
        if is_paragraph:
            continue
        expanded = fitz.Rect(main_bbox.x0 - proximity, main_bbox.y0 - proximity,
                              main_bbox.x1 + proximity, main_bbox.y1 + proximity)
        if expanded.intersects(rect):
            main_bbox |= rect

    return main_bbox


def extract_pdf_page_diagrams(
    pdf_path,
    padding_left=0,
    padding_right=0,
    padding_top=50,
    padding_bottom=342,
    output_dir=None
):
    """
    Extract the required PDF region using independent
    left/right/top/bottom padding.

    For the current 181995 Rev 4.pdf target:

        left   = 0
        right  = 0
        top    = 50
        bottom = 342

    Result:

        x0 = 0
        y0 = 50
        x1 = 1191
        y1 = 500
    """

    diagrams = []

    if not pdf_path or not os.path.isfile(pdf_path):
        logging.error(
            f"PDF not found: {pdf_path!r}"
        )
        return diagrams

    try:
        padding_left = max(0, int(padding_left))
    except (TypeError, ValueError):
        padding_left = 0

    try:
        padding_right = max(0, int(padding_right))
    except (TypeError, ValueError):
        padding_right = 0

    try:
        padding_top = max(0, int(padding_top))
    except (TypeError, ValueError):
        padding_top = 0

    try:
        padding_bottom = max(0, int(padding_bottom))
    except (TypeError, ValueError):
        padding_bottom = 0

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        logging.error(f"Failed to open PDF: {e}")
        return diagrams

    try:

        for page_num in range(len(doc)):

            page = doc[page_num]
            page_no = page_num + 1

            # ------------------------------------------------------
            # PAGE-BASED CROP
            # ------------------------------------------------------

            crop_x0 = max(
                page.rect.x0,
                page.rect.x0 + padding_left
            )

            crop_y0 = max(
                page.rect.y0,
                page.rect.y0 + padding_top
            )

            crop_x1 = min(
                page.rect.x1,
                page.rect.x1 - padding_right
            )

            crop_y1 = min(
                page.rect.y1,
                page.rect.y1 - padding_bottom
            )

            clip_rect = fitz.Rect(
                crop_x0,
                crop_y0,
                crop_x1,
                crop_y1
            )

            logging.info(
                f"PAGE {page_no} -> "
                f"LEFT={padding_left}, "
                f"RIGHT={padding_right}, "
                f"TOP={padding_top}, "
                f"BOTTOM={padding_bottom}"
            )

            logging.info(
                f"PAGE {page_no} CROP RECT -> "
                f"({clip_rect.x0:.1f}, "
                f"{clip_rect.y0:.1f}, "
                f"{clip_rect.x1:.1f}, "
                f"{clip_rect.y1:.1f})"
            )

            # ------------------------------------------------------
            # RENDER
            # ------------------------------------------------------

            pix = page.get_pixmap(
                matrix=fitz.Matrix(3, 3),
                clip=clip_rect,
                alpha=False
            )

            filename = (
                f"NEW_EDO_DIAGRAM_"
                f"p{page_no}_"
                f"L{padding_left}_"
                f"R{padding_right}_"
                f"T{padding_top}_"
                f"B{padding_bottom}.png"
            )

            data = pix.tobytes("png")

            if output_dir:
                os.makedirs(output_dir, exist_ok=True)

                filepath = os.path.join(
                    output_dir,
                    filename
                )

                with open(filepath, "wb") as f:
                    f.write(data)

                logging.info(
                    f"PAGE {page_no} CAPTURED -> {filepath}"
                )

            diagrams.append({
                "name": filename,
                "bytes": data,
                "extension": ".png",
                "page": page_no,
                "rect": (
                    round(clip_rect.x0, 1),
                    round(clip_rect.y0, 1),
                    round(clip_rect.x1, 1),
                    round(clip_rect.y1, 1)
                ),
                "padding_left": padding_left,
                "padding_right": padding_right,
                "padding_top": padding_top,
                "padding_bottom": padding_bottom
            })

    finally:
        doc.close()

    logging.info(
        f"TOTAL DIAGRAMS CAPTURED: {len(diagrams)}"
    )

    return diagrams

# ==================================================================
# ---- Web UI -------------------------------------------------------
# ==================================================================

app = Flask(__name__)

PAGE_TEMPLATE = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>PDF Diagram Extraction - Padding Check</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 2rem; background: #f7f7f9; color: #222; }
    h1 { font-size: 1.4rem; }
    form { background: #fff; padding: 1rem 1.2rem; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,.1); margin-bottom: 1.5rem; }
    label { display: block; margin-top: .6rem; font-weight: 600; font-size: .9rem; }
    input[type=text], input[type=number] { width: 100%; padding: .4rem; margin-top: .2rem; box-sizing: border-box; }
    button { margin-top: 1rem; padding: .5rem 1.2rem; background: #2563eb; color: #fff; border: none; border-radius: 6px; cursor: pointer; }
    button:hover { background: #1d4ed8; }
    .error { color: #b91c1c; font-weight: 600; }
    .gallery { display: flex; flex-wrap: wrap; gap: 1rem; }
    .card { background: #fff; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,.1); padding: .6rem; max-width: 420px; }
    .card img { max-width: 100%; border: 1px solid #ddd; }
    .meta { font-size: .8rem; color: #555; margin-top: .3rem; }
    .count { margin-bottom: 1rem; color: #444; }
  </style>
</head>
<body>
  <h1>PDF Diagram Extraction — Padding Check</h1>
  <form method="get" action="/">
    <label>PDF path</label>
    <input type="text" name="pdf" value="{{ pdf_path }}" placeholder="/full/path/to/file.pdf">
    <label>Padding X (points)</label>
    <input type="number" name="padx" value="{{ padding_x }}" min="0" step="1">
    <label>Padding Y (points)</label>
    <input type="number" name="pady" value="{{ padding_y }}" min="0" step="1">
    <button type="submit">Extract</button>
  </form>

  {% if error %}
    <p class="error">{{ error }}</p>
  {% endif %}

  {% if diagrams is not none %}
    <p class="count">{{ diagrams|length }} diagram(s) extracted with padding_x = {{ padding_x }}, padding_y = {{ padding_y }}.</p>
    <div class="gallery">
      {% for d in diagrams %}
        <div class="card">
          <img src="data:image/png;base64,{{ d.b64 }}" alt="{{ d.name }}">
          <div class="meta">
            Page {{ d.page }} — {{ d.name }}<br>
            Crop rect: {{ d.rect }}
          </div>
        </div>
      {% endfor %}
    </div>
  {% endif %}
</body>
</html>
"""


@app.route("/", methods=["GET"])
def index():
    pdf_path = request.args.get("pdf", DEFAULT_PDF_PATH or "").strip()
    padding_x_raw = request.args.get("padx", str(DEFAULT_PADDING_X))
    padding_y_raw = request.args.get("pady", str(DEFAULT_PADDING_Y))
    try:
        padding_x = max(0, int(padding_x_raw))
    except ValueError:
        padding_x = DEFAULT_PADDING_X
    try:
        padding_y = max(0, int(padding_y_raw))
    except ValueError:
        padding_y = DEFAULT_PADDING_Y

    diagrams = None
    error = None

    # Only run extraction once a pdf path has actually been given.
    if pdf_path:
        if not os.path.isfile(pdf_path):
            error = f"File not found: {pdf_path}"
        else:
            try:
                results = extract_pdf_page_diagrams(
    pdf_path,
    padding_left=0,
    padding_right=0,
    padding_top=50,
    padding_bottom=342
)
                diagrams = [
                    {
                        "name": d["name"],
                        "page": d["page"],
                        "rect": d["rect"],
                        "b64": base64.b64encode(d["bytes"]).decode("ascii"),
                    }
                    for d in results
                ]
                if not diagrams:
                    error = "No diagrams detected in this PDF (no drawing regions found on any page)."
            except Exception as e:
                error = f"Extraction failed: {e}"

    return render_template_string(
        PAGE_TEMPLATE, pdf_path=pdf_path, padding_x=padding_x, padding_y=padding_y, diagrams=diagrams, error=error
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Check PDF diagram-extraction padding in a web page.")
    parser.add_argument("--pdf", type=str, default="", help="Path to the PDF file to extract from.")
    parser.add_argument("--padx", type=int, default=DEFAULT_PADDING_X, help="Initial crop padding in points for the x-axis.")
    parser.add_argument("--pady", type=int, default=DEFAULT_PADDING_Y, help="Initial crop padding in points for the y-axis.")
    parser.add_argument("--port", type=int, default=5050, help="Port to run the web server on.")
    parser.add_argument("--host", type=str, default="127.0.0.1", help="Host to bind the web server to.")
    args = parser.parse_args()

    DEFAULT_PDF_PATH = args.pdf
    DEFAULT_PADDING_X = args.padx
    DEFAULT_PADDING_Y = args.pady

    print(f"Open http://{args.host}:{args.port}/ in your browser.")
    app.run(host=args.host, port=args.port, debug=True)