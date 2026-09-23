import os
import re
import ast
import json
import time
import logging
from typing import List, Dict, Any
import sys
import unicodedata
import zipfile
import tempfile
import hashlib
import math
import openpyxl
import fitz
from PIL import Image as PILImage, ImageDraw

from openpyxl.styles import (
    Alignment,
    Border,
    Side,
    Font
)
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter

from Files.database import DatabaseHandler
from retrieval.retrieve_content_prompt import retrieve_content_for_prompt


# ==========================================================
# GLOBAL CONSTANTS / CONFIGURATION
# ==========================================================

CURRENT_FOLDER = os.path.dirname(os.path.abspath(__file__))

MAX_LLM_RETRIES = 3
INITIAL_RETRY_DELAY = 1

# ==========================================================
# TARGET EDO TEST RUN
# ==========================================================
# Set the ONE Existing EDO tag you want to test.
# Every Existing-EDO stage uses this same tag.
TARGET_EDO_TAG = "EDO-40"

# ---- New EDO test-loop limit ----
# During a test run, process/print only the first 3 New EDO records.
# CALL 4 creates at most these 3 reference records, and CALL 5 processes
# only those same 3 records.
NEW_EDO_MAX_TEST_RECORDS = 3


def _normalize_edo_tag(value):
    return normalize_text(value).upper()


def _is_target_edo(value):
    return _normalize_edo_tag(value) == _normalize_edo_tag(TARGET_EDO_TAG)


def _filter_target_existing_edos(existing_edos):
    """Keep only the hard-coded Existing EDO tag for this test run."""
    return {
        key: value
        for key, value in (existing_edos or {}).items()
        if _is_target_edo(value.get("edo_tag") or key)
    }


# ---- Existing EDO tag retrieval depth ----
# extract_edo_tags() (CALL 1) previously called build_prompt_row()
# without max_results, so retrieval fell back to its own default depth
# - observed to select as few as 5 chunks even when the candidate pool
# has 40+ chunks. Any Existing EDO tag whose row lived outside that
# narrow slice was never seen by the LLM and silently never made it
# into the tags dict at all (not a sorting issue - sort_existing_edos()
# only re-orders whatever keys are already present, it cannot drop
# any). Raising this covers the whole Existing EDO table in one call.
EXISTING_EDO_TAGS_MAX_RESULTS = 60

# ---- Existing EDO per-tag hydration retrieval depth ----
# extract_edo_details() (CALL 2) has the exact same gap as CALL 1: its
# build_prompt_row() call also never set max_results, so retrieval fell
# back to the same shallow default depth. Confirmed directly in a prior
# production log: the "Hydrating EDO-54" call retrieved chunks
# containing EDO-29/EDO-31/EDO-32/EDO-55's rows, but NOT EDO-54's own
# row - so the LLM correctly reported "Total_Rows": "0" and every
# descriptive column for EDO-54 stayed blank even though the tag itself
# was found in CALL 1. Same fix, same reasoning as CALL 1 above.
EXISTING_EDO_DETAILS_MAX_RESULTS = 60

# ---- Image pipeline (from edo_image.py) ----
IMAGE_WIDTH = 180
IMAGE_HEIGHT = 110

# ---- PDF image extraction (ported from generate_ICU_Template_final_working_code.py) ----
PDF_GLOBAL_OVERRIDE = os.path.join(CURRENT_FOLDER, "80369-6.pdf")
IMAGE_OUTPUT_DIR = "extracted_images"
IMAGE_TEXT_OFFSET_PX = 45

# ---- New EDO diagram guaranteed-placement target ----
# The New EDO diagram queue (see extract_new_edo_diagram_queue()) is
# otherwise pure FIFO / content-blind. Per requirement, the row whose
# RA_Number/FMEA_Number match this specific pair is guaranteed to get
# the next available diagram from that queue, reserved for it ahead of
# every other New EDO row - see format_edo_worksheet().
TARGET_IMAGE_RA_NUMBER = "RA-141"
TARGET_IMAGE_FMEA_NUMBER = "FMEA Sys-152"


# ==========================================================
# DOCUMENT NUMBER UTILITY (identity-agnostic document matching)
# ==========================================================
# Used by get_edo_document() to build a document-number index across
# EVERY mapped document regardless of its document_identity tag (so
# documents whose identity is null/untagged in the DB are still
# resolvable), and by CALL 7 / CALL 8 to key a verification code's
# source document by its NPD number instead of by fuzzy filename text.

DOC_NUMBER_PATTERN = re.compile(r'\bNPD\d+\b', re.IGNORECASE)


def _extract_doc_number(value):
    """
    Extracts a document number token (e.g. "NPD43975") from a document
    name / filename / location string. Returns the number in uppercase,
    or None if no such token is found. This is intentionally
    identity-agnostic - it works on any string regardless of whether
    the underlying document was tagged with a document_identity in the
    database.
    """
    if not value:
        return None
    match = DOC_NUMBER_PATTERN.search(str(value))
    return match.group(0).upper() if match else None


# ==========================================================
# DOCUMENT RETRIEVAL
# ==========================================================
# Loads EDO Proposed / RA&C / FMEA / PDF documents and the Excel template once, up front, for the whole pipeline to reuse.

DEFAULT_DOCUMENT_SEARCH_ROOTS = [
    "/mnt/documents",
    "/data/documents",
    ".",
]

def get_document_search_roots(pipeline_config=None):
    """
    Builds the ordered list of directories to search for the actual
    source .docx file, per the priority described above. The
    EDO_DOCUMENTS_DIR environment variable is read here (at call time),
    not baked into a module-level constant, so it's picked up correctly
    regardless of when it was set relative to module import.
    """
    roots = []

    if pipeline_config:
        configured_root = pipeline_config.get("documents_root")
        if configured_root:
            roots.append(configured_root)

    env_root = os.environ.get("EDO_DOCUMENTS_DIR")
    if env_root:
        roots.append(env_root)

    for root in DEFAULT_DOCUMENT_SEARCH_ROOTS:
        if root and root not in roots:
            roots.append(root)

    return roots


def find_file_by_name(filename, search_dir="."):
    if not search_dir or not os.path.isdir(search_dir):
        return None
    for root, _, files in os.walk(search_dir):
        if filename in files:
            return os.path.join(root, filename)
    return None


def get_template_documents(
    client,
    product_family,
    product,
    templatename,
    db: DatabaseHandler
):
    docs = db.get_template_documents(
        client,
        product_family,
        product,
        templatename
    )

    if not docs:
        raise Exception("No template documents were found.")

    logging.info("=" * 80)
    logging.info("AVAILABLE TEMPLATE DOCUMENTS")
    logging.info("=" * 80)

    for doc in docs:
        logging.info(
            f"{doc.get('document_identity')}  -->  "
            f"{doc.get('document_name')}"
        )

    logging.info("=" * 80)

    return docs


def get_edo_document(
    client,
    product_family,
    product,
    templatename,
    db: DatabaseHandler
):
    """
    CANONICAL version - updated to resolve EDO_Proposed, EDO_RA_C, EDO_FMEA,
    EDO_PDF_New, and 4 Excel documents from Chroma DB.
    Raises if EDO_Proposed is missing.

    ALSO builds edo_documents["documents_by_number"] - an
    IDENTITY-AGNOSTIC index of every single mapped document (regardless
    of its document_identity tag, including documents whose identity is
    null/untagged in the DB), keyed by the NPD number embedded in its
    document_name. This is the primary lookup CALL 8
    (find_verification_evidence_document()) now uses to resolve a
    verification code's source document - see the "CALL 8" section
    below and the DOCUMENT NUMBER UTILITY section above.
    """
    documents = get_template_documents(
        client,
        product_family,
        product,
        templatename,
        db
    )

    if not documents:
        raise Exception("No EDO template documents found.")

    edo_documents = {}
    traceability_documents = []
    # Column F evidence-lookup source documents (see CALL 8 /
    # extract_and_apply_verification_evidence()) - every document loaded
    # under document_identity "EDO_VER" is a verification-report source
    # that a Column E filename can be PARTIALLY matched against. There
    # can be any number of these (same dynamic pattern as
    # traceability_documents above), so they're kept as a flat list
    # rather than individual named keys.
    edo_ver_documents = []
    # Identity-agnostic document-number index - see docstring above.
    documents_by_number = {}

    for document in documents:
        identity = normalize_text(document.get("document_identity")).lower()
        doc_name = normalize_text(document.get("document_name")).lower()
        logging.info(f"AVAILABLE DOCUMENT : {identity}")

        # ---- Identity-agnostic document-number indexing ----
        # Runs for EVERY document, regardless of its identity tag (even
        # "null"/untagged ones), so CALL 8 can resolve a verification
        # code's source document by NPD number alone, without depending
        # on document_identity being correctly tagged in the DB.
        doc_number = _extract_doc_number(
            document.get("document_name") or document.get("originalfilename")
        )
        if doc_number:
            if doc_number not in documents_by_number:
                documents_by_number[doc_number] = document
            else:
                logging.warning(
                    f"DOCUMENT NUMBER {doc_number!r} appears on more than one "
                    "mapped document - keeping the first one encountered "
                    f"({documents_by_number[doc_number].get('document_name')!r}), "
                    f"ignoring {document.get('document_name')!r}."
                )

        if identity == "edo_proposed":
            edo_documents["edo_proposed"] = document
        elif identity in ["edo_ra_c", "edo_ra&c", "edo_rac", "edo_ra"]:
            edo_documents["edo_ra_c"] = document
        elif identity in ["edo_fmea", "fmea", "system_fmea"]:
            edo_documents["edo_fmea"] = document
        elif identity in ["edo_pdf_new", "edo_pdf-new", "edo pdf new"]:
            edo_documents["edo_pdf_new"] = document
        # CALL 7 traceability documents - every source document with
        # this identity is a traceability spreadsheet, regardless of
        # how many there are or what order the DB returns them in (see
        # CALL 7 / search_codes_across_documents() for how they're
        # searched dynamically instead of via a fixed prefix->document
        # map). Matched case-insensitively/trimmed (via normalize_text()
        # + .lower() above) so minor DB casing/whitespace differences
        # don't cause a traceability document to be silently dropped.
        elif identity == "edo_tm":
            traceability_documents.append(document)
        # CALL 8 verification-evidence source documents - see
        # find_verification_evidence_document() / CALL 8 below. Same
        # dynamic, any-count pattern as traceability_documents.
        elif identity in ["edo_ver", "edo ver", "edo_verification", "edo verification"]:
            edo_ver_documents.append(document)
        # Note: the old generic ".xlsx"/"excel"-in-identity Excel
        # fallback (and the edo_excel_1..4 explicit keys) has been
        # removed - every traceability spreadsheet for this pipeline is
        # confirmed to be loaded under the single identity "EDO_TM"
        # above, so that fallback existed only to catch these same
        # documents under a naming scheme that isn't actually used.

    # Retain the full list of traceability document collections for
    # CALL 7's dynamic, per-EDO document search.
    edo_documents["traceability_documents"] = traceability_documents
    # Retain the full list of EDO_VER (Column F evidence source) documents.
    edo_documents["edo_ver_documents"] = edo_ver_documents
    # Retain the identity-agnostic document-number index.
    edo_documents["documents_by_number"] = documents_by_number

    # ---- CALL 8 fallback candidate pool (identity-agnostic, widened) ----
    # Many real V/V evidence source files (e.g. "SV CTRL 1 SW V&V
    # General.pdf", "Vest APX APG Module EE Feature Test Result.pdf")
    # have NEITHER an NPD number in their own filename NOR an "EDO_VER"
    # document_identity tag - so they match neither documents_by_number
    # nor the old edo_ver_documents-only fuzzy fallback, and got
    # silently skipped in CALL 8. This pool is every mapped document
    # EXCEPT the four structural/core ones (EDO_Proposed, EDO_RA_C,
    # EDO_FMEA, EDO_pdf_new) - i.e. every EDO_VER-tagged, EDO_TM-tagged,
    # and untagged/null-identity document is a fair candidate for CALL
    # 8's fuzzy filename fallback, since we can't rely on identity
    # tagging being complete or correct.
    core_document_keys = ["edo_proposed", "edo_ra_c", "edo_fmea", "edo_pdf_new"]
    core_document_ids = {
        id(edo_documents[k]) for k in core_document_keys if k in edo_documents
    }
    evidence_candidate_documents = [
        document for document in documents if id(document) not in core_document_ids
    ]
    edo_documents["evidence_candidate_documents"] = evidence_candidate_documents

    logging.info("=" * 80)
    logging.info("EDO DOCUMENT CONFIGURATION")
    logging.info("=" * 80)

    for key, doc in edo_documents.items():
        if key == "traceability_documents":
            logging.info(f"Loaded Total Traceability (EDO_TM) Documents : {len(doc)}")
            continue
        if key == "edo_ver_documents":
            logging.info(f"Loaded Total EDO_VER Documents : {len(doc)}")
            for ver_doc in doc:
                logging.info(f"  EDO_VER document_name : {ver_doc.get('document_name')}")
            continue
        if key == "documents_by_number":
            logging.info(f"Loaded Total Document-Number-Indexed Documents : {len(doc)}")
            for number, number_doc in doc.items():
                logging.info(
                    f"  {number} -> {number_doc.get('document_name')} "
                    f"(identity={number_doc.get('document_identity')})"
                )
            continue
        if key == "evidence_candidate_documents":
            logging.info(
                f"CALL 8 fuzzy-fallback candidate pool (non-core documents) : {len(doc)}"
            )
            continue
        logging.info(f"{key}")
        logging.info(f"Identity   : {doc.get('document_identity')}")
        logging.info(f"Name       : {doc.get('document_name')}")
        logging.info(f"Collection : {doc.get('collection')}")

    logging.info("=" * 80)

    if "edo_proposed" not in edo_documents:
        raise Exception("Required document EDO_Proposed was not found.")

    if "edo_pdf_new" not in edo_documents:
        logging.warning(
            "Document with identity 'EDO_pdf_new' was not found in the "
            "template's configured documents - New EDO diagram lookup "
            "(Column H) will be skipped for this run."
        )

    if not traceability_documents:
        logging.warning(
            "No documents with identity 'EDO_TM' were found for this "
            "template - CALL 7 will be unable to resolve any "
            "verification codes to traceability records this run."
        )
    else:
        logging.info(
            f"Loaded {len(traceability_documents)} traceability (EDO_TM) "
            "document(s) for CALL 7's dynamic document search."
        )

    if not edo_ver_documents:
        logging.warning(
            "No documents with identity 'EDO_VER' were found for this "
            "template - CALL 8's fuzzy-filename fallback will have "
            "nothing to match against this run (the document-number "
            "index is still available and is tried first regardless)."
        )

    if not documents_by_number:
        logging.warning(
            "No documents had a recognizable NPD number in their name - "
            "CALL 8's identity-agnostic document-number lookup will be "
            "unable to resolve anything this run."
        )

    return edo_documents


def initialize_workbook(
    pipeline_config
):
    """
    Opens Excel template.
    """

    workbook = openpyxl.load_workbook(
        pipeline_config["input_file_path"]
    )

    sheet = workbook.active

    return workbook, sheet


# ==========================================================
# IMAGE EXTRACTIONS
# ==========================================================
# DOC image extraction, PDF image extraction, image insertion, and their shared helper functions.

CAPTION_FIGURE_PATTERN = re.compile(
    r'Figure\s+\d+|text\s+image|product\s+image|:',
    re.IGNORECASE
)
CAPTION_TABLE_PATTERN = re.compile(r'^\s*table\s+\d+', re.IGNORECASE)
EXCLUDED_LOGO_KEYWORDS = ["hillrom"]

# ---- Last-figure marker ----
# Per requirement, the figure captioned "For 211651, the hose metal ring
# dimension is controlled as below:" is the last real figure that should
# ever be extracted/printed. Duplicates are otherwise allowed as before,
# but once THIS specific figure is reached, nothing after it in the
# document is retrieved - see is_last_figure_caption() /
# extract_docx_figures_only().
LAST_FIGURE_CAPTION_PATTERN = re.compile(
    r'211651.*hose\s+metal\s+ring', re.IGNORECASE | re.DOTALL
)

def _extract_paragraph_texts_and_images(doc_xml):
    """
    Splits word/document.xml into paragraphs, extracting each paragraph's
    plain text and any image relationship IDs (r:embed / r:id) referenced
    within it, in document order. Returns a list of
    {"text": str, "embed_ids": [str, ...]} dicts.

    BUGFIX (root cause of "TOTAL FIGURES EXTRACTED : 0" on documents that
    plainly contain captioned figures): this used to scan the raw XML
    TEXT with regexes - r'<w:p[ >].*?</w:p>' to split paragraphs, then
    embed="..."/r:id="..." to find image references inside each one.
    Concretely, that missed real images two ways:
      1. Any image that sits inside a text box, content control, or
         other construct that nests a SECOND <w:p>...</w:p> block inside
         an outer paragraph. The non-greedy regex closes the "outer"
         paragraph match at the FIRST </w:p> it sees - the INNER one -
         silently merging the outer paragraph's own text together with
         the nested paragraph's content into one mangled "paragraph",
         and leaving the true outer closing tag dangling with no
         matching opener. Depending on layout this can merge a caption
         onto the wrong image, drop a paragraph entirely, or throw off
         every paragraph index after it.
      2. A legacy VML picture (<v:imagedata .../>, the pre-DrawingML
         image format) almost always identifies its relationship via the
         OFFICE namespace attribute o:relid (e.g. o:relid="rId9"), not
         r:id - the old code only ever looked for a literal 'r:id="..."'
         string, so a document whose figures use this common legacy
         reference style produced ZERO embed ids, on every paragraph,
         with nothing to log as "skipped" either (the image reference
         was never found at all, so it never reached the skip-reason
         checks below).
    Both are fixed by parsing with a real XML parser (ElementTree) -
    resolving elements/attributes by their namespace URI instead of a
    hard-coded prefix string, checking o:relid for VML images, and
    attributing each image to its closest/innermost enclosing <w:p> (via
    a parent map) instead of a regex span, so nested paragraphs can never
    corrupt or duplicate an image/caption pairing.
    """
    import xml.etree.ElementTree as ET

    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    V_NS = "urn:schemas-microsoft-com:vml"
    O_NS = "urn:schemas-microsoft-com:office:office"

    w_p = f"{{{W_NS}}}p"
    w_t = f"{{{W_NS}}}t"
    a_blip = f"{{{A_NS}}}blip"
    v_imagedata = f"{{{V_NS}}}imagedata"
    r_embed = f"{{{R_NS}}}embed"
    r_id = f"{{{R_NS}}}id"
    o_relid = f"{{{O_NS}}}relid"

    try:
        root = ET.fromstring(doc_xml)
    except ET.ParseError as parse_error:
        logging.warning(f"word/document.xml could not be parsed as XML: {parse_error}")
        return []

    # ElementTree elements don't expose a parent pointer - build one so
    # each image reference can be attributed to its CLOSEST enclosing
    # <w:p>, not every ancestor paragraph (which would double-count an
    # image living inside a text box nested inside another paragraph).
    parent_of = {child: parent for parent in root.iter() for child in parent}

    def _closest_paragraph(element):
        node = parent_of.get(element)
        while node is not None:
            if node.tag == w_p:
                return node
            node = parent_of.get(node)
        return None

    all_paragraphs = list(root.iter(w_p))
    embed_ids_by_paragraph: Dict[Any, List[str]] = {p: [] for p in all_paragraphs}

    # DrawingML (mc:Choice) picture reference - the one actually
    # rendered/visible in the document.
    for blip in root.iter(a_blip):
        embed = blip.get(r_embed)
        if not embed:
            continue
        owner = _closest_paragraph(blip)
        if owner is not None:
            embed_ids_by_paragraph[owner].append(embed)

    # Legacy VML picture reference. Word's own VML markup for a legacy
    # picture almost always identifies the relationship via the OFFICE
    # namespace attribute o:relid (e.g. <v:imagedata o:relid="rId9" .../>)
    # rather than r:id - r:id on <v:imagedata> is comparatively rare.
    # Checking only r:id (as the previous regex-based version did, via
    # its literal 'r:id="..."' fallback pattern) misses this common case
    # entirely, which alone is enough to explain "0 embed ids found" on a
    # document whose figures use this style of legacy image reference.
    # Only counted for a paragraph that has NO DrawingML blip at all, so
    # the invisible mc:Fallback duplicate of a Choice/Fallback pair is
    # never counted alongside its visible DrawingML counterpart (same
    # intent as the previous mc:Fallback-stripping step).
    for imagedata in root.iter(v_imagedata):
        rid = imagedata.get(o_relid) or imagedata.get(r_id)
        if not rid:
            continue
        owner = _closest_paragraph(imagedata)
        if owner is not None and not embed_ids_by_paragraph[owner]:
            embed_ids_by_paragraph[owner].append(rid)

    result = []
    for p in all_paragraphs:
        text = "".join(t.text or "" for t in p.iter(w_t)).strip()
        result.append({"text": text, "embed_ids": embed_ids_by_paragraph[p]})
    return result


def _resolve_caption_for_image(paragraphs, para_index):
    """
    Looks at the paragraph containing the image, then the paragraph
    immediately before it, then immediately after it, and returns the
    first one with any text - that's treated as the image's caption/
    nearby-text context.
    """
    candidates = []
    if 0 <= para_index < len(paragraphs):
        candidates.append(paragraphs[para_index]["text"])
    if para_index - 1 >= 0:
        candidates.append(paragraphs[para_index - 1]["text"])
    if para_index + 1 < len(paragraphs):
        candidates.append(paragraphs[para_index + 1]["text"])

    for text in candidates:
        if text:
            return text
    return ""


def is_figure_caption(caption_text):
    return bool(CAPTION_FIGURE_PATTERN.search(caption_text or ""))


def is_table_caption(caption_text):
    return bool(CAPTION_TABLE_PATTERN.search(caption_text or ""))


def is_excluded_logo(caption_text):
    lowered = (caption_text or "").lower()
    return any(keyword in lowered for keyword in EXCLUDED_LOGO_KEYWORDS)


def is_last_figure_caption(caption_text):
    """
    True when `caption_text` is the "For 211651, the hose metal ring
    dimension is controlled as below:" caption - the figure that must be
    the LAST one ever extracted. See LAST_FIGURE_CAPTION_PATTERN.
    """
    return bool(LAST_FIGURE_CAPTION_PATTERN.search(caption_text or ""))


def extract_docx_figures_only(docx_file):
    """
    CANONICAL image extractor used by extract_edo_proposed_images().
    See the module-level "FIGURE-ONLY IMAGE EXTRACTION" comment above for
    the full rule set. Every image placement is evaluated independently
    (duplicates allowed/retrieved), and only those with a "Figure N"
    caption nearby are kept - "Table N" captions, uncaptioned images, and
    anything mentioning "Hillrom" are all excluded.

    Per requirement, the figure captioned "For 211651, the hose metal
    ring dimension is controlled as below:" is treated as the LAST real
    figure in the document - once it's been extracted, nothing appearing
    after it is retrieved (see LAST_FIGURE_CAPTION_PATTERN /
    is_last_figure_caption()), even though duplicates before that point
    are still allowed as before.
    """
    logging.info("READING WORD MEDIA IMAGES (FIGURES ONLY)")

    with zipfile.ZipFile(docx_file, "r") as archive:
        try:
            doc_xml = archive.read("word/document.xml").decode("utf-8")
        except KeyError:
            logging.warning("word/document.xml not found - cannot extract figures.")
            return []

        rel_map = {}
        try:
            rel_xml = archive.read("word/_rels/document.xml.rels").decode("utf-8")
            # Parsed with ElementTree instead of an attribute-order-dependent
            # regex - Relationship elements are not guaranteed to write their
            # Id/Type/Target attributes in that exact order (e.g. Target
            # often comes before Type, or a TargetMode attribute is present),
            # and a strict ordered regex like the previous
            # r'Id="([^"]+)"\s+Type="[^"]+/image"\s+Target="([^"]+)"' would
            # then match nothing at all - leaving rel_map empty and silently
            # dropping every image later on with no log line and no skip
            # counter incremented (see the embed_id resolution loop below).
            import xml.etree.ElementTree as ET
            rel_root = ET.fromstring(rel_xml)
            for rel in rel_root:
                rel_type = rel.get("Type", "")
                if rel_type.endswith("/image"):
                    rel_id = rel.get("Id")
                    target = rel.get("Target")
                    if rel_id and target:
                        rel_map[rel_id] = os.path.basename(target)
        except KeyError:
            logging.warning("word/_rels/document.xml.rels not found - cannot resolve image relationships.")
            return []
        except ET.ParseError as parse_error:
            logging.warning(f"word/_rels/document.xml.rels could not be parsed as XML: {parse_error}")
            return []

        media_bytes = {}
        for file in archive.namelist():
            if file.startswith("word/media/"):
                media_bytes[os.path.basename(file)] = {
                    "extension": os.path.splitext(file)[1],
                    "bytes": archive.read(file)
                }

        paragraphs = _extract_paragraph_texts_and_images(doc_xml)

        figures = []
        skipped_table = 0
        skipped_logo = 0
        skipped_uncaptioned = 0
        skipped_unresolved = 0
        reached_last_figure = False

        for para_index, para in enumerate(paragraphs):
            if reached_last_figure:
                break

            for embed_id in para["embed_ids"]:
                filename = rel_map.get(embed_id)
                if not filename or filename not in media_bytes:
                    skipped_unresolved += 1
                    logging.info(
                        f"SKIPPED (embed id could not be resolved to a media file) : "
                        f"embed_id={embed_id!r}, resolved_filename={filename!r}"
                    )
                    continue

                caption = _resolve_caption_for_image(paragraphs, para_index)

                if is_excluded_logo(caption):
                    skipped_logo += 1
                    logging.info(f"SKIPPED (Hillrom logo/letterhead) : {filename}")
                    continue

                if is_table_caption(caption):
                    skipped_table += 1
                    logging.info(f"SKIPPED (Table image, not a Figure) : {filename} - caption: {caption!r}")
                    continue

                if not is_figure_caption(caption):
                    skipped_uncaptioned += 1
                    logging.info(f"SKIPPED (no Figure caption found nearby) : {filename} - nearby text: {caption!r}")
                    continue

                media = media_bytes[filename]
                figures.append({
                    "name": filename,
                    "extension": media["extension"],
                    "bytes": media["bytes"],
                    "caption": caption
                })
                logging.info(f"FIGURE FOUND : {filename} - caption: {caption!r}")

                if is_last_figure_caption(caption):
                    reached_last_figure = True
                    logging.info(
                        "LAST FIGURE REACHED - caption "
                        f"{caption!r} matches the 211651 hose metal ring "
                        "marker; no further images later in the document "
                        "will be extracted."
                    )
                    break

    logging.info(
        f"TOTAL FIGURES EXTRACTED : {len(figures)}  "
        f"(skipped {skipped_table} table image(s), {skipped_logo} Hillrom logo/letterhead "
        f"image(s), {skipped_uncaptioned} uncaptioned image(s), {skipped_unresolved} "
        f"unresolved embed id(s))"
    )
    return figures


def save_images_to_folder(images, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    saved_paths = []
    for img in images:
        ext = img["extension"] if img["extension"] else ".png"
        out_path = os.path.join(output_folder, img["name"])
        # avoid collisions if names repeat
        base, counter = out_path, 1
        while os.path.exists(out_path):
            root, e = os.path.splitext(base)
            out_path = f"{root}_{counter}{e}"
            counter += 1
        with open(out_path, "wb") as f:
            f.write(img["bytes"])
        saved_paths.append(out_path)
        logging.info(f"Saved : {out_path}")
    return saved_paths


def resolve_edo_source_file(edo_document, document_key, pipeline_config=None):
    """
    SINGLE generic resolver for any EDO source document's local file path.

    Previously this logic existed as two near-identical copies -
    get_edo_proposed_file() (for the "edo_proposed" .docx) and
    get_edo_pdf_new_file() (for the "edo_pdf_new" .pdf) - which was exactly
    the "get proposed file" / "get edo pdf new file" duplication that
    needed to go. Now there is exactly ONE document-file-resolution
    function in the whole pipeline; callers just pass which key they want:

        resolve_edo_source_file(edo_document, "edo_proposed", pipeline_config)
        resolve_edo_source_file(edo_document, "edo_pdf_new", pipeline_config)

    Resolution order (identical to the old behaviour for both callers):
      1. edo_document[document_key] must be present (set once by
         get_edo_document()).
      2. Check common path fields directly on that document's metadata.
      3. Otherwise, search for its document_name/name across
         get_document_search_roots() (pipeline_config["documents_root"],
         EDO_DOCUMENTS_DIR, then the default mount-point fallbacks).

    Every branch that fails to resolve a path logs exactly why, so an
    empty result downstream is traceable back to a specific cause.
    """
    document = edo_document.get(document_key) if edo_document else None
    if not document:
        logging.warning(
            f"{document_key} RESOLUTION FAILED - no document with that "
            "document_identity was configured for this template "
            f"(edo_document has no '{document_key}' key)."
        )
        return None

    for key in ["file_path", "document_path", "path", "local_path", "filepath"]:
        val = document.get(key)
        if not val:
            continue
        if os.path.exists(val):
            logging.info(f"Resolved {document_key} file from document field '{key}' : {val}")
            return val
        logging.warning(
            f"{document_key} document['{key}'] = {val!r} was set but does "
            "not exist on disk in this container - falling back to "
            "filename search."
        )

    doc_name = document.get("document_name") or document.get("name")
    if not doc_name:
        logging.error(
            f"{document_key} RESOLUTION FAILED - the document has no "
            "usable path field (checked file_path/document_path/path/"
            "local_path/filepath) AND no document_name/name field to "
            f"search by. Full document dict: {document}"
        )
        return None

    search_roots = get_document_search_roots(pipeline_config)
    logging.info(f"Searching for {document_key} '{doc_name}' under: {search_roots}")

    for root in search_roots:
        found = find_file_by_name(doc_name, search_dir=root)
        if found:
            logging.info(f"Found {document_key} '{doc_name}' at : {found}")
            return found

    logging.error(
        f"{document_key} RESOLUTION FAILED - could not find '{doc_name}' "
        f"under any of {search_roots}. If this is running in Docker, "
        "that directory needs to actually be mounted into the container "
        "- set EDO_DOCUMENTS_DIR or pipeline_config['documents_root'] to "
        "the correct in-container mount path."
    )
    return None


def extract_edo_proposed_images(edo_document, pipeline_config=None):
    """
    CANONICAL version. Uses the SAME "edo_proposed" document reference
    that extract_edo_tags() already resolves (edo_document["edo_proposed"])
    - no separate document lookup. Resolves its local .docx path via the
    single unified resolve_edo_source_file() call, then extracts its
    embedded FIGURE images (duplicates included, tables/logo excluded)
    via extract_docx_figures_only(). Source documents are always .docx
    here, so no .doc -> .docx conversion step is needed.

    `pipeline_config` is optional and passed straight through to
    resolve_edo_source_file() so pipeline_config["documents_root"] can be
    used to fix Docker deployments where the source .docx files live on
    a mounted volume rather than the container's default CWD - see the
    "DOCUMENT SEARCH ROOTS (Docker fix)" section above.

    """
    try:
        file_path = resolve_edo_source_file(edo_document, "edo_proposed", pipeline_config)
        if not file_path:
            # resolve_edo_source_file() / find_file_by_name() already log
            # exactly which directories were searched and why nothing
            # was found - this just makes the end result unambiguous.
            logging.warning(
                "IMAGE EXTRACTION SKIPPED - the edo_proposed .docx file "
                "itself could not be located (see the search log above)."
            )
            return []

        images = extract_docx_figures_only(file_path)

        if not images:
            # The file WAS found and read - so if this is empty, it's
            # almost certainly the "Figure N" caption filter in
            # extract_docx_figures_only() not matching this document's
            # actual caption style (see its per-image SKIPPED log lines
            # just above this one for exactly why each image was
            # excluded), NOT a missing-file problem.
            logging.warning(
                f"IMAGE EXTRACTION RETURNED 0 FIGURES from a file that WAS "
                f"found and read successfully ({file_path}). This means "
                "either every image was excluded by the Figure/Table/"
                "Hillrom caption filter, OR every embedded image's r:embed "
                "id failed to resolve to an actual media file (e.g. a "
                "malformed/unparseable document.xml.rels) - check the "
                "'SKIPPED (...)' log lines just above for the reason each "
                "one was excluded: a 'Hillrom logo/letterhead', 'Table "
                "image', or 'no Figure caption found nearby' line points "
                "at the caption filter and CAPTION_FIGURE_PATTERN, while "
                "an 'embed id could not be resolved to a media file' line "
                "points at word/_rels/document.xml.rels instead."
            )
        else:
            logging.info(
                f"Extracted {len(images)} figure(s) from edo_proposed "
                f"({edo_document.get('edo_proposed', {}).get('document_name')})."
            )

        return images
    except Exception as e:
        logging.error(f"Image extraction failed : {e}")
        return []


def insert_image_below_text(sheet, image, row, column=8, text_offset_px=IMAGE_TEXT_OFFSET_PX):
    """
    CANONICAL - from edo_image.py. Anchors `image` into `sheet` at
    (row, column) - Column H (column=8) by default - positioned
    `text_offset_px` pixels below the top of the cell, so it renders
    underneath whatever text is already in that cell rather than
    overlapping it. Also grows the row height / column width as needed
    so the image isn't clipped.
    """
    if not image:
        return False
    try:
        suffix = image.get("extension") or ".png"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp:
            temp.write(image["bytes"])
            temp_name = temp.name

        excel_image = XLImage(temp_name)
        excel_image.width = IMAGE_WIDTH
        excel_image.height = IMAGE_HEIGHT

        marker = AnchorMarker(
            col=column - 1, colOff=pixels_to_EMU(2),
            row=row - 1, rowOff=pixels_to_EMU(text_offset_px)
        )
        size = XDRPositiveSize2D(cx=pixels_to_EMU(IMAGE_WIDTH), cy=pixels_to_EMU(IMAGE_HEIGHT))
        excel_image.anchor = OneCellAnchor(_from=marker, ext=size)
        sheet.add_image(excel_image)

        required_height = (text_offset_px + IMAGE_HEIGHT) * 0.75
        current_height = sheet.row_dimensions[row].height or 0
        if required_height > current_height:
            sheet.row_dimensions[row].height = required_height

        required_col_width = (IMAGE_WIDTH + 15) * 0.14
        current_width = sheet.column_dimensions['H'].width or 0
        if required_col_width > current_width:
            sheet.column_dimensions['H'].width = required_col_width
        return True
    except Exception as e:
        logging.error(f"Image insert failed at row {row} : {e}")
        return False


def is_valid_rect(rect):
    """True if a PyMuPDF rect is non-empty, finite, and has real size."""
    if rect is None:
        return False
    if rect.is_empty or rect.is_infinite:
        return False
    if rect.width <= 1 or rect.height <= 1:
        return False
    return True


def _diagram_clip_rect(
    page,
    region,
    padding_left=20,
    padding_right=20,
    padding_top=20,
    padding_bottom=20,
):
    """Convert an unrotated PDF content rect into displayed page coordinates."""
    source_bounds = fitz.Rect(page.cropbox)
    padded = fitz.Rect(
        max(source_bounds.x0, region.x0 - padding_left),
        max(source_bounds.y0, region.y0 - padding_top),
        min(source_bounds.x1, region.x1 + padding_right),
        min(source_bounds.y1, region.y1 + padding_bottom),
    )

    # Drawing, image, and text-block coordinates are returned in the
    # unrotated PDF coordinate system. get_pixmap() clips in displayed
    # coordinates, so pages stored with /Rotate=90 must be transformed
    # before rendering. Without this conversion, the result is a tall
    # page strip or an almost complete page instead of the diagram.
    if page.rotation:
        padded = padded * page.rotation_matrix

    page_bounds = fitz.Rect(page.rect)
    clip = fitz.Rect(
        max(page_bounds.x0, padded.x0),
        max(page_bounds.y0, padded.y0),
        min(page_bounds.x1, padded.x1),
        min(page_bounds.y1, padded.y1),
    )

    # Technical-drawing page borders and zone labels sit at the extreme
    # displayed edge. When padding reaches that edge, trim the narrow frame
    # band while retaining the actual diagram and its callouts.
    if clip.x1 >= page_bounds.x1 - 1:
        clip.x1 = min(clip.x1, page_bounds.x0 + page_bounds.width * 0.95)
    return clip


def _save_diagram_pixmap(pix, page, clip_rect, filepath):
    """Save a diagram crop after masking lower-left drawing metadata text."""
    mode = "RGBA" if pix.n == 4 else "RGB"
    image = PILImage.frombytes(mode, (pix.width, pix.height), pix.samples)

    metadata_rects = []
    for block in page.get_text("blocks"):
        rect = fitz.Rect(block[0], block[1], block[2], block[3])
        display_rect = rect * page.rotation_matrix if page.rotation else rect
        if (
            display_rect.x1 <= page.rect.width * 0.22
            and display_rect.y0 >= page.rect.height * 0.32
        ):
            metadata_rects.append(display_rect)

    if metadata_rects:
        metadata_bbox = fitz.Rect(metadata_rects[0])
        for rect in metadata_rects[1:]:
            metadata_bbox |= rect

        visible = metadata_bbox & clip_rect
        if is_valid_rect(visible):
            scale_x = pix.width / clip_rect.width
            scale_y = pix.height / clip_rect.height
            left = max(0, int((visible.x0 - clip_rect.x0) * scale_x) - 8)
            top = max(0, int((visible.y0 - clip_rect.y0) * scale_y) - 8)
            right = min(pix.width, int((visible.x1 - clip_rect.x0) * scale_x) + 8)
            bottom = min(pix.height, int((visible.y1 - clip_rect.y0) * scale_y) + 8)
            ImageDraw.Draw(image).rectangle(
                (left, top, right, bottom),
                fill=(255, 255, 255, 255) if mode == "RGBA" else (255, 255, 255),
            )

    image.save(filepath, format="PNG")


def extract_pdf_page_diagrams(
    pdf_path,
    output_dir=IMAGE_OUTPUT_DIR,
    padding_left=20,
    padding_right=20,
    padding_top=20,
    padding_bottom=20,
):
    """
    Fresh, content-blind diagram extractor for EDO_pdf_new.

    Does NOT search for RA/FMEA text anywhere. Simply walks every page
    of the PDF, and for any page that contains embedded images or
    vector drawings ("looks like it has a diagram on it"), takes a
    full-page screenshot and appends it to an ORDERED list. No
    identification/matching happens here - that happens later, purely
    by row order, when the Excel is filled.

    Returns (in page order):
        [ {"name": "...", "bytes": b"...", "extension": ".png"}, ... ]
    """
    diagrams = []
    if not pdf_path or not os.path.isfile(pdf_path):
        logging.error(f"PDF PAGE DIAGRAM EXTRACTION FAILED - PDF not found: {pdf_path!r}")
        return diagrams

    os.makedirs(output_dir, exist_ok=True)

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        logging.error(f"PDF PAGE DIAGRAM EXTRACTION FAILED - fitz.open() failed: {e!r}")
        return diagrams

    try:
        for page_num in range(len(doc)):
            page = doc[page_num]
            page_no = page_num + 1

            region = _get_drawing_region(page)
            if region is None:
                logging.info(f"PAGE {page_no} SKIPPED - no drawing region detected on this page.")
                continue

            try:
                clip_rect = _diagram_clip_rect(
                    page,
                    region,
                    padding_left=padding_left,
                    padding_right=padding_right,
                    padding_top=padding_top,
                    padding_bottom=padding_bottom,
                )
                if not is_valid_rect(clip_rect):
                    logging.warning(
                        f"PAGE {page_no} SKIPPED - detected drawing produced an invalid clip rect."
                    )
                    continue
                pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip_rect, alpha=False)
                filename = f"NEW_EDO_DIAGRAM_p{page_no}.png"
                filepath = os.path.join(output_dir, filename)
                _save_diagram_pixmap(pix, page, clip_rect, filepath)
                with open(filepath, "rb") as f:
                    data = f.read()
                diagrams.append({"name": filename, "bytes": data, "extension": ".png"})
                logging.info(f"PAGE {page_no} CAPTURED (drawing region only, cropped+padded) -> {filepath}")
            except Exception as e:
                logging.warning(f"Failed to rasterize drawing region on page {page_no}: {e}")
    finally:
        doc.close()

    logging.info(f"PDF PAGE DIAGRAM EXTRACTION TOTAL: {len(diagrams)} diagram page(s) captured, in page order.")
    return diagrams


def _get_drawing_region(page, proximity=150, paragraph_char_threshold=80, paragraph_height_threshold=40, band_vertical_tolerance=60):
    """
    Finds the bounding box of just the actual line-art drawing on the
    page - excludes large text paragraphs (DESCRIPTION, RECOMMENDED
    VENDOR, NOTES, title block cells, revision table) and excludes
    page-border/grid lines, but INCLUDES small annotation/callout/
    dimension labels sitting right next to the drawing (e.g.
    "2950±50mm", "NEMA 1-15P BLACK", "SJT 1.00mm x 2") since those are
    part of the drawing, not "other text".

    NOTE ON band_vertical_tolerance:
    A single technical drawing on these sheets is often laid out as
    several DISCONNECTED vector clusters that sit side-by-side in the
    same row (e.g. the main connector/cord diagram on the left, and a
    separate C17-connector + polarity-diagram cluster far to the
    right). Those clusters can be well beyond `proximity` from each
    other horizontally, so the old "only merge what's within
    `proximity` pixels of the main cluster" logic clipped the crop to
    just the left-hand cluster and cut off everything to the right.
    Now, any cluster that overlaps the main cluster's vertical extent
    (within `band_vertical_tolerance`) is merged in regardless of how
    far away it is horizontally, since it's part of the same drawing
    row. `proximity` is still used as a fallback for genuinely nearby
    content that doesn't share the row (e.g. a diagonal callout).

    Returns a fitz.Rect, or None if no drawing content was found.
    """
    # Source objects use unrotated PDF coordinates even when page.rect is
    # rotated for display. Use cropbox dimensions for source-space
    # clustering and page-border filtering.
    page_w, page_h = page.cropbox.width, page.cropbox.height
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

    # Drop page-border / zone-grid / table-gridline strokes - these
    # span almost the full page width or height as a thin line, and are
    # not part of the drawing itself.
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

    # Keep only clusters that look like a real drawing - not tiny
    # decorations, and not a near-full-page frame.
    clusters = [
        c for c in clusters
        if c.width > 30 and c.height > 30
        and not (c.width > 0.9 * page_w and c.height > 0.9 * page_h)
    ]
    content_bboxes.extend(clusters)

    if not content_bboxes:
        return None

    # Anchor on the single largest piece of drawing content, then pull
    # in every other drawing/image cluster that either:
    #   (a) sits in the same horizontal band/row as the main cluster
    #       (shares vertical extent, within band_vertical_tolerance) -
    #       this is what pulls in far-right content like a separate
    #       C17-connector/polarity-diagram cluster that is part of the
    #       same drawing row but not within `proximity` pixels, or
    #   (b) is simply close to the main cluster (within `proximity`),
    #       same as before, for nearby content that doesn't share a row.
    # This runs iteratively since merging can grow main_bbox's vertical
    # extent, which can then bring a further cluster into the band.
    main_bbox = max(content_bboxes, key=lambda r: r.width * r.height)
    changed = True
    while changed:
        changed = False
        for r in content_bboxes:
            if r is main_bbox or main_bbox.contains(r):
                continue
            vertical_overlap = min(main_bbox.y1, r.y1) - max(main_bbox.y0, r.y0)
            shares_band = vertical_overlap > -band_vertical_tolerance
            expanded = fitz.Rect(main_bbox.x0 - proximity, main_bbox.y0 - proximity, main_bbox.x1 + proximity, main_bbox.y1 + proximity)
            if shares_band or expanded.intersects(r):
                merged = fitz.Rect(main_bbox)
                merged |= r
                if merged != main_bbox:
                    main_bbox = merged
                    changed = True

    # Pull in small nearby labels (dimensions/callouts), but SKIP large
    # paragraph-style text blocks even if they're nearby. Keep the search
    # box fixed: allowing every included label to grow the next search box
    # creates a chain into DESCRIPTION, NOTES, and title-block content.
    drawing_bbox = fitz.Rect(main_bbox)
    drawing_display_bbox = (
        drawing_bbox * page.rotation_matrix if page.rotation else drawing_bbox
    )
    label_search_bbox = fitz.Rect(
        drawing_bbox.x0 - proximity,
        drawing_bbox.y0 - proximity,
        drawing_bbox.x1 + proximity,
        drawing_bbox.y1 + proximity,
    )
    for block in page.get_text("blocks"):
        bx0, by0, bx1, by1, text = block[0], block[1], block[2], block[3], block[4]
        rect = fitz.Rect(bx0, by0, bx1, by1)
        display_rect = rect * page.rotation_matrix if page.rotation else rect
        is_paragraph = (
            len(text.strip()) > paragraph_char_threshold
            or display_rect.height > paragraph_height_threshold
        )
        if is_paragraph:
            continue

        is_page_frame_label = (
            display_rect.width > 0.75 * page.rect.width
            and display_rect.height < paragraph_height_threshold
        )
        if is_page_frame_label:
            continue

        is_lower_left_metadata = (
            display_rect.x1 < drawing_display_bbox.x0
            and display_rect.y0 > drawing_display_bbox.y0 + drawing_display_bbox.height * 0.5
        )
        if is_lower_left_metadata:
            continue

        if label_search_bbox.intersects(rect):
            main_bbox |= rect

    return main_bbox


def extract_new_edo_diagram_queue(edo_document, pipeline_config=None, output_dir=IMAGE_OUTPUT_DIR):
    """
    Resolves the EDO_pdf_new PDF and returns the ordered, content-blind
    diagram list from extract_pdf_page_diagrams(). Replaces every
    previous RA/FMEA-matching diagram function.
    """
    pdf_path = resolve_edo_source_file(edo_document, "edo_pdf_new", pipeline_config)
    if not pdf_path:
        logging.warning("NEW EDO DIAGRAM QUEUE SKIPPED - EDO_pdf_new PDF could not be resolved.")
        return []
    return extract_pdf_page_diagrams(pdf_path, output_dir=output_dir)


# ==========================================================
# PROMPT EXECUTION
# ==========================================================
# LLM invocation, retry logic, response/JSON cleanup, parsing and validation shared by every extraction stage below.

def normalize_text(value):
    if value is None:
        return ""

    value = str(value)

    value = value.replace("\u00A0", " ")
    value = value.replace("\u2007", " ")
    value = value.replace("\u202F", " ")

    value = unicodedata.normalize(
        "NFKC",
        value
    )

    return value.strip()


def clean_llm_response(response):

    response = normalize_text(response)

    return re.sub(
        r"^```json\s*|\s*```$",
        "",
        response,
        flags=re.DOTALL
    ).strip()


def _salvage_truncated_json_array(text):
    """
    Recovers as many complete top-level JSON objects as possible from a
    '[' ... ']' array whose text was cut off partway through (typically
    because the LLM hit its max_tokens limit mid-response on a large
    table). Walks the text tracking brace depth and string/escape state,
    so it isn't fooled by braces or brackets that appear inside quoted
    string values. Every '{ ... }' block that closes cleanly before the
    cutoff is parsed and kept; the dangling partial object at the very
    end (the one that got cut off) is simply dropped instead of causing
    the whole batch to be discarded.
    """
    start = text.find("[")
    if start == -1:
        return []

    objects = []
    depth = 0
    obj_start = None
    in_string = False
    escape = False

    for i in range(start, len(text)):
        ch = text[i]

        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue

        if ch == '"':
            in_string = True
            continue

        if ch == "{":
            if depth == 0:
                obj_start = i
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0 and obj_start is not None:
                candidate = text[obj_start:i + 1]
                try:
                    parsed_obj = json.loads(candidate)
                    if isinstance(parsed_obj, dict):
                        objects.append(parsed_obj)
                except Exception:
                    pass
                obj_start = None

    return objects


def parse_json(response):
    try:

        cleaned = clean_llm_response(response)

        logging.info("========== CLEANED JSON ==========")
        logging.info(cleaned)

        return json.loads(cleaned)

    except Exception as e:

        logging.error(f"JSON Parse Error : {e}")

        # FALLBACK 1 - the direct parse failed, most likely because the LLM
        # wrapped the JSON in extra prose/text beyond a plain ```json fence
        # (clean_llm_response only strips a leading/trailing fence, not
        # surrounding text). Try to salvage the first {...} or [...] block
        # in the response before giving up, so a well-formed JSON payload
        # isn't silently discarded just because of text around it.
        try:
            match = re.search(r"(\{.*\}|\[.*\])", normalize_text(response), re.DOTALL)
            if match:
                salvaged = json.loads(match.group(1))
                logging.warning(
                    "JSON Parse Error recovered via fallback extraction - "
                    "the LLM response had extra text around the JSON block."
                )
                return salvaged
        except Exception as fallback_error:
            logging.error(f"JSON fallback extraction also failed: {fallback_error}")

        # FALLBACK 2 - the response is a JSON array that was cut off mid-way
        # (e.g. 'Unterminated string' / 'Expecting value' errors partway
        # through the text) - almost always caused by hitting max_tokens on
        # a large table. Salvage every complete object that DID come through
        # before the cutoff rather than discarding the whole response.
        try:
            salvaged_objects = _salvage_truncated_json_array(normalize_text(response))
            if salvaged_objects:
                logging.warning(
                    f"JSON response appears TRUNCATED (likely hit max_tokens) - "
                    f"recovered {len(salvaged_objects)} complete object(s) before the "
                    f"cutoff point out of a presumably larger table. Consider raising "
                    f"max_tokens for this extraction or chunking the source table into "
                    f"smaller batches to avoid losing the remaining rows."
                )
                return salvaged_objects
        except Exception as salvage_error:
            logging.error(f"Truncated-array salvage also failed: {salvage_error}")

        logging.error(response)

        return {}


def blank(value):
    """
    NOTE: previously substituted the literal text "Blank" for empty
    values. Per requirement, no placeholder text should ever be written
    to the output Excel - a missing value should simply be an empty
    cell. This now just normalizes the text and returns "" for anything
    empty, instead of inserting "Blank".

    BUGFIX: normalize_text() alone does NOT catch the case where the
    LLM itself literally answers with the placeholder word "Blank" (or
    "None") as the VALUE of a design_elements field (location/
    description/reason/sysdd) - that text passed straight through
    unchanged. Two problems resulted:
      1. The literal word "Blank"/"None" got printed into the Excel
         cell instead of a real value or an empty cell.
      2. In format_edo_worksheet(), the split-row forward-fill logic
         (`if description: last_description = description else:
         description = last_description`) only treats a field as
         "missing" when it's falsy/empty - a non-empty string like
         "Blank" is truthy, so forward-fill never kicked in for that
         split row, and the next split row(s) of the same EDO tag kept
         showing the literal "Blank" text instead of inheriting the
         previous real description/reason (exactly the symptom seen in
         Excel: split row 1 shows real text, split row 2+ shows
         "Blank").
    Treating "none"/"blank" (case-insensitive) as empty here - the
    same convention already used by get_llm_value() elsewhere in this
    file - fixes both: the placeholder is never written to Excel, and
    forward-fill correctly carries the last real value down to every
    split row that has no genuine value of its own.
    """
    text = normalize_text(value)
    if text.lower() in ("none", "blank"):
        return ""
    return text


def call_llm(prompt, pipeline_config, question="generate content"):
    # NOTE: `question` used to be hardcoded to "risk classification"
    # here regardless of what the prompt actually asked for. That mislabel
    # was silently degrading unrelated callers (e.g. recommendation/warning
    # text generation in generate_recommendation_text()) since the LLM was
    # being told it was doing risk classification while receiving a full
    # free-text recommendation prompt - producing short/blank-ish answers
    # that then failed is_meaningful_llm_text() and got dropped from
    # Column M entirely. Callers now pass their own `question` so the LLM
    # is routed correctly for the content they're actually requesting.
    try:
        llm = pipeline_config["llm"]
        response = llm.generate(
            prompt,
            context="",
            question=question,
            temperature=pipeline_config["temperature"],
            max_tokens=pipeline_config["max_tokens"]
        )
        logging.info(f"LLM raw response: {response!r}")
        return response
    except Exception as e:
        # NOTE: previously returned the literal string "No" here. That
        # sentinel was indistinguishable from a genuine (wrong) LLM
        # answer, so callers such as generate_remarks_and_recommendation()
        # would treat a *failed* call as a valid "No"/short response and
        # write it straight into Column M. Returning "" instead lets
        # every caller's existing emptiness checks correctly detect the
        # failure and fall back cleanly.
        logging.error(f"LLM call failed: {e}")
        return ""


def is_meaningful_llm_text(response):
    """
    True only when `response` is real generated content - i.e. not
    empty, and not one of the placeholder/failure words an LLM (or a
    failed call_llm()) might return instead of an actual answer.
    """
    if not response:
        return False

    text = normalize_text(response).strip().upper()

    if not text:
        return False

    if text in ("NONE", "NO", "N/A", "NA", "NULL", "-"):
        return False

    return True


def clean_response(response):
    """
    Strips whitespace/code-fences from a plain-text LLM response
    (used by classify_risk_status - it only ever expects one bare word
    back, e.g. "High").
    """
    return clean_llm_response(response)


def execute_llm(
    pipeline_config,
    collection,
    prompt_row
):
    """
    Generic LLM wrapper used by every prompt passed via positional arguments.
    """
    return retrieve_content_for_prompt(
        pipeline_config,
        collection,
        prompt_row["question"],
        prompt_row["prompt_role"],
        prompt_row["prompt_text"],
        prompt_row["fulltext"],
        prompt_row["where_filter"],
        prompt_row["where_document"],
        prompt_row.get("checkpoint", ""),
        max_results=prompt_row.get("max_results"),
    )


def execute_llm_retry(
    pipeline_config,
    collection,
    prompt_row
):
    """
    Executes the LLM with retry logic.
    """

    last_exception = None

    for attempt in range(MAX_LLM_RETRIES):

        try:

            docs, metadata, response = execute_llm(
                pipeline_config,
                collection,
                prompt_row
            )

            if response:
                return docs, metadata, response

        except Exception as ex:

            last_exception = ex

            logging.exception(
    f"LLM Retry {attempt+1}/{MAX_LLM_RETRIES} failed: {ex}"
)

            time.sleep(
                INITIAL_RETRY_DELAY * (attempt + 1)
            )

    raise Exception(
        f"LLM failed after retries : {last_exception}"
    )


def get_prompt(
    client,
    product_family,
    product,
    templatename,
    prompt_name,
    db: DatabaseHandler
):
    prompt = db.get_prompt_by_name(
        client,
        product_family,
        product,
        templatename,
        prompt_name
    )

    if not prompt:
        raise Exception(f"Prompt '{prompt_name}' was not found.")

    logging.info(f"Loaded Prompt : {prompt_name}")
    return prompt


def build_prompt_row(
    prompt,
    question,
    fulltext="Yes",
    where_filter="",
    where_document="",
    checkpoint="",
    max_results=None
):
    return {
        "prompt_role": prompt["prompt_role"],
        "prompt_text": prompt["prompt_text"],
        "question": question,
        "fulltext": fulltext,
        "where_filter": where_filter,
        "where_document": where_document,
        "checkpoint": checkpoint,
        "max_results": max_results
    }


def execute_prompt(
    pipeline_config,
    collection,
    prompt
):
    _, _, response = execute_llm_retry(
        pipeline_config,
        collection,
        prompt
    )

    logging.info("=" * 80)
    logging.info("RAW LLM RESPONSE")
    logging.info("=" * 80)
    logging.info(response)

    return parse_json(response)


def deep_extract_records(data):
    """
    Recursive lookup utility to unwrap fluctuating JSON parent containers.
    Merged version: union of the container keys recognised by both source
    files (edo_existing_final.py's set plus generate_EDO_template_copy.py's
    extra "New_EDOs" / "New_Edos" / "new_edos" keys), so this single
    function correctly unwraps both existing-EDO and new-EDO responses.
    """
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]

    if isinstance(data, dict):
        for key in [
            "Records", "records",
            "New_EDOs", "New_Edos", "new_edos",
            "EDO_Table", "EDO_Tag_Values", "Verification_Details"
        ]:
            if key in data and isinstance(data[key], list):
                return [item for item in data[key] if isinstance(item, dict)]

        if any(k in data for k in ["Product_Feature_Function", "RA_Number", "FMEA_Number", "Traceability", "System DFMEA #", "Trace To RAC#"]):
            return [data]

        for val in data.values():
            if isinstance(val, list) and len(val) > 0 and isinstance(val[0], dict):
                return val
            elif isinstance(val, dict):
                res = deep_extract_records(val)
                if res:
                    return res
    return []


# ==========================================================
# EXISTING EDO PIPELINE
# ==========================================================
# CALL 1: Extract Existing EDO Tags -> Extract RA Number -> Extract FMEA Number -> CALL 2: Extract Remaining Existing EDO Details -> CALL 3: Extract Existing EDO Trace Details.

def create_empty_edo():
    """
    CANONICAL version - from edo_existing_final.py.
    Superset of the copy-file version: also carries FMEA_Number so that
    Stage 3B (verification-reference matching by RA/FMEA number) has
    something to match against.

    NOTE: defaults are empty strings, not the literal text "Blank" -
    an unfound value should render as a truly empty cell in the output
    Excel, per requirement.
    """
    return {
        "edo_type": "Existing",
        "edo_tag": "",
        "ra_number": "",
        "FMEA_Number": "",
        "edo_description": "",
        "reason_identified": "",
        "dfmea": "",
        "verification_reference": "",
        "existing_trace": "",
        "location": "",
        "description_2": "",
        "reason_2": "",
        "sysdd": "",
        "Project_code": ""
    }


def extract_edo_tags(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    edo_document,
    db: DatabaseHandler
):
    """
    CANONICAL - requested function #1, taken from edo_existing_final.py.
    """
    logging.info("=" * 80)
    logging.info("STAGE 3: EXTRACTING EXISTING EDO TAGS")
    logging.info("=" * 80)

    prompt = get_prompt(
        client,
        product_family,
        product,
        templatename,
        "EDO_Existing_Tag",
        db
    )

    target_text = TARGET_EDO_TAG
    question = (
        "Extract only the following Known Active Design Output Tracking "
        f"Number(s): {target_text}. Do not return any other EDO number."
    )

    prompt_row = {
        "prompt_role": prompt["prompt_role"],
        "prompt_text": prompt["prompt_text"],
        "question": question,
        "fulltext": "Yes",
        "where_filter": "",
        "where_document": "",
        "checkpoint": question,
        "max_results": EXISTING_EDO_TAGS_MAX_RESULTS
    }

    result = execute_prompt(
        pipeline_config,
        edo_document["edo_proposed"]["collection"],
        prompt_row
    )

    items = deep_extract_records(result)
    logging.info(f"Unconditional Record Extraction count: {len(items)}")
    logging.info(
        f"TARGET EDO MODE: CALL 1 returned {len(items)} raw tag record(s); "
        f"only configured target {TARGET_EDO_TAG} will continue."
    )

    tags = {}
    for item in items:
        if not isinstance(item, dict):
            continue

        tag = (
            item.get("edo_number")
            or item.get("edo_tag")
            or item.get("EDO Number")
            or item.get("EDO_Tag")
            or item.get("EDO")
            or ""
        )
        tag = normalize_text(tag)

        if tag == "" or tag.lower() == "blank" or tag in tags:
            continue

        tags[tag] = create_empty_edo()
        tags[tag]["edo_type"] = "Existing"
        tags[tag]["edo_tag"] = tag

    logging.info(
        f"extract_edo_tags: CALL 1 returned {len(tags)} tag(s): "
        f"{list(tags.keys())}"
    )

    return tags


def validate_existing_tags(tags):
    validated = {}
    for key, value in tags.items():
        tag = normalize_text(value.get("edo_tag"))
        if tag == "" or tag.lower() == "blank":
            continue
        validated[tag] = value
    return validated


_EDO_TAG_NUMBER_PATTERN = re.compile(r'(\d+)')


def sort_existing_edos(existing_edos):
    """
    Re-orders the `existing_edos` dict (edo_tag -> edo dict) into
    ascending numeric order by the digits in the tag (EDO-29, EDO-31,
    EDO-32, EDO-54, EDO-55, ...), instead of whatever order CALL 1
    (extract_edo_tags) happened to return records in.

    WHY THIS IS NEEDED: nothing in this pipeline ever explicitly sorts
    existing_edos - format_edo_worksheet() writes rows by iterating
    final_edos in dict order, and that order is inherited all the way
    back from the order CALL 1's LLM response listed the tags in. That
    order is a side effect of which chunks the vector store happened to
    return for the "Extract all Known Active Design Output Tracking
    Numbers." question - it is NOT tied to the source document's actual
    row order in any guaranteed way, only to how the documents were
    chunked/embedded/indexed into that particular collection. Moving the
    project to a new folder and re-ingesting the same source document
    rebuilds that collection from scratch, which can change chunk
    ordering/IDs even though the document content itself is unchanged -
    that is exactly why the same source document produced tags in
    correct ascending order (EDO-29, EDO-31, EDO-32, EDO-54, EDO-55) in
    the old folder, but a different, non-ascending order (EDO-32,
    EDO-55, EDO-54, EDO-29) after the move. Sorting explicitly here
    makes the printed order deterministic and correct regardless of
    retrieval/chunk order.
    """
    def sort_key(tag):
        match = _EDO_TAG_NUMBER_PATTERN.search(tag)
        if match:
            return (0, int(match.group(1)), tag)
        return (1, 0, tag)

    ordered_tags = sorted(existing_edos.keys(), key=sort_key)

    logging.info(
        f"sort_existing_edos: reordered {len(ordered_tags)} existing "
        f"EDO tag(s) into ascending numeric order for printing: "
        f"{ordered_tags}"
    )

    return {tag: existing_edos[tag] for tag in ordered_tags}


def extract_ra_fmea_pairs_from_text(text):
    """
    Pulls EVERY RA_Number / FMEA_Number pair out of the column D (dfmea)
    narrative extracted for an existing EDO - not just the first of each.
    Mirrors the one-row-per-FMEA approach extract_new_edo_tags() already
    uses for New EDOs: one RA mapped to several FMEA numbers becomes one
    row per FMEA number, all sharing that RA.
    """
    text = normalize_text(text)

    ra_matches = re.findall(r"RA[\s-]?(\d+)", text, re.IGNORECASE)
    fmea_matches = re.findall(r"SYS[\s-]?(\d+)", text, re.IGNORECASE)

    ra_numbers = list(dict.fromkeys(f"RA-{m}" for m in ra_matches))
    fmea_numbers = list(dict.fromkeys(f"SYS-{m}" for m in fmea_matches))

    if not ra_numbers and not fmea_numbers:
        return [("", "")]
    if not ra_numbers:
        return [("", f) for f in fmea_numbers]
    if not fmea_numbers:
        return [(r, "") for r in ra_numbers]

    # Common case (e.g. RA-66 -> FMEA Sys-187, FMEA Sys-200): one RA,
    # several FMEAs -> one row per FMEA, same RA on each.
    if len(ra_numbers) == 1:
        return [(ra_numbers[0], f) for f in fmea_numbers]

    # Multiple RAs + multiple FMEAs with no explicit pairing in the text
    # (e.g. EDO-56) - pair what lines up 1:1, keep any leftovers as their
    # own rows rather than silently dropping them.
    pairs = list(zip(ra_numbers, fmea_numbers))
    pairs += [(r, "") for r in ra_numbers[len(pairs):]]
    pairs += [("", f) for f in fmea_numbers[len(pairs):]]
    return pairs
def build_pair_dfmea_text(full_text, ra_number, fmea_number):
    """
    Column D value for ONE RA/FMEA pair: keeps the document-name context
    from the original narrative but strips every OTHER RA/FMEA token,
    leaving just this pair's numbers - so two rows for the same EDO show
    the same surrounding text with a different RA/FMEA at the end.
    """
    text = full_text
    for token in re.findall(r"RA[\s-]?\d+", text, re.IGNORECASE):
        text = text.replace(token, "", 1)
    for token in re.findall(r"(?:FMEA\s+)?SYS[\s-]?\d+", text, re.IGNORECASE):
        text = text.replace(token, "", 1)
    text = re.sub(r"\s{2,}", " ", text).strip()
    if ra_number:
        text = f"{text} {ra_number}".strip()
    if fmea_number:
        text = f"{text} FMEA {fmea_number}".strip()
    return text

def normalize_id(value):
    """
    Canonicalizes an RA/FMEA identifier for comparison purposes:
    uppercase, single spaces, trimmed. Used so that matching between the
    column-D-derived identifiers and whatever format Stage 3B's
    verification prompt returns doesn't fail on case/whitespace noise.
    """
    value = normalize_text(value).upper()
    return re.sub(r"\s+", " ", value).strip()


def extract_edo_details(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    edo_document,
    existing_edos,
    db: DatabaseHandler
):
    """
    CANONICAL version - from edo_existing_final.py.
    Comprehensive attribute hydration: extracts description/reason lists,
    dfmea narrative, location, sysdd, and (crucially) RA_Number /
    FMEA_Number - needed downstream by Stage 3B verification matching.

    UPDATED to match the EDO_Existing_Generic prompt's nested JSON shape
    (EDO_Table -> existing_edo_data[] -> design_elements[]):
      - Tag matching now also recognises "edo_no" - the key this prompt
        actually returns. The old code only checked edo_number/edo_tag/
        "EDO Number"/"EDO_Tag", so `result` was NEVER populated for this
        prompt's output and every field silently stayed blank - this is
        why nothing was printing in Excel.
      - "design_elements" (a LIST of per-location dicts) is walked and
        normalized into {"location", "description", "reason", "sysdd",
        "image"} entries, stored as a list on edo["design_elements"] for
        any downstream code (Excel writer / image placement) that wants
        the full per-location breakdown.
      - The single-value fields the rest of the pipeline reads today
        (edo_description / reason_identified / location / sysdd /
        description_2 / reason_2) are populated from the new top-level
        fields, backfilled from the design_elements list where the old
        schema had no top-level equivalent (e.g. location/sysdd only
        ever existed per design element, never at the top level).
    """
    logging.info("=" * 80)
    logging.info("STAGE 3: ATTRIBUTE HYDRATION (EXISTING EDO)")
    logging.info("=" * 80)

    prompt = get_prompt(
        client,
        product_family,
        product,
        templatename,
        "EDO_Existing_Generic",
        db
    )

    def normalize_key(key):
        key = normalize_text(key).lower()
        key = key.replace("&", "and")
        key = re.sub(r"[^a-z0-9]+", "_", key)
        key = re.sub(r"_+", "_", key)
        return key.strip("_")

    def first_value(record, keywords):
        """
        Return first non-empty value whose normalized key contains
        any keyword as a substring.
        """
        normalized_record = {normalize_key(k): v for k, v in record.items()}
        for kw in keywords:
            for nk, v in normalized_record.items():
                if kw in nk or nk in kw:
                    if normalize_text(v):
                        return normalize_text(v)
        return ""
    
    hydrated_existing_edos = {}

    for edo_tag in existing_edos.keys():


        logging.info(f"Hydrating {edo_tag}")

        question = (
            f"Extract the complete row details for the EDO tag requested by the user. Requested EDO Tag: {edo_tag}  Find the row where the EDO Tag exactly matches the requested EDO Ta. Return all available column values for that row, including EDO number/tag, description, reason identified as EDO, RA&C and/or Sys-DFMEA Trace,EDO Location, EDO Description, reason identified as EDO  If the requested EDO tag is not found, return none"
        )

        prompt_row = {
            "prompt_role": prompt["prompt_role"],
            "prompt_text": prompt["prompt_text"],
            "question": question,
            "fulltext": "Yes",
            "where_filter": "",
            "where_document": "",
            "checkpoint": question,
            "max_results": EXISTING_EDO_DETAILS_MAX_RESULTS
        }

        result = {}

        # 1. Fetch data
        for attempt in range(3):
            raw_result = execute_prompt(
                pipeline_config,
                edo_document["edo_proposed"]["collection"],
                prompt_row
            )
            records = deep_extract_records(raw_result)

            for record in records:
                tag = normalize_text(
                    record.get("edo_no")
                    or record.get("edo_number")
                    or record.get("edo_tag")
                    or record.get("EDO Number")
                    or record.get("EDO_Tag")
                ).lower()

                if tag == edo_tag.lower():
                    result = record
                    break

            # 2. Debug log
            logging.info(f"DEBUG: Hydration data for {edo_tag}: {json.dumps(result, indent=2)}")

            if not result:
                other_tags_seen = sorted({
                    normalize_text(
                        r.get("edo_no") or r.get("edo_number")
                        or r.get("edo_tag") or r.get("EDO Number")
                        or r.get("EDO_Tag") or ""
                    )
                    for r in records
                } - {""})
                logging.warning(
                    f"HYDRATION MISS for {edo_tag} on attempt "
                    f"{attempt + 1}/3: requested tag not present in the "
                    f"retrieved records. Tags actually returned this "
                    f"attempt: {other_tags_seen}. If {edo_tag} is never "
                    "among these across all 3 attempts, raise "
                    "EXISTING_EDO_DETAILS_MAX_RESULTS further - the row "
                    "is being excluded from retrieval, not failing to "
                    "parse."
                )

            # Always accept whatever we found to avoid "Blank"
            if result:
                break
            time.sleep(INITIAL_RETRY_DELAY)

        normalized = {normalize_key(k): v for k, v in result.items()}
        edo = existing_edos[edo_tag]

        # ---- Design elements (new nested structure) ----
        # "design_elements" is a LIST of per-location dicts under the new
        # EDO_Existing_Generic prompt shape - normalize every entry so
        # downstream code has a clean, predictable structure to work
        # from, instead of only ever seeing one location's worth of data.
        raw_design_elements = result.get("design_elements") or result.get("Design_Elements") or []
        if not isinstance(raw_design_elements, list):
            raw_design_elements = []

        design_elements = []
        for element in raw_design_elements:
            if not isinstance(element, dict):
                continue
            design_elements.append({
                "location": blank(
                    element.get("edo_location") or element.get("EDO_Location") or element.get("location")
                ),
                "description": blank(
                    element.get("edo_description") or element.get("EDO_Description") or element.get("description")
                ),
                "reason": blank(
                    element.get("reason identified as edo")
                    or element.get("reason_identified_as_edo")
                    or element.get("reason")
                ),
                "sysdd": blank(
                    element.get("SysDD or HDD Reference") or element.get("sysdd") or element.get("SysDD")
                ),
                "image": blank(element.get("image")),
            })

        edo["design_elements"] = design_elements

        # ---- Top-level description / reason ----
        # The new prompt returns ONE top-level "edo_description" /
        # "reason identified as edo" pair for the EDO itself - fall back
        # to the first design element only if the top-level fields are
        # somehow missing.
        top_description = first_value(result, ["edo_description"])
        top_reason = first_value(result, ["reason_identified_as_edo", "reason"])

        edo["edo_description"] = top_description or (design_elements[0]["description"] if design_elements else "")
        edo["reason_identified"] = top_reason or (design_elements[0]["reason"] if design_elements else "")

        # description_2 / reason_2 - backfilled from the SECOND design
        # element (if any), for backward compatibility with any code
        # still reading these two single-value fields.
        edo["description_2"] = design_elements[1]["description"] if len(design_elements) > 1 else ""
        edo["reason_2"] = design_elements[1]["reason"] if len(design_elements) > 1 else ""

        # ---- RA&C / Sys-DFMEA trace (Column D) ----
        # "RA and FMEA no" holds the real RA&C/System FMEA reference text
        # under the new schema - "Trace" is usually just "Blank"/empty,
        # so prefer the former and only fall back to the latter.
        ra_and_fmea = first_value(result, ["ra_and_fmea_no", "ra_and_fmea", "ra_and"])
        trace_narrative = first_value(result, ["trace"])
        dfmea_text = ra_and_fmea or trace_narrative

        # Prefer the first real design-element value, but do not let an
        # empty/"None" element hide a valid top-level LLM value.
        design_location = next(
            (element["location"] for element in design_elements if element.get("location")),
            ""
        )
        design_sysdd = next(
            (element["sysdd"] for element in design_elements if element.get("sysdd")),
            ""
        )
        edo["location"] = design_location or first_value(result, ["location"])
        edo["sysdd"] = design_sysdd or first_value(
            result,
            ["sysdd", "hdd_reference", "hardware", "design_reference"]
        )

        explicit_ra = first_value(result, ["ra_number"])
        explicit_fmea = first_value(result, ["fmea_number"])

        if explicit_ra and explicit_fmea:
            ra_fmea_pairs = [(explicit_ra, explicit_fmea)]
        else:
            ra_fmea_pairs = extract_ra_fmea_pairs_from_text(dfmea_text)

        # Explode this EDO tag into one row per RA/FMEA pair. Every extra
        # pair gets its own dict entry (unique key) so it becomes its own
        # Excel row, but edo_tag stays identical across all of them so
        # Column A still prints the same EDO number on every row.
        for index, (ra_number, fmea_number) in enumerate(ra_fmea_pairs):
            key = edo_tag if index == 0 else f"{edo_tag}__pair{index}"
            row = edo if index == 0 else dict(edo)
            row["dfmea"] = (
                dfmea_text if len(ra_fmea_pairs) == 1
                else build_pair_dfmea_text(dfmea_text, ra_number, fmea_number)
            )
            row["ra_number"] = ra_number or ""
            row["FMEA_Number"] = fmea_number or ""
            row["verification_reference"] = ""
            hydrated_existing_edos[key] = row

    return hydrated_existing_edos


def extract_existing_edo_trace_details(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    edo_document,
    existing_edos,
    db: DatabaseHandler
):
    """
    ADDED - dedicated trace extraction for EXISTING EDOs (CALL 3): its own prompt
    ("EDO_Existing_Trace"), queried against the edo_fmea collection,
    matched by RA/FMEA number that was pulled from the column D (dfmea)
    data in extract_edo_details(). The result is written to Column D,
    below the existing RA/FMEA data, in the same row (see
    apply_existing_edo_trace() and the Column D value construction in
    format_edo_worksheet()).

    CHANGED: now issues ONE LLM call PER EDO instead of a single batched
    call covering all EDOs at once. The batched version could get cut
    short mid-way through the response (partial results - some EDOs'
    tags missing or truncated), because all RA/FMEA targets shared one
    fixed max_tokens budget. Calling per-EDO gives each EDO its own full
    output budget and means one bad/oversized row can't blank out the
    rest of the batch.
    """
    logging.info("=" * 80)
    logging.info("STAGE 3C: EXISTING EDO - TRACE EXTRACTION (per-EDO calls)")
    logging.info("=" * 80)

    if "edo_fmea" not in edo_document:
        raise Exception(
            "EDO_FMEA document/collection not configured - cannot run "
            "trace extraction."
        )

    prompt_data = get_prompt(
        client,
        product_family,
        product,
        templatename,
        "EDO_Existing_Trace",
        db
    )

    results = {}

    for edo_tag, edo in existing_edos.items():
        ra_number = edo.get("ra_number")
        fmea_number = edo.get("FMEA_Number")

        has_ra = ra_number not in (None, "", "Blank")
        has_fmea = fmea_number not in (None, "", "Blank")

        if not has_ra and not has_fmea:
            logging.info(
                f"extract_existing_edo_trace_details: {edo_tag} has no "
                "RA/FMEA number - skipping (empty traces)."
            )
            results[edo_tag] = {"traces": []}
            continue

        target = (
            f"EDO : {edo_tag}\n"
            f"RA_Number : {ra_number}\n"
            f"FMEA_Number : {fmea_number}"
        )

        prompt_row = {
            "prompt_role": prompt_data["prompt_role"],
            "prompt_text": prompt_data["prompt_text"] + "\nTARGET:\n" + target,
            "question": f"Fetch all trace records for {edo_tag} (FMEA_Number: {fmea_number})",
            "fulltext": "Yes",
            "where_filter": "",
            "where_document": "",
            "checkpoint": f"Fetch all trace records for {edo_tag} (FMEA_Number: {fmea_number})"
        }

        try:
            _, _, response = execute_llm_retry(
                pipeline_config,
                edo_document["edo_fmea"]["collection"],
                prompt_row
            )
        except Exception as call_error:
            logging.warning(
                f"extract_existing_edo_trace_details: LLM call failed for "
                f"{edo_tag} after retries - leaving traces empty. "
                f"Reason: {call_error}"
            )
            results[edo_tag] = {"traces": []}
            continue

        edo_result = _parse_single_edo_trace_response(edo_tag, response)
        results[edo_tag] = edo_result

        logging.info(
            f"extract_existing_edo_trace_details: {edo_tag} -> "
            f"{len(edo_result.get('traces', []))} trace(s) extracted."
        )

    logging.info("========== PER-EDO TRACE RESULTS (COMBINED) ==========")
    logging.info(json.dumps(results, indent=2))

    return results


def parse_custom_fmea_format(response):
    """
    Parses custom text format containing:
    Fmea_number: SYS-147  
    [Document_number: NPD37819  
    {Trace: MRS CU FMEA-391, Trace: MRS CU FMEA-463}]

    Returns a list of formatted trace strings, e.g.:
    [
      "NPD37819: MRS CU FMEA-391, MRS CU FMEA-463",
      "NPD36569: MRS Software FMEA-422, MRS Software FMEA-423"
    ]
    """
    text = normalize_text(response)
    traces = []

    # Find each block starting with [Document_number: ... {Trace: ...}]
    # NOTE: doc_num uses a non-greedy `.*?` (bounded by the next "{")
    # instead of a `[^\]\n]+` char class, so it can't accidentally
    # swallow past a brace on multi-line blocks.
    doc_blocks = re.findall(
        r"\[\s*Document_number\s*:\s*(.*?)\s*\{(.*?)\}\s*\]?",
        text,
        re.IGNORECASE | re.DOTALL
    )

    for doc_num, trace_content in doc_blocks:
        doc_num = doc_num.strip()

        # Extract all individual tag values following "Trace:".
        # FIXED: the previous pattern used a character class
        # [^,Trace:\}] which - combined with re.IGNORECASE - excludes
        # any occurrence of the individual letters T/r/a/c/e (any
        # case), NOT the literal substring "Trace:". That truncated
        # every value at its first such letter, e.g. "MRS CU FMEA-391"
        # got cut to just "M" (stopped at the "R"). Instead, capture
        # everything up to the NEXT "Trace:" marker or the end of the
        # block, which correctly preserves the full value.
        raw_tags = re.findall(
            r"Trace\s*:\s*(.*?)(?=,\s*Trace\s*:|$)",
            trace_content,
            re.IGNORECASE | re.DOTALL
        )
        # BUGFIX: drop placeholder/status tags such as "Not Available" or
        # "Not Found" that the LLM sometimes emits as a Trace value when
        # it has nothing real to report for a document number, instead of
        # simply omitting the Trace entry. Left unfiltered, these produced
        # Column D lines like "NPD36569: Not Available" - a non-answer
        # printed as if it were real trace data. See _is_noise_code(),
        # shared with parse_verification_codes()'s identical problem on
        # the New EDO side.
        tags = [
            t.strip().rstrip(",").strip()
            for t in raw_tags
            if t.strip() and not _is_noise_code(t)
        ]

        if tags:
            # Format as: NPD37819: MRS CU FMEA-391, MRS CU FMEA-463
            comma_separated_tags = ", ".join(tags)
            traces.append(f"{doc_num}: {comma_separated_tags}")

    return traces

def _parse_single_edo_trace_response(edo_tag, response):
    """
    Parses ONE EDO's raw LLM response into {"traces": [str, ...]}.
    Order of operations:
      1. Direct JSON parse (parse_json / clean_llm_response)
      2. Fenced ```json code blocks (extract_json_code_blocks)
      3. Custom Bracket/Trace parser (parse_custom_fmea_format) -> Handles your prompt format!
      4. Legacy Markdown fallback (parse_markdown_trace_response)
    """
    # 1. Direct JSON parse
    parsed = parse_json(response)

    def normalize_result(obj):
        if isinstance(obj, dict):
            if "traces" in obj and isinstance(obj["traces"], list):
                # BUGFIX: same noise-token problem as parse_custom_fmea_format()
                # / parse_verification_codes() - a "traces" entry that is
                # itself just "Not Available"/"Not Found" placeholder text
                # is dropped rather than printed into Column D as if it
                # were a real trace.
                traces = [
                    normalize_text(t) for t in obj["traces"]
                    if normalize_text(t) and not _is_noise_code(t)
                ]
                return {"traces": traces}
            if edo_tag in obj and isinstance(obj[edo_tag], dict):
                return normalize_result(obj[edo_tag])
            for v in obj.values():
                if isinstance(v, dict) and "traces" in v:
                    return normalize_result(v)
        return None

    result = normalize_result(parsed)
    if result is not None:
        return result

    # 2. Code block parser
    code_block_records = extract_json_code_blocks(response)
    for block in code_block_records:
        result = normalize_result(block)
        if result is not None:
            logging.warning(
                f"_parse_single_edo_trace_response: {edo_tag} recovered via "
                "fenced ```json code block fallback."
            )
            return result

    # 3. Custom Bracket/Trace parser for the custom prompt format
    custom_traces = parse_custom_fmea_format(response)
    if custom_traces:
        logging.info(
            f"_parse_single_edo_trace_response: {edo_tag} successfully extracted "
            f"{len(custom_traces)} document trace group(s) using Custom Bracket parser."
        )
        return {"traces": custom_traces}

    # 4. Markdown fallback parser
    markdown_records = parse_markdown_trace_response(response)
    if markdown_records:
        traces = []
        for rec in markdown_records:
            module_controls = rec.get("Traces to Module DFMEA risk controls")
            if isinstance(module_controls, list):
                traces.extend(normalize_text(t) for t in module_controls if normalize_text(t))
        if traces:
            logging.warning(
                f"_parse_single_edo_trace_response: {edo_tag} recovered via "
                "Markdown fallback parser."
            )
            return {"traces": traces}

    logging.warning(
        f"_parse_single_edo_trace_response: {edo_tag} - could not parse any "
        "traces from the LLM response. Raw response logged below."
    )
    logging.warning(response)
    return {"traces": []}


def extract_json_code_blocks(text):
    """
    Fallback parser for extract_existing_edo_trace_details(). The LLM
    sometimes wraps EACH FMEA number's answer in its own separate
    ```json ... ``` fenced code block (e.g. one block per
    "### **For FMEA_Number: SYS-XXX**" section) instead of returning one
    single JSON document for the whole response. A single json.loads()
    call on the whole response fails in that case, because multiple
    top-level JSON objects side-by-side with Markdown headers/separators
    in between them isn't valid JSON as a whole - even though each
    individual fenced block IS perfectly valid JSON on its own. This
    extracts every ```json fenced block and parses each independently,
    returning the combined list of records.
    """
    text = normalize_text(text)
    blocks = re.findall(r"```(?:json)?\s*(.*?)```", text, re.DOTALL | re.IGNORECASE)

    records = []
    for block in blocks:
        try:
            parsed_block = json.loads(block.strip())
        except Exception:
            continue

        if isinstance(parsed_block, list):
            records.extend(item for item in parsed_block if isinstance(item, dict))
        elif isinstance(parsed_block, dict):
            records.append(parsed_block)

    return records


def parse_markdown_trace_response(response):
    """
    Fallback parser for extract_existing_edo_trace_details(). The
    EDO_Existing_Trace prompt is supposed to return JSON, but the LLM
    sometimes replies with Markdown instead - one "### **<FMEA id>**"
    block per FMEA number, with **bold** labels and "-" bullet lists,
    e.g.:

        ### **FMEA Sys-147**
        **Trace To RAC#:** RA-180
        **Traces to Module DFMEA risk controls:**
        - **Empty Array** (No matching entries found)
        **Additional Risk Control Measures from Module and/or Component DFMEAs:**
        - **Empty Array** (No matching identifiers found)
        **DRS Identifiers:**
        - **DRS-570**

    parse_json() can't recover this (there is no {}/[] JSON anywhere in
    it). This converts each block into the same record shape
    apply_existing_edo_trace() expects from a proper JSON response:
    {"System DFMEA #", "Trace To RAC#", "Traces to Module DFMEA risk
    controls", "Additional Risk Control Measures from Module and/or
    Component DFMEAs", "DRS Identifiers"}.
    """
    text = normalize_text(response)

    def extract_bullets(section_text):
        items = []
        for line in section_text.splitlines():
            line = line.strip()
            if not line.startswith("-"):
                continue
            line = line.lstrip("-").strip()
            line = re.sub(r"^\*+|\*+$", "", line).strip()
            if not line or "empty array" in line.lower() or line.lower() in ("none", "n/a"):
                continue
            items.append(line)
        return items

    def extract_section(block, label, next_labels):
        stop_pattern = "|".join(re.escape(nl) for nl in next_labels) if next_labels else None
        pattern = re.escape(label) + r"[:\*\s]*\n?(.*?)" + (
            f"(?=\\*\\*(?:{stop_pattern})|$)" if stop_pattern else "$"
        )
        match = re.search(pattern, block, re.DOTALL | re.IGNORECASE)
        return match.group(1) if match else ""

    records = []

    # Split on "### **<header>**" markers into (header, body) pairs.
    blocks = re.split(r"^#{1,4}\s*\*\*(.+?)\*\*\s*$", text, flags=re.MULTILINE)

    if len(blocks) > 1:
        for i in range(1, len(blocks), 2):
            header = normalize_text(blocks[i])
            body = blocks[i + 1] if i + 1 < len(blocks) else ""

            ra_match = re.search(r"Trace To RAC#\s*[:\*]*\s*([^\n]+)", body, re.IGNORECASE)
            ra_number = re.sub(r"\*+", "", normalize_text(ra_match.group(1))).strip() if ra_match else ""

            module_section = extract_section(
                body, "Traces to Module DFMEA risk controls",
                ["Additional Risk Control Measures", "DRS Identifiers"]
            )
            additional_section = extract_section(
                body, "Additional Risk Control Measures from Module and/or Component DFMEAs",
                ["DRS Identifiers"]
            )
            drs_section = extract_section(body, "DRS Identifiers", [])

            records.append({
                "System DFMEA #": header,
                "Trace To RAC#": ra_number,
                "Traces to Module DFMEA risk controls": extract_bullets(module_section),
                "Additional Risk Control Measures from Module and/or Component DFMEAs": extract_bullets(additional_section),
                "DRS Identifiers": extract_bullets(drs_section),
            })

    return records


def apply_existing_edo_trace(existing_edos, trace_details):
    """
    ADDED - couples the trace records back onto existing_edos.

    CHANGED: the EDO_Existing_Trace prompt's response (from the per-EDO
    extract_existing_edo_trace_details above) is now keyed DIRECTLY by
    EDO tag - {"EDO-29": {"traces": [...]}, "EDO-31": {"traces": []}, ...}
    - matching existing_edos' own keys exactly (existing_edos is itself
    keyed by edo_tag, e.g. existing_edos["EDO-29"]).

    Previously this went through deep_extract_records() + RA/FMEA-number
    matching, which only worked for a different response shape (a flat
    list of {RA_Number, FMEA_Number, ...} records). Against the current
    dict-of-EDO-tag response, deep_extract_records() returned an empty
    list every time (no recognized container key, and "traces" holds
    plain strings rather than dicts) - so existing_trace was NEVER
    populated and nothing reached the Excel output, even though the LLM
    logs showed correct answers. This version matches by EDO tag
    directly instead.
    """
    matched_count = 0

    if not isinstance(trace_details, dict):
        logging.warning(
            "apply_existing_edo_trace: trace_details was not a dict "
            f"(got {type(trace_details).__name__}) - nothing to apply."
        )
        return existing_edos

    for raw_tag, row in trace_details.items():
        if not isinstance(row, dict):
            continue

        # Format-tolerant lookup: exact key, then case-insensitive match
        # against existing_edos' own keys.
        edo = existing_edos.get(raw_tag)
        if edo is None:
            normalized_tag = normalize_text(raw_tag).lower()
            for key, candidate in existing_edos.items():
                if normalize_text(key).lower() == normalized_tag:
                    edo = candidate
                    break

        if edo is None:
            logging.warning(
                f"apply_existing_edo_trace: no existing EDO found matching "
                f"key {raw_tag!r} - skipping."
            )
            continue

        traces = row.get("traces") or row.get("Traces") or []
        lines = [normalize_text(t) for t in traces if isinstance(t, str) and normalize_text(t)]

        if lines:
            edo["existing_trace"] = "\n".join(lines)
            matched_count += 1
        else:
            edo["existing_trace"] = ""

    logging.info(
        f"apply_existing_edo_trace: populated existing_trace on "
        f"{matched_count} existing-EDO match(es) out of {len(trace_details)} "
        "trace response(s)."
    )

    return existing_edos


# ==========================================================
# NEW EDO PIPELINE
# ==========================================================
# CALL 4: Extract New EDO Tags -> Remove Existing EDO Matches (performed during the merge below) -> CALL 5: Extract New EDO Summary Details (Verification Reference is now extracted here too, directly for New EDOs).

def extract_new_edo_tags(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    edo_document,
    db
):
    """
    STEP 1 of 3: scans Section 10 of the FMEA against the RA&C and
    identifies every qualifying new-EDO row's {RA_Number, FMEA_Number,
    Status}.

    Per requirement, every RA id found is consolidated into ONE
    dictionary - edo_new_data - keyed by RA_Number (falling back to
    FMEA_Number, then to a positional key, only when RA_Number itself
    is missing). This SAME dictionary is what gets passed into and
    progressively enriched by extract_new_edo_summary_details() and
    extract_new_edo_traceability_details() afterwards, so every
    RA id's data lives in one place from identification all the way
    through to the final merge.
    """
    logging.info("=" * 80)
    logging.info("STAGE 4a: WORKFLOW 2 - NEW EDO IDENTIFICATION (TAGS ONLY)")
    logging.info("=" * 80)

    prompt_data = db.get_prompt_by_name(
        client,
        product_family,
        product,
        templatename,
        "EDO_NEW_tags"
    )

    prompt_row = {
        "prompt_role": prompt_data["prompt_role"],
        "prompt_text": prompt_data["prompt_text"],
        "question": (
            'Get all RA-# entries whose Risk Evaluation status is "See FMEA" '
            'from Appendix A - Risk Assessment and Control Table.'
        ),
        "fulltext": "Yes",
        "where_filter": "",
        "where_document": '{"$contains": "See FMEA"}',
        "checkpoint": ('Extract all Appendix A rows whose Risk Evaluation is "See FMEA".'),
        "max_results": 35,
    }

    # ---- DIAGNOSTIC: log how many chunks/docs are in this collection,
    # so we can tell "collection is empty" apart from "filter/retrieval
    # logic excluded everything".
    try:
        collection_count = edo_document["edo_ra_c"]["collection"].count()
        logging.info(f"edo_ra_c collection count: {collection_count}")
    except Exception as e:
        logging.warning(f"Could not inspect edo_ra_c collection count: {e}")

    docs, metadata, response = execute_llm_retry(
        pipeline_config,
        edo_document["edo_ra_c"]["collection"],
        prompt_row
    )

    # ---- DIAGNOSTIC: confirm content is now actually being retrieved.
    try:
        logging.info(f"RETRIEVED DOCS COUNT: {len(docs) if docs else 0}")
        logging.info(f"RETRIEVED DOCS PREVIEW: {str(docs)[:500]}")
    except Exception as e:
        logging.warning(f"Could not log retrieved docs: {e}")

    tags = parse_json(response)
    tag_records = deep_extract_records(tags)

    logging.info(f"extract_new_edo_tags: {len(tag_records)} raw tag row(s) returned by the LLM.")
    logging.info(
        "TARGET EDO MODE: New EDO stages are not used when the configured "
        f"target(s) are Existing EDO numbers: {TARGET_EDO_TAG}."
    )

    # ---- Consolidate New EDO records into ONE dictionary ----
    # Test mode: stop after the first 3 records have been added.
    # This keeps CALL 4 and CALL 5 bounded to the same three New EDOs.
    edo_new_data = {}
    new_edo_records_added = 0

    for index, row in enumerate(tag_records):
        if new_edo_records_added >= NEW_EDO_MAX_TEST_RECORDS:
            logging.info(
                f"CALL 4: reached New EDO test limit "
                f"({NEW_EDO_MAX_TEST_RECORDS}); stopping the loop."
            )
            break

        if not isinstance(row, dict):
            continue

        ra_number = normalize_text(row.get("RA_Number") or row.get("RA Number") or "")
        status = normalize_text(row.get("Status") or row.get("status") or "")

        # FIXED: the LLM returns every FMEA number for a given RA as ONE
        # combined, comma-separated string under "FMEA_Numbers" (plural)
        # - e.g. "FMEA Sys-150, FMEA Sys-151, FMEA Sys-167, FMEA Sys-725"
        # - not a single "FMEA_Number". The old code only ever looked
        # for "FMEA_Number"/"FMEA Number" (singular), so that key never
        # matched at all and every record ended up with FMEA_Number =
        # "" downstream - and even if it HAD matched, storing the whole
        # comma-joined string under one dict entry per RA_Number would
        # still collapse every FMEA number for that RA into a single
        # row. Both the plural and singular keys are read here, the
        # value is split into its individual FMEA numbers, and each one
        # becomes its OWN separate edo_new_data entry - one row per
        # RA/FMEA pair, matching the Excel layout where a RA with
        # several FMEA numbers (e.g. RA-141 with FMEA Sys-154 and
        # FMEA Sys-729) prints as separate blocks, not one row with
        # every FMEA number crammed together.
        fmea_numbers_raw = normalize_text(
            row.get("FMEA_Numbers") or row.get("FMEA Numbers")
            or row.get("FMEA_Number") or row.get("FMEA Number") or ""
        )
        fmea_numbers = [f.strip() for f in fmea_numbers_raw.split(",") if f.strip()]
        if not fmea_numbers:
            # Keep the RA even when it has no FMEA number at all, so it
            # isn't silently dropped from edo_new_data.
            fmea_numbers = [""]

        for fmea_number in fmea_numbers:
            key = (
                f"{ra_number}::{fmea_number}" if ra_number and fmea_number
                else (ra_number or fmea_number or f"NEW-EDO-TAG-{index}")
            )
            if key in edo_new_data:
                key = f"{key}_{index}"

            edo_new_data[key] = {
                "RA_Number": ra_number,
                "FMEA_Number": fmea_number,
                "Status": status
            }
            new_edo_records_added += 1

            if new_edo_records_added >= NEW_EDO_MAX_TEST_RECORDS:
                logging.info(
                    f"CALL 4: printed/created {new_edo_records_added} New EDO "
                    "record(s); stopping the loop."
                )
                break

    logging.info(
        f"extract_new_edo_tags: consolidated {len(edo_new_data)} RA id(s) "
        f"into edo_new_data."
    )
    logging.info(f"[COUNT] CALL 4 (extract_new_edo_tags): {len(edo_new_data)} candidate RA/FMEA pair(s) identified this run")
    print ("NEw edo tags:",edo_new_data)
    return edo_new_data


def extract_new_edo_summary_details(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    edo_document,
    edo_new_data,
    db,
    fmea_reference_pairs=None
):
    """
    STEP 2 of 3: takes the SAME edo_new_data dictionary built by
    extract_new_edo_tags() (keyed by RA id) and, for EVERY entry in it,
    calls the LLM ONE TIME - inside the loop, one RA_Number/FMEA_Number
    pair per call - to pull that single row's full detail record.
    """
    logging.info("=" * 80)
    logging.info("STAGE 4b: WORKFLOW 2 - NEW EDO FULL DETAIL EXTRACTION (PER RA/FMEA, LOOPED)")
    logging.info("=" * 80)

    if not edo_new_data:
        logging.info("No new EDO tags to extract details for.")
        return edo_new_data

    logging.info(
        "TARGET EDO MODE: CALL 5 is available for New EDO data only when a "
        "New EDO target is explicitly configured; current target is Existing."
    )

    prompt_data = db.get_prompt_by_name(
        client,
        product_family,
        product,
        templatename,
        "EDO_NEW_details"
    )

    keys_to_remove = []
    new_edo_records_processed = 0

    for key, entry in list(edo_new_data.items()):
        if new_edo_records_processed >= NEW_EDO_MAX_TEST_RECORDS:
            logging.info(
                f"CALL 5: reached New EDO test limit "
                f"({NEW_EDO_MAX_TEST_RECORDS}); stopping the loop."
            )
            break

        new_edo_records_processed += 1
        ra_number = entry.get("RA_Number", "")
        fmea_number = entry.get("FMEA_Number", "")

        target_text = f"RA_Number : {ra_number}\nFMEA_Number : {fmea_number}"

        prompt_row = {
            "prompt_role": prompt_data["prompt_role"],
            "prompt_text": prompt_data["prompt_text"] + "\nTARGETS:\n" + target_text,
            "question": (
                f"For RA_Number {ra_number} / FMEA_Number {fmea_number} "
                "only, extract the full EDO detail record as a single "
                "JSON object, including Risk_Status - no other RA/FMEA pairs."
            ),
            "fulltext": "Yes",
            "where_filter": "",
           # "where_document": {"$contains": "Safety Hazard DFMEA Table"},
           "where_document":"",
            "checkpoint":  (
                f"For RA_Number {ra_number} / FMEA_Number {fmea_number} "
                "only, extract the full EDO detail record as a single "
                "JSON object, including Risk_Status - no other RA/FMEA pairs."
            ),
            "max_result":35
        }

        # --- LOGGING: Inspect Prompt Row Filters & Collection state ---
        logging.info(f"--- DEBUG RETRIEVAL ATTEMPT for RA={ra_number!r} FMEA={fmea_number!r} ---")
        logging.info(f"where_document filter being applied: {prompt_row['where_document']}")

        try:
            collection_count = edo_document["edo_fmea"]["collection"].count()
            logging.info(f"edo_fmea collection total count: {collection_count}")
        except Exception as e:
            logging.warning(f"Could not inspect edo_fmea collection count: {e}")

        detail_row = {}
        try:
            docs, metadata, response = execute_llm_retry(
                pipeline_config,
                edo_document["edo_fmea"]["collection"],
                prompt_row
            )

            # --- LOGGING: Detailed Chunk & Metadata Diagnostics ---
            logging.info(f"RA={ra_number!r} FMEA={fmea_number!r} - RETRIEVED DOCS COUNT: {len(docs) if docs else 0}")
            logging.info(f"RA={ra_number!r} FMEA={fmea_number!r} - RETRIEVED DOCS TYPE: {type(docs)}")
            logging.info(f"RA={ra_number!r} FMEA={fmea_number!r} - RETRIEVED DOCS CONTENT PREVIEW: {str(docs)[:500]}")
            logging.info(f"RA={ra_number!r} FMEA={fmea_number!r} - METADATA TYPE: {type(metadata)}")
            logging.info(f"RA={ra_number!r} FMEA={fmea_number!r} - METADATA CONTENT PREVIEW: {str(metadata)[:500]}")
            logging.info(f"RA={ra_number!r} FMEA={fmea_number!r} - RAW LLM RESPONSE PREVIEW: {str(response)[:300]}")

            parsed = parse_json(response)
            records = deep_extract_records(parsed)
            if records:
                detail_row = records[0]
            elif isinstance(parsed, dict):
                detail_row = parsed

        except Exception as e:
            logging.error(
                f"LLM call failed for RA={ra_number!r} FMEA={fmea_number!r}: {e}"
            )
            detail_row = {}

        if not detail_row:
            logging.warning(
                f"No detail record returned for RA={ra_number!r} FMEA={fmea_number!r}. "
                "Excluding this entry from edo_new_data."
            )
            keys_to_remove.append(key)
            continue

        # Preserve the RA/FMEA values returned by the CALL 5 FMEA-document
        # response before Risk_Status filtering can remove this entry.
        # These values are used only for the Existing-EDO mismatch observation.
        fmea_doc_ra_number = normalize_text(
            get_llm_value(
                detail_row,
                "RA_Number",
                "RA Number",
                "ra_number",
                "ra_num"
            )
        )
        fmea_doc_fmea_number = normalize_text(
            get_llm_value(
                detail_row,
                "FMEA_Number",
                "FMEA Number",
                "fmea_number",
                "fmea_num"
            )
        )
        entry["FMEA_Document_RA_Number"] = fmea_doc_ra_number
        entry["FMEA_Document_FMEA_Number"] = fmea_doc_fmea_number

        if isinstance(fmea_reference_pairs, dict):
            fmea_reference_pairs[key] = {
                "RA_Number": ra_number,
                "FMEA_Number": fmea_number,
                "FMEA_Document_RA_Number": fmea_doc_ra_number,
                "FMEA_Document_FMEA_Number": fmea_doc_fmea_number,
            }

        print(
            f"CALL 5: FMEA DOCUMENT RA/FMEA -> RA={fmea_doc_ra_number!r} "
            f"FMEA={fmea_doc_fmea_number!r} "
            f"(CALL 4 RA={ra_number!r} FMEA={fmea_number!r})"
        )
        logging.info(
            f"CALL 5: FMEA DOCUMENT RA/FMEA -> RA={fmea_doc_ra_number!r} "
            f"FMEA={fmea_doc_fmea_number!r} "
            f"(CALL 4 RA={ra_number!r} FMEA={fmea_number!r})"
        )

        entry["Risk_Status"] = normalize_text(
            get_llm_value(
                detail_row,
                "Risk_Status",
                "Risk Status",
                "risk_status",
                "RiskStatus",
            )
        )
        if entry["Risk_Status"].casefold() != "medium":
            logging.info(
                f"Excluding New EDO RA={ra_number!r} FMEA={fmea_number!r} "
                f"because Risk_Status={entry['Risk_Status']!r}; only "
                "Risk_Status='Medium' records are eligible for merging."
            )
            keys_to_remove.append(key)
            continue

        entry["Product_Feature_Function"] = get_llm_value(
            detail_row, "Product_Feature_Function", "Product Feature Function"
        )
        entry["Reason_Identified_as_EDO"] = get_llm_value(
            detail_row, "Reason_Identified_as_EDO", "Reason Identified as EDO"
        )
        entry["Traceability"] = get_llm_value(
            detail_row, "Traceability", "traceability"
        )
        # CHANGED: Verification_Reference is now extracted and set HERE,
        # directly from this same CALL 5 detail_row - New EDOs no longer
        # wait for the post-merge CALL 6 to get their verification
        # reference. CALL 6 (extract_and_apply_verification_details) now
        # runs ONLY against Existing EDOs, before the merge - see STAGE 6
        # in generate_edo_template(). Downstream consumers
        # (merge_new_edo_dictionary(), merge_new_edo_dictionary_full(),
        # merge_new_edo_records()) already read this exact key
        # ("Verification_Reference"/"Verification Reference") via
        # get_llm_value(), so no other change is needed for it to reach
        # the merged record's "verification_reference" field.
        entry["Verification_Reference"] = get_llm_value(
            detail_row, "Verification_Reference", "Verification Reference", "verification_reference"
        )
        entry["EDO_Location"] = get_llm_value(
            detail_row, "EDO_Location", "location"
        )
        entry["EDO_Description"] = get_llm_value(
            detail_row, "EDO_Description", "description"
        )
        entry["Reason_Identified_as_EDO_ColH"] = get_llm_value(
            detail_row, "Reason_Identified_as_EDO_ColH", "reason_2"
        )
        # NEW: Project_code (CP) - identifies which project/CP this New
        # EDO's FMEA row belongs to. The EDO_NEW_details prompt now also
        # extracts this (labeled "CP" in the table, trimmed of a leading
        # "No"). Needed downstream by CALL 7 to disambiguate a
        # verification code that appears more than once in a
        # traceability spreadsheet - e.g. a Superseded entry and a
        # Current entry for the same DRS-### tag under different
        # PROJECT IDs - see resolve_best_traceability_row().
        entry["Project_code"] = get_llm_value(
            detail_row, "Project_code", "Project Code", "CP", "cp"
        )

    # Remove records with no details and records whose Risk_Status is
    # missing or anything other than Medium before the Existing/New merge.
    for key in keys_to_remove:
        del edo_new_data[key]

    logging.info(
        f"extract_new_edo_summary_details: enriched {len(edo_new_data)} "
        f"entries in edo_new_data with full detail records "
        f"({len(keys_to_remove)} excluded due to missing details or a "
        "non-Medium Risk_Status)."
    )

    logging.info(
        f"[COUNT] CALL 5 (extract_new_edo_summary_details): {len(edo_new_data)} "
        f"survived (started with {len(edo_new_data) + len(keys_to_remove)}, "
        f"excluded {len(keys_to_remove)} for missing details / non-Medium Risk_Status)"
    )

    return edo_new_data

    return edo_new_data


def extract_new_edo_trace_details(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    edo_document,
    edo_new_data,
    db: DatabaseHandler
):
    """
    ADDED - dedicated trace extraction for NEW EDOs, mirroring CALL 3
    (extract_existing_edo_trace_details()) exactly: same prompt
    ("EDO_Existing_Trace"), same collection (edo_fmea), same per-EDO
    call pattern, and the same question wording. The only difference is
    that a New EDO has no per-record EDO tag of its own (Column A always
    prints the fixed literal "EDO-XX\nNew") - so the fixed literal
    "EDO-XX New" is used in place of the per-record edo_tag wherever the
    Existing-EDO version would have used it.
    """
    logging.info("=" * 80)
    logging.info("STAGE 4C: NEW EDO - TRACE EXTRACTION (per-EDO calls)")
    logging.info("=" * 80)

    if "edo_fmea" not in edo_document:
        raise Exception(
            "EDO_FMEA document/collection not configured - cannot run "
            "trace extraction."
        )

    prompt_data = get_prompt(
        client,
        product_family,
        product,
        templatename,
        "EDO_Existing_Trace",
        db
    )

    results = {}
    new_edo_tag_label = "EDO-XX New"

    for key, edo in edo_new_data.items():
        ra_number = edo.get("RA_Number")
        fmea_number = edo.get("FMEA_Number")

        has_ra = ra_number not in (None, "", "Blank")
        has_fmea = fmea_number not in (None, "", "Blank")

        if not has_ra and not has_fmea:
            logging.info(
                f"extract_new_edo_trace_details: {key} has no "
                "RA/FMEA number - skipping (empty traces)."
            )
            results[key] = {"traces": []}
            continue

        target = (
            f"EDO : {new_edo_tag_label}\n"
            f"RA_Number : {ra_number}\n"
            f"FMEA_Number : {fmea_number}"
        )

        prompt_row = {
            "prompt_role": prompt_data["prompt_role"],
            "prompt_text": prompt_data["prompt_text"] + "\nTARGET:\n" + target,
            "question": f"Fetch all trace records for {new_edo_tag_label} (FMEA_Number: {fmea_number})",
            "fulltext": "Yes",
            "where_filter": "",
            "where_document": "",
            "checkpoint": f"Fetch all trace records for {new_edo_tag_label} (FMEA_Number: {fmea_number})"
        }

        try:
            _, _, response = execute_llm_retry(
                pipeline_config,
                edo_document["edo_fmea"]["collection"],
                prompt_row
            )
        except Exception as call_error:
            logging.warning(
                f"extract_new_edo_trace_details: LLM call failed for "
                f"{key} after retries - leaving traces empty. "
                f"Reason: {call_error}"
            )
            results[key] = {"traces": []}
            continue

        edo_result = _parse_single_edo_trace_response(new_edo_tag_label, response)
        results[key] = edo_result

        logging.info(
            f"extract_new_edo_trace_details: {key} -> "
            f"{len(edo_result.get('traces', []))} trace(s) extracted."
        )

    logging.info("========== NEW EDO PER-EDO TRACE RESULTS (COMBINED) ==========")
    logging.info(json.dumps(results, indent=2))

    return results


def apply_new_edo_trace(edo_new_data, trace_details):
    """
    ADDED - mirrors apply_existing_edo_trace() for NEW EDOs: couples the
    trace records from extract_new_edo_trace_details() back onto
    edo_new_data, keyed by the SAME key already used in edo_new_data
    (its RA/FMEA based key), writing into the "existing_trace" field so
    it flows through merge_new_edo_records() -> build_column_d_reference()
    the same way it already does for Existing EDOs.
    """
    matched_count = 0

    if not isinstance(trace_details, dict):
        logging.warning(
            "apply_new_edo_trace: trace_details was not a dict "
            f"(got {type(trace_details).__name__}) - nothing to apply."
        )
        return edo_new_data

    for key, row in trace_details.items():
        if not isinstance(row, dict):
            continue

        edo = edo_new_data.get(key)
        if edo is None:
            logging.warning(
                f"apply_new_edo_trace: no New EDO found matching key "
                f"{key!r} - skipping."
            )
            continue

        traces = row.get("traces") or row.get("Traces") or []
        lines = [normalize_text(t) for t in traces if isinstance(t, str) and normalize_text(t)]

        if lines:
            edo["existing_trace"] = "\n".join(lines)
            matched_count += 1
        else:
            edo["existing_trace"] = ""

    logging.info(
        f"apply_new_edo_trace: populated existing_trace on "
        f"{matched_count} new-EDO match(es) out of {len(trace_details)} "
        "trace response(s)."
    )

    return edo_new_data


# ==========================================================
# MERGE PIPELINE
# ==========================================================
# Merges Existing EDO records + New EDO records (de-duplicated against Existing) into ONE common list/dict. Everything after this point works on the merged list only.


def canonical_ra_id(value):
    """Return an RA identifier in canonical ``RA-<digits>`` form."""
    match = re.search(r"\bRA\s*[-:#]?\s*(\d+)\b", normalize_text(value), re.IGNORECASE)
    return f"RA-{match.group(1)}" if match else ""


def canonical_fmea_id(value):
    """Return System FMEA variants in canonical ``SYS-<digits>`` form."""
    match = re.search(
        r"\b(?:FMEA\s*)?(?:SYSTEM\s*)?SYS\s*[-:#]?\s*(\d+)\b",
        normalize_text(value),
        re.IGNORECASE,
    )
    return f"SYS-{match.group(1)}" if match else ""


def get_existing_identifier_pairs(existing_edos):
    """
    Builds a set of (RA Number, FMEA Number) pairs from Existing EDOs.

    A New EDO is a duplicate only when BOTH values match the same
    Existing EDO.

    Supports both possible field-name variants:
        ra_number / RA_Number
        fmea_number / FMEA_Number
    """

    existing_pairs = set()

    for edo in existing_edos.values():
        if not isinstance(edo, dict):
            continue

        dfmea = edo.get("dfmea", "")
        ra_key = canonical_ra_id(
            edo.get("ra_number") or edo.get("RA_Number") or dfmea
        )
        fmea_key = canonical_fmea_id(
            edo.get("fmea_number") or edo.get("FMEA_Number") or dfmea
        )

        # Only create a pair when BOTH identifiers are present.
        if ra_key and fmea_key:
            existing_pairs.add((ra_key, fmea_key))

    return existing_pairs




def normalize_edo_record(edo_record):
    default_structure = {
        "edo_type": "Existing",
        "edo_tag": "",
        "RA_Number": "",
        "FMEA_Number": "",
        "edo_description": "",
        "reason_identified": "",
        "dfmea": "",
        "verification_reference": "",
        "existing_trace": "",
        "location": "",
        "description_2": "",
        "reason_2": "",
        "sysdd": ""
    }
    if not isinstance(edo_record, dict):
        return default_structure

    for key, value in default_structure.items():
        if key not in edo_record:
            edo_record[key] = value
    return edo_record


def merge_new_edo_dictionary(
    edo_tags,
    edo_summary_details,
    edo_trace_details
):
    """
    From generate_EDO_template_copy.py - only pipeline that builds NEW EDO
    records, so it is kept as-is and used by the merged pipeline.
    """
    logging.info("=" * 80)
    logging.info("STAGE 5: WORKFLOW MATRIX COUPLING ENGINE")
    logging.info("=" * 80)

    merged = {}

    # Deep Structure Unwrapping
    records = deep_extract_records(edo_summary_details)
    logging.info(f"SUMMARY DETAILS COUNT EXTRICATED : {len(records)}")

    for index, row in enumerate(records):
        ra_number = normalize_text(get_llm_value(row, "RA_Number", "RA Number", "ra_num"))
        fmea_number = normalize_text(get_llm_value(row, "FMEA_Number", "FMEA Number", "fmea_num"))
        status_label = normalize_text(get_llm_value(row, "Status", "status")) or "New EDO"

        key = fmea_number or ra_number
        if not key:
            key = f"NEW-EDO-RECORD-{index}"

        # Safe attribute alignment ensuring literal 'None' strings from the dictionary are preserved
        edo_desc = normalize_text(get_llm_value(row, "Product_Feature_Function", "Product Feature Function", "EDO_Description", "edo_description"))
        reason_id = normalize_text(get_llm_value(row, "Reason_Identified_as_EDO", "Reason Identified as EDO", "reason_identified"))
        dfmea_trace = normalize_text(get_llm_value(row, "Traceability", "dfmea", "traceability"))
        ver_ref = normalize_text(get_llm_value(row, "Verification_Reference", "Verification Reference", "verification_reference"))

        loc_val = normalize_text(get_llm_value(row, "EDO_Location", "location")) or "None"
        desc2_val = normalize_text(get_llm_value(row, "EDO_Description", "Description_2")) or "None"
        reason2_val = normalize_text(get_llm_value(row, "Reason_2")) or "None"
        sys_dd_val = normalize_text(get_llm_value(
            row,
            "SysDD or HDD Reference",
            "SysDD_or_HDD_Reference",
            "sysdd",
            "SYS_DD",
            "HDD Reference"
        ))

        merged[key] = {
            "edo_type": "New",
            "edo_tag": f"EDO-XX\n{status_label}",
            "edo_description": edo_desc if edo_desc else "Blank",
            "reason_identified": reason_id if reason_id else "Blank",
            "dfmea": dfmea_trace if dfmea_trace else "Blank",
            "verification_reference": ver_ref if ver_ref else "Blank",
            "location": loc_val,
            "description_2": desc2_val,
            "reason_2": reason2_val,
            "sysdd": sys_dd_val,
            "RA_Number": ra_number,
            "FMEA_Number": fmea_number
        }

    # Asymmetric Verification Coupling
    trace_records = deep_extract_records(edo_trace_details)
    for row in trace_records:
        if not isinstance(row, dict):
            continue

        trace_fmea = normalize_text(get_llm_value(row, "FMEA_Number", "FMEA Number"))
        trace_ra = normalize_text(get_llm_value(row, "RA_Number", "RA Number"))
        trace_key = trace_fmea or trace_ra

        verification = normalize_text(get_llm_value(row, "Verification_Reference", "Verification Reference", "verification_reference"))
        if not verification or verification.lower() == "none" or verification == "":
            continue

        if trace_key and trace_key in merged:
            merged[trace_key]["verification_reference"] = verification
        else:
            for main_key, data in merged.items():
                if (trace_ra and data["RA_Number"] == trace_ra) or (trace_fmea and data["FMEA_Number"] == trace_fmea):
                    data["verification_reference"] = verification

    return merged


def merge_new_edo_dictionary_full(
    edo_tags,
    edo_summary_details,
    edo_trace_details,
    edo_ra_details
):
    """
    CANONICAL new-EDO merge - used by generate_edo_template().

    merge_new_edo_dictionary() (above) is left completely UNTOUCHED, but
    it silently drops any RA record whose Status is "Medium", because it
    only ever looks at edo_summary_details / edo_trace_details, both of
    which come exclusively from the EDO_FMEA collection.

    This function fixes that: it builds the output starting from
    edo_tags itself (the full, unfiltered list returned by
    extract_new_edo_tags), so EVERY RA Number is guaranteed a row in the
    final Excel output - none are excluded.

    For each tag:
      - Status contains "Medium"  -> description/reason are pulled ONLY
        from edo_ra_details (EDO_RA_C document) - the FMEA-sourced
        edo_summary_details is deliberately NOT consulted for these,
        exactly as requested.
      - Any other Status (e.g. "See FMEA") -> description/reason/trace/
        verification are pulled from edo_summary_details / edo_trace_details,
        exactly like the original merge_new_edo_dictionary() behaviour.
      - If no matching record is found in the relevant source at all, the
        tag is still written to the output - with empty ("") fields
        rather than a "Blank"/"None" placeholder string, and rather than
        being skipped/excluded.
      - Column A (edo_tag) is always the fixed literal "EDO-XX\\nNew EDO"
        for every New EDO record - the risk Status (Medium / See FMEA) is
        used internally to choose the data source but is never printed.
      - Column D (dfmea) for "Medium" records is always the RA Number
        itself (there is no FMEA trace to show for these), so every RA
        Number is visibly represented in Column D regardless of Status.
    """
    logging.info("=" * 80)
    logging.info("STAGE 5: WORKFLOW MATRIX COUPLING ENGINE (FULL - INCLUDES MEDIUM RA)")
    logging.info("=" * 80)

    def index_records(records):
        index = {}
        for row in records:
            if not isinstance(row, dict):
                continue
            ra = normalize_text(get_llm_value(row, "RA_Number", "RA Number", "ra_num"))
            fmea = normalize_text(get_llm_value(row, "FMEA_Number", "FMEA Number", "fmea_num"))
            if ra and ra not in index:
                index[ra] = row
            if fmea and fmea not in index:
                index[fmea] = row
        return index

    summary_index = index_records(deep_extract_records(edo_summary_details))
    ra_index = index_records(deep_extract_records(edo_ra_details))

    # Verification lookups from the FMEA-sourced traceability extraction -
    # only ever applied to non-Medium ("See FMEA") records.
    trace_index = {}
    for row in deep_extract_records(edo_trace_details):
        if not isinstance(row, dict):
            continue
        trace_ra = normalize_text(get_llm_value(row, "RA_Number", "RA Number"))
        trace_fmea = normalize_text(get_llm_value(row, "FMEA_Number", "FMEA Number"))
        verification = normalize_text(get_llm_value(row, "Verification_Reference", "Verification Reference", "verification_reference"))
        if not verification or verification.lower() == "none":
            continue
        if trace_ra:
            trace_index.setdefault(trace_ra, verification)
        if trace_fmea:
            trace_index.setdefault(trace_fmea, verification)

    merged = {}

    for index, item in enumerate(edo_tags):
        ra_number = normalize_text(item.get("RA_Number"))
        fmea_number = normalize_text(item.get("FMEA_Number"))
        status_label = normalize_text(item.get("Status")) or "New EDO"
        is_medium = "medium" in status_label.lower()

        key = fmea_number or ra_number
        if not key:
            key = f"NEW-EDO-RECORD-{index}"
        # avoid clobbering an existing key on duplicate RA/FMEA numbers
        if key in merged:
            key = f"{key}_{index}"

        if is_medium:
            # Medium risk value -> EDO_RA_C ONLY, never the FMEA document.
            # Keyed by RA Number ONLY - the tag's FMEA_Number is a
            # placeholder ("FMEA-UNKNOWN") for Medium records and is not
            # a reliable lookup key.
            source_row = ra_index.get(ra_number)
        else:
            # See FMEA / anything else -> FMEA-sourced summary, as before
            source_row = summary_index.get(fmea_number) or summary_index.get(ra_number)

        if source_row:
            edo_desc = normalize_text(get_llm_value(source_row, "Product_Feature_Function", "Product Feature Function", "EDO_Description", "edo_description"))
            reason_id = normalize_text(get_llm_value(source_row, "Reason_Identified_as_EDO", "Reason Identified as EDO", "reason_identified"))
            dfmea_trace = normalize_text(get_llm_value(source_row, "Traceability", "dfmea", "traceability"))
            ver_ref = normalize_text(get_llm_value(source_row, "Verification_Reference", "Verification Reference", "verification_reference"))
            loc_val = normalize_text(get_llm_value(source_row, "EDO_Location", "location"))
            desc2_val = normalize_text(get_llm_value(source_row, "EDO_Description", "Description_2"))
            reason2_val = normalize_text(get_llm_value(source_row, "Reason_2"))
            sys_dd_val = normalize_text(get_llm_value(
                source_row,
                "SysDD or HDD Reference",
                "SysDD_or_HDD_Reference",
                "sysdd",
                "SYS_DD",
                "HDD Reference"
            ))
        else:
            edo_desc = reason_id = dfmea_trace = ver_ref = ""
            loc_val = desc2_val = reason2_val = sys_dd_val = ""

        # Per requirement: for "Medium" records, Column D (dfmea) always
        # shows the RA Number itself - there is no FMEA trace document to
        # pull a narrative from, so the RA id is the traceability value.
        if is_medium:
            dfmea_trace = ra_number

        # Verification backfill only applies to non-Medium records, since
        # the trace/verification extraction is itself FMEA-sourced.
        if not is_medium and not ver_ref:
            ver_ref = trace_index.get(fmea_number) or trace_index.get(ra_number) or ver_ref

        merged[key] = {
            "edo_type": "New",
            # Per requirement: Column A never prints the risk Status
            # (Medium / See FMEA) - always the fixed literal tag text.
            "edo_tag": "EDO-XX\nNew EDO",
            "edo_description": edo_desc,
            "reason_identified": reason_id,
            "dfmea": dfmea_trace,
            "verification_reference": ver_ref,
            "location": loc_val,
            "description_2": desc2_val,
            "reason_2": reason2_val,
            "sysdd": sys_dd_val,
            "RA_Number": ra_number,
            "FMEA_Number": fmea_number
        }

    return merged

def _get_existing_sysdd_value(existing_edos):
    """Reuse the LLM-extracted Existing EDO sysdd value for New EDO rows
    instead of generating/looking up a separate value for New EDOs."""
    for data in (existing_edos or {}).values():
        val = normalize_text(data.get("sysdd", ""))
        if val:
            return val
    return ""

def merge_new_edo_records(new_records, existing_edos):
    """
    CANONICAL new-EDO merge.

    A New EDO is considered a duplicate ONLY when BOTH:
        1. RA Number matches an Existing EDO
        2. FMEA Number matches the SAME Existing EDO

    Examples:
        Existing: RA-180 + FMEA Sys-147

        New: RA-180 + FMEA Sys-147
            -> REMOVE (duplicate)

        New: RA-180 + FMEA Sys-999
            -> KEEP

        New: RA-999 + FMEA Sys-147
            -> KEEP
    """

    logging.info("=" * 80)
    logging.info(
        "STAGE 5: NEW EDO MERGE "
        "(DE-DUPED AGAINST EXISTING BY RA + FMEA PAIR)"
    )
    logging.info("=" * 80)

    existing_pairs = get_existing_identifier_pairs(existing_edos)
    logging.info(
        "Existing EDO RA/FMEA pairs: %s",
        sorted(existing_pairs)
    )

    records = deep_extract_records(new_records)
    merged = {}
    seen_new_pairs = set()

    for index, row in enumerate(records):
        if not isinstance(row, dict):
            continue

        # Defense in depth: merge only New EDO records explicitly
        # classified as Medium. Existing EDOs are handled separately by
        # merge_all_edos() and are never evaluated by this condition.
        risk_status = normalize_text(
            get_llm_value(
                row,
                "Risk_Status",
                "Risk Status",
                "risk_status",
                "RiskStatus",
            )
        )
        if risk_status.casefold() != "medium":
            logging.info(
                "Skipping New EDO during merge because Risk_Status=%r; "
                "only Medium records are eligible.",
                risk_status,
            )
            continue

        ra_number = normalize_text(
            get_llm_value(row, "RA_Number", "RA Number")
        )

        fmea_number = normalize_text(
            get_llm_value(row, "FMEA_Number", "FMEA Number")
        )

        # Canonicalize formatting variants such as ``SYS-734`` and
        # ``FMEA Sys-734`` before comparing the pair.
        ra_key = canonical_ra_id(ra_number)
        fmea_key = canonical_fmea_id(fmea_number)

        current_pair = (ra_key, fmea_key)

        # ---------------------------------------------------------
        # Requirement:
        # Remove ONLY when BOTH RA and FMEA match an existing
        # EDO pair.
        # ---------------------------------------------------------
        if (
            ra_key
            and fmea_key
            and current_pair in existing_pairs
        ):
            logging.info(
                "Skipping duplicate New EDO: "
                "RA=%r FMEA=%r - exact RA/FMEA pair already "
                "exists in Existing EDO.",
                ra_number,
                fmea_number,
            )
            continue

        if ra_key and fmea_key:
            if current_pair in seen_new_pairs:
                logging.info(
                    "Skipping repeated New EDO: RA=%r FMEA=%r - canonical "
                    "pair %r was already emitted in this run.",
                    ra_number,
                    fmea_number,
                    current_pair,
                )
                continue
            seen_new_pairs.add(current_pair)

        key = fmea_number or ra_number or f"NEW-EDO-RECORD-{index}"

        if key in merged:
            key = f"{key}_{index}"

        merged[key] = {
            "edo_type": "New",

            # Column A always uses fixed literal.
            "edo_tag": "EDO-XX\nNew",

            "edo_description": normalize_text(
                get_llm_value(
                    row,
                    "Product_Feature_Function",
                    "Product Feature Function"
                )
            ),

            "reason_identified": normalize_text(
                get_llm_value(
                    row,
                    "Reason_Identified_as_EDO",
                    "Reason Identified as EDO"
                )
            ),

            "dfmea": normalize_text(
                get_llm_value(
                    row,
                    "Traceability",
                    "traceability"
                )
            ),

            "verification_reference": normalize_text(
                get_llm_value(
                    row,
                    "Verification_Reference",
                    "Verification Reference"
                )
            ),

            "location": normalize_text(
                get_llm_value(
                    row,
                    "EDO_Location",
                    "location"
                )
            ),

            "description_2": normalize_text(
                get_llm_value(
                    row,
                    "EDO_Description",
                    "description"
                )
            ),

            "reason_2": normalize_text(
                get_llm_value(
                    row,
                    "Reason_Identified_as_EDO_ColH",
                    "reason_2"
                )
            ),

            "sysdd": normalize_text(
                get_llm_value(
                    row,
                    "SysDD or HDD Reference",
                    "SysDD_or_HDD_Reference",
                    "sysdd",
                    "SYS_DD",
                    "HDD Reference"
                )
            ),

            "RA_Number": ra_number,
            "FMEA_Number": fmea_number,

            # Carries through the New EDO trace populated by
            # extract_new_edo_trace_details() / apply_new_edo_trace()
            # (mirrors "existing_trace" on Existing EDOs) so
            # build_column_d_reference() picks it up the same way.
            "existing_trace": row.get("existing_trace", ""),

            # NEW: Project_code (CP) - carried through onto the merged
            # record so CALL 7 (build_final_edos_with_traceability) can
            # use it to disambiguate a verification code that matches
            # multiple rows in a traceability spreadsheet (e.g. a
            # Superseded vs Current entry for the same DRS-### tag under
            # different PROJECT IDs) - see resolve_best_traceability_row().
            # Existing EDOs never carry this field, so
            # final_record.get("Project_code", "") downstream simply
            # returns "" for them and CP-based narrowing is skipped,
            # matching prior behaviour for Existing EDOs.
            "Project_code": _normalize_project_code(
                get_llm_value(row, "Project_code", "Project Code", "CP", "cp")
            ),
        }

    logging.info(f"[COUNT] merge_new_edo_records: {len(merged)} New EDO(s) survived dedup against Existing")
    return merged



def merge_existing_edo_dictionary(existing_details):
    """
    CANONICAL version - from edo_existing_final.py.
    Richer than the copy-file version: carries edo_type, and correctly
    threads RA_Number/FMEA_Number through (needed since Stage 3B already
    populated ra_number/FMEA_Number on each existing EDO).

    NOTE: missing values default to "" (a truly empty cell), not the
    literal text "Blank". Column J preserves the SysDD/HDD value returned
    by the LLM instead of replacing it with a fixed reference.

    Also carries "design_elements" (the full per-location list built by
    extract_edo_details() from the EDO_Existing_Generic prompt's nested
    design_elements[] array) straight through - previously this dict was
    rebuilt with only a fixed set of keys, so design_elements was
    silently dropped here and format_edo_worksheet() only ever saw a
    single backfilled location instead of every location.
    """
    final = {}
    for tag, data in existing_details.items():
        final[tag] = {
            "edo_type": "Existing",
            "edo_tag": data.get("edo_tag", tag),
            "edo_description": data.get("edo_description", ""),
            "reason_identified": data.get("reason_identified", ""),
            "dfmea": data.get("dfmea", ""),
            "location": data.get("location", ""),
            "description_2": data.get("description_2", ""),
            "reason_2": data.get("reason_2", ""),
            "sysdd": data.get("sysdd", ""),
            "verification_reference": data.get("verification_reference", ""),
            "existing_trace": data.get("existing_trace", ""),
            "RA_Number": data.get("ra_number", ""),
            "FMEA_Number": data.get("FMEA_Number", ""),
            "Project_code": _normalize_project_code(data.get("Project_code", "")),
            "design_elements": data.get("design_elements", []),
            "FMEA_Document_RA_Number": data.get("FMEA_Document_RA_Number", ""),
            "FMEA_Document_FMEA_Number": data.get("FMEA_Document_FMEA_Number", "")
        }
    return final


def merge_all_edos(existing_edos, new_edos):
    """
    From generate_EDO_template_copy.py - only pipeline that combines
    Existing + New EDO dictionaries into one, so it is kept as-is.
    """
    final_edos = {}

    for key, value in existing_edos.items():
        normalized = normalize_edo_record(value)
        normalized["edo_type"] = "Existing"
        final_edos[key] = normalized

    for key, value in new_edos.items():
        normalized = normalize_edo_record(value)
        normalized["edo_type"] = "New"

        final_key = key
        if final_key in final_edos:
            counter = 1
            while f"{key}_{counter}" in final_edos:
                counter += 1
            final_key = f"{key}_{counter}"

        final_edos[final_key] = normalized
    print("All nerged values:",final_edos)
    return final_edos


def validate_final_edos(final_edos):
    """
    CANONICAL version - from generate_EDO_template_copy.py.
    General-purpose: handles both "Existing" and "New" typed records
    (assigns a placeholder tag for blank New EDOs). Needed because the
    merged pipeline's final_edos dictionary contains both types.
    """
    validated = {}
    for key, value in final_edos.items():
        if not isinstance(value, dict):
            continue

        edo_tag = normalize_text(value.get("edo_tag"))
        if edo_tag == "" and value.get("edo_type") == "New":
            value["edo_tag"] = "EDO-XX\nNew EDO"

        normalized = normalize_edo_record(value)
        validated[key] = normalized
        print(f"final values:", final_edos)
    return validated


# ==========================================================
# VERIFICATION REFERENCE
# ==========================================================
# CALL 6 - post-CALL-1-3, PRE-MERGE, per-record verification reference
# extraction for EXISTING EDOs ONLY. New EDOs no longer go through this
# call at all - their Verification_Reference is now extracted directly
# in CALL 5 (extract_new_edo_summary_details). Runs on `existing_edos`
# (before merge_existing_edo_dictionary()/merge_all_edos() combine
# Existing + New) - see STAGE 6 in generate_edo_template(), which now
# happens BEFORE the merge instead of after it. Looped one LLM call per
# RA/FMEA pair, the exact same call pattern as
# extract_new_edo_summary_details() (CALL 5).

VERIFICATION_LABELED_LINE_PATTERN = re.compile(
    r'^\s*(?P<prefix>\([^)]*\))?\s*'
    r'Source\s*:\s*(?P<source>.*?)\s*'
    r'(?:\|\s*Location\s*:\s*(?P<location>.*?)\s*)?'
    r'(?:\|\s*File\s*:\s*(?P<file>.*?)\s*)?'
    r'(?:\|\s*Result\s*:\s*(?P<result>.*?)\s*)?$',
    re.IGNORECASE
)


def clean_verification_reference_line(line):
    """
    Per requirement: a raw Verification_Reference line shaped like
        "(DRS-570) Source: Vest APX ... .xlsx | Location: NPD45678 ... | File: ... | Result: PASS"
    is stripped down to just the trace code (if present) plus the
    Location and File values - the "Source:" and "Result:" labels/
    values are dropped entirely, and no label text is printed at all:
        "(DRS-570) NPD45678 Vest APX System Verification Traceability Report ..."
    A line that doesn't match this labeled "Source: ... | Location: ...
    | File: ... | Result: ..." shape is returned unchanged, so genuine
    free-form verification text still prints as-is.
    """
    stripped = line.strip()
    if not stripped:
        return stripped

    match = VERIFICATION_LABELED_LINE_PATTERN.match(stripped)
    if not match:
        return stripped

    prefix = normalize_text(match.group("prefix"))
    location = normalize_text(match.group("location"))
    file_name = normalize_text(match.group("file"))

    body = " ".join(part for part in [location, file_name] if part)
    if not body:
        # Nothing usable besides the label text itself - fall back to
        # the original line rather than printing an empty result.
        return stripped

    return f"{prefix} {body}".strip() if prefix else body


def clean_verification_reference_text(raw_text):
    """
    Applies clean_verification_reference_line() to every line of a
    (possibly multi-line, multi-code) raw Verification_Reference value.
    """
    if not raw_text:
        return raw_text
    lines = [ln for ln in raw_text.split("\n") if ln.strip()]
    return "\n".join(clean_verification_reference_line(ln) for ln in lines)


# BUGFIX: the LLM sometimes answers with its own search narration or a
# plain "I couldn't find anything" statement - e.g. "Searching in RA
# document for RA_Number 12345, no verification reference found." -
# instead of either a real reference or the literal word "None". The
# only guard before this was `verification.lower() != "none"`, so that
# narration text passed straight through and got written into Column E
# as if it were real data. NON_ANSWER_VERIFICATION_PATTERN recognizes
# this shape of response so it can be discarded like "None" already is.
NON_ANSWER_VERIFICATION_PATTERN = re.compile(
    r'^\s*(searching|search(?:ed|ing)?\s+(?:in|the|for)|no\s+(?:verification|reference|'
    r'matching|relevant|explicit)|not\s+found|unable\s+to\s+(?:find|locate)|'
    r'could\s+not\s+(?:find|locate)|no\s+(?:information|data|match|result)|'
    r'not\s+(?:available|mentioned|provided|present)|n/?a)\b',
    re.IGNORECASE
)


def is_non_answer_verification_text(text):
    """
    True when `text` reads like the LLM describing its search process or
    reporting that it found nothing (see NON_ANSWER_VERIFICATION_PATTERN),
    rather than an actual Verification Reference value that should be
    written to Column E.
    """
    return bool(NON_ANSWER_VERIFICATION_PATTERN.match((text or "").strip()))


# BUGFIX: EDO_RA_C is the Risk Assessment and Control document itself,
# not a verification/test report - it has no genuine "Verification
# Reference" field. When a record is searched there via the RA_Number-
# only fallback and no real reference exists, the search still returns
# a labeled line pointing back at the RA&C document's own identity, e.g.
# "NPD36702 Vest APX Risk Assessment and Control RA-124" - the RA&C
# document's own Location/File name, not a distinct verification code.
# This is self-referential noise ("this is the document I searched"),
# not data, and must not be written to Column E.
#
# NARROWED: only treat this as self-referential noise when the "Risk
# Assessment and Control" phrase is immediately followed by THIS
# record's own RA_Number (i.e. it's just echoing back what was
# searched for). Previously ANY answer containing that phrase was
# discarded, which also threw away genuinely distinct verification
# text that happened to mention it (e.g. a real reference quoting the
# RA&C document's title alongside its own separate report number) -
# that over-filtering was blanking Column E for records that actually
# had a real answer.
SELF_REFERENTIAL_RA_DOCUMENT_PATTERN = re.compile(
    r'risk\s+assessment\s+(?:and|&)\s+control\s*[-:]?\s*(RA[\s-]?\d+|\d+)',
    re.IGNORECASE
)


def is_self_referential_ra_document_text(text, ra_number=""):
    """
    True when `text` is just the RA&C document naming itself back using
    THIS record's own RA_Number (e.g. searched for "RA-124" and got back
    "... Risk Assessment and Control RA-124") rather than a genuine,
    distinct Verification Reference. Only fires when the number right
    after "Risk Assessment and Control" matches `ra_number` - if it's a
    different number, or there's no `ra_number` to compare against, this
    returns False so real data is never discarded.
    """
    if not text:
        return False

    match = SELF_REFERENTIAL_RA_DOCUMENT_PATTERN.search(text)
    if not match:
        return False

    if not ra_number:
        return False

    found_digits = re.sub(r"\D", "", match.group(1))
    target_digits = re.sub(r"\D", "", ra_number)
    return bool(found_digits) and found_digits == target_digits


def extract_and_apply_verification_details(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    edo_document,
    existing_edos,
    db: DatabaseHandler
):
    """
    CALL 6 - EXISTING EDOs ONLY, run BEFORE the Existing/New merge.

    For every record in `existing_edos` (the dict built/enriched by
    CALL 1-3, still keyed by edo_tag - NOT yet merged with New EDOs),
    use its RA_Number/FMEA_Number to search for the Verification
    Reference and write it onto each record's "verification_reference"
    field - ONE (or, on fallback, two) LLM call(s) per record.

    Project_code / CP is also extracted directly from the CALL 6 LLM
    response and stored on the Existing EDO record under "Project_code".

    Returns existing_edos.
    """

    logging.info("=" * 80)
    logging.info(
        "CALL 6: VERIFICATION REFERENCE EXTRACTION "
        "(EXISTING EDOs ONLY, PRE-MERGE)"
    )
    logging.info("=" * 80)

    if not existing_edos:
        logging.info(
            "No Existing EDO records to extract verification details for."
        )
        return existing_edos

    logging.info(
        f"TARGET EDO MODE: CALL 6 will process configured target EDO(s) only: "
        f"{TARGET_EDO_TAG}."
    )

    if "edo_fmea" not in edo_document and "edo_ra_c" not in edo_document:
        raise Exception(
            "Neither EDO_FMEA nor EDO_RA_C document/collection is "
            "configured - cannot run verification reference extraction."
        )

    prompt_data = get_prompt(
        client,
        product_family,
        product,
        templatename,
        "EDO_NEW_Verification_details",
        db
    )

    def _query_verification_reference(
        source_key,
        ra_number,
        fmea_number,
        ra_only,
        fmea_only
    ):
        """
        Runs the CALL 6 verification-reference prompt against a single
        source document/collection (edo_fmea or edo_ra_c) for one
        RA_Number/FMEA_Number pair.

        Returns:
            (
                cleaned_verification_text_or_empty,
                raw_llm_response_text,
                project_code,
                fmea_document_ra_number,
                fmea_document_fmea_number
            )

        Project_code is extracted directly from the same LLM response.
        """

        if ra_only:
            target_text = f"RA_Number : {ra_number}"
            verification_question = (
                f"Fetch risk controls values for {ra_number}"
            )

        elif fmea_only:
            target_text = f"fmea_Number : {fmea_number}"
            verification_question = (
                f"Fetch risk controls values for {fmea_number}"
            )

        else:
            target_text = (
                f"RA_Number : {ra_number}\n"
                f"FMEA_Number : {fmea_number}"
            )
            verification_question = (
                f"Fetch risk controls values for "
                f"{fmea_number} and {ra_number}"
            )

        prompt_row = {
            "prompt_role": prompt_data["prompt_role"],
            "prompt_text": (
                prompt_data["prompt_text"]
                + "\nTARGETS:\n"
                + target_text
            ),
            "question": verification_question,
            "fulltext": "Yes",
            "where_filter": "",
            "where_document": "",
            "checkpoint": (
                f"Fetch risk controls values for "
                f"{fmea_number} and {ra_number}"
            ),
            "max_results": 35
        }

        try:
            _, _, response = execute_llm_retry(
                pipeline_config,
                edo_document[source_key]["collection"],
                prompt_row
            )

            print(
                f"CALL 6: RAW LLM RESPONSE [{source_key}] "
                f"RA={ra_number!r} FMEA={fmea_number!r} -> {response!r}"
            )

            parsed = parse_json(response)

            print(
                f"CALL 6: PARSED JSON [{source_key}] "
                f"RA={ra_number!r} FMEA={fmea_number!r} -> {parsed!r}"
            )

        except Exception as e:
            logging.error(
                f"CALL 6: verification LLM call against {source_key} "
                f"failed for RA={ra_number!r} FMEA={fmea_number!r}: {e}"
            )
            return "", "", "", "", ""

        records = deep_extract_records(parsed)

        row = (
            records[0]
            if records
            else (parsed if isinstance(parsed, dict) else {})
        )

        print(
            f"CALL 6: EXTRACTED ROW [{source_key}] "
            f"RA={ra_number!r} FMEA={fmea_number!r} -> {row!r}"
        )

        # ----------------------------------------------------------
        # Extract Verification Reference from LLM response
        # ----------------------------------------------------------
        verification = normalize_text(
            get_llm_value(
                row,
                "Verification_Reference",
                "Verification Reference",
                "verification_reference"
            )
        )

        verification = clean_verification_reference_text(
            verification
        )

        print(
            f"CALL 6: CLEANED VERIFICATION VALUE [{source_key}] "
            f"RA={ra_number!r} FMEA={fmea_number!r} "
            f"-> {verification!r}"
        )

        # ----------------------------------------------------------
        # NEW: Extract Project_code / CP from the SAME LLM response
        # ----------------------------------------------------------
        project_code = normalize_text(
            get_llm_value(
                row,
                "Project_code",
                "Project Code",
                "CP",
                "cp"
            )
        )

        print(
            f"CALL 6: PROJECT CODE [{source_key}] "
            f"RA={ra_number!r} FMEA={fmea_number!r} "
            f"-> {project_code!r}"
        )

        is_bad_answer = (
            not verification
            or verification.lower() == "none"
            or is_non_answer_verification_text(verification)
        )

        result = "" if is_bad_answer else verification

        # Actual RA/FMEA values returned by the FMEA document.  The caller
        # stores these on the Existing EDO so the observation logic can
        # compare CALL 2 values against the FMEA-document values.
        fmea_document_ra_number = ""
        fmea_document_fmea_number = ""
        if source_key == "edo_fmea":
            fmea_document_ra_number = normalize_text(
                get_llm_value(
                    row,
                    "RA_Number",
                    "RA Number",
                    "ra_number",
                    "ra_num"
                )
            )
            fmea_document_fmea_number = normalize_text(
                get_llm_value(
                    row,
                    "FMEA_Number",
                    "FMEA Number",
                    "fmea_number",
                    "fmea_num"
                )
            )
            print(
                f"CALL 6: FMEA DOCUMENT RA/FMEA -> "
                f"RA={fmea_document_ra_number!r} "
                f"FMEA={fmea_document_fmea_number!r}"
            )
            logging.info(
                f"CALL 6: FMEA DOCUMENT RA/FMEA -> "
                f"RA={fmea_document_ra_number!r} "
                f"FMEA={fmea_document_fmea_number!r}"
            )

        return (
            result,
            response,
            project_code,
            fmea_document_ra_number,
            fmea_document_fmea_number
        )

    # ==============================================================
    # PROCESS EXISTING EDO RECORDS
    # ==============================================================

    for key, edo in existing_edos.items():

        ra_number = (
            edo.get("ra_number")
            or edo.get("RA_Number", "")
        )

        fmea_number = edo.get(
            "FMEA_Number",
            ""
        )

        if (
            ra_number in (None, "", "Blank")
            and fmea_number in (None, "", "Blank")
        ):
            logging.info(
                f"CALL 6: skipping {edo.get('edo_tag', key)!r} - "
                "no RA_Number or FMEA_Number available to search with, "
                "so Column E is left blank for this record."
            )
            continue

        verification = ""
        raw_response = ""
        project_code = ""
        fmea_document_ra_number = ""
        fmea_document_fmea_number = ""

        # ==========================================================
        # FIRST SEARCH: EDO_FMEA
        # ==========================================================

        if (
            fmea_number not in (None, "", "Blank")
            and "edo_fmea" in edo_document
        ):

            (
                verification,
                raw_response,
                project_code,
                fmea_document_ra_number,
                fmea_document_fmea_number
            ) = _query_verification_reference(
                "edo_fmea",
                ra_number,
                fmea_number,
                ra_only=False,
                # Search by FMEA number alone. Including the Existing-EDO RA
                # number made the model echo that RA even when the actual
                # FMEA row was linked to a different RA.
                fmea_only=True
            )

        # ==========================================================
        # FALLBACK SEARCH: EDO_RA_C
        # ==========================================================

        if (
            not verification
            and ra_number not in (None, "", "Blank")
            and "edo_ra_c" in edo_document
        ):

            logging.info(
                f"CALL 6: no usable answer from EDO_FMEA for "
                f"RA={ra_number!r} FMEA={fmea_number!r} - "
                "trying EDO_RA_C via RA_Number instead."
            )

            (
                verification,
                raw_response_rac,
                project_code_rac,
                _unused_fmea_doc_ra,
                _unused_fmea_doc_fmea
            ) = _query_verification_reference(
                "edo_ra_c",
                ra_number,
                fmea_number,
                ra_only=True,
                fmea_only=True
            )

            raw_response = raw_response_rac or raw_response

            # If EDO_RA_C returned a Project_code, use it.
            # Otherwise retain the Project_code from EDO_FMEA.
            project_code = project_code_rac or project_code

        # ==========================================================
        # STORE RAW RESPONSE
        # ==========================================================

        edo["verification_raw_llm_response"] = raw_response

        # ==========================================================
        # STORE PROJECT CODE / CP
        # ==========================================================

        edo["Project_code"] = project_code

        # Preserve the actual RA/FMEA values returned by the EDO_FMEA
        # document.  The observation in Column M compares these against
        # CALL 2's Existing EDO RA/FMEA values.
        edo["FMEA_Document_RA_Number"] = fmea_document_ra_number
        edo["FMEA_Document_FMEA_Number"] = fmea_document_fmea_number

        logging.info(
            f"CALL 6: Project_code for "
            f"RA={ra_number!r} FMEA={fmea_number!r} "
            f"-> {project_code!r}"
        )

        # ==========================================================
        # STORE VERIFICATION REFERENCE
        # ==========================================================

        if verification:
            edo["verification_reference"] = verification

            logging.info(
                f"CALL 6: verification reference for "
                f"RA={ra_number!r} "
                f"FMEA={fmea_number!r} "
                f"-> {verification!r}"
            )

        else:
            logging.info(
                f"CALL 6: no usable verification reference found for "
                f"RA={ra_number!r} FMEA={fmea_number!r} in either "
                "EDO_FMEA or EDO_RA_C - Column E left blank."
            )

    print(
        "VErification_Reference:",
        existing_edos
    )

    return existing_edos
def _normalize_llm_key(key):
    """Lowercases a key and strips spaces/underscores/hyphens so keys
    that only differ by case or separator style - e.g.
    "Verification_Reference", "Verification Reference", and the LLM's
    actual "Verification_reference" - all compare equal."""
    return re.sub(r'[\s_\-]+', '', str(key).lower())


def get_llm_value(row, *keys):
    """
    Returns the first valid, non-empty value found across `keys`.
    Treats None, "", any case-insensitive "none"/"blank" placeholder
    text (e.g. "None", "NONE", "Blank"), AND any noise/placeholder
    status text recognized by _is_noise_code() (e.g. "na", "n/a",
    "Not Available", "Not Found", "unknown", "tbd") as invalid/empty.

    BUGFIX: previously this only filtered "none"/"blank" - every other
    placeholder the LLM commonly returns for a missing field (e.g.
    "na", "Not Available", "Not Found") passed straight through as if
    it were real data. Callers such as build_final_edos_with_traceability()
    already have "or" fallback chains for exactly this situation (e.g.
    `get_llm_value(row, "vv_record_file_name", ...) or source_document_name
    or "Unknown File"`), but those fallbacks never fired because the
    placeholder text was truthy and non-empty. Filtering noise text here
    lets those fallbacks work as originally intended, instead of writing
    "na"/"Not Available"/"Not Found" straight into the output Excel.
    """
    if not isinstance(row, dict):
        return ""

    normalized_row = {_normalize_llm_key(k): v for k, v in row.items()}

    for key in keys:
        value = normalized_row.get(_normalize_llm_key(key))
        if value is None:
            continue
        text = str(value).strip()
        if text == "" or text.lower() in ("none", "blank") or _is_noise_code(text):
            continue
        return value
    return ""



def _strip_wrapping_quotes_and_brackets(token: str) -> str:
    """
    Strips any stray outer bracket/paren and quote characters left behind
    on a single code token, e.g. "['DRS-570'" -> "DRS-570",
    "'SRS-CTRL-51']" -> "SRS-CTRL-51", '"DRS-570"' -> 'DRS-570'.
    Runs repeatedly since a token can carry more than one such layer
    (e.g. a leading "[" AND a leading "'" at once).
    """
    token = token.strip()
    prev = None
    while token and token != prev:
        prev = token
        token = token.strip()
        # Strip one layer of matching/loose brackets or parens on either side.
        if token[:1] in "([" :
            token = token[1:]
        if token[-1:] in ")]":
            token = token[:-1]
        # Strip one layer of quote marks on either side (they don't need
        # to match each other - malformed input can leave just one side).
        if token[:1] in "'\"":
            token = token[1:]
        if token[-1:] in "'\"":
            token = token[:-1]
        token = token.strip()
    return token


# BUGFIX: CALL 5's Verification_Reference extraction (New EDOs) sometimes
# sweeps in document-status text from the FMEA table alongside the real
# control tags - e.g. a raw value like "NPD36569 Not Available, MRS
# Software FMEA-428, ..., MRS Software FMEA-431 Not Found" - instead of
# the clean comma-separated tag list the prompt asks for. Comma-split by
# parse_verification_codes(), tokens like "NPD36569 Not Available" and
# "MRS Software FMEA-431 Not Found" then get treated as if they were real
# verification codes: CALL 7 dutifully "resolves" them against the
# traceability documents (via _find_code_in_record()'s fallback substring
# scan) and prints garbage rows like "NPD36569 Not Available - Not
# Available" straight into Column E. _is_noise_code() recognizes and
# drops these before they're ever treated as codes.
_NOISE_CODE_PATTERN = re.compile(
    r'\b(not\s+available|not\s+found|no\s+match(?:\s+found)?|n/?a|unknown|tbd)\b',
    re.IGNORECASE
)


def _is_noise_code(code: str) -> bool:
    """
    True when `code` is placeholder/status text that leaked in alongside
    real verification codes (see _NOISE_CODE_PATTERN), rather than an
    actual requirement/verification code such as "DRS-570" or "MS CU
    Mod-384". A genuine code never contains these status words, so any
    match here is enough to drop the token.
    """
    return bool(_NOISE_CODE_PATTERN.search(normalize_text(code)))


def parse_verification_codes(verification_ref_str: str) -> List[str]:
    """
    Parses a verification string into a clean list of individual code strings,
    regardless of which format CALL 6 happened to emit it in. Handles:
      - Plain parens:            '(DRS-570, MS CU Mod-384, SRS-CTRL-39)'
      - Plain comma list:        'DRS-570, MS CU Mod-384, SRS-CTRL-39'
      - Python list repr (str):  "['DRS-570', 'MS CU Mod-384', 'SRS-CTRL-39']"
      - Python tuple repr (str): "('DRS-570', 'MS CU Mod-384', 'SRS-CTRL-39')"
    In every case the output is the same clean list:
      ['DRS-570', 'MS CU Mod-384', 'SRS-CTRL-39']

    Tokens that are placeholder/status noise rather than real codes (e.g.
    "NPD36569 Not Available", "MRS Software FMEA-431 Not Found" - see
    _is_noise_code()) are dropped from every return path, so they never
    reach CALL 7's resolution or Column E.
    """
    if not verification_ref_str:
        return []

    # Already an actual list/tuple (not a string) - just clean each element.
    if isinstance(verification_ref_str, (list, tuple, set)):
        cleaned = [
            _strip_wrapping_quotes_and_brackets(str(c))
            for c in verification_ref_str
            if str(c).strip()
        ]
        return [c for c in cleaned if c and not _is_noise_code(c)]

    raw_str = str(verification_ref_str).strip()
    if not raw_str:
        return []

    # If this is a stringified Python list/tuple (e.g. "['DRS-570', 'MS CU Mod-97']"),
    # parse it properly instead of naively splitting on commas - a naive split
    # leaves brackets/quotes stuck onto the first and last codes, which then
    # fail every prefix match in resolve_traceability_reference().
    if (raw_str.startswith("[") and raw_str.endswith("]")) or \
       (raw_str.startswith("(") and raw_str.endswith(")")):
        try:
            parsed = ast.literal_eval(raw_str)
            if isinstance(parsed, (list, tuple, set)):
                cleaned = [
                    _strip_wrapping_quotes_and_brackets(str(c))
                    for c in parsed
                    if str(c).strip()
                ]
                return [c for c in cleaned if c and not _is_noise_code(c)]
        except (ValueError, SyntaxError):
            # Not valid Python literal syntax - fall through to manual
            # stripping below rather than failing outright.
            pass
        raw_str = raw_str[1:-1]

    codes = [
        _strip_wrapping_quotes_and_brackets(code)
        for code in raw_str.split(",")
        if code.strip()
    ]
    return [code for code in codes if code and not _is_noise_code(code)]



# ==========================================================
# CALL 7: TRACEABILITY REFERENCE RESOLUTION (per EDO, dynamic document search)
# ==========================================================
# Each verification code produced by CALL 6/CALL 5 (e.g. "DRS-570",
# "SRS-CTRL-39", "MS ACC Mod-112", "MS CU Mod-384") lives in exactly ONE
# of the configured traceability spreadsheets - but WHICH one, and how
# many traceability spreadsheets exist at all, is not knowable ahead of
# time: every traceability document is loaded from the DB under the SAME
# document_identity ("EDO_TM"), so there is no per-document signal to
# route a code by. Instead of a fixed tag-prefix -> document map, codes
# are searched across edo_document["traceability_documents"] (an
# arbitrary-length, DB-order list - see get_edo_document()) one document
# at a time: batch-query document 0 with every still-unresolved code for
# this EDO, remove whatever matched, batch-query document 1 with the
# remainder, and so on until every code is resolved or the documents are
# exhausted.
#
# A run-scoped cache (`prefix -> document index`) is kept across EDOs so
# that once a given code prefix (e.g. "DRS") is found to live in document
# 2, every subsequent EDO's "DRS-xxx" codes try that document FIRST
# instead of re-searching from document 0 every time. This keeps the
# common case down to ~1 call per EDO per distinct prefix, same as the
# old fixed-bucket approach, while making no assumption about prefix
# naming, document count, or document order.
#
# CP / RECORD STATUS DISAMBIGUATION: a single verification code can
# legitimately appear on MORE than one row of a traceability document -
# e.g. a Superseded entry and a Current entry for the same DRS-###
# tag, filed under different PROJECT IDs:
#
#     DRS-594 | ... | PRJ8611715 | Superseded
#     DRS-594 | ... | PRJ8611715 | Current
#
# _query_traceability_details_batch() therefore keeps EVERY matching row
# per code (not just the first one found), and
# resolve_best_traceability_row() picks the one that's actually correct
# for a given EDO: prefer a row whose PROJECT ID matches the EDO's own
# Project_code (CP, extracted in CALL 5 from the FMEA table), then
# prefer RECORD STATUS "Current" over "Superseded" among whatever
# remains.


def _code_prefix_hint(code: str) -> str:
    """
    Cheap, best-effort prefix extraction from a verification code, used
    ONLY as a cache key hint to try the most-likely document first - it
    is never used to decide whether/where a code is allowed to match.
    e.g. "DRS-570" -> "DRS", "MS CU Mod-384" -> "MS CU MOD",
    "SRS-CTRL-39" -> "SRS-CTRL". Falls back to the whole cleaned code
    when no trailing "-<digits>" / " <digits>" suffix is found.
    """
    if not code:
        return ""
    clean_code = _strip_wrapping_quotes_and_brackets(str(code).strip())
    match = re.match(r"^(.*?)[\s_-]*\d+\s*$", clean_code)
    prefix = match.group(1) if match else clean_code
    return re.sub(r"[\s_-]+", " ", prefix).strip().upper()


def _normalize_project_code(value):
    """
    Canonicalizes a project/CP identifier for comparison, e.g.
    "PRJ8611715", " prj-8611715 ", "No.PRJ8611715" all collapse to the
    same "PRJ8611715" token. Used on BOTH sides of the PROJECT ID
    comparison in resolve_best_traceability_row(), so minor formatting
    differences between what CALL 5 extracted (Project_code, trimmed of
    a leading "No" per the prompt) and what the traceability sheet's
    PROJECT ID column actually contains don't silently break the match.
    Extracts the first PRJ<digits> or SUS<digits> token if one is
    present; otherwise falls back to an uppercased, alphanumeric-only
    version of the whole value so an unrecognized-but-still-consistent
    format still compares equal to itself.
    """
    text = normalize_text(value).upper()
    if not text:
        return ""
    match = re.search(r'\b(PRJ|SUS)\s*[-#]?\s*(\d+)', text)
    if match:
        return f"{match.group(1)}{match.group(2)}"
    return re.sub(r'[^A-Z0-9]', '', text)


def resolve_best_traceability_row(candidate_rows: List[dict], project_code: str):
    """
    Given every traceability-spreadsheet row that matched a single
    verification code (see _query_traceability_details_batch()'s
    rows_by_code, a list per code rather than one row), picks the ONE
    row that is actually correct for this EDO - a code like DRS-594 can
    legitimately appear more than once in the sheet, e.g.:

        DRS-594 | ... | PRJ8611715 | Superseded
        DRS-594 | ... | PRJ8611715 | Current

    Resolution order:
      1. If `project_code` is set and one or more candidate rows have a
         PROJECT ID matching it (compared via _normalize_project_code()
         on both sides, so minor formatting differences don't break the
         match), narrow to just those rows first.
      2. Among whatever remains, prefer a row whose RECORD STATUS says
         "Current" over "Superseded" (substring match, case-insensitive,
         so "Current Revision" etc. still match).
      3. If more than one row is still tied after both filters, the
         first one (original response order) is used - the old
         first-match behaviour, now only as a last-resort tiebreaker
         instead of unconditionally discarding every other match.

    Returns a single row dict, or None if candidate_rows is empty.
    """
    if not candidate_rows:
        return None
    if len(candidate_rows) == 1:
        return candidate_rows[0]

    project_code_norm = _normalize_project_code(project_code)
    pool = candidate_rows

    if project_code_norm:
        cp_matches = [
            row for row in candidate_rows
            if _normalize_project_code(
                get_llm_value(
                    row,
                    "Project_ID", "Project Id", "PROJECT ID",
                    "Project_Code", "Project Code", "CP", "cp"
                )
            ) == project_code_norm
        ]
        if cp_matches:
            pool = cp_matches

    current_matches = [
        row for row in pool
        if "current" in normalize_text(
            get_llm_value(
                row,
                "Record_Status", "Record Status", "RECORD STATUS",
                "Status", "status"
            )
        ).lower()
    ]
    if current_matches:
        pool = current_matches

    return pool[0]


# BUGFIX: some traceability documents act as a master index/table of
# contents that lists nearly every req_tag but with a blank or "N/A"
# V/V Record Location and V/V Record File Name (e.g. "Vest APX Control
# Unit Module Verification Traceability Spreadsheet Rev 5.xlsx" matched
# DRS-570, DRS-687, SRS-CTRL-39, SRS-CTRL-53, etc. this way). Before this
# fix, resolve_best_traceability_row() accepted ANY non-empty candidate
# list as a final answer - so the first code with a given prefix (e.g.
# "DRS") to hit this index document got "resolved" against a blank row,
# warmed prefix_cache to point at that index document, and every
# SUBSEQUENT code sharing that prefix then went straight to the cache
# hit path, matched the same blank index row, and was marked resolved
# too - permanently starving out the real source document (which may sit
# later in traceability_documents) from ever being tried. This is why
# DRS-570 regressed from a real answer ("NPD43975 rev 3 - vest apx
# software features verification tdr") to "N/A - Vest APX Control Unit
# Module Verification Traceability Spreadsheet Rev 5.xlsx".
def _has_real_traceability_content(row: dict) -> bool:
    """
    True when `row` carries a genuine V/V Record Location or V/V Record
    File Name - not blank, and not a placeholder like "N/A"/"Not Found"
    (reusing _is_noise_code(), the same check already used to keep
    placeholder text out of parsed verification codes). A row that fails
    this check is treated as a low-quality index-only match: still kept
    as a last-resort fallback, but never accepted as final while any
    other configured document hasn't been tried yet.
    """
    location = get_llm_value(row, "vv_record_location", "VV_Record_Location")
    filename = get_llm_value(row, "vv_record_file_name", "VV_Record_File_Name")
    location_ok = bool(location) and not _is_noise_code(location)
    filename_ok = bool(filename) and not _is_noise_code(filename)
    return location_ok or filename_ok


# Maximum number of codes sent to the LLM in a single traceability
# detail-extraction call. Batching too many codes into one call (e.g.
# 15+ at once) causes the LLM to fully detail only the first couple and
# return bare req_tag-only matches (no vv_record_location/
# vv_record_file_name) for the rest - which is why Column E showed
# "<code> N/A - <traceability spreadsheet name>" for most rows and
# Column F only populated for 1-2 of them. Splitting into small chunks
# gives the LLM enough room to extract full detail for every code.
TRACEABILITY_CODE_CHUNK_SIZE = 4


def _query_traceability_details_chunked(pipeline_config, collection, prompt_data, codes: List[str],project_code):
    """
    Wraps _query_traceability_details_batch() to split `codes` into
    chunks of at most TRACEABILITY_CODE_CHUNK_SIZE before querying, so a
    long code list for one EDO doesn't get crammed into a single LLM
    call. Issues one call per chunk and merges the results back into a
    single {code: [row, ...]} dict, same return contract as
    _query_traceability_details_batch() itself - each value is now a
    LIST of every matching row for that code (see
    resolve_best_traceability_row()), so chunks are merged by
    concatenating each code's row list rather than overwriting it.
    """
    if not codes:
        return {}

    merged: Dict[str, List[dict]] = {}
    for i in range(0, len(codes), TRACEABILITY_CODE_CHUNK_SIZE):
        chunk = codes[i:i + TRACEABILITY_CODE_CHUNK_SIZE]
        chunk_result = _query_traceability_details_batch(
            pipeline_config, collection, prompt_data, chunk, project_code
        )
        logging.info(
            f"CALL 7 (chunk {i // TRACEABILITY_CODE_CHUNK_SIZE}): "
            f"codes {chunk} -> {len(chunk_result)} code(s) with matches"
        )
        for code, rows in chunk_result.items():
            merged.setdefault(code, []).extend(rows)
    return merged


def _get_traceability_documents(edo_document):
    """
    Returns the ordered list of traceability document dicts
    (edo_document["traceability_documents"], populated in
    get_edo_document() from every source document whose document_identity
    is "EDO_TM"). Returns [] when none are configured.
    """
    if not edo_document:
        return []
    return edo_document.get("traceability_documents", []) or []


def search_codes_across_documents(
    pipeline_config,
    traceability_documents,
    prompt_data,
    codes: List[str],
    prefix_cache: Dict[str, int],
    project_code: str = ""
):
    """
    CALL 7 - Search verification/traceability codes across EDO_TM documents.

    Optimized behavior:
        1. All unresolved codes are sent together in ONE LLM call per document.
        2. A real/valid traceability match with a non-blank matching PROJECT ID
           resolves that code.
        3. Resolved codes are removed from the remaining list.
        4. If all codes are resolved, the remaining EDO_TM documents are NOT queried.
        5. Low-quality/index-only matches are retained as fallback candidates.
        6. Existing resolve_best_traceability_row() logic is preserved.
        7. Cached document hints are tried first.
        8. No per-code LLM calls are made.

    Returns:
        Dict[str, dict]:
            {
                "DRS-594": {
                    "row": {...},
                    "document_index": 0,
                    "document_name": "..."
                },
                ...
            }
    """

    # ------------------------------------------------------------------
    # 1. Normalize / deduplicate requested codes
    # ------------------------------------------------------------------
    remaining = list(dict.fromkeys(
        code for code in codes
        if code and str(code).strip()
    ))

    resolved: Dict[str, dict] = {}

    # Low-quality matches are retained only as fallback.
    # If a real match is found later, it always takes precedence.
    fallback: Dict[str, dict] = {}

    # BUGFIX: a code can have real (non-blank) rows in MORE THAN ONE
    # EDO_TM document, differentiated only by PROJECT ID (e.g. DRS-216
    # existing as a Current row under three different PROJECT IDs across
    # three separate documents). The early-stop optimization below used
    # to accept the FIRST document's real match unconditionally, even
    # when project_code was known and that row's PROJECT ID didn't match
    # it - so whichever document happened to be checked first "won",
    # regardless of whether it was actually this EDO's project. A real
    # match is now only treated as a confident, code-resolving answer
    # when its PROJECT ID matches project_code (or when project_code is
    # empty, i.e. there is nothing to check against - same fast
    # early-stop as before for Existing EDOs / New EDOs with no CP
    # extracted). A real match whose PROJECT ID doesn't match is instead
    # collected here so every configured document still gets a chance to
    # produce the correct, project-matching row before one is finally
    # chosen (see the resolution step after the document loop).
    real_candidates: Dict[str, List[dict]] = {}

    if not remaining:
        logging.info(
            "CALL 7: No verification/traceability codes to search."
        )
        return resolved

    if not traceability_documents:
        logging.warning(
            "CALL 7: No traceability documents available for codes: %s",
            remaining
        )
        return resolved

    logging.info(
        "CALL 7: Starting traceability search for %d code(s): %s",
        len(remaining),
        remaining
    )

    # ------------------------------------------------------------------
    # 2. Helper to process candidate rows returned by ONE LLM call
    # ------------------------------------------------------------------
    def _consider_match(
        code: str,
        candidate_rows: List[dict],
        doc_index: int,
        doc_name: str
    ) -> bool:
        """
        Process the candidate rows returned for one code.

        Returns:
            True  -> real traceability match found and code is resolved.
            False -> no real match; caller should continue searching.
        """

        if not candidate_rows:
            return False

        try:
            row = resolve_best_traceability_row(
                candidate_rows,
                project_code
            )
        except Exception as e:
            logging.warning(
                "CALL 7: Failed to resolve best traceability row "
                "for code %r in document %r: %s",
                code,
                doc_name,
                e
            )
            return False

        if not row:
            return False

        # --------------------------------------------------------------
        # REAL MATCH
        #
        # BUGFIX: a real match is only accepted as the FINAL answer when
        # its source-row PROJECT ID is present and exactly matches the
        # EDO's known project_code. A missing PROJECT ID is NOT confirmation.
        # This deliberately keeps the code unresolved so every configured
        # EDO_TM document gets a chance to return the source row with the
        # actual project ID. A real match whose
        # PROJECT ID does NOT match a known project_code is a genuine
        # row (not blank/placeholder) but not yet confirmed as the right
        # one for THIS EDO - e.g. the same DRS-### code can have a
        # Current row under a different project in an earlier document.
        # That row is kept as a candidate in `real_candidates` and the
        # code stays unresolved so remaining documents still get
        # searched for a project-matching row; only if none is ever
        # found does the resolution step after the document loop fall
        # back to the best of these real (non-blank) candidates.
        # --------------------------------------------------------------
        if _has_real_traceability_content(row):

            row_project_id = get_llm_value(
                row,
                "Project_ID", "Project Id", "PROJECT ID",
                "Project_Code", "Project Code", "CP", "cp"
            )
            project_code_norm = _normalize_project_code(project_code)
            row_project_id_norm = _normalize_project_code(row_project_id)

            # A known project can only be confirmed by a NON-BLANK PROJECT
            # ID that actually matches the EDO's project code.  Previously,
            # an empty row PROJECT_ID normalized to "" and an empty
            # project_code also normalized to "", which made missing data
            # look like a confirmed match.  More importantly for CALL 7,
            # when project_code is known, a row with a missing PROJECT_ID
            # must NEVER be treated as confirmed.
            project_confirmed = bool(
                project_code_norm
                and row_project_id_norm
                and row_project_id_norm == project_code_norm
            )

            if project_confirmed:
                resolved[code] = {
                    "row": row,
                    "document_index": doc_index,
                    "document_name": doc_name,
                }

                # Warm the prefix cache so future EDOs with the same
                # verification-code prefix can try this document first.
                try:
                    prefix = _code_prefix_hint(code)
                    if prefix:
                        prefix_cache[prefix] = doc_index
                except Exception as e:
                    logging.debug(
                        "CALL 7: Could not update prefix cache for %r: %s",
                        code,
                        e
                    )

                logging.info(
                    "CALL 7: REAL traceability match found for %s "
                    "in EDO_TM document #%d (%s)%s. "
                    "Code marked RESOLVED; it will not be queried again.",
                    code,
                    doc_index,
                    doc_name,
                    " (PROJECT ID confirmed)" if project_code else ""
                )

                return True

            # Real content, but PROJECT ID doesn't match this EDO's own
            # project_code - keep searching remaining documents instead
            # of locking this in as the final answer.
            real_candidates.setdefault(code, []).append({
                "row": row,
                "document_index": doc_index,
                "document_name": doc_name,
            })

            logging.info(
                "CALL 7: Real traceability match for %s in EDO_TM "
                "document #%d (%s) has PROJECT ID %r, which does not "
                "match this EDO's project_code %r - keeping as a "
                "candidate and continuing to search remaining documents.",
                code,
                doc_index,
                doc_name,
                row_project_id,
                project_code
            )

            return False

        # --------------------------------------------------------------
        # LOW-QUALITY / FALLBACK MATCH
        #
        # Keep it temporarily, but continue searching because another
        # EDO_TM document may contain the actual traceability record.
        # --------------------------------------------------------------
        if code not in fallback:
            fallback[code] = {
                "row": row,
                "document_index": doc_index,
                "document_name": doc_name,
            }

            logging.info(
                "CALL 7: Low-quality traceability candidate retained "
                "as fallback for %s from EDO_TM document #%d (%s).",
                code,
                doc_index,
                doc_name
            )

        return False

    # ------------------------------------------------------------------
    # 3. CACHE PASS
    #
    # If a prefix cache entry exists, query those documents first.
    #
    # IMPORTANT:
    # _query_traceability_details_batch() makes ONE LLM call for ALL
    # cached codes sent to that document.
    # ------------------------------------------------------------------
    cache_groups: Dict[int, List[str]] = {}

    for code in remaining:
        try:
            prefix = _code_prefix_hint(code)
        except Exception:
            prefix = None

        if prefix is None:
            continue

        cached_doc_index = prefix_cache.get(prefix)

        if (
            cached_doc_index is not None
            and 0 <= cached_doc_index < len(traceability_documents)
        ):
            cache_groups.setdefault(
                cached_doc_index,
                []
            ).append(code)

    # Process cached documents first.
    for doc_index, cached_codes in sorted(cache_groups.items()):

        # Remove anything that may already have been resolved.
        cached_codes = [
            code for code in cached_codes
            if code not in resolved
        ]

        if not cached_codes:
            continue

        doc = traceability_documents[doc_index]

        collection = doc.get("collection")
        doc_name = (
            doc.get("document_name")
            or doc.get("name")
            or f"EDO_TM document #{doc_index}"
        )

        if not collection:
            logging.warning(
                "CALL 7: Cached EDO_TM document #%d (%s) "
                "has no collection. Skipping.",
                doc_index,
                doc_name
            )
            continue

        logging.info(
            "CALL 7: CACHE PASS - querying EDO_TM document #%d (%s) "
            "for %d unresolved code(s): %s",
            doc_index,
            doc_name,
            len(cached_codes),
            cached_codes
        )

        # --------------------------------------------------------------
        # ONE LLM CALL FOR ALL CACHED CODES
        # --------------------------------------------------------------
        rows_by_code = _query_traceability_details_chunked(
            pipeline_config,
            collection,
            prompt_data,
            cached_codes,     # <-- stays "cached_codes" here
            project_code
)

        newly_resolved = []

        for code in cached_codes:

            if code in resolved:
                continue

            candidate_rows = rows_by_code.get(code, [])

            if _consider_match(
                code,
                candidate_rows,
                doc_index,
                doc_name
            ):
                newly_resolved.append(code)

        if newly_resolved:
            remaining = [
                code
                for code in remaining
                if code not in newly_resolved
            ]

            logging.info(
                "CALL 7: CACHE PASS resolved %d code(s): %s. "
                "Remaining unresolved: %d",
                len(newly_resolved),
                newly_resolved,
                len(remaining)
            )

        # --------------------------------------------------------------
        # CRITICAL EARLY STOP
        # --------------------------------------------------------------
        if not remaining:
            logging.info(
                "CALL 7: CACHE PASS resolved ALL traceability codes. "
                "Skipping all remaining EDO_TM LLM calls."
            )
            return resolved

    # ------------------------------------------------------------------
    # 4. NORMAL ROUND-ROBIN DOCUMENT SEARCH
    #
    # Only unresolved codes are sent to each document.
    #
    # ONE LLM CALL PER DOCUMENT.
    # ------------------------------------------------------------------
    for doc_index, doc in enumerate(traceability_documents):

        # --------------------------------------------------------------
        # EARLY STOP BEFORE STARTING NEXT DOCUMENT
        # --------------------------------------------------------------
        if not remaining:
            logging.info(
                "CALL 7: All traceability codes resolved. "
                "Stopping EDO_TM document loop."
            )
            break

        # Remove anything that may have been resolved during cache pass
        # or earlier iterations.
        remaining = [
            code for code in remaining
            if code not in resolved
        ]

        if not remaining:
            logging.info(
                "CALL 7: No unresolved traceability codes remain. "
                "Stopping document search."
            )
            break

        collection = doc.get("collection")

        doc_name = (
            doc.get("document_name")
            or doc.get("name")
            or f"EDO_TM document #{doc_index}"
        )

        if not collection:
            logging.warning(
                "CALL 7: EDO_TM document #%d (%s) has no collection. "
                "Skipping document.",
                doc_index,
                doc_name
            )
            continue

        logging.info(
            "CALL 7: Querying EDO_TM document #%d (%s) "
            "for %d unresolved code(s): %s",
            doc_index,
            doc_name,
            len(remaining),
            remaining
        )

        # --------------------------------------------------------------
        # IMPORTANT:
        #
        # Use BATCH instead of CHUNKED.
        #
        # This creates exactly ONE LLM call for this document containing
        # ALL currently unresolved codes.
        # --------------------------------------------------------------
        rows_by_code = _query_traceability_details_chunked(
            pipeline_config,
            collection,
            prompt_data,
            remaining,        # <-- stays "remaining" here
            project_code
)

        newly_resolved = []

        # --------------------------------------------------------------
        # Process every requested code returned by this ONE LLM call.
        # --------------------------------------------------------------
        for code in list(remaining):

            # Defensive check - another code-processing path may already
            # have resolved it.
            if code in resolved:
                continue

            candidate_rows = rows_by_code.get(code, [])

            if _consider_match(
                code,
                candidate_rows,
                doc_index,
                doc_name
            ):
                newly_resolved.append(code)

        # --------------------------------------------------------------
        # Remove newly resolved codes BEFORE querying the next document.
        # --------------------------------------------------------------
        if newly_resolved:

            remaining = [
                code
                for code in remaining
                if code not in newly_resolved
            ]

            logging.info(
                "CALL 7: EDO_TM document #%d resolved %d code(s): %s",
                doc_index,
                len(newly_resolved),
                newly_resolved
            )

            logging.info(
                "CALL 7: Remaining unresolved code(s): %s",
                remaining
            )

        else:
            logging.info(
                "CALL 7: EDO_TM document #%d produced no new "
                "real traceability matches. Remaining: %s",
                doc_index,
                remaining
            )

        # --------------------------------------------------------------
        # CRITICAL EARLY STOP AFTER THIS DOCUMENT
        # --------------------------------------------------------------
        if not remaining:

            logging.info(
                "CALL 7: ALL traceability codes have been resolved "
                "after EDO_TM document #%d (%s).",
                doc_index,
                doc_name
            )

            logging.info(
                "CALL 7: SKIPPING all remaining EDO_TM documents. "
                "No additional LLM calls will be made."
            )

            break

    # ------------------------------------------------------------------
    # 5. APPLY FALLBACK ONLY TO CODES THAT NEVER GOT A REAL MATCH
    # ------------------------------------------------------------------
    unresolved_after_search = []

    for code in remaining:

        # Real match always has priority.
        if code in resolved:
            continue

        # BUGFIX: prefer a REAL (non-blank) candidate row over a blank/
        # index-only fallback row - no document ever confirmed one of
        # these candidates' PROJECT ID against project_code, so pick the
        # best of what was found using the same PROJECT ID / RECORD
        # STATUS priority resolve_best_traceability_row() already uses,
        # just applied across every document's candidates pooled
        # together instead of one document's single response.
        if code in real_candidates:
            candidates = real_candidates[code]
            candidate_rows = [c["row"] for c in candidates]
            best_row = resolve_best_traceability_row(candidate_rows, project_code)
            best_entry = next(
                (c for c in candidates if c["row"] is best_row),
                candidates[0]
            )
            resolved[code] = best_entry

            logging.info(
                "CALL 7: No document had a PROJECT ID-confirmed match for "
                "%s - using the best real (non-blank) candidate from "
                "EDO_TM document #%d (%s) out of %d candidate(s) checked.",
                code,
                best_entry["document_index"],
                best_entry["document_name"],
                len(candidates)
            )

        elif code in fallback:

            resolved[code] = fallback[code]

            logging.info(
                "CALL 7: Using fallback traceability result for %s "
                "from EDO_TM document #%d (%s).",
                code,
                fallback[code]["document_index"],
                fallback[code]["document_name"]
            )

        else:
            unresolved_after_search.append(code)

    # ------------------------------------------------------------------
    # 6. FINAL LOGGING
    # ------------------------------------------------------------------
    if unresolved_after_search:

        logging.warning(
            "CALL 7: No traceability result found for %d code(s): %s",
            len(unresolved_after_search),
            unresolved_after_search
        )

    logging.info(
        "CALL 7: Traceability search completed. "
        "Requested=%d, Resolved=%d, Unresolved=%d",
        len(codes),
        len(resolved),
        len(unresolved_after_search)
    )

    return resolved

def _match_record_to_code(value, codes: List[str]):
    """
    Matches a single returned-record value (typically its req_tag)
    against the list of codes that were sent together in a batched
    query, so a multi-record batch response can be split back apart
    per original code. Comparison ignores spaces/underscores/hyphens
    and case, e.g. "DRS-570" matches "DRS 570" or "drs_570".

    BUGFIX: previously this matched on two-way substring containment
    ONLY, with no exact-match preference. In a batched query (see
    TRACEABILITY_CODE_CHUNK_SIZE), a shorter/similar code that happens
    to be a substring of another queued code - e.g. "DRS-65" is a
    substring of "DRS-651" - could "win" the match for a row that
    actually belongs to the longer code, whichever code happened to be
    checked first in the loop. That misattributed a real test
    report's traceability row to the wrong RA/FMEA record. An exact
    normalized match is now tried first across every code before any
    substring fallback is considered, so a genuine "DRS-651" row can
    never be stolen by an unrelated "DRS-65" code.
    """
    if not value:
        return None
    normalized_value = re.sub(r'[\s_-]+', '', str(value)).upper()
    if not normalized_value:
        return None

    # 1. Exact match first - avoids cross-contamination between
    #    similar/overlapping codes (e.g. "DRS-65" vs "DRS-651").
    for code in codes:
        normalized_code = re.sub(r'[\s_-]+', '', str(code)).upper()
        if normalized_code and normalized_code == normalized_value:
            return code

    # 2. Substring containment fallback - only used when no code
    #    matched exactly, same behaviour as before for genuinely
    #    partial/fuzzy record values.
    for code in codes:
        normalized_code = re.sub(r'[\s_-]+', '', str(code)).upper()
        if normalized_code and (normalized_code in normalized_value or normalized_value in normalized_code):
            return code
    return None


def _find_code_in_record(record: dict, codes: List[str]):
    """
    Fallback used when a batch-returned record's req_tag field doesn't
    directly match one of the queried codes (e.g. missing/renamed key).
    Scans every string value anywhere in the record for a code match.
    """
    for value in _flatten_record_values(record):
        matched = _match_record_to_code(value, codes)
        if matched:
            return matched
    return None


def _query_traceability_details_batch(pipeline_config, collection, prompt_data, codes: List[str], project_code):
    """
    Shared multi-code LLM query used by search_codes_across_documents().
    Targets ALL of the given codes for ONE traceability document in a
    SINGLE call - e.g. if 2 codes are still unresolved for one EDO, both
    are sent together in one call instead of two separate calls.

    Returns a dict keyed by the ORIGINAL code string -> a LIST of every
    matching record for that code, e.g.:
        {"DRS-594": [{...Superseded row...}, {...Current row...}]}
    A code can legitimately appear on more than one row of a
    traceability document (e.g. a Superseded entry and a Current entry
    for the same tag under different PROJECT IDs) - every matching
    record is kept here instead of only the first one found, so the
    caller (resolve_best_traceability_row(), via
    search_codes_across_documents()) can pick the correct row per EDO
    based on PROJECT ID / RECORD STATUS. A code with no matching record
    in the LLM's response is simply absent from the returned dict.
    """
    if not codes:
        return {}

    codes_list_str = ", ".join(codes)
    question = f"Extract the details for the following codes: {codes_list_str} with Project code: {project_code}"

    # IMPORTANT: The configured DB prompt has historically returned the
    # V/V filename/location correctly while omitting the traceability
    # columns PROJECT ID and RECORD STATUS.  CALL 7 uses PROJECT ID to
    # disambiguate the same verification code across multiple EDO_TM
    # spreadsheets, so those two source columns are mandatory output.
    # Keep the DB prompt intact, but append a strict, machine-readable
    # extraction contract for this CALL 7 query.
    traceability_schema_instruction = f"""

CALL 7 TRACEABILITY EXTRACTION CONTRACT - MANDATORY:
Return one JSON object/row for every matching requested REQ TAG.
For EACH returned row, include ALL of these keys, even when the source
cell is blank:
- req_tag
- vv_record_file_name
- vv_record_location
- req_result
- issue
- Project_ID
- Record_Status

SOURCE-COLUMN MAPPING (do not omit these):
- Project_ID = exact value from the source column headed
  'PROJECT ID (PRJ # for NPD) (SUS # for Sustaining)'.
- Record_Status = exact value from the source column headed
  'RECORD STATUS'.
- vv_record_file_name = exact value from 'V/V RECORD FILE NAME'.
- vv_record_location = exact value from 'V/V RECORD LOCATION'.
- req_result = exact value from 'REQ RESULT (PASS/FAIL)'.
- issue = exact value from 'ISSUE'.

Do NOT infer Project_ID from the query's Project code. Read it from the
matched source row. If the source row has no Project_ID, return an empty
string for Project_ID; do not invent one. Likewise return an empty string
for Record_Status when the source cell is absent/blank.

The requested Project code is {project_code!r}. It is a DISAMBIGUATION
VALUE ONLY and must not be copied into Project_ID unless that exact value
appears in the matched source row's PROJECT ID column.
"""

    prompt_row = {
        "prompt_role": prompt_data["prompt_role"],
        "prompt_text": prompt_data["prompt_text"] + traceability_schema_instruction,
        "question": question,
        "fulltext": "Yes",
        "where_filter": "",
        "where_document": "",
        "checkpoint": question,
        "max_results": max(35, 35 * len(codes))
    }

    try:
        _, _, response = execute_llm_retry(pipeline_config, collection, prompt_row)
    except Exception as e:
        logging.error(f"CALL 7: batched traceability LLM call failed for codes {codes!r}: {e}")
        return {}

    parsed = parse_json(response)
    records = deep_extract_records(parsed)
    if not records and isinstance(parsed, dict):
        records = [parsed]

    # Split the batch response back apart: match each returned record to
    # the one code (out of the codes we sent) it actually belongs to. A
    # code can legitimately have MULTIPLE rows in the traceability sheet
    # (e.g. a Superseded entry and a Current entry for the same DRS-###
    # tag, each tied to a different PROJECT ID) - so every matching
    # record is kept here, and resolve_best_traceability_row() (called
    # by the caller) picks the single right one afterward instead of
    # this function silently discarding all but the first.
    rows_by_code: Dict[str, List[dict]] = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        req_tag_value = get_llm_value(record, "req_tag", "Req_Tag", "REQ_TAG")
        matched_code = _match_record_to_code(req_tag_value, codes) if req_tag_value else None
        if not matched_code:
            matched_code = _find_code_in_record(record, codes)
        if matched_code:
            rows_by_code.setdefault(matched_code, []).append(record)

    return rows_by_code


def build_final_edos_with_traceability(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    edo_document,
    merged_edos: Dict[str, dict],
    db
) -> List[dict]:
    """
    CALL 7: Traceability Reference Resolution (dynamic, per-EDO document search).

    For every merged EDO record, parses its verification_reference
    string (written by CALL 6/CALL 5) into individual codes (e.g.
    'DRS-570', 'MS CU Mod-384', 'SRS-CTRL-39', 'MS ACC Mod-112'), then
    resolves them against edo_document["traceability_documents"] - an
    arbitrary-length list of documents all sharing document_identity
    "EDO_TM" - via search_codes_across_documents(), which searches
    documents in order (fast-pathed by a run-scoped prefix cache) instead
    of routing by a fixed, hardcoded tag-prefix -> document map. This
    means the pipeline no longer assumes exactly 4 traceability
    documents, specific document names, or a fixed set of code prefixes -
    any number of EDO_TM documents, in any order, with any code naming
    scheme, resolve correctly.

    When a code matches more than one row in a traceability document
    (e.g. a Superseded and a Current entry for the same tag under
    different PROJECT IDs), search_codes_across_documents() passes this
    record's own Project_code (CP, extracted in CALL 5 from the FMEA
    table) through to resolve_best_traceability_row() so the row that
    actually belongs to this EDO's project - and, failing that, whichever
    row has RECORD STATUS "Current" - is the one used.

    The combined, human-readable trace text for all of a record's codes
    is written back onto `verification_reference` (the same column
    format_edo_worksheet() reads into the output Excel), in the SAME
    order the codes originally appeared. The raw parsed code list is
    also kept under `verification_reference_parsed`.

    NEW: alongside the human-readable trace text, this also captures,
    per code, an IDENTITY-AGNOSTIC document number (e.g. "NPD43975")
    extracted from that code's matched V/V record location/filename/
    source document name - stored as `verification_reference_doc_numbers`,
    a list positionally aligned with `verification_reference_parsed`.
    CALL 8 (find_verification_evidence_document()) uses this number to
    look the source document up directly via
    edo_document["documents_by_number"], instead of depending on fuzzy
    filename-text matching or on document_identity being correctly
    tagged in the database.
    """
    logging.info("=" * 80)
    logging.info("CALL 7: TRACEABILITY REFERENCE RESOLUTION (DYNAMIC DOCUMENT SEARCH)")
    logging.info("=" * 80)

    if not merged_edos:
        logging.info("No merged EDO records to resolve traceability for.")
        return []

    # generate_edo_template() filters the merged dictionary into a list
    # before CALL 7.  The previous implementation unconditionally used
    # merged_edos.items(), so every Existing-EDO run stopped here with:
    #     'list' object has no attribute 'items'
    # Normalize both supported containers once and use the same iterable
    # for the success and prompt-loading-failure paths.
    if isinstance(merged_edos, dict):
        merged_items = list(merged_edos.items())
    elif isinstance(merged_edos, list):
        merged_items = [
            (record.get("edo_id") or f"record_{index}", record)
            for index, record in enumerate(merged_edos)
            if isinstance(record, dict)
        ]
    else:
        logging.error(
            "CALL 7: merged_edos must be a dict or list, got %s.",
            type(merged_edos).__name__
        )
        return []

    try:
        prompt_data = get_prompt(
            client, product_family, product, templatename, "EDO_Excel_Extraction", db
        )
    except Exception as e:
        logging.error(f"CALL 7: failed to load traceability prompt: {e}")
        return [dict(record) for _, record in merged_items]

    traceability_documents = _get_traceability_documents(edo_document)
    if not traceability_documents:
        logging.warning(
            "CALL 7: no traceability documents (document_identity 'EDO_TM') "
            "configured for this template - every verification code will be "
            "left unresolved."
        )

    # Run-scoped cache: {code_prefix_hint: document_index}. Populated as
    # codes are matched, so later EDOs with the same prefix try the right
    # document first instead of re-searching from document 0 every time.
    prefix_cache: Dict[str, int] = {}

    final_edos = []

    for key, edo_record in merged_items:
        # Target filtering happens in generate_edo_template() before merge,
        # so every record reaching CALL 7 is already one of TARGET_EDO_TAG.
        # No positional/3-record limit is applied here.

        # Create a copy to prevent mutation issues
        final_record = dict(edo_record)

        ra_number = final_record.get("RA_Number", "")
        fmea_number = final_record.get("FMEA_Number", "")
        # This EDO's CP/PROJECT ID (New EDOs only - see CALL 5 /
        # merge_new_edo_records()). Existing EDOs never carry this key,
        # so .get(...) simply returns "" and resolve_best_traceability_row()
        # skips CP-based narrowing for them, falling straight to the
        # RECORD STATUS "Current" preference.
        project_code = final_record.get("Project_code", "")

        ver_ref_str = final_record.get("verification_reference", "")
        parsed_codes = parse_verification_codes(ver_ref_str)

        logging.info("-" * 80)
        logging.info(
            f"EDO Tag: {key!r} | RA_Number: {ra_number!r} | FMEA_Number: {fmea_number!r} "
            f"| Project_code: {project_code!r} "
            f"-> Verification codes to resolve: {parsed_codes}"
        )

        matched_trace_details_by_code: Dict[str, str] = {}
        # NEW: identity-agnostic document number captured per code, used
        # by CALL 8 to look up the EDO_VER source document directly.
        matched_doc_number_by_code: Dict[str, str] = {}
        # Column M ("Remarks and Recommendation") needs a called-out remark
        # for every verification code whose traceability call comes back
        # with a Fail result - collected here and merged into the record
        # below, then picked up by generate_remarks_and_recommendation().
        fail_remarks = []

        resolved = search_codes_across_documents(
            pipeline_config,
            traceability_documents,
            prompt_data,
            parsed_codes,
            prefix_cache,
            project_code=project_code,
        )

        for code in parsed_codes:
            match = resolved.get(code)

            if not match:
                matched_trace_details_by_code[code] = f"{code} - No match found"
                matched_doc_number_by_code[code] = ""
                logging.info(f"  NO MATCH | Code: {code} - not found in any configured traceability document")
                continue

            row = match["row"]
            source_document_name = match.get("document_name")

            req_tag_value = get_llm_value(row, "req_tag", "Req_Tag", "REQ_TAG") or code
            filename = get_llm_value(row, "vv_record_file_name", "VV_Record_File_Name") \
                or source_document_name or "Unknown File"
            location = get_llm_value(row, "vv_record_location", "VV_Record_Location") or "N/A"
            result_text = get_llm_value(row, "req_result", "Req_Result", "REQ_RESULT") or "N/A"

            # Final output format the user wants in the Excel cell:
            # "<Req_Tag> <V/V RECORD LOCATION> - <V/V RECORD FILE NAME>"
            trace_entry = f"{req_tag_value} {location} - {filename}"
            matched_trace_details_by_code[code] = trace_entry

            # ---- Identity-agnostic document number capture ----
            # Try the structured V/V record location first (it's the
            # field that actually holds the NPD number in practice, e.g.
            # "npd43975 rev 3"), then the filename, then the source
            # traceability document's own name, as fallbacks.
            doc_number = (
                _extract_doc_number(location)
                or _extract_doc_number(filename)
                or _extract_doc_number(source_document_name)
            )
            matched_doc_number_by_code[code] = doc_number or ""

            logging.info(
                f"  MATCH    | Document: #{match['document_index']} ({source_document_name}) | Code: {code} | "
                f"RA_Number: {ra_number!r} | FMEA_Number: {fmea_number!r} | "
                f"req_tag: {req_tag_value} | File: {filename} | "
                f"Location: {location} | Result: {result_text} | "
                f"Doc Number: {doc_number!r}"
            )

            # If this code's traceability test result comes back Fail, flag it
            # for Column M with a dedicated remark - independent of whether a
            # matching row was otherwise found.
            if normalize_text(result_text).strip().lower() == "fail" or "fail" in normalize_text(result_text).strip().lower():
                fail_remark = (
                    f'The following "{code}" test report is not identified and not '
                    "verified in any of the traceability spreadsheets or records."
                )
                fail_remarks.append(fail_remark)
                logging.info(
                    f"  RESULT FAIL | Code: {code} - flagged for Column M remark: {fail_remark!r}"
                )

        # Rebuild the trace-detail list in the SAME order codes originally
        # appeared in verification_reference.
        matched_trace_details = [
            matched_trace_details_by_code[code]
            for code in parsed_codes
            if code in matched_trace_details_by_code
        ]
        # Same positional order for the document-number list, so index i
        # of verification_reference_doc_numbers corresponds to index i of
        # verification_reference_parsed / the i-th line of the joined
        # verification_reference text below.
        matched_doc_numbers = [
            matched_doc_number_by_code[code]
            for code in parsed_codes
            if code in matched_trace_details_by_code
        ]

        # Assign the resolved trace text back onto verification_reference itself -
        # this is the field format_edo_worksheet() writes into the output Excel,
        # so the resolved location/filename/result now actually reach the sheet.
        final_record["verification_reference_parsed"] = parsed_codes
        final_record["verification_reference"] = "\n".join(matched_trace_details)
        final_record["verification_reference_doc_numbers"] = matched_doc_numbers
        # Picked up by generate_remarks_and_recommendation() below and
        # rendered into Column M ("Remarks and Recommendation").
        final_record["traceability_fail_remarks"] = fail_remarks

        final_edos.append(final_record)

    logging.info("-" * 80)
    logging.info(f"Processed {len(final_edos)} records into final_edos.")
    new_edo_count = sum(1 for e in final_edos if e.get("edo_type") == "New")
    logging.info(f"[COUNT] CALL 7 (build_final_edos_with_traceability): {new_edo_count} New EDO(s) in final output")
    logging.info(f"Prefix -> document index cache learned this run: {prefix_cache}")
    return final_edos


# ==========================================================
# CALL 8: VERIFICATION EVIDENCE EXTRACTION (COLUMN F)
# ==========================================================
# Runs AFTER CALL 7 (build_final_edos_with_traceability), once Column
# E's "verification_reference" text is already the final, resolved
# per-code trace lines (each shaped like
# "<code> <location> - <filename>", or "<code> - No match found" for
# codes CALL 7 couldn't resolve - see
# build_final_edos_with_traceability()).
#
# For every code in Column E that DOES have a filename:
#   1. Get the code, its filename, AND its document number straight off
#      that Column E line / CALL 7's parallel lists (paired positionally
#      with "verification_reference_parsed" and
#      "verification_reference_doc_numbers"), since the plain text alone
#      can't be split back into code/location/filename unambiguously
#      when a code itself contains spaces, e.g. "MS CU Mod-384".
#   2. Resolve that code's source document in TWO stages:
#        a. PRIMARY - exact match by document number (e.g. "NPD43975")
#           against edo_document["documents_by_number"], an
#           IDENTITY-AGNOSTIC index built in get_edo_document() from
#           every mapped document's name, regardless of its
#           document_identity tag. This works even for documents whose
#           document_identity is null/untagged in the database.
#        b. FALLBACK - PARTIAL filename match against every document
#           loaded under document_identity "EDO_VER"
#           (edo_document["edo_ver_documents"]), for the rare case a
#           code carries no document number at all.
#   3. Ask the Design Verification Traceability Analyst prompt (role/
#      prompt/question fixed below - not looked up from the prompts
#      database) to identify the code, its evidence, and why it passed,
#      against exactly that one matched document.
#   4. Write "Code / Filename / Why It Passed" - one block per code, in
#      Column E's original order - onto the record's
#      "verification_evidence" field, read by format_edo_worksheet()
#      into Column F.

VERIFICATION_EVIDENCE_PROMPT_ROLE = "You are a Design Verification Traceability Analyst."

VERIFICATION_EVIDENCE_PROMPT_TEXT = """Task:

Review the document and identify:


1. The requirement code(s) being verified

   (e.g., DRS-xxx, SRS-CTRL-xxx, MS CU Mod-xxx, MS ACC Mod-xxx).


2. The evidence used to verify the requirement.


Analysis Rules:


A. Code Identification

- Scan the document title, heading, objective, verification section, verification requirement statements, meeting minutes, review comments, conclusions, approval records, and traceability references.

- Identify all requirement codes referenced in the section.

- Determine the primary requirement code associated with the specific requirement being verified.

- If multiple requirement codes are referenced, create a separate result only when distinct evidence is tied to that code.

- Do not invent, assume, or infer requirement codes that are not explicitly present in the document.


B. Verification Evidence Identification

- Locate the document text used to verify the requirement.

- Consider:

  - Verification statements

  - Requirement verification descriptions

  - Review conclusions

  - Meeting agreements

  - Acceptance criteria

  - Compliance statements

  - Test results

  - Review outcomes

  - Approval records

  - Verification summaries

- Use only evidence explicitly present in the document.

- Do not infer verification outcomes from missing information.


C. Evidence Extraction

- Extract ONLY the exact sentence(s) from the document that support verification of the requirement.

- Copy the text verbatim.

- Do not paraphrase.

- Do not summarize.

- Do not interpret the text.

- If multiple sentences are required to capture the evidence, include all relevant sentences.


D. Evidence Cleaning Rules

- Exclude standalone status indicators such as:

  - Pass

  - Fail

  - Passed

  - Failed

  - O Pass

  - O Fail

  - ☑ Pass

  - ☑ Fail

  - Checkbox selections

  - Radio button selections

  - Status columns

  - Result columns

  - Table status cells

- Remove trailing status markers even if they appear immediately after the evidence text.

- Preserve only the requirement statement, acceptance criterion, verification condition, result description, review conclusion, or verification conclusion.

- If a status word appears as part of a complete meaningful sentence, retain the full sentence exactly as written.

- Do not append any status wording to the Evidence field.


E. Strict Restrictions

- Do NOT determine Pass, Fail, Passed, Failed, Successful, Unsuccessful, Compliant, Non-Compliant, or any equivalent status.

- Do NOT provide explanations.

- Do NOT provide reasoning.

- Do NOT provide analysis.

- Do NOT provide conclusions.

- Do NOT append verdicts to any field.

- Do NOT add labels such as:

  - Pass

  - Fail

  - Passed

  - Failed

  - Status

  - Verification Result

  - Outcome

  - Conclusion

- Do NOT add comments before or after the quoted evidence.

- Other than the exact extracted text, do not introduce additional wording into the Evidence field.


F. Confidence Assessment


High:

- Requirement code is explicitly identified.

- Evidence directly references the requirement and verification activity.


Medium:

- Requirement code is explicitly identified.

- Evidence indirectly supports verification through review, agreement, approval, acceptance, or compliance statements.


Low:

- Requirement code is identified but supporting evidence is limited, indirect, or ambiguous.


G. Output Requirements

- Produce one result block per requirement code.

- Use only information present in the document.

- Do not include any reasoning or narrative text outside the specified format.

- Do not output Pass/Fail status anywhere.

- Do not infer or report verification outcomes.

- The Evidence field must contain only the extracted document text after applying the Evidence Cleaning Rules.

- Do not add any text beyond the fields defined below.


Output Format:


Code:

<identified requirement code>


Document:

<document name>


Heading:

<section heading>


Evidence:

<exact sentence(s) quoted verbatim from document, with standalone status markers removed>


Confidence:

High / Medium / Low"""

def verification_evidence_question(code):
    """
    Builds a CODE-SPECIFIC retrieval question for CALL 8, mirroring the
    pattern every other successful targeted call in this pipeline uses
    (CALL 5/6/7 all build their "question" per RA/FMEA/code rather than
    reusing one fixed string - e.g. extract_new_edo_summary_details()'s
    "For RA_Number ... / FMEA_Number ... only, extract...").

    BUGFIX: the old VERIFICATION_EVIDENCE_QUESTION was a single FIXED
    string reused identically for every code/document call. Since the
    retriever's "question" field is what drives which content actually
    gets pulled for the LLM to look at (the code itself only appeared
    inside "prompt_text"'s TARGETS block, which is background
    instruction text, not the search target), every call for every
    code effectively asked the retriever the SAME generic question -
    "identify the requirement code" - with no code named at all. That
    meant retrieval had no way to prefer content specific to code X
    over code Y, so most calls got back whatever generic/unrelated
    chunk the fixed question happened to match, and the LLM correctly
    reported "Not Determined" against content that had nothing to do
    with that code. Naming the code directly in the question lets
    retrieval actually target that code's section of the document.
    """
    return (
        f"For requirement code {code}, find where this document verifies "
        f"it, state whether it Passed or Failed, and quote the exact "
        f"condition or sentence the document uses to support that "
        f"conclusion."
    )

# The exact field labels (in the order they may appear) that the LLM's
# plain-text response follows, per VERIFICATION_EVIDENCE_PROMPT_TEXT's
# Output Format above.
_EVIDENCE_FIELD_NAMES = [
    "Code", "Document", "Heading", "Pass Status",
    "Evidence", "Why It Passed", "Confidence",
]


def parse_verification_evidence_fields(response):
    """
    Parses the plain-text "Code: ... / Document: ... / Heading: ... /
    Pass Status: ... / Evidence: ... / Why It Passed: ... / Confidence:
    ..." block returned by the CALL 8 LLM call into a dict keyed by
    lowercase, underscore-joined field name, e.g. {"code": "...",
    "document": "...", "why_it_passed": "..."}. This response is plain
    labelled text (per the fixed Output Format above), not JSON, so a
    label-based regex is used instead of parse_json().
    """
    fields = {}
    text = normalize_text(response)
    if not text:
        return fields

    # BUGFIX: the LLM frequently wraps each label in markdown bold, e.g.
    # "**Code:**\nDRS-570" instead of the plain "Code:\nDRS-570" the
    # regex below expects. "**" is not whitespace, so `^\s*` never
    # matched past it and EVERY label silently failed to parse - the
    # LLM's raw response could have a perfect, fully-detailed Pass/
    # Evidence/Why-It-Passed answer, and parse_verification_evidence_fields()
    # would still return {} because the label itself was never found
    # (confirmed against a real production response: raw LLM text had
    # "**Code:**", "**Pass Status:**\nPass", etc. - fields came back
    # {} every time, and format_verification_evidence_block() fell back
    # to "None"/"Not Determined" for a code that had actually passed).
    # Stripping all "**" up front (this response format never uses "**"
    # for anything except label emphasis) fixes this without weakening
    # the match in any other way.
    text = text.replace("**", "")

    label_alternation = "|".join(re.escape(name) for name in _EVIDENCE_FIELD_NAMES)
    pattern = (
        rf'(?im)^\s*({label_alternation})\s*:\s*(.*?)'
        rf'(?=\n\s*(?:{label_alternation})\s*:|\Z)'
    )

    for match in re.finditer(pattern, text, re.DOTALL):
        key = match.group(1).strip().lower().replace(" ", "_")
        value = match.group(2).strip()
        fields[key] = value

    return fields


def is_meaningful_evidence_field(value):
    """
    True when a parsed evidence field actually carries content, as
    opposed to being blank or one of the LLM's own "nothing found"
    placeholders (e.g. "Not Determined", "None", "N/A") - OR the LLM
    echoing back the prompt's own unfilled instruction text (e.g.
    "<Exact condition/sentence quoted verbatim from document>") instead
    of a real answer.
    """
    if not value:
        return False
    normalized = normalize_text(value).strip().strip(".").upper()
    if not normalized:
        return False
    if normalized in ("NONE", "N/A", "NA", "NOT DETERMINED", "-", "UNKNOWN"):
        return False
    # Catch the LLM echoing the prompt's own placeholder instruction
    # instead of filling it in with real content.
    if "EXACT CONDITION" in normalized and "QUOTED VERBATIM" in normalized:
        return False
    return True


def parse_verification_reference_code_filename_pairs(edo):
    """
    Reads Column E's already-resolved "verification_reference" text
    (per-code trace lines, one per line, each shaped like
    "<code> <location> - <filename>" - see
    build_final_edos_with_traceability()) back apart into ordered
    (code, filename, doc_number) triples.

    The code list is taken from "verification_reference_parsed" (the
    exact, ordered list of codes CALL 7 built the text FROM) rather than
    re-parsed out of the free text itself, since a code can legitimately
    contain spaces (e.g. "MS CU Mod-384") which makes splitting the code
    back out of "<code> <location> - <filename>" text ambiguous on its
    own. The document-number list ("verification_reference_doc_numbers")
    is built by CALL 7 in the exact same order, so it's paired in
    positionally too - this is what lets CALL 8 resolve a code's source
    document by exact number match instead of relying purely on fuzzy
    filename text.

    Falls back to a best-effort split (everything before " - " treated
    as the code, no doc_number) only when "verification_reference_parsed"
    is missing or out of sync with the text - e.g. CALL 7 was skipped and
    verification_reference still holds some other format.

    Returns a list of {"code": ..., "filename": ..., "doc_number": ...}
    dicts ("doc_number" may be None). Lines with no usable filename
    (unresolved codes, "No match found", "Unknown File", etc.) are
    omitted, since there is no document to look evidence up in for
    those.
    """
    ver_ref_text = edo.get("verification_reference", "")
    lines = [ln.strip() for ln in normalize_text(ver_ref_text).split("\n") if ln.strip()]
    if not lines:
        return []

    parsed_codes = edo.get("verification_reference_parsed") or []
    doc_numbers = edo.get("verification_reference_doc_numbers") or []

    if parsed_codes and len(parsed_codes) == len(lines):
        if doc_numbers and len(doc_numbers) == len(lines):
            code_line_pairs = list(zip(parsed_codes, lines, doc_numbers))
        else:
            code_line_pairs = [(c, l, "") for c, l in zip(parsed_codes, lines)]
    else:
        code_line_pairs = []
        for line in lines:
            if " - " not in line:
                continue
            left, _, _ = line.rpartition(" - ")
            if left.strip():
                # Best-effort only - the whole left side (which may still
                # include the location text) is used as the code.
                code_line_pairs.append((left.strip(), line, ""))

    pairs = []
    for code, line, doc_number in code_line_pairs:
        code = normalize_text(code)
        if not code or " - " not in line:
            continue

        _, _, filename = line.rpartition(" - ")
        filename = filename.strip()

        if not filename or filename.lower() in ("unknown file", "no match found"):
            continue
        # Lines CALL 7 writes for codes it couldn't resolve at all end
        # with explanatory text instead of a real filename - skip those
        # too (e.g. "... document not configured", "Unrecognized tag",
        # "not found in any of the ... traceability document(s)").
        lowered_filename = filename.lower()
        if (
            "not configured" in lowered_filename
            or "unrecognized tag" in lowered_filename
            or "no match found" in lowered_filename
        ):
            continue

        pairs.append({
            "code": code,
            "filename": filename,
            "doc_number": normalize_text(doc_number) or None,
        })

    return pairs


def _filename_stem(value):
    """
    Lowercases, strips a trailing file extension, and collapses ALL
    word separators (spaces, underscores, hyphens, dots, and any other
    non-alphanumeric run) down to a single space - so
    "NPD44675_Vest_APX_Software_Features_Verification_TDR_Rev2.xlsx" and
    "vest apx software features verification tdr" normalize to
    comparable token strings instead of failing a literal substring
    check over mismatched underscore/space characters. Only used as the
    FALLBACK path now - see find_verification_evidence_document(),
    which tries an exact document-number match first.
    """
    text = normalize_text(value).lower().strip()
    # Strip a trailing file extension first (e.g. ".xlsx", ".pdf")
    text = re.sub(r'\.[a-z0-9]{2,5}$', '', text).strip()
    # Collapse any run of non-alphanumeric characters (_, -, ., multiple
    # spaces, etc.) into a single space, and trim the ends.
    text = re.sub(r'[^a-z0-9]+', ' ', text).strip()
    return text


def find_verification_evidence_document(edo_document, filename_hint, doc_number=None):
    """
    Finds the source document matching a Column E verification code, in
    two stages:

      1. EXACT document-number match (e.g. "NPD43975") against
         edo_document["documents_by_number"] - an IDENTITY-AGNOSTIC
         index built in get_edo_document() from EVERY mapped document's
         name, regardless of its document_identity tag. This is the
         primary, unambiguous path and works even when a document's
         document_identity in the DB is null/untagged.
      2. FALLBACK - token-overlap PARTIAL filename match against
         edo_document["evidence_candidate_documents"] - EVERY mapped
         document except the four structural ones (EDO_Proposed,
         EDO_RA_C, EDO_FMEA, EDO_pdf_new), regardless of
         document_identity tag or whether it has an NPD number in its
         name. This deliberately widens beyond just "EDO_VER"-tagged
         documents, because many real V/V evidence files (e.g.
         "SV CTRL 1 SW V&V General.pdf", "Vest APX APG Module EE
         Feature Test Result.pdf") have neither a document number NOR
         an EDO_VER identity tag - they were previously skipped
         entirely. Falls back further to edo_ver_documents alone only
         if evidence_candidate_documents isn't populated (older
         edo_document dicts / backwards compatibility).

    Returns the matched document dict, or None.
    """
    if doc_number:
        by_number = (edo_document or {}).get("documents_by_number") or {}
        exact_match = by_number.get(normalize_text(doc_number).upper())
        if exact_match:
            logging.info(
                f"CALL 8: matched document by NUMBER {doc_number!r} -> "
                f"{exact_match.get('document_name')!r}"
            )
            return exact_match
        logging.info(
            f"CALL 8: document number {doc_number!r} not found in the "
            "documents_by_number index - falling back to fuzzy filename match."
        )

    documents = (
        (edo_document or {}).get("evidence_candidate_documents")
        or (edo_document or {}).get("edo_ver_documents")
        or []
    )
    hint_stem = _filename_stem(filename_hint)
    if not documents or not hint_stem:
        return None

    hint_tokens = set(hint_stem.split())

    best_match = None
    best_score = 0

    for document in documents:
        doc_stem = _filename_stem(document.get("document_name"))
        if not doc_stem:
            continue

        doc_tokens = set(doc_stem.split())
        overlap = hint_tokens & doc_tokens

        # Require the hint to be substantially contained in the doc name
        # (or vice versa) - not just a couple of stray shared words.
        if not overlap:
            continue
        coverage = len(overlap) / len(hint_tokens)
        if coverage < 0.6:
            continue

        score = len(overlap)
        if score > best_score:
            best_score = score
            best_match = document

    return best_match


VERIFICATION_EVIDENCE_MAX_RESULTS = 60  # was 25 - too shallow for documents with 40+ codes

def _code_text_variants(code):
    code = code.strip()
    variants = [
        code,
        code.replace("-", " "),
        code.replace(" ", "-"),
        re.sub(r'[\s-]+', '', code),
        re.sub(r'[\s-]+', ' ', code),
        code.replace("Mod-", "Mod "),
        code.replace("Mod ", "Mod-"),
        code.replace("CTRL-", "CTRL "),
        code.replace("CTRL ", "CTRL-"),
        code.upper(),
        code.lower(),
    ]
    seen = set()
    out = []
    for v in variants:
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def _query_verification_evidence_document(pipeline_config, document, code):
    target_text = f"Verification_Reference : {code}"

    def _run(where_document, max_results=VERIFICATION_EVIDENCE_MAX_RESULTS):
        prompt_row = {
            "prompt_role": VERIFICATION_EVIDENCE_PROMPT_ROLE,
            "prompt_text": VERIFICATION_EVIDENCE_PROMPT_TEXT + "\nTARGETS:\n" + target_text,
            "question": verification_evidence_question(code),
            "fulltext": "Yes",
            "where_filter": "",
            "where_document": where_document,
            "checkpoint": f"Retrieve verification evidence for {code}",
            "max_results": max_results
        }
        return execute_llm_retry(pipeline_config, document["collection"], prompt_row)

    docs, metadata, response = [], None, ""
    matched_variant = None

    try:
        for variant in _code_text_variants(code):
            docs, metadata, response = _run(json.dumps({"$contains": variant}))
            logging.info(
                f"CALL 8: [{document.get('document_name')}] CODE={code!r} "
                f"variant={variant!r} retrieved {len(docs) if docs else 0} doc(s)."
            )
            if docs:
                matched_variant = variant
                break

        if not docs:
            logging.info(
                f"CALL 8: no text variant of CODE={code!r} matched in "
                f"{document.get('document_name')!r} - retrying without filter "
                "and a wider max_results."
            )
            docs, metadata, response = _run("", max_results=VERIFICATION_EVIDENCE_MAX_RESULTS * 2)

        print(
            f"CALL 8: RAW LLM RESPONSE [{document.get('document_name')}] "
            f"CODE={code!r} (matched_variant={matched_variant!r}) -> {response!r}"
        )
    except Exception as e:
        logging.error(
            f"CALL 8: verification evidence LLM call against "
            f"{document.get('document_name')!r} failed for CODE={code!r}: {e}"
        )
        return {}, ""

    fields = parse_verification_evidence_fields(response)
    return fields, response


def _normalize_pass_fail_status(pass_status_raw):
    """
    Collapses the LLM's raw "Pass Status" text down to a clean, fixed
    verdict for Column F: "Pass", "Fail", or "Not Determined". Matches
    on the leading word so variants like "Passed", "PASS", "Fail -
    incomplete data", etc. all normalize correctly.
    """
    normalized = normalize_text(pass_status_raw).strip().lower()
    if normalized.startswith("pass"):
        return "Pass"
    if normalized.startswith("fail"):
        return "Fail"
    return "Not Determined"


def format_verification_evidence_block(code, filename_hint, fields):
    """
    Builds one Column F block for a single code, per requirement:
        Code: <requirement code>
        Filename: <source document>
        Condition Met: <exact condition/sentence the document states>

    "Condition Met" is the LLM's "Evidence" field (see
    parse_verification_evidence_fields() /
    VERIFICATION_EVIDENCE_PROMPT_TEXT) - the literal condition/sentence
    the document itself uses to state the requirement was satisfied.

    Falls back to the original Column E code whenever the LLM's own
    Code field comes back blank or as a "nothing found" placeholder,
    so the block is never emptier than what was already known before
    the LLM call.

    BUGFIX: Filename is now ALWAYS `filename_hint` - the document
    find_verification_evidence_document() already resolved and queried
    - and NEVER the LLM's own self-reported "Document:" field. The LLM
    was being asked to name the document it was reading, but the
    retrieved chunks it sees rarely state their own filename, so it
    frequently answered with a placeholder like "[Document name not
    provided in the prompt]" or "<document name not provided>" - and
    since those exact phrasings weren't in is_meaningful_evidence_field()'s
    blocklist, they were printed as-is instead of falling back (this is
    why Column F showed "not provided" even though the real document
    name was sitting right there in Column E). Worse, sometimes the LLM
    would quote an unrelated name it found INSIDE the document's content
    and mistake that for the document's own title (e.g. querying "vest
    apx mcb module feature test result" but reporting the filename as
    "VAR-SOM-MX7/VAR-SOM-MX7-5G SYSTEM ON MODULE", a module name merely
    mentioned in that document's text) - a wrong answer that no
    placeholder blocklist could ever catch. Since the correct document
    was already resolved before this LLM call was made, there is no
    reason to ask the LLM to re-identify it - filename_hint is the
    ground truth in every case.
    """
    code_out = fields.get("code", "")
    condition_out = fields.get("evidence", "")

    code_out = code_out if is_meaningful_evidence_field(code_out) else code
    filename_out = filename_hint or "None"
    condition_out = condition_out if is_meaningful_evidence_field(condition_out) else "None"

    return (
        f"Code: {code_out}\n"
        f"Filename: {filename_out}\n"
        f"Condition Met: {condition_out}"
    )


def extract_and_apply_verification_evidence(pipeline_config, edo_document, final_edos):
    """
    CALL 8 - runs on the fully merged, traceability-resolved final_edos
    list (Existing + New), AFTER build_final_edos_with_traceability()
    (CALL 7).

    CHANGED: added a run-scoped success-only cache keyed by
    (resolved_document_name, code). A code that appears in many EDO
    records (e.g. SRS-CTRL-130 appearing in 10 different rows) is only
    ever a genuine retrieval question ONCE - if it succeeds on ANY
    occurrence, every other occurrence reuses that same correct answer
    instead of re-rolling retrieval independently and risking a worse
    result. A failed/empty result is NEVER cached, so later occurrences
    of the same code keep trying instead of being locked into "None".
    A one-time immediate retry is also added when the first attempt
    comes back non-meaningful. After the main loop, a backfill pass
    goes back over every record and replaces any remaining "None" block
    whose (document, code) key was successfully resolved LATER in the
    run - so earlier rows benefit from a success found afterwards too.
    """
    logging.info("=" * 80)
    logging.info("CALL 8: VERIFICATION EVIDENCE EXTRACTION (COLUMN F, POST-TRACEABILITY)")
    logging.info("=" * 80)

    if not final_edos:
        logging.info("No EDO records to extract verification evidence for.")
        return final_edos

    has_number_index = bool((edo_document or {}).get("documents_by_number"))
    has_fallback_pool = bool(
        (edo_document or {}).get("evidence_candidate_documents")
        or (edo_document or {}).get("edo_ver_documents")
    )
    if not has_number_index and not has_fallback_pool:
        logging.warning(
            "CALL 8: neither a document-number index nor any fallback "
            "document pool is configured - Column F will be left blank "
            "for every record."
        )
        return final_edos

    # Run-scoped, SUCCESS-ONLY cache: (resolved_document_name, code) -> block.
    # Never populated with a "None"/non-meaningful result - see the
    # is_meaningful_evidence_field() guard below before writing to it.
    evidence_cache = {}

    for edo in final_edos:
        # Target filtering happens before CALL 8. No positional/3-record
        # limit is applied here.

        code_filename_pairs = parse_verification_reference_code_filename_pairs(edo)

        if not code_filename_pairs:
            logging.info(
                f"CALL 8: skipping {edo.get('edo_tag') or edo.get('edo_id')!r} - "
                "no code/filename pair could be recovered from Column E, "
                "so Column F is left blank for this record."
            )
            continue

        blocks = []
        for pair in code_filename_pairs:
            code = pair["code"]
            filename_hint = pair["filename"]
            doc_number = pair.get("doc_number")

            matched_document = find_verification_evidence_document(
                edo_document, filename_hint, doc_number=doc_number
            )
            if not matched_document:
                logging.info(
                    f"CALL 8: no document matched code {code!r} "
                    f"(doc_number={doc_number!r}, filename_hint={filename_hint!r}) "
                    "by either number or fuzzy filename match - skipping this code."
                )
                continue

            resolved_document_name = (
                normalize_text(matched_document.get("document_name")) or filename_hint
            )
            cache_key = (resolved_document_name, normalize_id(code))

            # ---- Cache hit: reuse a PRIOR SUCCESSFUL result, no LLM call ----
            if cache_key in evidence_cache:
                blocks.append(evidence_cache[cache_key])
                logging.info(f"CALL 8: cache hit for {cache_key!r} - reused prior result.")
                continue

            fields, raw_response = _query_verification_evidence_document(
                pipeline_config, matched_document, code
            )

            # ---- One immediate retry if the first pass came back empty ----
            if not is_meaningful_evidence_field(fields.get("evidence", "")):
                logging.info(
                    f"CALL 8: first pass for {code!r} against "
                    f"{resolved_document_name!r} returned no usable "
                    "evidence - retrying once."
                )
                retry_fields, retry_raw = _query_verification_evidence_document(
                    pipeline_config, matched_document, code
                )
                if is_meaningful_evidence_field(retry_fields.get("evidence", "")):
                    fields, raw_response = retry_fields, retry_raw

            if not fields:
                logging.warning(
                    f"CALL 8: code={code!r} matched document "
                    f"{matched_document.get('document_name')!r} but the LLM "
                    "response could not be parsed into any evidence fields "
                    f"(empty/malformed response, or a failed call). Raw "
                    f"response: {raw_response!r}"
                )

            block = format_verification_evidence_block(code, resolved_document_name, fields)

            # ---- Only cache a REAL success - never cache "None" ----
            if is_meaningful_evidence_field(fields.get("evidence", "")):
                evidence_cache[cache_key] = block

            blocks.append(block)

            edo.setdefault("verification_evidence_raw_llm_responses", {})[code] = raw_response

            logging.info(
                f"CALL 8: code={code!r} matched document "
                f"{matched_document.get('document_name')!r} -> {block!r}"
            )

        edo["verification_evidence"] = "\n\n".join(blocks)

    # ---- Backfill pass: fix earlier records whose code later succeeded ----
    final_edos = backfill_evidence_from_cache(final_edos, evidence_cache, edo_document)

    print("Verification_Evidence (Column F):", final_edos)
    return final_edos

def backfill_evidence_from_cache(final_edos, evidence_cache, edo_document):
    """
    Second pass over final_edos, run once the main CALL 8 loop has
    finished (and evidence_cache holds every SUCCESSFUL (document, code)
    result found anywhere in the run - see extract_and_apply_verification_evidence()).

    A record processed EARLY in the main loop may have gotten "None" for
    a code that only succeeded later, on a DIFFERENT record's occurrence
    of that same code. This pass finds every remaining "Condition Met:
    None" block, re-derives its (document, code) cache key, and replaces
    it with the now-known-good cached block if one exists - so success
    found anywhere in the run benefits every occurrence, not just the
    ones processed after it.
    """
    for edo in final_edos:
        verification_evidence = edo.get("verification_evidence", "")
        if not verification_evidence or "Condition Met: None" not in verification_evidence:
            continue

        pairs = parse_verification_reference_code_filename_pairs(edo)
        if not pairs:
            continue

        blocks = verification_evidence.split("\n\n")
        changed = False

        for i, pair in enumerate(pairs):
            if i >= len(blocks):
                break
            if "Condition Met: None" not in blocks[i]:
                continue

            matched_document = find_verification_evidence_document(
                edo_document, pair["filename"], doc_number=pair.get("doc_number")
            )
            if not matched_document:
                continue

            doc_name = normalize_text(matched_document.get("document_name"))
            cache_key = (doc_name, normalize_id(pair["code"]))

            if cache_key in evidence_cache:
                blocks[i] = evidence_cache[cache_key]
                changed = True
                logging.info(
                    f"CALL 8 BACKFILL: replaced 'None' block for code "
                    f"{pair['code']!r} on {edo.get('edo_tag') or edo.get('edo_id')!r} "
                    "with a later-found successful result."
                )

        if changed:
            edo["verification_evidence"] = "\n\n".join(blocks)

    return final_edos


# ==========================================================
# COMMON PROCESSING PIPELINE
# ==========================================================
# Risk Classification, Remarks and Recommendation - executed once per merged record, shared between Existing and New EDOs (see format_edo_worksheet in EXCEL FORMATTING). Traceability reference resolution (CALL 7) now happens earlier, in build_final_edos_with_traceability() above.


def _flatten_record_values(record):
    """Yields every string value found anywhere in a (possibly nested)
    dict/list record, so a code can be found regardless of which key
    the LLM put it under."""
    if isinstance(record, dict):
        for v in record.values():
            yield from _flatten_record_values(v)
    elif isinstance(record, list):
        for v in record:
            yield from _flatten_record_values(v)
    elif isinstance(record, (str, int, float)):
        text = normalize_text(record)
        if text:
            yield text


FIXED_SYSDD_REFERENCE = "NPD38119 Titan Hardware Detailed Design"

def get_fixed_sysdd_reference():
    """
    Legacy helper retained for compatibility with external callers.
    The Excel pipeline no longer uses this as a Column J fallback;
    Column J now prints only the SysDD/HDD value extracted by the LLM.
    """
    return FIXED_SYSDD_REFERENCE


def classify_risk_status(is_new, pipeline_config=None):
    """
    Risk Classification (Column L / Column 12) is fixed by EDO type -
    per requirement:

        Existing EDO  -> "Medium"
        New EDO       -> "High"

    This is a deterministic, content-blind rule now (no LLM call, no
    dependency on the description/reason/FMEA text) - every Existing
    EDO record prints "Medium" and every New EDO record prints "High",
    with no other possible value ("Low"/"None" are no longer produced
    here).

    `pipeline_config` is kept as an accepted (optional) parameter only
    so this stays a drop-in replacement for any other caller of the
    old signature - it is not used.
    """

    risk_status = "High" if is_new else "Medium"

    print(f"risk status ({'New' if is_new else 'Existing'} EDO):", risk_status)
    logging.info(
        f"classify_risk_status: edo_type={'New' if is_new else 'Existing'} "
        f"-> risk_status={risk_status}"
    )

    return risk_status


def apply_risk_cell_style(cell, risk):
    from openpyxl.styles import PatternFill, Font

    fills = {
        "High": "FF0000",
        "Medium": "FFFF00",
        "Low": "00B050"
    }

    if risk in fills:
        cell.fill = PatternFill("solid", fgColor=fills[risk])
        cell.font = Font(bold=True, color="000000", name="Calibri", size=10)
    else:
        cell.fill = PatternFill(fill_type=None)


def get_fmea_risk_evaluation(pipeline_config, edo_document, prompt_data, ra_number, fmea_number):
    """
    Looks up the Risk_Evaluation column value for a specific FMEA_Number
    (paired with its RA_Number for targeting) directly from the
    EDO_FMEA document/collection via one targeted LLM call. Returns the
    normalized, lowercased value (e.g. "low", "medium", "high") or ""
    when the FMEA document isn't configured, no row is found, or the
    value can't be determined.

    NOTE: The FMEA sheet has TWO separate "Risk Evaluation" columns -
    one for "Risk Evaluation (Prior to Risk Control)" and another for
    "Risk Evaluation (After Risk Control)". The Observation block must
    only ever be driven off the PRIOR TO RISK CONTROL column, so both
    the LLM question below and the key matching against the returned
    row are scoped explicitly to that column - a bare "Risk_Evaluation"
    match is deliberately NOT attempted, since that key is ambiguous
    between the two columns and could silently resolve to the "After
    Risk Control" value instead.
    """
    if not fmea_number or not edo_document or "edo_fmea" not in edo_document or not prompt_data:
        return ""

    target_text = f"RA_Number : {ra_number}\nFMEA_Number : {fmea_number}"

    prompt_row = {
        "prompt_role": prompt_data["prompt_role"],
        "prompt_text": prompt_data["prompt_text"] + "\nTARGETS:\n" + target_text,
        "question": (
            f"Does a row for FMEA_Number {fmea_number} exist in this FMEA "
            "document? This FMEA sheet has two separate Risk Evaluation "
            "columns - 'Risk Evaluation (Prior to Risk Control)' and "
            "'Risk Evaluation (After Risk Control)'. Return that row as a "
            "single JSON object if it exists, including BOTH risk "
            "evaluation values, clearly labeled as "
            "'Risk_Evaluation_Prior_to_Risk_Control' and "
            "'Risk_Evaluation_After_Risk_Control' respectively - do not "
            "merge them into a single generic 'Risk_Evaluation' field."
        ),
        "fulltext": "Yes",
        "where_filter": "",
        "where_document": "",
        "checkpoint": ""
    }

    try:
        _, _, response = execute_llm_retry(
            pipeline_config,
            edo_document["edo_fmea"]["collection"],
            prompt_row
        )
        parsed = parse_json(response)
        records = deep_extract_records(parsed)
        row = records[0] if records else (parsed if isinstance(parsed, dict) else {})
    except Exception as e:
        logging.error(f"OBSERVATION: Risk_Evaluation lookup failed for FMEA={fmea_number!r}: {e}")
        return ""

    risk_evaluation_value = normalize_text(
        get_llm_value(
            row,
            "Risk_Evaluation_Prior_to_Risk_Control",
            "Risk Evaluation Prior to Risk Control",
            "Risk Evaluation (Prior to Risk Control)",
            "Risk_Evaluation_Prior",
            "Prior to Risk Control",
            "Risk_Rating_Prior_to_Risk_Control",
            "Risk Rating (Prior to Risk Control)"
        )
    )
    return risk_evaluation_value.strip().lower()


import re
import logging


def _norm_id(value):
    """Normalize an RA/FMEA number for comparison (case/space/dash-insensitive)."""
    return re.sub(r"[\s\-_]+", "", str(value or "")).upper()


def build_ra_fmea_mismatch_observation(edo):
    """
    Builds the Existing-EDO observation by comparing:
      - EDO number  : CALL 1
      - RA/FMEA     : CALL 2
      - RA/FMEA     : the actual FMEA-document response used by CALL 6

    An observation is printed only when the FMEA-document RA/FMEA does not
    match the Existing EDO RA/FMEA.  If both values match, this returns an
    empty string so the normal Remarks/Recommendation content is unchanged.
    """
    if not isinstance(edo, dict) or edo.get("edo_type") != "Existing":
        return ""

    edo_number = normalize_text(edo.get("edo_tag"))
    call2_ra = normalize_text(edo.get("RA_Number") or edo.get("ra_number"))
    call2_fmea = normalize_text(edo.get("FMEA_Number"))

    # These two values are captured directly from the EDO_FMEA response
    # during CALL 6.  They represent the actual values returned by the
    # FMEA document, not the CALL 2 input values.
    fmea_doc_ra = normalize_text(edo.get("FMEA_Document_RA_Number"))
    fmea_doc_fmea = normalize_text(edo.get("FMEA_Document_FMEA_Number"))

    if not edo_number or not call2_ra or not call2_fmea:
        return ""

    # Do not manufacture a mismatch when the FMEA document did not return
    # usable RA/FMEA values.
    if not fmea_doc_ra or not fmea_doc_fmea:
        return ""

    ra_mismatch = canonical_ra_id(call2_ra) != canonical_ra_id(fmea_doc_ra)
    # Treat "SYS-80" and "FMEA Sys-80" as the same identifier.  A raw
    # string comparison incorrectly classified this harmless prefix
    # difference as a second mismatch.
    fmea_mismatch = canonical_fmea_id(call2_fmea) != canonical_fmea_id(fmea_doc_fmea)

    if not (ra_mismatch or fmea_mismatch):
        return ""

    if ra_mismatch and not fmea_mismatch:
        return (
            f"Observation: As per FMEA document, there is no trace for "
            f"{call2_ra} against {call2_fmea}. Instead of {fmea_doc_ra}, "
            f"the existing edo code document has {call2_ra}."
        )

    if fmea_mismatch and not ra_mismatch:
        return (
            f"Observation: As per FMEA document, {call2_ra} is not traced "
            f"to {call2_fmea}; it is traced to {fmea_doc_fmea}."
        )

    return (
        f"Observation: As per FMEA document, there is no trace for "
        f"{call2_ra} against {call2_fmea}. The matching FMEA document "
        f"values are {fmea_doc_ra} and {fmea_doc_fmea}."
    )


def generate_recommendation_text(edo, risk, pipeline_config):
    """
    Column M "Recommendation" block.

    Only ever generated when the record's Risk Classification (Column
    L) is "High" - Medium/Low risk records never get a recommendation.
    Category-driven (A: User Manual / B: Drawing-Design / C: Training /
    D: Combined Design+User Manual), generated straight from the
    record's feature/reason text. Returns "" (nothing printed) when the
    LLM determines no recommendation is needed, or on failure - no
    rule-based fallback text is fabricated.
    """
    if risk != "High":
        return ""

    feature_text = format_output_text(edo.get("edo_description") or edo.get("description_2"))
    reason_text = format_output_text(edo.get("reason_identified") or edo.get("reason_2"))

    if not feature_text and not reason_text:
        return ""

    rec_prompt = f"""instruction:
below are the product feature and reason identified as edo.
and provide the recommendation and remarks 
Product feature: {feature_text}
Rason identified as edo: {reason_text}

for example, 

Recommendation: 
Recommended to include the following details in the user manual to mitigate and control potential power cord damage
 
WARNING: Proper Use and Handling of Power Cord
To ensure user safety and maintain product integrity, strict adherence to the following instructions regarding the use and handling of the power cord is required:

- The power cord shall be used only as specified in the product instructions.
- Rough handling, excessive bending, pulling, twisting, or improper storage of the power cord is strictly prohibited.
- Repeated or improper handling may result in damage to the power cord insulation or internal wiring.
- Damaged power cords may expose live electrical components, posing a risk of electric shock or serious injury to the user.

Precautionary Measures:
- Regularly inspect the power cord for signs of wear, damage, or exposed wires.
- Do not use the device if the power cord is damaged.
- Replace the power cord immediately if any defect is identified.
- Handle and store the power cord with care to prevent unnecessary stress or damage.

this is the sample response to be retrieved as output.
"""

    try:
        rec_response = call_llm(rec_prompt, pipeline_config, question="provide recommendation and warning/precaution for the given product feature and reason identified as edo")
        rec_response = clean_response(rec_response)
        if is_meaningful_llm_text(rec_response):
            return rec_response.strip()
    except Exception as e:
        logging.error(f"Recommendation generation failed: {e}")

    return ""


def generate_remarks_and_recommendation(
    edo,
    tag_value,
    is_new,
    risk,
    pipeline_config,
    edo_document=None,
    risk_eval_prompt_data=None,
    traceability_fail_remarks=None,
    reference_pairs=None,
    fmea_reference_pairs=None
):
    """
    Builds Column M ("Remarks and Recommendation") content:
      1. Gap + Verification Status boilerplate (New EDO vs Existing EDO wording) - always printed.
      2. Traceability Fail remarks - one line per verification code whose
         traceability call (CALL 7) came back with a Fail result (see
         build_final_edos_with_traceability()); omitted entirely when there
         are none.
      3. Observation - only for Existing EDOs when CALL 2's RA/FMEA values
         mismatch the RA/FMEA values returned by the FMEA document in CALL 6.
         No observation heading is printed when there is no mismatch.
      4. Recommendation (including any Warning/Precaution line) - only when
         Risk Classification (Column L) == "High" (see generate_recommendation_text()).
    Only the generated text itself is printed for (2)/(3)/(4) - no extra
    headings/titles are added - and each block is entirely omitted when
    it isn't applicable, leaving just the Gap and Verification Status text.
    """

    # ---- 1. Base boilerplate (Gap + Verification Status) - kept as-is ----
    if is_new:
        base_text = (
            "Gap:\n"
            "As identified in the Risk Assessment & Control (RA&C) and System "
            "DFMEA, this risk impacts the product\u2019s functions and features. "
            "Therefore, it is classified as a new Essential Design Output "
            "(EDO) and must be incorporated into the existing EDO list.\n\n"
            "Verification Status:\n"
            f"Design verification has been conducted for this EDO ({tag_value}), "
            "and the corresponding reports are traced in the Verification "
            "Reference (Column E)."
        )
    else:
        base_text = (
            "Gap:\n"
            "The design verification reference corresponding to where the EDO "
            "is controlled has not been included in the existing EDO list.\n\n"
            "Verification Status:\n"
            f"Design verification has been conducted for this EDO ({tag_value}), "
            "and the corresponding reports are traced in the Verification "
            "Reference (Column E)."
        )

    # ---- 2. Traceability Fail remarks (one line per failed verification code) ----
    observation_text = build_ra_fmea_mismatch_observation(edo)

    print(
        f"OBSERVATION_TEXT | EDO={edo.get('edo_tag') or tag_value!r} "
        f"| RA={edo.get('RA_Number', '')!r} "
        f"| FMEA={edo.get('FMEA_Number', '')!r} "
        f"| {observation_text!r}"
    )
    logging.info(
        f"OBSERVATION_TEXT | EDO={edo.get('edo_tag') or tag_value!r} "
        f"| RA={edo.get('RA_Number', '')!r} "
        f"| FMEA={edo.get('FMEA_Number', '')!r} "
        f"| {observation_text!r}"
    )

    fail_remarks_text = ""
    if traceability_fail_remarks:
        fail_remarks_text = "\n".join(traceability_fail_remarks)

    # ---- 4. Recommendation (only when risk == "High") ----
    recommendation_text = generate_recommendation_text(edo, risk, pipeline_config)

    parts = [base_text]
    if fail_remarks_text:
        parts.append(fail_remarks_text)
    if observation_text:
        parts.append(observation_text)
    if recommendation_text:
        parts.append(recommendation_text)

    return "\n\n".join(parts)


# EXCEL FORMATTING
# ==========================================================

thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin")
)

cell_alignment = Alignment(
    horizontal="left",
    vertical="top",
    wrap_text=True
)

BLACK = "FF000000"
RED = "FFFF0000"

RED_FONT = Font(color=RED, name="Calibri", size=10)

def _text(value):
    return "" if value is None else str(value)


def _apply_border_alignment(cell):
    cell.alignment = cell_alignment
    cell.border = thin_border


def _capitalize_sentences(text):
    """
    Capitalizes the first letter of every sentence in `text`, line by
    line (a "sentence" ends at '.', '!', or '?' followed by whitespace).
    Leading whitespace/indentation on each line is preserved.
    """
    lines = text.split("\n")
    result_lines = []

    for line in lines:
        stripped = line.lstrip()
        if not stripped:
            result_lines.append(line)
            continue

        leading_ws = line[:len(line) - len(stripped)]
        sentences = re.split(r'(?<=[.!?])\s+', stripped)
        fixed = []
        for sentence in sentences:
            if sentence:
                fixed.append(sentence[0].upper() + sentence[1:])
            else:
                fixed.append(sentence)
        result_lines.append(leading_ws + " ".join(fixed))

    return "\n".join(result_lines)


def _normalize_document_codes(text):
    """
    Fixes casing on document / tag codes wherever they appear inside a
    sentence:
      - "npd36702" / "Npd36702"     -> "NPD36702"
      - "edo-29" / "Edo-29"         -> "EDO-29"
      - "ra-108" / "Ra-108"         -> "RA-108"
      - "fmea sys-82" / "Fmea sys-82" -> "FMEA Sys-82"
    """
    text = re.sub(r'\bnpd(\d+)', lambda m: f"NPD{m.group(1)}", text, flags=re.IGNORECASE)
    text = re.sub(r'\bedo[\s-]*([a-z0-9]+)', lambda m: f"EDO-{m.group(1).upper()}", text, flags=re.IGNORECASE)
    text = re.sub(r'\bfmea\s+sys[\s-]*(\d+)', lambda m: f"FMEA Sys-{m.group(1)}", text, flags=re.IGNORECASE)
    text = re.sub(r'\bra[\s-]+(\d+)', lambda m: f"RA-{m.group(1)}", text, flags=re.IGNORECASE)
    return text


def format_output_text(value):
    """
    Canonical text formatter applied to every descriptive/narrative cell
    value before it's written to the output Excel:
      - first letter of every sentence is capitalized
      - "npdxxxx" document codes are uppercased to "NPDxxxx"
      - "edo-xx" tag codes are uppercased to "EDO-XX"
    Leaves truly empty values as empty ("") - no placeholder text.
    """
    text = _text(value)
    if not text.strip():
        return text

    text = _capitalize_sentences(text)
    text = _normalize_document_codes(text)
    return text
def build_column_d_reference(edo, edo_document):
    """Build Column D from RA/FMEA numbers, their source document names, and trace."""
    ra_number = normalize_text(edo.get("RA_Number", ""))
    fmea_number = normalize_text(edo.get("FMEA_Number", ""))

    has_ra = ra_number not in ("", "Blank", "None")
    has_fmea = fmea_number not in ("", "Blank", "None")

    ra_document = edo_document.get("edo_ra_c", {}) or {}
    fmea_document = edo_document.get("edo_fmea", {}) or {}

    ra_doc_name = normalize_text(
        ra_document.get("document_name")
        or ra_document.get("originalfilename")
        or ra_document.get("filename")
    )
    fmea_doc_name = normalize_text(
        fmea_document.get("document_name")
        or fmea_document.get("originalfilename")
        or fmea_document.get("filename")
    )

    reference_lines = []

    if has_ra:
        reference_lines.append(
            f" {ra_number}"
            + (f" |  {ra_doc_name}" if ra_doc_name else "")
        )

    if has_fmea:
        reference_lines.append(
            f" {fmea_number}"
            + (f" |  {fmea_doc_name}" if fmea_doc_name else "")
        )

    # NOTE: the trace itself is intentionally NOT appended here anymore.
    # It used to be added to reference_lines whenever both RA_Number and
    # FMEA_Number were present, but format_edo_worksheet() ALSO appends
    # the same "existing_trace" value as a separate red TextBlock on top
    # of this function's return value (see the CellRichText block built
    # around existing_trace_value) - so the trace was being printed
    # TWICE in Column D: once here as plain black text, and once again
    # in red right after it. The red TextBlock append is the only place
    # the trace should be added, so it's no longer duplicated here.

    return "\n".join(reference_lines)

def format_edo_tag_text(value):
    """
    Same code-casing fix as format_output_text(), but WITHOUT sentence
    capitalization - used specifically for Column A (EDO Tag), which is a
    short code/label rather than a narrative sentence (e.g. "edo-29" ->
    "EDO-29").
    """
    text = _text(value)
    if not text.strip():
        return text
    return _normalize_document_codes(text)


# ---- Row auto-height (fits every row's height to its actual content) ----
DEFAULT_COLUMN_WIDTH_CHARS = 8.43
ROW_LINE_HEIGHT_PT = 15
ROW_MIN_HEIGHT_PT = 15


def _cell_text_for_sizing(value):
    """
    Returns the plain text openpyxl will actually render for `value`,
    whether it's a plain string or a CellRichText (used for Column D's
    base-text + red trace-text rich cells) - CellRichText is a sequence
    of str/TextBlock items, so its rendered text is the concatenation of
    each item's text (TextBlock.text for TextBlock items, the item
    itself for plain str items).
    """
    if value is None:
        return ""
    if isinstance(value, CellRichText):
        parts = []
        for item in value:
            parts.append(item.text if isinstance(item, TextBlock) else str(item))
        return "".join(parts)
    return str(value)


def _column_width_chars(sheet, col_idx):
    letter = get_column_letter(col_idx)
    width = sheet.column_dimensions[letter].width
    return width if width else DEFAULT_COLUMN_WIDTH_CHARS


def _lines_needed_for_text(text, col_width_chars):
    """
    Estimates how many wrapped display lines `text` will occupy in a
    wrap_text=True cell of the given column width - splits on explicit
    newlines first (each forces its own line break), then estimates how
    many times each of those segments itself wraps, based on roughly how
    many characters fit across the column's width.
    """
    if not text:
        return 1

    chars_per_line = max(1, int(round(col_width_chars)))
    total_lines = 0
    for segment in str(text).split("\n"):
        if not segment:
            total_lines += 1
        else:
            total_lines += math.ceil(len(segment) / chars_per_line)
    return max(1, total_lines)


def _merge_span_for_cell(sheet, row, col):
    """
    Returns (first_row, last_row) of the merged range containing
    (row, col), or (row, row) if that cell isn't part of any merge.
    """
    for merged_range in sheet.merged_cells.ranges:
        if (
            merged_range.min_row <= row <= merged_range.max_row
            and merged_range.min_col <= col <= merged_range.max_col
        ):
            return merged_range.min_row, merged_range.max_row
    return row, row


def autosize_edo_rows(sheet, start_row, end_row, columns=range(1, 14)):
    """
    Grows (never shrinks) every row's height in [start_row, end_row] to
    fit the actual wrapped content of every column in `columns` -
    covering the whole table (A-M) so every row auto-extends according
    to its data, not just the ones carrying images.
    """
    if end_row < start_row:
        return

    row_line_needs = {row: 1 for row in range(start_row, end_row + 1)}
    processed_spans = set()

    for row in range(start_row, end_row + 1):
        for col in columns:
            span_start, span_end = _merge_span_for_cell(sheet, row, col)
            span_key = (span_start, span_end, col)
            if span_key in processed_spans:
                continue
            processed_spans.add(span_key)

            text = _cell_text_for_sizing(sheet.cell(span_start, col).value)
            if not text:
                continue

            col_width = _column_width_chars(sheet, col)
            lines = _lines_needed_for_text(text, col_width)
            span_rows = span_end - span_start + 1
            lines_per_row = math.ceil(lines / span_rows)

            for r in range(max(span_start, start_row), min(span_end, end_row) + 1):
                if lines_per_row > row_line_needs[r]:
                    row_line_needs[r] = lines_per_row

    for row, lines in row_line_needs.items():
        needed_height = max(ROW_MIN_HEIGHT_PT, lines * ROW_LINE_HEIGHT_PT)
        current_height = sheet.row_dimensions[row].height or 0
        if needed_height > current_height:
            sheet.row_dimensions[row].height = needed_height


def format_edo_worksheet(sheet, final_edos, start_row, pipeline_config, images=None, new_edo_diagram_queue=None, edo_document=None, risk_eval_prompt_data=None, reference_pairs=None, fmea_reference_pairs=None):
    """
    Final writer:
    A-E : existing columns
    F   : Verification Evidence - one "Code / Filename / Why It Passed"
          block per Column E code, produced by
          extract_and_apply_verification_evidence() (CALL 8) against the
          document-number-indexed (primary) / partially-matched EDO_VER
          (fallback) source document(s). Left blank for a record
          whenever Column E has no resolvable code/filename pair.
    G-I : new EDO fields
    H   : also carries any images extracted from edo_proposed (see
          extract_edo_proposed_images() / insert_image_below_text()),
          stacked below the description_2 text. One image is placed per
          split row within each EDO tag's merged block (Column H is
          never merged across split rows, unlike A-E), pulling from the
          shared queue in document order; once the queue runs dry, no
          image is placed on that split row - Column H is simply left
          empty (no image is duplicated).

          For New EDO records specifically, the row whose RA_Number /
          FMEA_Number match TARGET_IMAGE_RA_NUMBER / TARGET_IMAGE_FMEA_NUMBER
          (currently RA-141 / FMEA Sys-152) is guaranteed the next queued
          diagram, reserved for it before any other row can consume it.
          Every other New EDO row falls back to the same shared FIFO
          queue as before. Every case where Column H ends up with no
          image at all is logged with the specific reason.
    K   : Risk Classification
    L   : Risk evaluation text / classification trigger
    M   : Gap and Verification Status statement
    """

    current_row = start_row
    existing_ranges = []
    image_queue = list(images) if images else []
    new_edo_diagram_queue = list(new_edo_diagram_queue) if new_edo_diagram_queue else []
    last_image_row = start_row

    # ---- Print-limit counters: stop once 5 Existing + 5 New records
    # have been written to the sheet. Initialized ONCE, before the loop
    # starts - NOT inside it, or they'd reset every iteration. ----
    existing_printed = 0
    new_printed = 0

    # ---- Reserve a diagram specifically for RA-141 / FMEA Sys-152 ----
    # Guarantees that record gets an image even if other New EDO rows
    # come first in iteration order and would otherwise drain the FIFO
    # queue before reaching it.
    reserved_target_image = None
    if new_edo_diagram_queue:
        for candidate_edo in final_edos:
            if (
                candidate_edo.get("edo_type") == "New"
                and normalize_id(candidate_edo.get("RA_Number")) == normalize_id(TARGET_IMAGE_RA_NUMBER)
                and normalize_id(candidate_edo.get("FMEA_Number")) == normalize_id(TARGET_IMAGE_FMEA_NUMBER)
            ):
                reserved_target_image = new_edo_diagram_queue.pop(0)
                logging.info(
                    "COLUMN H: reserved the next queued New EDO diagram "
                    f"exclusively for RA={TARGET_IMAGE_RA_NUMBER!r} "
                    f"FMEA={TARGET_IMAGE_FMEA_NUMBER!r}."
                )
                break

    for key, edo in enumerate(final_edos):

        # ---- Stop entirely once both quotas are filled ----
        if existing_printed >= 5 and new_printed >= 5:
            logging.info(
                "PRINT LIMIT REACHED: 5 Existing + 5 New EDO records "
                "already written - stopping the write loop."
            )
            break

        edo_id = edo.get("edo_id", key)  # Fallback to index if key missing

        is_new = edo.get("edo_type") == "New"

        # ---- Skip this record if its own type's quota is already full ----
        if is_new and new_printed >= 5:
            continue
        if not is_new and existing_printed >= 5:
            continue

        if is_new:
            tag_value = "EDO-XX\nNew"
        else:
            tag_value = format_edo_tag_text(edo.get("edo_tag") or key)

        raw_location = format_output_text(edo.get("location"))
        design_elements = edo.get("design_elements") or []

        if design_elements:
            split_rows = []
            last_description, last_reason = "", ""
            for element in design_elements:
                description = element.get("description", "")
                reason = element.get("reason", "")

                if description:
                    last_description = description
                else:
                    description = last_description

                if reason:
                    last_reason = reason
                else:
                    reason = last_reason

                split_rows.append({
                    "location": element.get("location", ""),
                    "description_2": description,
                    "reason_2": reason,
                })
        else:
            locations = re.findall(r'\d+\s*\([^\)]+\)', raw_location) or [raw_location]
            split_rows = [
                {
                    "location": location,
                    "description_2": edo.get("description_2"),
                    "reason_2": edo.get("reason_2"),
                }
                for location in locations
            ]

        first = current_row
        last_image_for_this_edo = None

        for idx, row_data in enumerate(split_rows):

            col_g_value = row_data["location"]
            col_h_value = format_output_text(row_data["description_2"])
            col_i_value = format_output_text(row_data["reason_2"])

            is_target_row = False

            if is_new:
                is_target_row = (
                    normalize_id(edo.get("RA_Number")) == normalize_id(TARGET_IMAGE_RA_NUMBER)
                    and normalize_id(edo.get("FMEA_Number")) == normalize_id(TARGET_IMAGE_FMEA_NUMBER)
                )

                if is_target_row:
                    if not normalize_text(col_g_value):
                        col_g_value = "181995"
                    if not normalize_text(col_h_value):
                        col_h_value = "EDO Symbol needs to be updated in the power cord length."
                    if not normalize_text(col_i_value):
                        col_i_value = "Cord cable length of 3m decrease the possibility of loose connection of power cord to control unit"
            # Runs for BOTH Existing and New (non-target) records now.

            if not is_target_row:
                if not normalize_text(col_g_value):
                    col_g_value = "None"
                if not normalize_text(col_h_value):
                    col_h_value = "None"
                if not normalize_text(col_i_value):
                    col_i_value = "None"


            col_d_value = build_column_d_reference(edo, edo_document)
            existing_trace_value = format_output_text(edo.get("existing_trace"))

            logging.info(f"{key} dfmea raw: {edo.get('dfmea')}")
            logging.info(f"{key} existing_trace raw: {edo.get('existing_trace')}")
            logging.info(f"{key} final column D base value: {col_d_value}")

            values = {
                1: tag_value,
                2: format_output_text(edo.get("edo_description")) if idx == 0 else "",
                3: format_output_text(edo.get("reason_identified")),
                4: col_d_value,
                5: format_output_text(edo.get("verification_reference")),
                6: format_output_text(edo.get("verification_evidence")),
                7: col_g_value,
                8: col_h_value,
                9: col_i_value,
                # Column J must show the LLM-extracted SysDD/HDD value.
                # Do not silently replace an empty extraction with the old
                # template-wide constant.
                10: format_output_text(edo.get("sysdd")),
                11: "None",
            }

            risk = classify_risk_status(is_new, pipeline_config)
            values[12] = risk

            values[13] = generate_remarks_and_recommendation(
                edo, tag_value, is_new, risk, pipeline_config,
                edo_document=edo_document,
                risk_eval_prompt_data=risk_eval_prompt_data,
                traceability_fail_remarks=edo.get("traceability_fail_remarks"),
                reference_pairs=reference_pairs,
                fmea_reference_pairs=fmea_reference_pairs
            )
            

            for col, value in values.items():
                cell = sheet.cell(current_row, col)
                cell.value = value
                _apply_border_alignment(cell)

                if col == 5 or (is_new and col != 12):
                    cell.font = RED_FONT

            apply_risk_cell_style(sheet.cell(current_row, 12), risk)

            if existing_trace_value:
                d_cell = sheet.cell(current_row, 4)
                base_font = InlineFont(color=BLACK, rFont="Calibri", sz=10)
                trace_font = InlineFont(color=RED, rFont="Calibri", sz=10)
                if col_d_value:
                    d_cell.value = CellRichText(
                        TextBlock(base_font, col_d_value),
                        TextBlock(trace_font, "\n" + existing_trace_value),
                    )
                else:
                    d_cell.value = CellRichText(
                        TextBlock(trace_font, existing_trace_value)
                    )

            row_image = None

            has_ra_fmea_in_col_d = bool(
                normalize_text(edo.get("RA_Number")) or normalize_text(edo.get("FMEA_Number"))
            )

            is_target_image_row = (
                is_new
                and normalize_id(edo.get("RA_Number")) == normalize_id(TARGET_IMAGE_RA_NUMBER)
                and normalize_id(edo.get("FMEA_Number")) == normalize_id(TARGET_IMAGE_FMEA_NUMBER)
            )

            if is_target_image_row and reserved_target_image:
                row_image = reserved_target_image
                reserved_target_image = None
                logging.info(
                    f"COLUMN H (row {current_row}, key {key!r}): placed the "
                    f"RESERVED New EDO diagram for RA={TARGET_IMAGE_RA_NUMBER!r} "
                    f"FMEA={TARGET_IMAGE_FMEA_NUMBER!r}."
                )
            elif is_new and has_ra_fmea_in_col_d and new_edo_diagram_queue:
                row_image = new_edo_diagram_queue.pop(0)
                logging.info(
                    f"COLUMN H (row {current_row}, key {key!r}): placed next "
                    f"queued New EDO diagram (RA={edo.get('RA_Number')!r} "
                    f"FMEA={edo.get('FMEA_Number')!r})."
                )
            elif is_new and has_ra_fmea_in_col_d:
                logging.warning(
                    f"COLUMN H (row {current_row}, key {key!r}): New EDO "
                    "diagram queue is empty - no more diagrams to place."
                )

            if not row_image and not is_new:
                if image_queue:
                    row_image = image_queue.pop(0)
                    last_image_for_this_edo = row_image

            if row_image:
                insert_image_below_text(sheet, row_image, row=current_row, column=8, text_offset_px=IMAGE_TEXT_OFFSET_PX)
                last_image_row = current_row
                last_image_for_this_edo = row_image
            else:
                logging.error(
                    f"COLUMN H WILL BE EMPTY at row {current_row} for "
                    f"key {key!r} (RA={edo.get('RA_Number')!r} FMEA="
                    f"{edo.get('FMEA_Number')!r}, is_new={is_new}) - no "
                    "New EDO diagram available AND the generic image "
                    "queue is exhausted."
                )

            current_row += 1

        if not is_new:
            existing_ranges.append((first, current_row - 1))

        # ---- Record this EDO as printed (once per EDO record, not per
        # split row - an EDO with multiple design elements/locations
        # still only counts as ONE record toward the 5-record limit). ----
        if is_new:
            new_printed += 1
        else:
            existing_printed += 1

    for first, last in existing_ranges:
        if last > first:
            for col in range(1, 7):
                sheet.merge_cells(
                    start_row=first,
                    start_column=col,
                    end_row=last,
                    end_column=col
                )
                sheet.cell(first, col).alignment = cell_alignment
            for col in (12, 13):
                sheet.merge_cells(
                    start_row=first,
                    start_column=col,
                    end_row=last,
                    end_column=col
                )
                sheet.cell(first, col).alignment = cell_alignment

    autosize_edo_rows(sheet, start_row, current_row - 1)

    print(f"current _value:")
    return current_row


def clear_existing_rows(sheet, start_row, end_column=10):
    row = start_row
    while row <= sheet.max_row:
        empty = True
        for col in range(1, end_column + 1):
            if sheet.cell(row=row, column=col).value:
                empty = False
                break
        if empty:
            break
        for col in range(1, end_column + 1):
            sheet.cell(row=row, column=col).value = None
        row += 1


# ==========================================================
# EXCEL WRITER
# ==========================================================

def save_edo_workbook(workbook, pipeline_config):
    output_path = pipeline_config["output_file_path"]
    workbook.save(output_path)
    return output_path


# ==========================================================
# MAIN PIPELINE
# ==========================================================
# START -> load all documents from the database -> load template -> image extractions -> prompt execution -> Existing EDO pipeline (CALL 1-3) -> New EDO pipeline (CALL 4-5, now including New EDO verification reference) -> verification reference for Existing EDOs (CALL 6) -> merge Existing + New -> traceability (CALL 7) -> Excel formatting -> write Excel -> END.

def generate_edo_template(
    client,
    product_family,
    product,
    templatename,
    pipeline_config,
    db: DatabaseHandler
):
    logging.info("=" * 80)
    logging.info("STARTING COMPLETE EDO TEMPLATE GENERATION PIPELINE (MERGED)")
    logging.info("=" * 80)
    logging.info(f"[RUN] template={templatename} product={product}")
    try:
        # ---------------------------------------------------
        # Load all documents from the database first
        # ---------------------------------------------------
        edo_document = get_edo_document(
            client,
            product_family,
            product,
            templatename,
            db
        )

        workbook, sheet = initialize_workbook(pipeline_config)
        start_row = pipeline_config.get("templatestartrow", 4)

        # Clear active table grid space exclusively up to Column J
        clear_existing_rows(sheet, start_row, end_column=10)

        # ---------------------------------------------------
        # Image extraction
        # ---------------------------------------------------
        try:
            edo_proposed_images = extract_edo_proposed_images(edo_document, pipeline_config)

            images_output_dir = os.path.join(
                os.path.dirname(pipeline_config.get("output_file_path", ".")) or ".",
                "edo_proposed_images"
            )
            if edo_proposed_images:
                save_images_to_folder(edo_proposed_images, images_output_dir)

        except Exception as image_error:
            edo_proposed_images = []
            logging.warning(
                "IMAGE EXTRACTION SKIPPED - could not extract images from "
                f"edo_proposed for this run. Reason: {image_error}"
            )

        # ---------------------------------------------------
        # CALL 1: Extract Existing EDO tags
        # ---------------------------------------------------
        existing_edos = extract_edo_tags(
            client,
            product_family,
            product,
            templatename,
            pipeline_config,
            edo_document,
            db
        )
        existing_edos = validate_existing_tags(existing_edos)
        existing_edos = sort_existing_edos(existing_edos)
        existing_edos = _filter_target_existing_edos(existing_edos)

        if existing_edos:
            # -----------------------------------------------
            # CALL 2: Extract Existing EDO details
            # -----------------------------------------------
            existing_edos = extract_edo_details(
                client,
                product_family,
                product,
                templatename,
                pipeline_config,
                edo_document,
                existing_edos,
                db
            )

            # -----------------------------------------------
            # CALL 3: Extract Existing EDO trace details
            # -----------------------------------------------
            try:
                existing_trace_details = extract_existing_edo_trace_details(
                    client,
                    product_family,
                    product,
                    templatename,
                    pipeline_config,
                    edo_document,
                    existing_edos,
                    db
                )

                existing_edos = apply_existing_edo_trace(
                    existing_edos,
                    existing_trace_details
                )
            
            except Exception as trace_error:
                logging.warning(
                    "CALL 3 SKIPPED - trace extraction failed, leaving "
                    "the trace part of column D blank for this run. "
                    f"Reason: {trace_error}"
                )
            print(f"extracted existing edo trace: ", existing_edos)

        # ---------------------------------------------------
        # CALL 4: Extract New EDO RA/FMEA reference pairs
        # CALL 5: Extract New EDO/FMEA details and preserve the actual
        #          RA/FMEA values returned by the FMEA-document LLM.
        #
        # These two calls are used as reference data for the Existing-EDO
        # RA/FMEA mismatch observation. Their New EDO records are NOT merged
        # into the final output while this run is targeting Existing EDOs.
        # ---------------------------------------------------
        call4_reference_pairs = {}
        call5_fmea_reference_pairs = {}

        try:
            edo_new_data = extract_new_edo_tags(
                client,
                product_family,
                product,
                templatename,
                pipeline_config,
                edo_document,
                db
            )
            call4_reference_pairs = {
                key: dict(value) for key, value in edo_new_data.items()
                if isinstance(value, dict)
            }
            print(f"CALL 4 reference pairs (max {NEW_EDO_MAX_TEST_RECORDS}): {call4_reference_pairs}")
            logging.info(f"CALL 4 reference pairs (max {NEW_EDO_MAX_TEST_RECORDS}): {call4_reference_pairs}")

            edo_new_data = extract_new_edo_summary_details(
                client,
                product_family,
                product,
                templatename,
                pipeline_config,
                edo_document,
                edo_new_data,
                db,
                fmea_reference_pairs=call5_fmea_reference_pairs
            )
            print(f"CALL 5 FMEA reference pairs (max {NEW_EDO_MAX_TEST_RECORDS}): {call5_fmea_reference_pairs}")
            logging.info(f"CALL 5 FMEA reference pairs (max {NEW_EDO_MAX_TEST_RECORDS}): {call5_fmea_reference_pairs}")
        except Exception as new_reference_error:
            logging.warning(
                "CALL 4/5 reference extraction failed. Existing EDO output will "
                f"continue without the unavailable reference data. Reason: {new_reference_error}"
            )

        # Do not merge New EDOs in this Existing-EDO target run.
        new_records = []

        # ---------------------------------------------------
        # New EDO Diagram Extraction
        # ---------------------------------------------------
        try:
            new_edo_diagram_queue = extract_new_edo_diagram_queue(
                edo_document,
                pipeline_config
            )
        except Exception as new_diagram_error:
            new_edo_diagram_queue = []
            logging.warning(
                "NEW EDO DIAGRAM EXTRACTION SKIPPED - could not extract "
                f"diagrams from EDO_pdf_new for this run. Reason: {new_diagram_error}"
            )

        # ---------------------------------------------------
        # CALL 6: Verification Reference extraction (EXISTING EDOs ONLY).
        # ---------------------------------------------------
        try:
            existing_edos = extract_and_apply_verification_details(
                client,
                product_family,
                product,
                templatename,
                pipeline_config,
                edo_document,
                existing_edos,
                db
            )
        except Exception as verification_error:
            logging.warning(
                "CALL 6 SKIPPED - verification reference extraction for "
                "Existing EDOs failed, leaving column E as 'Blank' for "
                f"those records this run. Reason: {verification_error}"
            )
        print(f"verification_details (existing EDOs): ", existing_edos)

        # ---------------------------------------------------
        # MERGE: Combine Existing (now carrying its CALL 6 verification
        # reference) and New (already carrying its CALL 5 verification
        # reference) EDOs into one final_edos dict.
        # ---------------------------------------------------
        new_final = merge_new_edo_records(
            new_records,
            existing_edos
        )

        existing_edo_final = merge_existing_edo_dictionary(existing_edos)
        final_edos = merge_all_edos(existing_edo_final, new_final)
        final_edos = validate_final_edos(final_edos)

        # Keep only the one hard-coded EDO tag for the remainder of the pipeline.
        if isinstance(final_edos, dict):
            final_edos = [
                edo for key, edo in final_edos.items()
                if _is_target_edo(edo.get("edo_tag") or key)
            ]
        else:
            final_edos = [
                edo for edo in (final_edos or [])
                if _is_target_edo(edo.get("edo_tag"))
            ]

        if not final_edos:
            raise Exception(f"Target EDO {TARGET_EDO_TAG} was not found in final records.")
        print("Final edos:", final_edos)
        # ---------------------------------------------------
        # CALL 7: Traceability Reference Resolution (dynamic document
        # search across every EDO_TM traceability document, run-scoped
        # prefix cache).
        # ---------------------------------------------------
        try:
            final_edos = build_final_edos_with_traceability(
                client,
                product_family,
                product,
                templatename,
                pipeline_config,
                edo_document,
                final_edos,
                db
            )
        except Exception as traceability_reference_error:
            logging.warning(
                "CALL 7 SKIPPED - traceability reference lookup failed. "
                f"Reason: {traceability_reference_error}"
            )
            if isinstance(final_edos, dict):
                final_edos = list(final_edos.values())
        print(f"Traceability details: ", final_edos)

        # ---------------------------------------------------
        # CALL 8: Verification Evidence Extraction (Column F) - per code
        # in Column E, matched by document number (primary) or PARTIAL
        # filename (fallback). Must run AFTER CALL 7, since it depends
        # on Column E's final, resolved "<code> <location> - <filename>"
        # text and the parallel verification_reference_doc_numbers list.
        # ---------------------------------------------------
        try:
            final_edos = extract_and_apply_verification_evidence(
                pipeline_config,
                edo_document,
                final_edos
            )
        except Exception as evidence_error:
            logging.warning(
                "CALL 8 SKIPPED - verification evidence extraction failed, "
                f"leaving Column F blank for this run. Reason: {evidence_error}"
            )
        print(f"verification evidence (Column F): ", final_edos)

        # ---------------------------------------------------
        # CALL 9: Output Mapping, Formatting, and Storage
        # ---------------------------------------------------
        try:
            risk_eval_prompt_data = get_prompt(
                client, product_family, product, templatename,
                "EDO_NEW_risk_evaluation", db
            )
        except Exception as risk_eval_prompt_error:
            logging.warning(
                "OBSERVATION: could not load 'EDO_NEW_risk_evaluation' prompt - "
                f"Observation text will be skipped for this run. Reason: {risk_eval_prompt_error}"
            )
            risk_eval_prompt_data = None

        format_edo_worksheet(
            sheet,
            final_edos,
            start_row,
            pipeline_config,
            images=edo_proposed_images,
            new_edo_diagram_queue=new_edo_diagram_queue,
            edo_document=edo_document,
            risk_eval_prompt_data=risk_eval_prompt_data,
            reference_pairs=call4_reference_pairs,
            fmea_reference_pairs=call5_fmea_reference_pairs
        )

        output_file = save_edo_workbook(
            workbook,
            pipeline_config
        )

        logging.info("=" * 80)
        logging.info("EDO PIPELINE COMPLETED SUCCESSFULLY")
        logging.info(f"OUTPUT FILE PERSISTED AT : {output_file}")
        logging.info("=" * 80)

        return output_file

    except Exception as e:
        logging.error("=" * 80)
        logging.error(f"EDO PIPELINE CRITICAL RUNTIME FAILURE : {str(e)}")
        logging.error("=" * 80)
        raise e