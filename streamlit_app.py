import io
import os
import re
import math
import zipfile
import tempfile

import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


# ---------- CONFIG ----------
ANCHOR_SHAPE_NAME = "SCREENSHOT_BOX"
OUTPUT_PPTX_NAME = "QuestionPPT.pptx"

WHITE_THRESHOLD = 240
CROP_PAD_PX = 20

# Target occupancy of slide area (30% to 35%)
TARGET_FRACTION = 0.25
# ---------------------------


st.set_page_config(page_title="PPT Generator", layout="centered")


# ---------- AUTH GATE ----------
def require_login_and_allowlist():
    u = getattr(st, "user", None)

    if not u or not u.is_logged_in:
        st.header("Private tool")
        st.write("Please sign in with your Google Workspace account.")
        if st.button("Log in with Google"):
            st.login()
        st.stop()

    email = (u.get("email") or "").strip().lower()

    raw = os.environ.get("ALLOWED_EMAILS", "")
    allowed = {e.strip().lower() for e in raw.replace(",", "\n").splitlines() if e.strip()}

    if allowed and email not in allowed:
        st.error(f"Access denied for: {email}")
        if st.button("Log out"):
            st.logout()
        st.stop()

    st.caption(f"Signed in as: {email}")


require_login_and_allowlist()
# -----------------------------


# --- UI CSS ---
st.markdown(
    """
    <style>
      .block-container {
        padding-top: 2.2rem !important;
        padding-bottom: 4.8rem !important;
        max-width: 820px !important;
      }

      div[data-testid="stVerticalBlock"] > div {
        gap: 0.8rem !important;
      }

      [data-testid="stAppViewContainer"] {
        scrollbar-width: none;
      }

      [data-testid="stAppViewContainer"]::-webkit-scrollbar {
        display: none;
      }

      footer {visibility: hidden;}
      #MainMenu {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 PPT Generator Tool")
st.write("Upload a ZIP (images) and an XLSX (mapping) to generate a PPTX.")

zip_file = st.file_uploader("Upload ZIP file", type=["zip"])
xlsx_file = st.file_uploader("Upload XLSX file", type=["xlsx"])

include_solutions = st.checkbox("Include Solution Images")


def remove_white(img: Image.Image, thr=WHITE_THRESHOLD) -> Image.Image:
    """Make near-white pixels transparent."""
    rgba = np.array(img.convert("RGBA"))
    rgb = rgba[:, :, :3]
    mask = (rgb >= thr).all(axis=2)
    rgba[mask, 3] = 0
    return Image.fromarray(rgba, "RGBA")


def crop_to_content_rgba(img_rgba: Image.Image, pad_px=CROP_PAD_PX) -> Image.Image:
    """Crop image to non-transparent content using alpha channel bounding box."""
    alpha = img_rgba.split()[-1]
    bbox = alpha.getbbox()

    if not bbox:
        return img_rgba

    x0, y0, x1, y1 = bbox
    x0 = max(0, x0 - pad_px)
    y0 = max(0, y0 - pad_px)
    x1 = min(img_rgba.width, x1 + pad_px)
    y1 = min(img_rgba.height, y1 + pad_px)

    return img_rgba.crop((x0, y0, x1, y1))


def extract_filename(cell):
    """Extract image filename from URLs/paths/cells."""
    if pd.isna(cell):
        return ""

    s = str(cell).strip()
    m = re.search(r"([^/\s]+?\.(png|jpg|jpeg|webp|bmp))", s, re.I)
    return m.group(1) if m else s


def iter_all_shapes(shapes):
    for shp in shapes:
        yield shp
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_all_shapes(shp.shapes)


def find_anchor_on_slide(slide, name):
    for shp in iter_all_shapes(slide.shapes):
        if getattr(shp, "name", "").strip() == name:
            return shp
    return None


def compute_size_for_area_fraction(slide_w, slide_h, left, top, aspect, target_fraction):
    """
    Compute (w,h) in EMU for picture aspect ratio (w/h),
    target area = target_fraction * slide_area,
    clamped to remain on-slide from anchor left/top.
    """
    slide_area = slide_w * slide_h
    desired_area = slide_area * target_fraction

    w = math.sqrt(desired_area * aspect)
    h = math.sqrt(desired_area / aspect)

    max_w = max(1, slide_w - left)
    max_h = max(1, slide_h - top)

    if w > max_w or h > max_h:
        scale = min(max_w / w, max_h / h)
        w *= scale
        h *= scale

    return int(w), int(h)


def safe_extract(zipf: zipfile.ZipFile, dest_dir: str):
    """Safely extract ZIP to prevent path traversal."""
    dest_abs = os.path.abspath(dest_dir)

    for member in zipf.infolist():
        if member.is_dir():
            continue

        member_path = member.filename.replace("\\", "/").lstrip("/")
        out_path = os.path.abspath(os.path.join(dest_dir, member_path))

        if not out_path.startswith(dest_abs + os.sep):
            raise ValueError(f"Unsafe zip entry blocked: {member.filename}")

        os.makedirs(os.path.dirname(out_path), exist_ok=True)

        with zipf.open(member) as src, open(out_path, "wb") as dst:
            dst.write(src.read())


def build_basename_index(root_dir: str):
    """Index files by basename lowercase."""
    idx = {}

    for root, _, files in os.walk(root_dir):
        for f in files:
            idx[f.lower()] = os.path.join(root, f)

    return idx


def find_src_path(idx, src_basename):
    """Find by exact basename; fallback fuzzy by stem ignoring spaces."""
    if not src_basename:
        return None

    src_basename = str(src_basename).strip()

    if src_basename.lower() in ["nan", "none", ""]:
        return None

    p = idx.get(src_basename.lower())

    if p:
        return p

    stem = os.path.splitext(src_basename)[0].lower().replace(" ", "")

    for name, pth in idx.items():
        if os.path.splitext(name)[0].replace(" ", "") == stem:
            return pth

    return None


def place_image_on_slide(slide, image_path, anchor_left, anchor_top, slide_w, slide_h):
    """Clean, crop, resize, and place image on slide."""
    with Image.open(image_path) as im:
        cleaned = remove_white(im, thr=WHITE_THRESHOLD)
        cleaned = crop_to_content_rgba(cleaned, pad_px=CROP_PAD_PX)

    w_px, h_px = cleaned.size

    if h_px == 0:
        return False

    aspect = w_px / h_px

    w_emu, h_emu = compute_size_for_area_fraction(
        slide_w,
        slide_h,
        anchor_left,
        anchor_top,
        aspect,
        TARGET_FRACTION,
    )

    buf = io.BytesIO()
    cleaned.save(buf, format="PNG")
    buf.seek(0)

    slide.shapes.add_picture(
        buf,
        anchor_left,
        anchor_top,
        width=w_emu,
        height=h_emu,
    )

    return True


if st.button("🚀 Generate PPT"):
    if not zip_file or not xlsx_file:
        st.error("Please upload both ZIP and XLSX files.")
        st.stop()

    try:
        with st.spinner("Processing..."):
            # Read mapping
            df_raw = pd.read_excel(io.BytesIO(xlsx_file.getvalue()))

            # Exact fixed column headers
            col_order = "Display Order*"
            col_image = "Question Image"
            col_solution = "Sol Image"

            required_cols = [col_order, col_image]

            if include_solutions:
                required_cols.append(col_solution)

            missing_cols = [c for c in required_cols if c not in df_raw.columns]

            if missing_cols:
                raise RuntimeError(
                    f"Missing required column(s): {missing_cols}. "
                    f"Found columns: {list(df_raw.columns)}"
                )

            needed_cols = [col_order, col_image]

            if include_solutions:
                needed_cols.append(col_solution)

            df = df_raw[needed_cols].copy()

            rename_map = {
                col_order: "display_order",
                col_image: "question_image",
            }

            if include_solutions:
                rename_map[col_solution] = "solution_image"

            df = df.rename(columns=rename_map)

            # Keep original behavior for questions
            df = df.dropna(subset=["display_order", "question_image"]).copy()

            df["display_order"] = df["display_order"].apply(
                lambda x: int(str(x).strip().split(".")[0])
            )

            df["question_src_name"] = df["question_image"].apply(extract_filename)

            if include_solutions:
                # Do not drop question rows if solution image is missing
                df["solution_src_name"] = df["solution_image"].apply(extract_filename)

            df = df.sort_values("display_order").reset_index(drop=True)

            # Load template
            template_path = "template.pptx"

            if not os.path.exists(template_path):
                raise FileNotFoundError("template.pptx not found in app directory.")

            prs = Presentation(template_path)

            # Use layout of slide 1
            template_layout = prs.slides[0].slide_layout

            def ensure_slide(i0):
                while len(prs.slides) <= i0:
                    prs.slides.add_slide(template_layout)
                return prs.slides[i0]

            # Anchor from slide 1
            anchor0 = find_anchor_on_slide(prs.slides[0], ANCHOR_SHAPE_NAME)

            if anchor0 is None:
                raise RuntimeError(
                    f"Anchor '{ANCHOR_SHAPE_NAME}' not found on template slide 1. "
                    f"Rename via Selection Pane."
                )

            ANCHOR_LEFT = anchor0.left
            ANCHOR_TOP = anchor0.top
            slide_w = prs.slide_width
            slide_h = prs.slide_height

            with tempfile.TemporaryDirectory(prefix="pptgen_") as tmpdir:
                src_dir = os.path.join(tmpdir, "src")
                os.makedirs(src_dir, exist_ok=True)

                # Unzip safely
                with zipfile.ZipFile(io.BytesIO(zip_file.getvalue()), "r") as zf:
                    safe_extract(zf, src_dir)

                idx = build_basename_index(src_dir)

                question_missing = 0
                solution_missing = 0
                question_placed = 0
                solution_placed = 0

                for row_position, row in df.iterrows():
                    question_src_name = row["question_src_name"]
                    question_src_path = find_src_path(idx, question_src_name)

                    if include_solutions:
                        question_slide_index = row_position * 2
                    else:
                        question_slide_index = int(row["display_order"]) - 1

                    question_slide = ensure_slide(question_slide_index)

                    if question_src_path:
                        ok = place_image_on_slide(
                            question_slide,
                            question_src_path,
                            ANCHOR_LEFT,
                            ANCHOR_TOP,
                            slide_w,
                            slide_h,
                        )

                        if ok:
                            question_placed += 1
                    else:
                        question_missing += 1

                    if include_solutions:
                        solution_src_name = row.get("solution_src_name", "")
                        solution_src_path = find_src_path(idx, solution_src_name)

                        solution_slide_index = question_slide_index + 1
                        solution_slide = ensure_slide(solution_slide_index)

                        if solution_src_path:
                            ok = place_image_on_slide(
                                solution_slide,
                                solution_src_path,
                                ANCHOR_LEFT,
                                ANCHOR_TOP,
                                slide_w,
                                slide_h,
                            )

                            if ok:
                                solution_placed += 1
                        else:
                            solution_missing += 1

            # Save PPTX to memory
            out = io.BytesIO()
            prs.save(out)
            out.seek(0)

        if include_solutions:
            st.success(
                f"✅ Questions placed: {question_placed}. "
                f"Solutions placed: {solution_placed}. "
                f"Question missing: {question_missing}. "
                f"Solution missing: {solution_missing}."
            )
        else:
            st.success(
                f"✅ Placed {question_placed} question image(s). "
                f"Missing: {question_missing}."
            )

        st.download_button(
            "⬇ Download PPTX",
            data=out.getvalue(),
            file_name=OUTPUT_PPTX_NAME,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

    except Exception as e:
        st.error(f"Error: {e}")


# ---- Footer ----
st.markdown(
    """
    <div class="neet-footer">
      Powered by NEET ARD Team | academics.innovation@pw.live
    </div>

    <style>
      .neet-footer{
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background: #f3f4f6;
        border-top: 1px solid #e5e7eb;
        padding: 10px 12px;
        text-align: center;
        font-size: 14px;
        color: #4b5563;
        z-index: 9999;
      }
    </style>
    """,
    unsafe_allow_html=True,
)