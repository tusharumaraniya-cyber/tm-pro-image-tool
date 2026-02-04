import streamlit as st
import pandas as pd
import re, io, zipfile
from rapidfuzz import process, fuzz
from PIL import Image

# ================= CONFIG =================
DEFAULT_MATCH = 75
DEFAULT_CHECK = 65
IMAGES_PER_ROW = 3
FIXED_FOLDER_NAME = "TM PRO"
RESIZE_W, RESIZE_H = 1200, 800

st.set_page_config(page_title="TM PRO Image Tool", layout="wide")

# ================= SIDEBAR =================
with st.sidebar:
    st.header("⚙️ Matching Settings")
    MATCH_MIN = st.slider("Perfect Match Score (%)", 0, 100, DEFAULT_MATCH)
    CHECK_MIN = st.slider("Review Match Score (%)", 0, 100, DEFAULT_CHECK)

# ================= HELPERS =================
def clean_text(t):
    t = str(t).lower()
    t = re.sub(r"\.(jpg|jpeg|png|webp)", "", t)
    t = re.sub(r"[^a-z0-9 ]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()

def base_name(name):
    name = clean_text(name)
    name = re.sub(r"_\d+$", "", name)
    return name

def resize_image(file):
    img = Image.open(file).convert("RGB")
    img = img.resize((RESIZE_W, RESIZE_H))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=95)
    buf.seek(0)
    return buf

# ================= UI =================
st.markdown(f"## 🟢 {FIXED_FOLDER_NAME} Matching Tool")

c1, c2 = st.columns(2)
with c1:
    uploaded_excel = st.file_uploader("📄 Upload Excel / CSV", ["xlsx", "xls", "csv"])
with c2:
    uploaded_images = st.file_uploader(
        "🖼️ Upload Images",
        ["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True
    )

if uploaded_excel and uploaded_images:

    df = pd.read_csv(uploaded_excel) if uploaded_excel.name.endswith(".csv") else pd.read_excel(uploaded_excel)

    sheet_items = (
        df[df.iloc[:, 3].isna()]
        .iloc[:, 2]
        .dropna()
        .astype(str)
        .str.strip()
        .tolist()
    )

    clean_map = {i: clean_text(i) for i in sheet_items}

    # ================= SESSION INIT =================
    if "results" not in st.session_state:
        st.session_state.results = {
            "MATCH": [],
            "CHECK": [],
            "DUPLICATE": []
        }
        st.session_state.used_base = set()

        for img in uploaded_images:
            bname = base_name(img.name)
            resized = resize_image(img)

            # ---- DUPLICATE FIRST ----
            if bname in st.session_state.used_base:
                st.session_state.results["DUPLICATE"].append({
                    "original": img.name
                })
                continue

            st.session_state.used_base.add(bname)

            # ---- MATCHING ----
            best = process.extractOne(
                clean_text(img.name),
                clean_map.values(),
                scorer=fuzz.token_sort_ratio
            )

            if best:
                match_txt, score, _ = best
                real_item = next(k for k, v in clean_map.items() if v == match_txt)
                target = "MATCH" if score >= MATCH_MIN else "CHECK"
            else:
                real_item, score, target = sheet_items[0], 0, "CHECK"

            st.session_state.results[target].append({
                "image": resized,
                "original": img.name,
                "final": real_item,
                "score": round(score, 2)
            })

    # ================= SUMMARY =================
    m, c, d = st.columns(3)
    m.metric("✅ MATCH", len(st.session_state.results["MATCH"]))
    c.metric("⚠️ CHECK", len(st.session_state.results["CHECK"]))
    d.metric("♻️ DUPLICATE", len(st.session_state.results["DUPLICATE"]))

    # ================= RENDER FUNCTION =================
    def render_section(title, key, allow_confirm):
        data = st.session_state.results[key]
        if not data:
            return

        st.markdown(f"### {title}")
        bulk_remove = []

        for i in range(0, len(data), IMAGES_PER_ROW):
            cols = st.columns(IMAGES_PER_ROW)

            for j in range(IMAGES_PER_ROW):
                if i + j >= len(data):
                    break

                item = data[i + j]
                idx = i + j

                with cols[j]:
                    st.image(item["image"], use_container_width=True)
                    st.caption(item["original"])

                    item["final"] = st.selectbox(
                        "Select Item",
                        sheet_items,
                        index=sheet_items.index(item["final"])
                        if item["final"] in sheet_items else 0,
                        key=f"{key}_sel_{idx}"
                    )

                    st.progress(item["score"] / 100)

                    if st.checkbox("Select Remove", key=f"{key}_chk_{idx}"):
                        bulk_remove.append(item)

                    if allow_confirm:
                        if st.button("✅ Confirm", key=f"{key}_conf_{idx}"):
                            st.session_state.results["MATCH"].append(item)
                            st.session_state.results[key].remove(item)
                            st.rerun()

                    if st.button("❌ Remove", key=f"{key}_rm_{idx}"):
                        st.session_state.results[key].remove(item)
                        st.rerun()

        if bulk_remove and st.button(f"❌ Remove Selected ({title})"):
            for r in bulk_remove:
                st.session_state.results[key].remove(r)
            st.rerun()

    # ================= RENDER =================
    render_section("✅ MATCH", "MATCH", allow_confirm=False)
    render_section("⚠️ CHECK", "CHECK", allow_confirm=True)
    # ❌ DUPLICATE SECTION HIDDEN (ONLY COUNT SHOWN)

    # ================= ZIP =================
    if st.session_state.results["MATCH"]:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for r in st.session_state.results["MATCH"]:
                zipf.writestr(
                    f"{FIXED_FOLDER_NAME}/{r['final']}.jpg",
                    r["image"].getvalue()
                )

        st.download_button(
            "📥 Download TM PRO Folder (ZIP)",
            zip_buffer.getvalue(),
            file_name=f"{FIXED_FOLDER_NAME}.zip",
            mime="application/zip"
        )
