
import os
import io
import time
import tempfile
from pathlib import Path

import streamlit as st

# Local imports
from ppt_generator import PPTGenerator

st.set_page_config(page_title="Beautiful PPT Generator", page_icon="📊", layout="centered")

st.title("📊 Beautiful PPT Generator")
st.caption("Gemini/Gemma + Pexels + python-pptx — with a clean design system")

with st.form("ppt_form"):
    topic = st.text_input("Topic *", placeholder="e.g., Generative AI for Product Teams")
    col1, col2 = st.columns(2)
    with col1:
        slide_count = st.slider("Slide count", 5, 25, 10, step=1)
        use_images = st.checkbox("Use images from Pexels", value=True)
    with col2:
        deck_title = st.text_input("Deck Title (optional)", placeholder="Leave blank to use topic")
        subtitle = st.text_input("Subtitle (optional)", placeholder="")
    author = st.text_input("Author", value="")
    logo_file = st.file_uploader("Logo (optional, PNG/JPG)", type=["png", "jpg", "jpeg"], accept_multiple_files=False)

    with st.expander("Advanced options"):
        audience = st.text_input("Audience", value="General")
        tone = st.text_input("Tone", value="Crisp, clear, professional")
        visual_style = st.text_input("Visual Style", value="Modern, clean, minimal")
        override_model = st.text_input("Override GEMINI_MODEL (optional)", value=os.getenv("GEMINI_MODEL", ""))

    submitted = st.form_submit_button("Generate Presentation", type="primary")

# Prepare temp logo path if uploaded
tmp_logo_path = None
if logo_file is not None:
    try:
        suffix = Path(logo_file.name).suffix.lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(logo_file.read())
            tmp_logo_path = tmp.name
    except Exception as e:
        st.warning(f"Logo upload failed: {e}")

if submitted:
    if not topic.strip():
        st.error("Please enter a topic.")
        st.stop()

    # Optional model override: the generator reads GEMINI_MODEL from env
    if override_model.strip():
        os.environ["GEMINI_MODEL"] = override_model.strip()

    out_name = f"presentation_{int(time.time())}.pptx"
    out_path = str(Path.cwd() / out_name)

    try:
        gen = PPTGenerator()
        result_path = gen.generate_presentation(
            topic=topic.strip(),
            audience=audience.strip(),
            tone=tone.strip(),
            visual_style=visual_style.strip(),
            slide_count=slide_count,
            title=(deck_title.strip() or topic.strip()),
            subtitle=subtitle.strip(),
            author=author.strip(),
            logo_path=tmp_logo_path,
            output_path=out_path,
            download_images=use_images,
        )

        st.success("Presentation generated successfully!")
        with open(result_path, "rb") as f:
            st.download_button(
                "⬇️ Download PPT",
                data=f,
                file_name="presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        st.caption("Tip: If images didn't appear, set your PEXELS_API_KEY in the .env and try again.")

    except Exception as e:
        st.error(f"Failed to generate presentation: {e}")
    finally:
        # Clean up temp logo
        if tmp_logo_path and Path(tmp_logo_path).exists():
            try:
                Path(tmp_logo_path).unlink(missing_ok=True)
            except Exception:
                pass

st.markdown("---")
st.markdown("Powered by **Google Gemini/Gemma**, **Pexels**, and **python-pptx**.")
