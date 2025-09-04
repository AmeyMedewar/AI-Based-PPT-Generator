import streamlit as st
import requests
import os

# Backend URL (FastAPI)
BACKEND_URL = "http://127.0.0.1:8000"
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

st.set_page_config(page_title="AI Slide Generator", layout="centered")
st.title("📊 AI Slide Generator")
st.write("Generate AI-powered presentation slides. Upload a document optionally, set parameters, and download the PPTX.")

# --------------------
# Upload Section (Optional)
# --------------------
uploaded_file = st.file_uploader("Upload a file (optional, .txt, .pdf, .docx, .csv)", type=["txt", "pdf", "docx", "csv"])
filename = None

if uploaded_file:
    save_path = os.path.join(UPLOAD_DIR, uploaded_file.name)
    with open(save_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success(f"✅ File '{uploaded_file.name}' saved for processing")
    filename = uploaded_file.name

# --------------------
# Presentation Settings Form (Always Visible)
# --------------------
with st.form("ppt_settings"):
    topic = st.text_input("Topic (required)")
    audience = st.text_input("Target Audience", placeholder="Optional")
    purpose = st.text_input("Purpose", placeholder="Optional")
    num_slides = st.number_input("Number of Slides", min_value=1, max_value=50, value=5, step=1, help="Optional")
    style = st.selectbox("Style/Tone", ["", "Professional", "Minimalist", "Casual"], index=0)
    complexity = st.selectbox("Content Complexity", ["", "Basic", "Intermediate", "Advanced"], index=0)
    language = st.selectbox("Language", ["", "English"], index=0)
    notes = st.text_area("Additional Notes", placeholder="Optional")
    ppt_type = st.selectbox("PPT Type", ["Silver", "Gold", "Platinum"])
    submitted = st.form_submit_button("✨ Generate Slides")

if submitted:
    if not topic.strip():
        st.error("❌ Topic is required")
    else:
        with st.spinner("Generating slides..."):
            # Prepare common params
            params = {
                "filename": filename or "",
                "topic": topic.strip(),
                "audience": audience.strip() or "",
                "purpose": purpose.strip() or "",
                "num_slides": num_slides,
                "style": style or "",
                "complexity": complexity or "",
                "language": language or "",
                "notes": notes.strip() or "",
                "ppt_type":ppt_type
            }

            # Choose endpoint based on PPT type
            if ppt_type == "Silver":
                endpoint = "generate_silver"
            elif ppt_type == "Gold":
                endpoint = "generate_gold"
            elif ppt_type == "Platinum":
                endpoint = "generate_platinum"
            else:
                endpoint = "generate_silver"  # fallback

            # Call the backend
            res = requests.post(f"{BACKEND_URL}/{endpoint}/", params=params)

        if res.status_code == 200:
            data = res.json()
            st.success("✅ Slides generated successfully!")

            # Show warning if content trimmed
            if "warning" in data:
                st.warning(data["warning"])

            # Download button
            pptx_path = data["pptx_path"]
            with open(pptx_path, "rb") as f:
                st.download_button(
                    label="📥 Download Slides",
                    data=f,
                    file_name=os.path.basename(pptx_path),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            # Optional: preview first slide bullets
            preview = data.get("preview", [])
            if preview:
                st.subheader("Preview of First Slide Bullets")
                for bullet in preview:
                    st.write(f"- {bullet}")
        else:
            st.error("❌ Slide generation failed. Please check the backend logs.")
