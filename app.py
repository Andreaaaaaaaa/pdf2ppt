import streamlit as st
from converter import PDFToPPTConverter

st.set_page_config(page_title="PDF to PPT Converter", layout="centered")

st.title("📄 PDF to PPT Converter")
st.markdown("Convert your PDF presentations to PowerPoint with high fidelity.")

uploaded_file = st.file_uploader("Upload your PDF file", type=["pdf"])

if uploaded_file is not None:
    st.success("File uploaded successfully!")
    
    # Configuration
    st.subheader("Configuration")
    
    # Mode Selection
    mode = st.radio(
        "Conversion Mode",
        ("Image-based (High Fidelity)", "Text-Image Separation (Editable)"),
        help="Image-based: Converts pages to images (Reliable, preserves layout/watermarks). Separation: Extracts text and images (Editable)."
    )
    
    dpi = 200
    if mode == "Image-based (High Fidelity)":
        dpi = st.slider("Image Quality (DPI)", min_value=72, max_value=400, value=200, step=10, help="Higher DPI means better quality but larger file size.")
    
    if st.button("Convert to PPT"):
        with st.spinner("Converting..."):
            try:
                # Reset file pointer just in case
                uploaded_file.seek(0)
                
                converter = PDFToPPTConverter(uploaded_file)
                
                if mode == "Text-Image Separation (Editable)":
                    ppt_file = converter.convert_separated()
                else:
                    ppt_file = converter.convert_to_images(dpi=dpi)
                
                st.success("Conversion complete!")
                
                st.download_button(
                    label="Download PPT",
                    data=ppt_file,
                    file_name="converted_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.exception(e)

st.markdown("---")
st.markdown("Created with ❤️ by Antigravity")
