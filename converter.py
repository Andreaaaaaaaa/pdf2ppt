import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from io import BytesIO
import io
from PIL import Image

class PDFToPPTConverter:
    def __init__(self, pdf_file):
        """
        Initialize with a PDF file (stream or path).
        """
        self.pdf_file = pdf_file
        self.doc = fitz.open(stream=pdf_file.read(), filetype="pdf")

    def convert_to_images(self, dpi=200):
        """
        Default Mode: Convert each page to an image and place on a slide.
        Returns a BytesIO object containing the PPTX file.
        """
        prs = Presentation()
        
        # Set slide dimensions to match the first page of PDF (assuming uniform size)
        if len(self.doc) > 0:
            page = self.doc[0]
            # PyMuPDF uses points (1/72 inch), pptx uses EMUs (914400 per inch)
            width_inch = page.rect.width / 72
            height_inch = page.rect.height / 72
            
            prs.slide_width = int(width_inch * 914400)
            prs.slide_height = int(height_inch * 914400)

        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            
            # Render page to image (high resolution)
            # Default PDF is 72 DPI. Zoom = dpi / 72
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            
            # Convert to PIL Image
            img_data = pix.tobytes("png")
            image_stream = BytesIO(img_data)
            
            # Add slide
            blank_slide_layout = prs.slide_layouts[6] # 6 is blank
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # Add image to slide
            slide.shapes.add_picture(image_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    def convert_separated(self):
        """
        Separation Mode: Extract images and text.
        Returns a BytesIO object containing the PPTX file.
        """
        prs = Presentation()
        
        # Set slide dimensions (standard 16:9 or based on PDF?)
        # Let's stick to PDF dimensions for consistency
        if len(self.doc) > 0:
            page = self.doc[0]
            width_inch = page.rect.width / 72
            height_inch = page.rect.height / 72
            prs.slide_width = int(width_inch * 914400)
            prs.slide_height = int(height_inch * 914400)

        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            
            # Create a slide for this page
            blank_slide_layout = prs.slide_layouts[6]
            slide = prs.slides.add_slide(blank_slide_layout)
            
            # 1. Extract and place images
            image_list = page.get_images(full=True)
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = self.doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_stream = BytesIO(image_bytes)
                
                # Try to find image location on page
                # This is tricky in PDF, get_images doesn't return rect.
                # We might need to use page.get_image_rects(xref) if available or iterate drawings.
                # For simplicity in this version, we'll tile them or just place them.
                # BETTER APPROACH for "Separation":
                # Render the page WITHOUT text to get the "background/images" layer?
                # Or just place extracted images in a grid?
                # User asked for "Text-Image Separation".
                # Let's try to place images roughly where they might be or just list them.
                # To keep it usable, let's place extracted images on the left/top and text on the right/bottom?
                # OR: Just render the page as an image for the background (visual context) and put text in editable boxes?
                # No, "Separation" implies they want the raw assets.
                
                # Let's go with: Slide 1 for Page 1 Images, Slide 2 for Page 1 Text?
                # Or: Slide 1 has images on one side, text on the other.
                
                # Let's try to preserve layout but separate elements.
                # Actually, PyMuPDF can give us image rects.
                try:
                    rects = page.get_image_rects(xref)
                    for rect in rects:
                        # Convert rect to pptx units
                        left = Inches(rect.x0 / 72)
                        top = Inches(rect.y0 / 72)
                        width = Inches(rect.width / 72)
                        height = Inches(rect.height / 72)
                        slide.shapes.add_picture(image_stream, left, top, width=width, height=height)
                except Exception as e:
                    print(f"Could not place image {img_index}: {e}")
                    # Fallback: just add it somewhere
                    slide.shapes.add_picture(image_stream, Inches(0.5), Inches(0.5), height=Inches(3))

            # 2. Extract and place text
            # get_text("dict") gives us blocks with bbox
            text_blocks = page.get_text("dict")["blocks"]
            for block in text_blocks:
                if block["type"] == 0: # Text
                    bbox = block["bbox"]
                    
                    # Add text box
                    left = Inches(bbox[0] / 72)
                    top = Inches(bbox[1] / 72)
                    width = Inches((bbox[2] - bbox[0]) / 72)
                    height = Inches((bbox[3] - bbox[1]) / 72)
                    
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    tf.word_wrap = True
                    
                    for line in block["lines"]:
                        p = tf.add_paragraph()
                        
                        for span in line["spans"]:
                            run = p.add_run()
                            run.text = span["text"]
                            
                            # Apply formatting
                            try:
                                # Font Size
                                run.font.size = Pt(span["size"])
                                
                                # Font Color
                                # PyMuPDF returns color as sRGB integer
                                color_int = span["color"]
                                r = (color_int >> 16) & 0xFF
                                g = (color_int >> 8) & 0xFF
                                b = color_int & 0xFF
                                run.font.color.rgb = RGBColor(r, g, b)
                                
                                # Bold/Italic (flags: 1=superscript, 2=italic, 4=serif, 8=monospaced, 16=bold)
                                # Note: PyMuPDF flags might vary by version, but usually:
                                # bit 0: superscript
                                # bit 1: italic
                                # bit 2: serif
                                # bit 3: monospaced
                                # bit 4: bold
                                flags = span["flags"]
                                if flags & 2**4: # Bold
                                    run.font.bold = True
                                if flags & 2**1: # Italic
                                    run.font.italic = True
                                    
                                # Font Name (Optional, might not map perfectly)
                                # run.font.name = span["font"] 
                                
                            except Exception as e:
                                # Ignore formatting errors, keep text
                                pass

        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    def extract_text_content(self):
        """
        Extract all text from the PDF.
        Returns a BytesIO object containing the text file.
        """
        text_output = io.StringIO()
        
        for page_num in range(len(self.doc)):
            page = self.doc[page_num]
            text = page.get_text()
            text_output.write(f"--- Page {page_num + 1} ---\n\n")
            text_output.write(text)
            text_output.write("\n\n")
            
        output = BytesIO(text_output.getvalue().encode('utf-8'))
        output.seek(0)
        return output
