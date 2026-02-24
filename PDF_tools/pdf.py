import io
import os
from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.pdfgen import canvas

def process_inventory_document_final(input_pdf, inv_number):
    if not os.path.exists(input_pdf):
        print(f"Error: {input_pdf} not found.")
        return

    output_folder = inv_number
    os.makedirs(output_folder, exist_ok=True)

    file_name = f"434-99-9999-588_V01_{inv_number}.pdf"
    output_path = os.path.join(output_folder, file_name)

    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    header_template = f"434-99-9999-588_V01_{inv_number}"

    for i, page in enumerate(reader.pages):
        mbox = page.mediabox
        cbox = page.cropbox
        rotation = page.get("/Rotate", 0) 
        
        m_width = float(mbox.width)
        m_height = float(mbox.height)
        
        c_top = float(cbox.top)
        c_left = float(cbox.left)
        c_width = float(cbox.width)

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(m_width, m_height))
        
        # Apply rotation to the canvas coordinate system
        if rotation == 90:
            can.rotate(90)
            can.translate(0, -m_width)
        elif rotation == 180:
            can.rotate(180)
            can.translate(-m_width, -m_height)
        elif rotation == 270:
            can.rotate(270)
            can.translate(-m_height, 0)

        # --- DRAWING ---
        # Header text
        can.setFillColorRGB(0, 0.388, 0.694)


        # Page 1 specific edits 
        if i == 0:

            can.setFont("Helvetica", 15)
            can.drawCentredString(c_top - 465, c_width - 25, header_template) # ADJUST if Header not placed correctly on front page

            can.setFont("Helvetica", 15)
            can.drawString(330, 620, inv_number) #ADJUST x, y to place above the line
            can.setFont("Helvetica-Bold", 12)
            can.drawString(242, 575, "x")  #ADJUST x, y to fit the check box of choice
        else:
            # Header for pages 2-12
            can.setFillColorRGB(1, 1, 1) 
            can.rect(c_left, c_top - 40, c_width, 40, fill=1, stroke=0)
            can.setFillColorRGB(0, 0.388, 0.694)
            can.setFont("Helvetica", 15)
            can.drawCentredString(c_width - 465, c_top - 25, header_template)

        can.save()
        packet.seek(0)

        overlay_page = PdfReader(packet).pages[0]
        
        if i == 0:
            # Stamping on top to bypass Master layer
            page.merge_page(overlay_page, over=True)
            writer.add_page(page)
        else:
            # Merging the page content onto the header canvas for stability
            overlay_page.merge_page(page)
            overlay_page.cropbox = cbox
            writer.add_page(overlay_page)

    with open(output_path, "wb") as f:
        writer.write(f)
    
    print(f"✅ Success! Hybrid merge complete. Saved to: {output_path}")

# --- RUN ---
# THIS ONLY WORKS FOR PDFS WITH ROTATION ON PAGES (Meaning a page may have been scanned and inserted)
# Input path for the pdf which you want to add header (ONE PDF AT A TIME)
# Insert the inventory number in the inv_numbers list
input_path = r"D:\VS\Js_And_Py\PDF_tools\434-99-9999-588_V01 Master.pdf"
inv_nums = ["04313210", "04313220", "04313230"]

for num in inv_nums:
    process_inventory_document_final(input_path, num)