import io
import os
from pypdf import PdfReader, PdfWriter
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
        box = page.cropbox
        width = float(box.width)
        height = float(box.height)
        
        # Check for rotation
        rotation = page.get("/Rotate", 0)

        packet = io.BytesIO()
        # Create canvas matching the visible dimensions
        can = canvas.Canvas(packet, pagesize=(width, height))
        
        # Correction for rotation, so if a page is rotated, coordinates don't need to be changed
        if rotation == 90:
            can.rotate(90)
            can.translate(0, -width)
        elif rotation == 180:
            can.rotate(180)
            can.translate(-width, -height)
        elif rotation == 270:
            can.rotate(270)
            can.translate(-height, 0)

        # White rectangle fill
        can.setFillColorRGB(1, 1, 1) 
        can.rect(0, height - 40, width, 40, fill=1, stroke=0)
        
        can.setFillColorRGB(0, 0, 0)
        can.setFont("Helvetica-Bold", 11)
        # Centered string pushed to the left
        can.drawCentredString(width - 465, height - 25, header_template)

        # Inventory number on line and checkbox on windows
        if i == 0:
            can.setFont("Helvetica", 12)
            can.drawString(350, 647, inv_number) # Gerate Nummer
            can.setFont("Helvetica-Bold", 12)
            can.drawString(239, 595, "x") # Checkbox

        can.save()
        packet.seek(0)

        header_reader = PdfReader(packet)
        new_page = header_reader.pages[0]
        
        # Merge content (canvas and pdf page)
        new_page.merge_page(page)
        new_page.cropbox = box
        writer.add_page(new_page)

    # Saves in folder named after inventory number
    with open(output_path, "wb") as f:
        writer.write(f)
    
    print(f"Success! Page 1 should now show edits in Portrait. Saved to: {output_path}")

# --- RUN ---
# Make sure file path is correct, and file sits within the folder of this pdf.py
input_path = r"D:\VS\Js_And_Py\PDF_tools\434-99-9999-588_V01_Plan_Bericht-Migration der OT Geräte.pdf"
# Place each inventory number in here 
inventory_numbers = ["04313210", "04313220", "04313230"]

for num in inventory_numbers:
    process_inventory_document_final(input_path, num)