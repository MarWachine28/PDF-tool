import io
import os
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

def process_inventory_document_final(input_pdf, inv_number):
    if not os.path.exists(input_pdf):
        print(f"Error: {input_pdf} not found.")
        return
    

    # Exist true so it doesnt crash if folder already exists
    output_folder = inv_number
    os.makedirs(output_folder, exist_ok=True)

    file_name = f"434-99-9999-588_V01_{inv_number}.pdf"
    output_path = os.path.join(output_folder, file_name)


    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    header_template = f"434-99-9999-588_V01_{inv_number}"

    for i, page in enumerate(reader.pages):
        # 1. Get the visible area (CropBox)
        box = page.cropbox
        width = float(box.width)
        height = float(box.height)

        # 2. Create the "Base Layer" (Header + Edits)
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(width, height))
        
        # --- PART A: THE HEADER (Inverse Style) ---
        # Draw a small white background strip to clear space for the header
        can.setFillColorRGB(1, 1, 1) # White
        can.rect(0, height - 40, width, 40, fill=1, stroke=0)
        
        can.setFillColorRGB(0, 0, 0) # Black
        can.setFont("Helvetica-Bold", 11)
        can.drawCentredString(width - 465, height - 25, header_template)

        # --- PART B: PAGE 1 EDITS ---
        if i == 0:
            # Fill the "Inventar/Gerate Nummer:" line
            can.setFont("Helvetica", 12)
            can.drawString(350, 647, inv_number)

            # Check the "Windows Environment" box
            can.setFont("Helvetica-Bold", 12)
            can.drawString(239, 595, "x")

        can.save()
        packet.seek(0)

        # 3. Create the new base page
        header_reader = PdfReader(packet)
        new_page = header_reader.pages[0]

        # 4. Merge ORIGINAL content ONTO the new page
        # This uses the inverse logic that worked for you before
        new_page.merge_page(page)
        
        # Sync the cropbox so the view remains the same
        new_page.cropbox = box
        writer.add_page(new_page)

    with open(output_path, "wb") as f:
        writer.write(f)
    
    print(f"✅ Process Complete! File saved in folder: {output_path}")


# --- RUN ---
# Input path for the pdf which you want to add header (ONE PDF AT A TIME)
# Make sure the file path is copied correctly
# Insert the inventory number in the inv_numbers list

input_path = r"D:\VS\Js_And_Py\PDF_tools\434-99-9999-588_V01_Plan_Bericht-Migration der OT Geräte.pdf"
inv_nums = ["04313210", "04313220", "04313230"]

for num in inv_nums:
    process_inventory_document_final(input_path, num)