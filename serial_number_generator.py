import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

def create_serials_excel(client_name, order_number, product_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Serial Numbers"

    # Safari Micro branding colors
    header_color = "004785"
    separator_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # Add Safari Micro Logo (Assuming logo file is present)
    img_path = "safari_micro_logo.png"  # Ensure the file is in the correct directory
    try:
        from openpyxl.drawing.image import Image
        logo = Image(img_path)
        ws.add_image(logo, "A1")
    except:
        pass  # Prevent crash if image is missing

    # Header formatting
    ws.merge_cells("A1:D3")
    ws["A4"] = "Safari Micro - Serial Number Report"
    ws["A4"].font = Font(size=16, bold=True, color=header_color)
    ws["A4"].alignment = Alignment(horizontal="left")

    # Customer message
    ws["A6"] = "Dear Valued Customer," 
    ws["A6"].font = Font(size=12, italic=True)
    ws["A7"] = "Please find below the serial numbers for your order. If you need any assistance, don't hesitate to reach out."
    ws["A7"].font = Font(size=12)

    # Order details
    ws["A9"], ws["B9"] = "Client Name:", client_name
    ws["A10"], ws["B10"] = "Order Number:", order_number

    # Table headers
    headers = ["Product", "Model Number", "Serial Number"]
    ws.append(headers)
    for col in range(1, 4):
        ws.cell(row=13, column=col).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=13, column=col).alignment = Alignment(horizontal="center")
        ws.cell(row=13, column=col).fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")

    # Insert serial numbers with separators for multiple products
    row_index = 14
    for product, model, serial_numbers in product_data:
        if row_index > 14:
            ws[f"A{row_index}"] = "---"
            ws[f"A{row_index}"].fill = separator_fill
            row_index += 1
        for serial in serial_numbers:
            ws[f"A{row_index}"] = product
            ws[f"B{row_index}"] = model
            ws[f"C{row_index}"] = serial
            row_index += 1
    
    # Adjust column width
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 40

    # Save to a BytesIO object
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# Streamlit UI
st.set_page_config(page_title="Safari Micro Serial Number to Excel Tool", page_icon="📄", layout="centered")
st.image("safari_micro_logo.png", width=300)  # Ensure logo is present in directory

st.title("Safari Micro Serial Number to Excel Tool")

client_name = st.text_input("Client Name", key="client_name")
order_number = st.text_input("Order Number", key="order_number")

st.write("Enter Product Details Below:")

product_data = []
num_products = st.number_input("Number of Line Items", min_value=1, step=1, key="num_products")

for i in range(num_products):
    product_name = st.text_input(f"Product {i+1} Name", key=f"product_name_{i}")
    model_number = st.text_input(f"Product {i+1} Model Number", key=f"model_number_{i}")
    serial_input = st.text_area(f"Enter Serial Numbers for Product {i+1} (comma-separated)", "", key=f"serial_input_{i}")
    serials = [s.strip() for s in serial_input.split(",") if s.strip()]
    product_data.append((product_name, model_number, serials))

if st.button("Generate Excel File", key="generate_excel"):
    excel_data = create_serials_excel(client_name, order_number, product_data)
    st.download_button(label="Download Excel File", data=excel_data, file_name="SafariMicro_Serials.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
