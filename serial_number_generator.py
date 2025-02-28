import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

def create_serials_excel(client_name, order_number, product_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Serial Numbers"

    # Header formatting
    ws.merge_cells("A1:D3")
    ws["A4"] = "Safari Micro - Serial Number Report"
    ws["A4"].font = Font(size=16, bold=True, color="004785")
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
        ws.cell(row=13, column=col).fill = PatternFill(start_color="004785", end_color="004785", fill_type="solid")

    # Insert serial numbers
    row_index = 14
    for product, model, serial_numbers in product_data:
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
st.title("Safari Micro - Serial Number Generator")

client_name = st.text_input("Client Name")
order_number = st.text_input("Order Number")

st.write("Enter Product Details Below:")

product_data = []
num_products = st.number_input("Number of Line Items", min_value=1, step=1)

for i in range(num_products):
    product_name = st.text_input(f"Product {i+1} Name")
    model_number = st.text_input(f"Product {i+1} Model Number")
    serial_input = st.text_area(f"Enter Serial Numbers for {product_name} (comma-separated)", "")
    serials = [s.strip() for s in serial_input.split(",") if s.strip()]
    product_data.append((product_name, model_number, serials))

if st.button("Generate Excel File"):
    excel_data = create_serials_excel(client_name, order_number, product_data)
    st.download_button(label="Download Excel File", data=excel_data, file_name="SafariMicro_Serials.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
