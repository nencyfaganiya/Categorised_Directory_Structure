import os
from datetime import datetime
import pandas as pd
import streamlit as st
from pathlib import Path
from openpyxl.styles import Font, Alignment  # <-- Add Font and Alignment here
from io import BytesIO
from docx import Document
from docx.shared import Pt
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import pyperclip


# Helper function to get files (excluding folders)
def get_files(path):
    items = []
    for root, dirs, files in os.walk(path):
        for name in files:
            full_path = os.path.join(root, name)
            modified_time = datetime.fromtimestamp(os.path.getmtime(full_path)).strftime('%Y-%m-%d')
            items.append((name, modified_time, full_path))
    return items


# Helper function to resolve paths
def resolve_path(path):
    """
    Resolves the path for both client and server environments.
    - If the app is running on the server, it expects UNC paths.
    """
    if path.startswith("Z:") or path.startswith("Y:"):  # Replace with your mapped drive letters
        # Map drive letter to its UNC equivalent
        drive_mappings = {
            "Z:": r"\\network-share\folder",  # Replace with actual UNC paths
            "Y:": r"\\another-share\folder"
        }
        unc_path = drive_mappings.get(path[:2], None) + path[2:].replace("\\", "/")
        if not unc_path or not os.path.exists(unc_path):
            raise ValueError(f"Invalid path or UNC path not accessible: {unc_path}")
        return unc_path

    # If already a UNC path
    if path.startswith("\\\\"):
        if not os.path.exists(path):
            raise ValueError(f"Invalid UNC path: {path}")
        return path

    # Local paths can be used as-is
    if os.path.exists(path):
        return path

    raise ValueError(f"Unrecognized or unsupported path format: {path}")


# Word generation in memory
def generate_word(categories):
    doc = Document()
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Category / File Name'
    hdr_cells[1].text = 'Last Modified'
    for cell in hdr_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(12)

    for category, items in categories.items():
        if items:
            category_row = table.add_row().cells
            category_row[0].text = category
            category_row[1].text = ""
            category_row[0].paragraphs[0].runs[0].font.bold = True
            for name, modified_time in items:
                file_row = table.add_row().cells
                file_row[0].text = name
                file_row[1].text = modified_time
                for cell in file_row:
                    cell.vertical_alignment = True

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# PDF generation in memory
def generate_pdf(categories):
    buffer = BytesIO()
    pdf = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    bold_style = styles["Heading4"]
    bold_style.fontSize = 10
    normal_style = styles["BodyText"]
    normal_style.fontSize = 10
    data = [[Paragraph("Category / File Name", bold_style), Paragraph("Last Modified", bold_style)]]

    for category, items in categories.items():
        if items:
            data.append([Paragraph(category, bold_style), ""])
            for name, modified_time in items:
                wrapped_name = Paragraph(name, normal_style)
                data.append([wrapped_name, modified_time])

    table = Table(data, colWidths=[250, 100])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
    ]))

    pdf.build([table])
    buffer.seek(0)
    return buffer


# Excel generation in memory
def generate_excel(categories):
    data = []
    for category, items in categories.items():
        if items:
            data.append({'Category / File Name': category, 'Last Modified': ''})
            for name, modified_time in items:
                data.append({'Category / File Name': f"   {name}", 'Last Modified': modified_time})

    df = pd.DataFrame(data)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Files', startrow=1)
        worksheet = writer.sheets['Files']
        headers = ['Category / File Name', 'Last Modified']
        for col_num, value in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=col_num, value=value)
            cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal='left')
        worksheet.column_dimensions['A'].width = 50
        worksheet.column_dimensions['B'].width = 20

        for index, row in df.iterrows():
            worksheet.cell(row=index + 2, column=1, value=row['Category / File Name'])
            worksheet.cell(row=index + 2, column=2, value=row['Last Modified'])

    buffer.seek(0)
    return buffer


# Streamlit UI
st.title("File Categorization Tool")
directory_path = st.text_input("Enter the directory path:", "")

# Session states for checkboxes and generated files
if 'generate_excel_option' not in st.session_state:
    st.session_state.generate_excel_option = False
if 'generate_word_option' not in st.session_state:
    st.session_state.generate_word_option = False
if 'generate_pdf_option' not in st.session_state:
    st.session_state.generate_pdf_option = False
if 'generated_files' not in st.session_state:
    st.session_state.generated_files = {}
if 'category_selection' not in st.session_state:
    st.session_state.category_selection = {}

# Display checkboxes horizontally using columns
col1, col2, col3 = st.columns(3)

with col1:
    st.session_state.generate_excel_option = st.checkbox("Generate Excel", value=st.session_state.generate_excel_option)
with col2:
    st.session_state.generate_word_option = st.checkbox("Generate Word", value=st.session_state.generate_word_option)
with col3:
    st.session_state.generate_pdf_option = st.checkbox("Generate PDF", value=st.session_state.generate_pdf_option)

# Get files in directory
if directory_path:
    try:
        resolved_path = resolve_path(directory_path)
        items = get_files(resolved_path)
        
        if items:
            categories = ["CONTRACTUAL", "ARCHITECTURAL", "STRUCTURAL", "SERVICES", "SAFETY"]
            category_selection = {}

            st.write("### Assign Categories")
            for index, item in enumerate(items):
                name, modified_time, full_path = item
                cols = st.columns([3, 1])
                
                with cols[0]:
                    if st.button(f"{index+1} {name}", key=f"copy_button_{index}"):
                        pyperclip.copy(full_path)
                        st.success(f"Path copied to clipboard: {full_path}")
                with cols[1]:
                    # Track category selections to detect changes
                    category = st.selectbox("Select category", options=categories, key=f"selectbox_{index}", label_visibility="collapsed")
                    category_selection[name] = (modified_time, category)

            # Detect if category selection has changed by comparing with session state
            if category_selection != st.session_state.category_selection:
                # Update session state with new selection and clear generated files to force regeneration
                st.session_state.category_selection = category_selection
                st.session_state.generated_files.clear()

            # Generate files if selected
            if st.button("Generate Selected Files"):
                categorized_data = {cat: [] for cat in categories}
                for name, (modified_time, category) in st.session_state.category_selection.items():
                    categorized_data[category].append((name, modified_time))

                # Generate files only if the file type is selected
                if st.session_state.generate_excel_option:
                    output_excel_buffer = generate_excel(categorized_data)
                    st.session_state.generated_files['excel'] = output_excel_buffer

                if st.session_state.generate_word_option:
                    output_word_buffer = generate_word(categorized_data)
                    st.session_state.generated_files['word'] = output_word_buffer

                if st.session_state.generate_pdf_option:
                    output_pdf_buffer = generate_pdf(categorized_data)
                    st.session_state.generated_files['pdf'] = output_pdf_buffer

            # Download generated files
            for file_type, file_buffer in st.session_state.generated_files.items():
                if file_type == 'excel':
                    st.download_button("Download Excel", data=file_buffer, file_name='output.xlsx', mime='application/vnd.ms-excel')
                elif file_type == 'word':
                    st.download_button("Download Word", data=file_buffer, file_name='output.docx', mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                elif file_type == 'pdf':
                    st.download_button("Download PDF", data=file_buffer, file_name='output.pdf', mime='application/pdf')

    except ValueError as e:
        st.error(str(e))
