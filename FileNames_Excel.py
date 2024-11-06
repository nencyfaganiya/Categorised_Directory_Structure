import os
from datetime import datetime
import pandas as pd
import streamlit as st
from pathlib import Path
from openpyxl.styles import Font, Alignment
import pyperclip

# Helper function to get only files (exclude folders)
def get_files(path):
    items = []
    for root, dirs, files in os.walk(path):
        for name in files:  # Only include files, exclude folders
            full_path = os.path.join(root, name)
            modified_time = datetime.fromtimestamp(os.path.getmtime(full_path)).strftime('%Y-%m-%d')
            items.append((name, modified_time))
    return items

# Function to generate Excel file
def generate_excel(categories, output_path):
    # Create a Pandas DataFrame
    data = []

    for category, items in categories.items():
        if items:  # Only include categories with files
            # Append the category name as a single entry
            data.append({'Category': category, 'Last Modified': ''})  # Empty for the first line
            for name, modified_time in items:
                data.append({'Category': f"   {name}", 'Last Modified': modified_time})  # Indent file names

    df = pd.DataFrame(data)

    # Create an Excel writer
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Files', startrow=1)
        workbook = writer.book
        worksheet = writer.sheets['Files']

        # Add a header
        headers = ['Category', 'Last Modified']
        for col_num, value in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=col_num, value=value)
            cell.font = Font(bold=True, size=12)  # Bold and increase font size
            cell.alignment = Alignment(horizontal='center')  # Center align header

        # Set column widths for better visibility
        worksheet.column_dimensions['A'].width = 50  # Category
        worksheet.column_dimensions['B'].width = 20  # Last Modified

        # Format the data
        for index, row in df.iterrows():
            category_cell = worksheet.cell(row=index + 2, column=1, value=row['Category'])
            category_cell.alignment = Alignment(horizontal='left', vertical='top')  # Align left and top
            
            last_modified_cell = worksheet.cell(row=index + 2, column=2, value=row['Last Modified'])
            last_modified_cell.alignment = Alignment(horizontal='center')  # Center align last modified date

            # Apply bold style for category names
            if not row['Last Modified']:  # This means it's a category row
                category_cell.font = Font(bold=True, size=12)  # Bold and increase font size

    return output_path

# Streamlit UI
st.title("File Categorization Tool")

# Directory path input
directory_path = st.text_input("Enter the directory path:", "")

# If directory is provided, proceed
if directory_path and Path(directory_path).exists():
    items = get_files(directory_path)
    
    if items:
        # Updated categories
        categories = ["CONTRACTUAL", "ARCHITECTURAL", "STRUCTURAL", "SERVICES", "SAFETY"]
        category_selection = {}

        st.write("### Assign Categories")
        for index, item in enumerate(items):
            name, modified_time, full_path = item

            # Create a horizontal layout for file name and category selection
            cols = st.columns([3, 1])  
            
            # Display file name clearly with modified time
            with cols[0]:
                # Button to copy path to clipboard
                if st.button(f"{index+1} {name}", key=f"copy_button_{index}"):
                    pyperclip.copy(full_path)
                    st.success(f"Path copied to clipboard: {full_path}")

            # Create a dropdown for each item with a unique key right next to the file name
            # Dropdown for each file's category selection with a non-empty placeholder label
            with cols[1]:
                category = st.selectbox(
                    label="Select category",  # Provide a placeholder label for accessibility
                    options=categories,
                    key=f"selectbox_{index}",
                    label_visibility="collapsed"  # Hide the label in the UI
                )
                category_selection[name] = (modified_time, category)

        # Generate Excel button
        if st.button("Generate Excel"):
            categorized_data = {cat: [] for cat in categories}
            for name, (modified_time, category) in category_selection.items():
                categorized_data[category].append((name, modified_time))

            # Generate and show Excel link
            output_excel_path = "Categorized_files.xlsx"
            excel_path = generate_excel(categorized_data, output_excel_path)
            st.success("Excel file generated successfully!")
            st.download_button("Download Excel", data=open(excel_path, "rb"), file_name=output_excel_path)
    else:
        st.warning("No files found in the selected directory.")
else:
    st.info("Please enter a valid directory path to start categorizing files.")
