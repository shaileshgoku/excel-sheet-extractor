import streamlit as st
import pandas as pd

# Function to adjust font size using HTML
def format_text(text, size):
    return f"<p style='font-size:{size}px;'>{text}</p>"

# Streamlit UI setup
st.title('Excel Sheet Extractor')

# Custom font-sized text
upload_instruction = format_text('Upload Excel File', 18)
sheet_instruction = format_text('Enter the sheet names you want to extract, separated by commas.', 18)



st.markdown(upload_instruction, unsafe_allow_html=True)

# Load the Excel file
uploaded_file = st.file_uploader('', type=['xlsx'])

# Display instructions with custom font size
st.markdown(sheet_instruction, unsafe_allow_html=True)
sheet_names_input = st.text_input('Sheet Names:', '')

if uploaded_file:
    # Read the Excel file
    excel_data = pd.ExcelFile(uploaded_file)

    if st.button('Generate New Excel File'):
        # Split the input into a list of sheet names and strip any extra whitespace
        sheet_names = [name.strip() for name in sheet_names_input.split(',')]
        
        # Filter existing sheet names
        existing_sheets = [name for name in sheet_names if name in excel_data.sheet_names]
        missing_sheets = [name for name in sheet_names if name not in excel_data.sheet_names]

        if existing_sheets:
            # Create a new Excel writer object
            output_file_path = 'output_excel_file.xlsx'
            with pd.ExcelWriter(output_file_path) as writer:
                for sheet_name in existing_sheets:
                    # Read the specified sheet
                    df = excel_data.parse(sheet_name)
                    # Write the DataFrame to the new sheet
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            st.success(f"The sheets '{', '.join(existing_sheets)}' have been written to {output_file_path}")
            st.download_button(label='Download New Excel File', data=open(output_file_path, 'rb').read(), file_name=output_file_path)
        
        if missing_sheets:
            st.warning(f"The following sheets were not found in the Excel file: {', '.join(missing_sheets)}")
        
        if not existing_sheets and not missing_sheets:
            st.error("No valid sheet names were entered.")
