import streamlit as st
import pandas as pd
import json
from fuzzywuzzy import fuzz
from io import BytesIO
import os

# Directory to save configurations
config_dir = 'configurations'
os.makedirs(config_dir, exist_ok=True)

# Initialize session state for pairs list if it doesn't exist
if 'pairs' not in st.session_state:
    st.session_state.pairs = []

# Function to add a new pair to the list
def add_pair():
    st.session_state.pairs.append({'text1': '', 'text2': '', 'toggle': False})

# Function to delete a pair from the list
def delete_pair(index):
    del st.session_state.pairs[index]

st.title('Excel Files Comparison App')

# Function to save pairs configuration to a file
def save_configuration(config_name):
    config_path = os.path.join(config_dir, f'{config_name}.json')
    config = {'pairs': st.session_state.pairs}
    with open(config_path, 'w') as f:
        json.dump(config, f)

# Function to load pairs configuration from a file
def load_configuration(config_path):
    with open(config_path, 'r') as f:
        config = json.load(f)
    st.session_state.pairs = config['pairs']

# Function to get the list of saved configurations
def get_saved_configurations():
    return [f.replace('.json', '') for f in os.listdir(config_dir) if f.endswith('.json')]

# Sidebar for saving and loading configurations
with st.sidebar:
    st.title("Configuration")
    config_name = st.text_input("New Configuration Name")
    if st.button("Save Configuration"):
        save_configuration(config_name)
        st.success(f"Configuration '{config_name}' saved!")

    # Dropdown to select a configuration to load
    saved_configs = get_saved_configurations()
    selected_config = st.selectbox("Select a configuration to load", saved_configs)
    if st.button("Load Selected Configuration"):
        config_path = os.path.join(config_dir, f'{selected_config}.json')
        load_configuration(config_path)
        st.success(f"Configuration '{selected_config}' loaded!")

# File uploaders
cols = st.columns(2)
with cols[0]:
    excel_file1 = st.file_uploader("Upload Excel 1 (.xlsx)", type=['xlsx'], key='file1')
with cols[1]:
    excel_file2 = st.file_uploader("Upload Excel 2 (.xlsx)", type=['xlsx'], key='file2')

df_excel_file1 = pd.DataFrame()
df_excel_file2 = pd.DataFrame()

# Display pairs of text inputs with bin icons for deletion
for i, pair in enumerate(st.session_state.pairs):
    cols = st.columns([3, 3, 0.2, 0.2, 0.6])  # Adjust the column widths as necessary
    with cols[0]:
        pair['text1'] = st.text_input(f"Text 1 for pair {i+1}", pair['text1'], label_visibility='collapsed')
    with cols[1]:
        pair['text2'] = st.text_input(f"Text 2 for pair {i+1}", pair['text2'], label_visibility='collapsed')
    with cols[2]:  # This is where the toggle button will go
        pair['toggle'] = st.checkbox("", value=pair['toggle'], key=f'toggle_{i}', label_visibility='collapsed')
    with cols[3]:  # This is where the bin icon button will go
        if st.button("🗑️", key=f'delete_{i}'):
            delete_pair(i)

# Button to add a new pair
st.button("Add new pair", on_click=add_pair)

def find_best_match(row, df_excel_file1, pairs):
    best_score = 0
    best_match = None

    for _, excel1_row in df_excel_file1.iterrows():
        current_score = 0
        num_none_toogle_pairs = 0
        for pair in pairs:
            col_excel1 = pair['text1'].strip()
            col_excel2 = pair['text2'].strip()

            if pair['toggle']:  # Only consider pairs with an enabled checkbox
                if str(row[col_excel2]).strip().lower() == str(excel1_row[col_excel1]).strip().lower():
                    best_score = 100
                    best_match = excel1_row
                    return pd.Series([best_match[pair['text1']] for pair in pairs] + [best_score],
                         index=[pair['text2'] for pair in pairs] + ['Confidence Level'])
            else:
                # Calculate fuzzy match score
                score = fuzz.ratio(str(row[col_excel2]).lower(), str(excel1_row[col_excel1]).lower())
                current_score += score
                num_none_toogle_pairs += 1

        # Average the score based on the number of active pairs
        active_pairs = num_none_toogle_pairs
        if active_pairs > 0:
            current_score /= active_pairs

        if current_score > best_score:
            best_score = current_score
            best_match = excel1_row

    if best_match is not None:
        return pd.Series([best_match[pair['text1']] for pair in pairs] + [best_score],
                         index=[pair['text2'] for pair in pairs] + ['Confidence Level'])
    else:
        return pd.Series([None] * active_pairs + [best_score],
                         index=[pair['text2'] for pair in pairs] + ['Confidence Level'])

# Function to start the comparison
def start_compare(df_excel_file1, df_excel_file2, pairs):
    # Apply the function to each row of EXCEL2
    matched = df_excel_file2.apply(find_best_match, axis=1, args=(df_excel_file1, pairs))

    # Combine the results with EXCEL2
    # result = pd.concat([df_excel_file2, matched], axis=1)
    result = pd.concat([matched], axis=1)

    # Save to a new Excel file
    result.to_excel('MATCHES.xlsx', index=False)
    return result

# Function to generate and download the Excel file
def generate_and_download_excel(df_excel_file1, df_excel_file2, pairs):
    # Apply the comparison function and get the result
    comparison_result = start_compare(df_excel_file1, df_excel_file2, pairs)
    
    # Convert the DataFrame to Excel and use BytesIO as a buffer
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        comparison_result.to_excel(writer, index=False)
        writer.close()
    excel_data = output.getvalue()

    # Provide a download button and return the buffer to the download button
    st.download_button(
        label="📁 Download Excel File",
        data=excel_data,
        file_name="MATCHES.xlsx",
        mime="application/vnd.ms-excel"
    )

# Button to start the comparison in the Streamlit app
if st.button('Start Compare'):
    if excel_file1 is not None and excel_file2 is not None:
        # Start the comparison process
        df_excel_file1 = pd.read_excel(excel_file1, skiprows=1)
        df_excel_file2 = pd.read_excel(excel_file2)
        generate_and_download_excel(df_excel_file1, df_excel_file2, st.session_state.pairs)
    else:
        st.error("Please upload both Excel files and select at least one pair to compare.")