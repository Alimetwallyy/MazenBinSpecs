import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill
import copy

# Set page configuration
st.set_page_config(page_title="Bin Divider Specification Generator", page_icon=":package:", layout="wide")

# Custom CSS to constrain app width
st.markdown("""
<style>
    .main .block-container {
        max-width: 800px;
        padding-left: 10%;
        padding-right: 10%;
    }
</style>
""", unsafe_allow_html=True)

# Title
st.title("Bin Divider Specification Generator")
st.write("Define a library of bin box types, then assign them to groups to generate an Excel file.")

# Initialize session state
if 'groups' not in st.session_state:
    st.session_state.groups = []
if 'bin_library' not in st.session_state:
    st.session_state.bin_library = {}

# Function to calculate derived fields (for app logic)
def calculate_fields(group_data, bin_data):
    bin_data = bin_data.copy()
    bin_data['# of Aisles'] = group_data['End Aisle'] - group_data['Start Aisle'] + 1
    bin_data['Qty Per Bay'] = bin_data['# of Shelves per Bay'] * bin_data['Qty bins per Shelf']
    bin_data['Total Quantity'] = bin_data['Qty Per Bay'] * group_data['# of Bays']
    bin_data['Bin Gross CBM'] = (bin_data['Depth (mm)'] * bin_data['Height (mm)'] * bin_data['Width (mm)']) / 1_000_000
    bin_data['Bin Net CBM'] = bin_data['Bin Gross CBM'] * bin_data['UT']
    return bin_data

# Function to generate Excel file
def generate_excel(groups):
    columns = [
        'Group Name', 'Floor', 'Mod', 'Depth', 'Start Aisle', 'End Aisle', '# of Aisles', '# of Bays',
        'Total # of Shelves per Bay', 'Bay Design', 'Bin Box Type', 'Depth (mm)',
        'Height (mm)', 'Width (mm)', 'Lip (cm)', '# of Shelves per Bay',
        'Qty bins per Shelf', 'Qty Per Bay', 'Total Quantity', 'UT',
        'Bin Gross CBM', 'Bin Net CBM'
    ]
    df = pd.DataFrame(columns=columns)
    group_row_counts = []
    for group in groups:
        group_data = group['group_data']
        bin_keys = group.get('bin_keys', [])
        bin_rows = [st.session_state.bin_library[k] for k in bin_keys]
        for bin_data in bin_rows:
            calculated = calculate_fields(group_data, bin_data)
            row = {**group_data, **calculated}
            row['Lip (cm)'] = '-' if row.get('Lip (cm)', 0.0) == 0.0 else row.get('Lip (cm)', 0.0)
            for col in columns:
                if col not in row:
                    row[col] = None
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        group_row_counts.append(len(bin_rows))

    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Bin Box"

    # Write DataFrame to Excel
    for r_idx, r in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(r)
        if r_idx == 1:  # Header row formatting
            for cell in ws[r_idx]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    # Merge and center cells for columns A to I
    current_row = 2
    for row_count in group_row_counts:
        if row_count > 0:
            for col_idx in range(1, 10):
                ws.merge_cells(start_row=current_row, start_column=col_idx, end_row=current_row + row_count - 1, end_column=col_idx)
                ws.cell(row=current_row, column=col_idx).alignment = Alignment(horizontal='center')
            current_row += row_count

    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ================= Bin Library Section =================
st.subheader("Bin Box Library")
if st.button("Add New Bin Type"):
    new_key = f"bin_{len(st.session_state.bin_library) + 1}"
    st.session_state.bin_library[new_key] = {
        'Bin Box Type': '',
        'Depth (mm)': 0.0,
        'Height (mm)': 0.0,
        'Width (mm)': 0.0,
        'Lip (cm)': 0.0,
        '# of Shelves per Bay': 1,
        'Qty bins per Shelf': 1,
        'UT': 0.0
    }

for key, bin_data in list(st.session_state.bin_library.items()):
    with st.expander(f"{bin_data['Bin Box Type'] or key}"):
        cols_bin = st.columns(2)
        with cols_bin[0]:
            bin_data['Bin Box Type'] = st.text_input("Bin Box Type", value=bin_data.get('Bin Box Type', ''), key=f"bin_type_{key}")
            bin_data['Depth (mm)'] = st.number_input("Depth (mm)", min_value=0.0, value=bin_data.get('Depth (mm)', 0.0), step=0.1, key=f"depth_mm_{key}")
            bin_data['Height (mm)'] = st.number_input("Height (mm)", min_value=0.0, value=bin_data.get('Height (mm)', 0.0), step=0.1, key=f"height_mm_{key}")
            has_lip = st.checkbox("Has Lip?", value=bin_data.get('Lip (cm)', 0) > 0, key=f"has_lip_{key}")
        with cols_bin[1]:
            bin_data['Width (mm)'] = st.number_input("Width (mm)", min_value=0.0, value=bin_data.get('Width (mm)', 0.0), step=0.1, key=f"width_mm_{key}")
            bin_data['Lip (cm)'] = (bin_data['Height (mm)'] * 0.2 / 10) if has_lip else 0.0
            st.number_input("Lip (cm)", value=bin_data['Lip (cm)'], disabled=True, key=f"lip_cm_{key}")
            bin_data['# of Shelves per Bay'] = st.number_input("# of Shelves per Bay", min_value=1, value=bin_data.get('# of Shelves per Bay', 1), step=1, key=f"shelves_per_bay_{key}")
            bin_data['Qty bins per Shelf'] = st.number_input("Qty bins per Shelf", min_value=1, value=bin_data.get('Qty bins per Shelf', 1), step=1, key=f"qty_bins_{key}")
            bin_data['UT'] = st.number_input("UT", min_value=0.0, max_value=1.0, value=bin_data.get('UT', 0.0), step=0.01, key=f"ut_{key}")

        # Delete option
        if st.button(f"Delete {bin_data['Bin Box Type'] or key}", key=f"delete_{key}"):
            del st.session_state.bin_library[key]
            st.experimental_rerun()

# ================= Groups Section =================
st.subheader("Manage Groups")
if st.button("Add New Group"):
    st.session_state.groups.append({
        'group_data': {
            'Group Name': '',
            'Floor': '',
            'Mod': '',
            'Depth': '',
            'Start Aisle': 1,
            'End Aisle': 1,
            '# of Bays': 1,
            'Total # of Shelves per Bay': 1,
            'Bay Design': ''
        },
        'bin_keys': [],
        'finalized': False
    })

# Display and edit groups
for group_idx, group in enumerate(st.session_state.groups):
    is_new_copy = group_idx == len(st.session_state.groups) - 1 and st.session_state.get('last_action') == f"copy_{group_idx-1}"
    with st.expander(f"Group {group_idx + 1}: {group['group_data']['Group Name'] or 'Untitled'} ({'Finalized' if group['finalized'] else 'Editing'})", expanded=not group['finalized'] or is_new_copy):
        if not group['finalized']:
            # Group inputs
            st.write("**Group Details**")
            cols = st.columns(2)
            with cols[0]:
                group['group_data']['Group Name'] = st.text_input("Group Name", value=group['group_data']['Group Name'], key=f"group_name_{group_idx}")
                group['group_data']['Floor'] = st.text_input("Floor", value=group['group_data']['Floor'], key=f"floor_{group_idx}")
                group['group_data']['Mod'] = st.text_input("Mod", value=group['group_data']['Mod'], key=f"mod_{group_idx}")
                group['group_data']['Depth'] = st.text_input("Depth", value=group['group_data']['Depth'], key=f"depth_{group_idx}")
            with cols[1]:
                group['group_data']['Start Aisle'] = st.number_input("Start Aisle", min_value=1, value=int(group['group_data']['Start Aisle']), step=1, key=f"start_aisle_{group_idx}")
                group['group_data']['End Aisle'] = st.number_input("End Aisle", min_value=1, value=int(group['group_data']['End Aisle']), step=1, key=f"end_aisle_{group_idx}")
                group['group_data']['# of Bays'] = st.number_input("# of Bays", min_value=1, value=int(group['group_data']['# of Bays']), step=1, key=f"bays_{group_idx}")
                group['group_data']['Total # of Shelves per Bay'] = st.number_input("Total # of Shelves per Bay", min_value=1, value=int(group['group_data']['Total # of Shelves per Bay']), step=1, key=f"shelves_bay_{group_idx}")
                group['group_data']['Bay Design'] = st.text_input("Bay Design", value=group['group_data']['Bay Design'], key=f"bay_design_{group_idx}")

            # Select bin types from library
            available_bins = {k: v['Bin Box Type'] or k for k, v in st.session_state.bin_library.items()}
            group['bin_keys'] = st.multiselect("Select Bin Box Types for this Group", options=list(available_bins.keys()), format_func=lambda k: available_bins[k], default=group['bin_keys'], key=f"bin_select_{group_idx}")

            if st.button(f"Finalize Group {group_idx + 1}", key=f"finalize_{group_idx}"):
                group['finalized'] = True
                st.success(f"Group {group_idx + 1} finalized!")
                st.experimental_rerun()
        else:
            if st.button(f"Edit Group {group_idx + 1}", key=f"edit_{group_idx}"):
                group['finalized'] = False
                st.success(f"Group {group_idx + 1} opened for editing!")
                st.experimental_rerun()

# Display finalized groups
if st.session_state.groups:
    st.subheader("All Groups")
    for i, group in enumerate(st.session_state.groups):
        with st.expander(f"Group {i + 1}: {group['group_data']['Group Name'] or 'Untitled'} ({'Finalized' if group['finalized'] else 'Editing'})"):
            st.write("**Group Details**")
            st.json(group['group_data'])
            if group['bin_keys']:
                st.write("**Selected Bin Box Types (with calculations)**")
                for key in group['bin_keys']:
                    calculated = calculate_fields(group['group_data'], st.session_state.bin_library[key])
                    st.json(calculated)
            if st.button(f"Copy Group {i + 1}", key=f"copy_{i}"):
                new_group = copy.deepcopy(group)
                new_group['finalized'] = False
                new_group['group_data']['Group Name'] = f"{new_group['group_data']['Group Name']} (Copy)" if new_group['group_data']['Group Name'] else "Untitled (Copy)"
                st.session_state.groups.append(new_group)
                st.session_state['last_action'] = f"copy_{i}"
                st.success(f"Group {i + 1} copied!")
                st.experimental_rerun()

# Download Excel file
if st.session_state.groups:
    excel_data = generate_excel(st.session_state.groups)
    st.download_button(
        label="Download Excel File",
        data=excel_data,
        file_name="Bin_Divider_Specs.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Clear all data
if st.button("Clear All Data"):
    st.session_state.groups = []
    st.session_state.bin_library = {}
    st.experimental_rerun()
