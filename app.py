import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill, NamedStyle
import copy

# ================= Page setup =================
st.set_page_config(page_title="Bin Divider Specification Generator", page_icon=":package:", layout="wide")

st.markdown(
    """
<style>
    .main .block-container { max-width: 900px; padding-left: 8%; padding-right: 8%; }
</style>
""",
    unsafe_allow_html=True,
)

st.title("Bin Divider Specification Generator")
st.write("Define a central Bin Box Library once, then assign bin types to groups and export a clean Excel spec.")

# ================= Session state =================
if "groups" not in st.session_state:
    st.session_state.groups = []
if "bin_library" not in st.session_state:
    st.session_state.bin_library = {}  # {bin_id: {fields}}
if "next_bin_id" not in st.session_state:
    st.session_state.next_bin_id = 1   # stable unique IDs for bins

# ================= Helpers =================
EXPORT_COLUMNS = [
    'Group Name', 'Floor', 'Mod', 'Depth', 'Start Aisle', 'End Aisle', '# of Aisles', '# of Bays',
    'Total # of Shelves per Bay', 'Bay Design', 'Bin Box Type', 'Depth (mm)',
    'Height (mm)', 'Width (mm)', 'Lip (cm)', '# of Shelves per Bay',
    'Qty bins per Shelf', 'Qty Per Bay', 'Total Quantity', 'UT',
    'Bin Gross CBM', 'Bin Net CBM'
]

def safe_float(x, default=0.0):
    try:
        return float(x)
    except Exception:
        return default


def calculate_fields(group_data: dict, bin_data: dict) -> dict:
    g = group_data
    b = bin_data
    out = {**b}
    shelves_per_bay = int(safe_float(b.get('# of Shelves per Bay', 1), 1))
    qty_per_shelf = int(safe_float(b.get('Qty bins per Shelf', 1), 1))
    ut = safe_float(b.get('UT', 0.0), 0.0)
    depth_mm = safe_float(b.get('Depth (mm)', 0.0), 0.0)
    height_mm = safe_float(b.get('Height (mm)', 0.0), 0.0)
    width_mm = safe_float(b.get('Width (mm)', 0.0), 0.0)
    lip_cm = safe_float(b.get('Lip (cm)', 0.0), 0.0)
    start_aisle = int(safe_float(g.get('Start Aisle', 1), 1))
    end_aisle = int(safe_float(g.get('End Aisle', 1), 1))
    bays = int(safe_float(g.get('# of Bays', 1), 1))

    out['Lip (cm)'] = '-' if lip_cm == 0 else lip_cm
    out['# of Aisles'] = end_aisle - start_aisle + 1
    out['Qty Per Bay'] = shelves_per_bay * qty_per_shelf
    out['Total Quantity'] = out['Qty Per Bay'] * bays
    out['Bin Gross CBM'] = (depth_mm * height_mm * width_mm) / 1_000_000
    out['Bin Net CBM'] = out['Bin Gross CBM'] * ut
    return out


def sync_bin_keys_with_library():
    valid_ids = set(st.session_state.bin_library.keys())
    for grp in st.session_state.groups:
        grp['bin_keys'] = [k for k in grp.get('bin_keys', []) if k in valid_ids]

# Safe rerun function
def rerun():
    try:
        st.rerun()
    except AttributeError:
        from streamlit import experimental_rerun
        experimental_rerun()

# ================= Excel export =================

def generate_excel(groups: list) -> bytes:
    df = pd.DataFrame(columns=EXPORT_COLUMNS)
    group_row_counts = []

    for group in groups:
        group_data = group['group_data']
        bin_keys = group.get('bin_keys', [])
        rows_added = 0
        for k in bin_keys:
            if k not in st.session_state.bin_library:
                continue
            base_bin = st.session_state.bin_library[k]
            calc = calculate_fields(group_data, base_bin)
            row = {**group_data, **calc}
            full_row = {c: row.get(c, None) for c in EXPORT_COLUMNS}
            df = pd.concat([df, pd.DataFrame([full_row])], ignore_index=True)
            rows_added += 1
        group_row_counts.append(rows_added)

    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Bin Box"

    for r_idx, r in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(r)
        if r_idx == 1:
            for cell in ws[r_idx]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    ws.freeze_panes = "A2"

    current_row = 2
    for row_count in group_row_counts:
        if row_count > 0:
            for col_idx in range(1, 10):
                ws.merge_cells(start_row=current_row, start_column=col_idx, end_row=current_row + row_count - 1, end_column=col_idx)
                ws.cell(row=current_row, column=col_idx).alignment = Alignment(horizontal='center', vertical='center')
            current_row += row_count

    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ================= Bin Library UI =================
st.subheader("Bin Box Library")
if st.button("➕ Add New Bin Type"):
    new_id = f"bin_{st.session_state.next_bin_id}"
    st.session_state.next_bin_id += 1
    st.session_state.bin_library[new_id] = {
        'Bin Box Type': '',
        'Depth (mm)': 0.0,
        'Height (mm)': 0.0,
        'Width (mm)': 0.0,
        'Lip (cm)': 0.0,
        '# of Shelves per Bay': 1,
        'Qty bins per Shelf': 1,
        'UT': 0.0,
    }
    rerun()

for bin_id, data in list(st.session_state.bin_library.items()):
    label = data.get('Bin Box Type') or bin_id
    with st.expander(f"{label}", expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            data['Bin Box Type'] = st.text_input("Bin Box Type", value=data.get('Bin Box Type', ''), key=f"name_{bin_id}")
            data['Depth (mm)'] = st.number_input("Depth (mm)", min_value=0.0, value=safe_float(data.get('Depth (mm)', 0.0)), step=0.1, key=f"depth_{bin_id}")
            data['Height (mm)'] = st.number_input("Height (mm)", min_value=0.0, value=safe_float(data.get('Height (mm)', 0.0)), step=0.1, key=f"height_{bin_id}")
            has_lip = st.checkbox("Has Lip?", value=safe_float(data.get('Lip (cm)', 0.0)) > 0, key=f"lip_chk_{bin_id}")
        with c2:
            data['Width (mm)'] = st.number_input("Width (mm)", min_value=0.0, value=safe_float(data.get('Width (mm)', 0.0)), step=0.1, key=f"width_{bin_id}")
            data['Lip (cm)'] = (safe_float(data['Height (mm)']) * 0.2 / 10) if has_lip else 0.0
            st.number_input("Lip (cm)", value=safe_float(data['Lip (cm)']), disabled=True, key=f"lip_val_{bin_id}")
            data['# of Shelves per Bay'] = int(st.number_input("# of Shelves per Bay", min_value=1, value=int(safe_float(data.get('# of Shelves per Bay', 1))), step=1, key=f"shelves_{bin_id}"))
            data['Qty bins per Shelf'] = int(st.number_input("Qty bins per Shelf", min_value=1, value=int(safe_float(data.get('Qty bins per Shelf', 1))), step=1, key=f"qty_{bin_id}"))
            data['UT'] = st.number_input("UT (0-1)", min_value=0.0, max_value=1.0, value=min(max(safe_float(data.get('UT', 0.0)), 0.0), 1.0), step=0.01, key=f"ut_{bin_id}")

        if st.button(f"🗑️ Delete {label}", key=f"del_{bin_id}"):
            del st.session_state.bin_library[bin_id]
            sync_bin_keys_with_library()
            rerun()

# Ensure groups are valid
sync_bin_keys_with_library()

# ================= Groups UI =================
st.subheader("Manage Groups")
if st.button("➕ Add New Group"):
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
        'finalized': False,
    })
    rerun()

for group_idx, group in enumerate(st.session_state.groups):
    with st.expander(f"Group {group_idx + 1}: {group['group_data']['Group Name'] or 'Untitled'} ({'Finalized' if group['finalized'] else 'Editing'})", expanded=not group['finalized']):
        if not group['finalized']:
            c1, c2 = st.columns(2)
            with c1:
                group['group_data']['Group Name'] = st.text_input("Group Name", value=group['group_data']['Group Name'], key=f"gname_{group_idx}")
                group['group_data']['Floor'] = st.text_input("Floor", value=group['group_data']['Floor'], key=f"gflr_{group_idx}")
                group['group_data']['Mod'] = st.text_input("Mod", value=group['group_data']['Mod'], key=f"gmod_{group_idx}")
                group['group_data']['Depth'] = st.text_input("Depth", value=group['group_data']['Depth'], key=f"gdepth_{group_idx}")
            with c2:
                group['group_data']['Start Aisle'] = int(st.number_input("Start Aisle", min_value=1, value=int(safe_float(group['group_data']['Start Aisle'], 1)), step=1, key=f"gstart_{group_idx}"))
                group['group_data']['End Aisle'] = int(st.number_input("End Aisle", min_value=1, value=int(safe_float(group['group_data']['End Aisle'], 1)), step=1, key=f"gend_{group_idx}"))
                group['group_data']['# of Bays'] = int(st.number_input("# of Bays", min_value=1, value=int(safe_float(group['group_data']['# of Bays'], 1)), step=1, key=f"gbays_{group_idx}"))
                group['group_data']['Total # of Shelves per Bay'] = int(st.number_input("Total # of Shelves per Bay", min_value=1, value=int(safe_float(group['group_data']['Total # of Shelves per Bay'], 1)), step=1, key=f"gshelves_{group_idx}"))
                group['group_data']['Bay Design'] = st.text_input("Bay Design", value=group['group_data']['Bay Design'], key=f"gbay_{group_idx}")

            available_ids = list(st.session_state.bin_library.keys())
            labels = {k: st.session_state.bin_library[k].get('Bin Box Type') or k for k in available_ids}
            default_vals = [k for k in group.get('bin_keys', []) if k in available_ids]
            group['bin_keys'] = st.multiselect("Select Bin Box Types for this Group", options=available_ids, default=default_vals, format_func=lambda k: labels[k], key=f"binsel_{group_idx}")

            if st.button(f"Finalize Group {group_idx + 1}", key=f"gfin_{group_idx}"):
                group['finalized'] = True
                rerun()
            if st.button(f"Delete Group {group_idx + 1}", key=f"gdel_{group_idx}"):
                st.session_state.groups.pop(group_idx)
                rerun()
        else:
            if st.button(f"Edit Group {group_idx + 1}", key=f"gedit_{group_idx}"):
                group['finalized'] = False
                rerun()

# ================= Download Excel =================
if st.session_state.groups:
    if not st.session_state.bin_library:
        st.warning("Add at least one bin type before exporting.")
    else:
        excel_data = generate_excel(st.session_state.groups)
        st.download_button(
            label="⬇️ Download Excel File",
            data=excel_data,
            file_name="Bin_Divider_Specs.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ================= Clear all =================
if st.button("Clear All Data"):
    st.session_state.groups = []
    st.session_state.bin_library = {}
    st.session_state.next_bin_id = 1
    rerun()
