import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re
import io

# Page Configuration
st.set_page_config(page_title="UNIMY Programme Analytics", layout="wide", page_icon="🎓")

# --- CSS for Styling ---
st.markdown("""
<style>
    .metric-card {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
    }
    .stDataFrame {
        border: 1px solid #e0e0e0;
        border-radius: 5px;
    }
</style>
""", unsafe_allow_html=True)

# --- Helper Functions ---

def find_header_row(df, keywords):
    """Locates the row index containing specific keywords."""
    for idx, row in df.iterrows():
        row_str = row.astype(str).str.cat(sep=' ').lower()
        if all(k.lower() in row_str for k in keywords):
            return idx
    return -1

def clean_percentage(val):
    """Converts percentage strings or decimals to float 0-100."""
    try:
        if isinstance(val, str):
            val = val.replace('%', '')
        num = float(val)
        return num * 100 if num <= 1.0 else num
    except:
        return 0.0

def process_subject_file(uploaded_file):
    """
    Parses a single subject Excel file to extract:
    1. Assessment Config (Setup)
    2. CLO-PLO Mapping (Table 2)
    3. Student Marks (Table 1)
    Returns a list of student records with calculated PLO scores.
    """
    data_extraction = []
    
    try:
        xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        
        # 1. IDENTIFY SHEETS
        sheet_setup = next((s for s in xls if "Setup" in s), None)
        sheet_marks = next((s for s in xls if "Table 1" in s or "Marks" in s), None)
        sheet_clo   = next((s for s in xls if "Table 2" in s or "CLO" in s), None)
        
        if not (sheet_setup and sheet_marks and sheet_clo):
            return [] # Missing critical sheets

        # 2. EXTRACT COURSE INFO & SETUP
        df_setup = xls[sheet_setup]
        course_code = "Unknown"
        course_name = "Unknown"
        
        # Find Course Code/Name in Setup
        for r in range(min(10, len(df_setup))):
            row_str = df_setup.iloc[r].astype(str).str.cat(sep=' ')
            if "Course Code" in row_str:
                match = re.search(r'([A-Z]{4}\d{4})', row_str)
                if match: course_code = match.group(1)
            if "Course Name" in row_str:
                 parts = str(df_setup.iloc[r, 1])
                 if parts != "nan": course_name = parts

        # Parse Assessments from Setup
        assessments = {} # { 'Assignment 1': {'clo': 'CLO 1', 'weight': 20, 'full': 100} }
        
        setup_header_idx = find_header_row(df_setup, ["Assessment Name", "Weightage"])
        if setup_header_idx != -1:
            df_setup.columns = df_setup.iloc[setup_header_idx]
            df_setup_data = df_setup.iloc[setup_header_idx+1:].reset_index(drop=True)
            
            for _, row in df_setup_data.iterrows():
                a_name = str(row.get("Assessment Name", "")).strip()
                if a_name and a_name.lower() != "nan":
                    try:
                        weight = float(row.get("Weightage (%)", 0))
                        full_m = float(row.get("Full Marks", 100))
                        clo_tag = str(row.get("CLO Tag", "")).strip()
                        assessments[a_name] = {'clo': clo_tag, 'weight': weight, 'full': full_m}
                    except: pass

        # 3. EXTRACT CLO-PLO MAPPING
        df_clo = xls[sheet_clo]
        clo_plo_map = {} # { 'CLO 1': 'PLO 1', 'CLO 2': 'PLO 2' }
        
        clo_header_idx = find_header_row(df_clo, ["CLO", "PLO"])
        if clo_header_idx != -1:
            df_clo.columns = df_clo.iloc[clo_header_idx]
            df_clo_data = df_clo.iloc[clo_header_idx+1:].reset_index(drop=True)
            for _, row in df_clo_data.iterrows():
                c_tag = str(row.get("CLO", "")).strip()
                p_tag = str(row.get("PLO", "")).strip()
                if "CLO" in c_tag and "PLO" in p_tag:
                    clo_plo_map[c_tag] = p_tag

        # 4. EXTRACT STUDENT MARKS
        df_marks = xls[sheet_marks]
        marks_header_idx = find_header_row(df_marks, ["STUDENT ID", "STUDENT NAME"])
        
        if marks_header_idx != -1:
            # Clean header to handle duplicates
            raw_header = df_marks.iloc[marks_header_idx].astype(str).tolist()
            seen = {}
            clean_header = []
            for h in raw_header:
                if h in seen:
                    seen[h] += 1
                    clean_header.append(f"{h}_{seen[h]}")
                else:
                    seen[h] = 0
                    clean_header.append(h)
            
            df_marks.columns = clean_header
            df_marks_data = df_marks.iloc[marks_header_idx+1:].reset_index(drop=True)
            
            # Map assessment columns
            col_map = {} # { 'Assign 1': 'Assignment 1' }
            for col in df_marks.columns:
                for setup_name in assessments.keys():
                    if setup_name.lower() == col.lower():
                        col_map[col] = setup_name
                    elif setup_name.lower() in col.lower() and "total" not in col.lower():
                         col_map[col] = setup_name

            # Process Students
            for _, row in df_marks_data.iterrows():
                s_id = str(row.get("STUDENT ID", "")).strip()
                s_name = str(row.get("STUDENT NAME", "")).strip()
                
                if len(s_id) > 2 and s_id.lower() != "nan":
                    
                    plo_totals = {} 
                    
                    for col_name, setup_name in col_map.items():
                        try:
                            raw_mark = pd.to_numeric(row[col_name], errors='coerce')
                            if pd.isna(raw_mark): raw_mark = 0
                            
                            config = assessments[setup_name]
                            clo = config['clo']
                            weight = config['weight']
                            full_marks = config['full']
                            
                            target_plo = clo_plo_map.get(clo, "Unmapped")
                            
                            if "PLO" in target_plo:
                                if target_plo not in plo_totals:
                                    plo_totals[target_plo] = {'earned': 0, 'total_weight': 0}
                                
                                contribution = (raw_mark / full_marks) * weight
                                plo_totals[target_plo]['earned'] += contribution
                                plo_totals[target_plo]['total_weight'] += weight
                                
                        except Exception as e:
                            pass
                            
                    final_plos = {}
                    for plo, vals in plo_totals.items():
                        if vals['total_weight'] > 0:
                            final = (vals['earned'] / vals['total_weight']) * 100
                            final_plos[plo] = final
                    
                    data_extraction.append({
                        'Student ID': s_id,
                        'Student Name': s_name,
                        'Course Code': course_code,
                        'Course Name': course_name,
                        'PLO_Data': final_plos
                    })

    except Exception as e:
        print(f"Error processing {uploaded_file.name}: {e}")
        
    return data_extraction

# --- Main App Logic ---

with st.sidebar:
    st.header("UNIMY Analytics")
    st.info("Upload all individual subject Excel files (e.g., DMIM1012.xlsx) to generate the master report.")
    st.caption("© 2026 UNIMY Programme Assessment")

st.title("🎓 UNIMY Programme Analytics")
st.markdown("**Master Dashboard for Student PLO Attainment (All Semesters)**")

uploaded_files = st.file_uploader("Upload Course Excel Files", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    all_records = []
    
    with st.status("Processing files...", expanded=True) as status:
        for f in uploaded_files:
            st.write(f"Reading {f.name}...")
            records = process_subject_file(f)
            all_records.extend(records)
        status.update(label="Processing Complete!", state="complete", expanded=False)
        
    if not all_records:
        st.error("No valid student data found. Please check your Excel files.")
    else:
        df_master = pd.DataFrame(all_records)
        
        tab1, tab2 = st.tabs(["Student Scorecard", "Programme Heatmap"])
        
        with tab1:
            st.subheader("Student Individual Performance")
            
            student_list = df_master[['Student ID', 'Student Name']].drop_duplicates().sort_values('Student Name')
            student_options = student_list.apply(lambda x: f"{x['Student Name']} ({x['Student ID']})", axis=1).tolist()
            
            selected_student_str = st.selectbox("Select Student:", student_options)
            
            if selected_student_str:
                selected_id = selected_student_str.split('(')[-1].replace(')', '')
                student_data = df_master[df_master['Student ID'] == selected_id]
                
                # 1. Subject Breakdown Table
                st.write("#### Subject Performance")
                
                table_rows = []
                plo_columns = set()
                
                for _, row in student_data.iterrows():
                    row_dict = {
                        'Course Code': row['Course Code'],
                        'Course Name': row['Course Name']
                    }
                    row_dict.update(row['PLO_Data'])
                    plo_columns.update(row['PLO_Data'].keys())
                    table_rows.append(row_dict)
                
                sorted_plos = sorted(list(plo_columns), key=lambda x: int(x.split(' ')[-1]) if x.split(' ')[-1].isdigit() else 99)
                
                df_student_table = pd.DataFrame(table_rows)
                cols = ['Course Code', 'Course Name'] + sorted_plos
                for c in cols:
                    if c not in df_student_table.columns: df_student_table[c] = np.nan
                
                st.dataframe(df_student_table[cols].style.format("{:.1f}", subset=sorted_plos), use_container_width=True)
                
                # 2. Average PLO Chart
                st.write("#### Overall PLO Achievement")
                
                avg_plo_scores = {}
                for plo in sorted_plos:
                    avg_plo_scores[plo] = df_student_table[plo].mean()
                
                if avg_plo_scores:
                    categories = list(avg_plo_scores.keys())
                    values = list(avg_plo_scores.values())
                    values += values[:1]
                    
                    angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
                    angles += angles[:1]
                    
                    fig, ax = plt.subplots(figsize=(6, 6), subplot_kw=dict(polar=True))
                    ax.fill(angles, values, color='#4A90E2', alpha=0.3)
                    ax.plot(angles, values, color='#4A90E2', linewidth=2)
                    
                    target_vals = [50] * len(values)
                    ax.plot(angles, target_vals, color='#FF4B4B', linestyle='--', linewidth=1, label='Target (50%)')
                    
                    ax.set_yticklabels([])
                    ax.set_xticks(angles[:-1])
                    ax.set_xticklabels(categories)
                    ax.legend(loc='upper right', bbox_to_anchor=(1.1, 1.1))
                    
                    col_chart, col_metric = st.columns([1, 1])
                    with col_chart:
                        st.pyplot(fig)
                    with col_metric:
                        st.write("**PLO Summary:**")
                        for plo, val in avg_plo_scores.items():
                            status = "✅ Achieved" if val >= 50 else "❌ Below Target"
                            st.write(f"- **{plo}:** {val:.1f}% ({status})")

        with tab2:
            st.subheader("Cohort Analysis (Programme Heatmap)")
            st.info("This view aggregates data from ALL students across ALL semesters.")
            
            all_plo_values = {plo: [] for plo in sorted_plos}
            
            for _, row in df_master.iterrows():
                for plo, val in row['PLO_Data'].items():
                    if plo in all_plo_values:
                        all_plo_values[plo].append(val)
            
            cohort_avgs = {k: sum(v)/len(v) if v else 0 for k, v in all_plo_values.items()}
            cohort_mean = np.mean(list(cohort_avgs.values())) if cohort_avgs else 0
            
            c1, c2 = st.columns(2)
            c1.metric("Total Students Tracked", len(student_list))
            c2.metric("Cohort Average PLO", f"{cohort_mean:.1f}%")
            
            fig2, ax2 = plt.subplots(figsize=(10, 5))
            plos = list(cohort_avgs.keys())
            scores = list(cohort_avgs.values())
            
            bars = ax2.bar(plos, scores, color=['#4CAF50' if s >= 50 else '#F44336' for s in scores])
            ax2.axhline(y=50, color='black', linestyle='--', label='Target (50%)')
            ax2.set_ylabel("Average Score (%)")
            ax2.set_title("Cohort Average PLO Attainment")
            ax2.legend()
            
            st.pyplot(fig2)
            st.dataframe(pd.DataFrame([cohort_avgs], index=["Average"]).style.format("{:.1f}"), use_container_width=True)

else:
    st.info("Waiting for Excel files... Drag and drop them above.")