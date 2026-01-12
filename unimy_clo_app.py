import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import io
import re

# Page Configuration
st.set_page_config(page_title="UNIMY CLO Analytics", layout="wide", page_icon="📊")

# --- CSS for Styling ---
st.markdown("""
<style>
    .metric-card {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
    }
    .report-box {
        background-color: #ffffff;
        padding: 20px;
        border: 1px solid #d0d0d0;
        border-radius: 5px;
        font-family: 'Times New Roman', serif;
        white-space: pre-wrap;
    }
    .success-box {
        padding: 10px;
        background-color: #d4edda;
        color: #155724;
        border-radius: 5px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# --- Helper Functions ---

def find_header_row(df, keywords):
    """Finds the row index containing specific keywords."""
    for idx, row in df.iterrows():
        row_str = row.astype(str).str.cat(sep=' ').lower()
        if all(k.lower() in row_str for k in keywords):
            return idx
    return -1

def get_smart_recommendation(clo_name, failure_rate):
    """Generates CQI actions based on failure rate and context."""
    if failure_rate < 15:
        return "Maintain current teaching methods."
    
    context = str(clo_name).lower()
    if "drawing" in context or "sketch" in context:
        return "Conduct extra studio sessions with live demonstrations."
    elif "calculation" in context or "math" in context:
        return "Provide remedial drills focusing on step-by-step methods."
    elif "software" in context or "tool" in context:
        return "Organize extra lab tutorials for software proficiency."
    elif "theory" in context or "history" in context:
        return "Use visual aids and mind maps to summarize key concepts."
    else:
        return "Review assessment difficulty and conduct revision classes."

def parse_campusone_file(uploaded_file):
    """
    Parses the specific 'UNIMY - Coursework Result' format.
    Returns: DataFrame of students, Course Info Dict, List of Potential Assessment Cols
    """
    try:
        df = pd.read_excel(uploaded_file, header=None)
        
        # 1. Extract Course Info (Metadata is usually in top rows)
        course_info = {"code": "Unknown", "name": "Unknown", "semester": "Unknown", "lecturer": "Unknown"}
        
        for r in range(min(15, len(df))):
            row_str = df.iloc[r].astype(str).str.cat(sep=' ')
            if "Subject" in row_str and ":" in row_str:
                parts = row_str.split(":", 1)[1].strip()
                if "-" in parts:
                    course_info["code"] = parts.split("-")[0].strip()
                    course_info["name"] = parts.split("-", 1)[1].strip()
                else:
                    course_info["name"] = parts
            if "Semester" in row_str and ":" in row_str:
                course_info["semester"] = row_str.split(":", 1)[1].strip()
            if "Lecturer" in row_str and ":" in row_str:
                course_info["lecturer"] = row_str.split(":", 1)[1].strip()

        # 2. Find Header Row
        header_idx = find_header_row(df, ["Student No.", "Student Name"])
        
        if header_idx == -1:
            return None, None, "Could not find header row (Student No., Student Name)"

        # 3. Set Header & Extract Data
        df.columns = df.iloc[header_idx]
        data_df = df.iloc[header_idx+1:].reset_index(drop=True)
        
        # Filter strictly for student rows
        data_df = data_df[pd.to_numeric(data_df.iloc[:, 0], errors='coerce').notna()]
        
        # Identify Potential Assessment Columns
        all_cols = list(data_df.columns)
        potential_assessments = []
        
        ignore_cols = ["No", "Student No.", "Student Name", "Enrollment No.", "Programme", "Intake", "Assessment Mark", "Total Mark", "Grade", "Point", "Result", "Hold", "Note"]
        
        for c in all_cols:
            c_str = str(c).strip()
            if c_str not in ignore_cols and "Unnamed" not in c_str:
                potential_assessments.append(c_str)
                
        return data_df, course_info, potential_assessments

    except Exception as e:
        return None, None, str(e)

def generate_evidence_excel(course_info, student_results, clo_stats, plo_stats, crr_data):
    """
    Generates an Excel file mimicking the 'Master Template' for audit evidence.
    Includes Setup, Table 1, Table 2, Table 3, and CRR sheets.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet: Setup
        setup_data = {
            'Field': ['Course Name', 'Course Code', 'Semester', 'Lecturer Name'],
            'Value': [course_info['name'], course_info['code'], course_info['semester'], course_info['lecturer']]
        }
        pd.DataFrame(setup_data).to_excel(writer, sheet_name='Setup', index=False)
        
        # Sheet: Table 1 - Marks
        # Flatten the dicts
        flat_results = []
        for res in student_results:
            row = {'Student ID': res['id'], 'Student Name': res['name']}
            # Add scores
            for k, v in res.items():
                if k not in ['id', 'name', 'raw_marks']:
                    row[k] = v
            flat_results.append(row)
        pd.DataFrame(flat_results).to_excel(writer, sheet_name='Table 1 - Marks', index=False)
        
        # Sheet: Table 2 - CLO Analysis
        pd.DataFrame(clo_stats).to_excel(writer, sheet_name='Table 2 - CLO Analysis', index=False)
        
        # Sheet: Table 3 - PLO Analysis
        pd.DataFrame(plo_stats).to_excel(writer, sheet_name='Table 3 - PLO Analysis', index=False)
        
        # Sheet: CRR (Audit Report)
        crr_rows = [
            {'Section': '1. COURSE PERFORMANCE', 'Detail': '', 'Value': ''},
            {'Section': 'Pass Rate', 'Detail': '', 'Value': f"{crr_data['pass_rate']:.2f}%"},
            {'Section': 'Average GPA', 'Detail': '', 'Value': f"{crr_data['gpa']:.2f}"},
            {'Section': '2. CLO ATTAINMENT', 'Detail': '', 'Value': ''}
        ]
        for c in clo_stats:
            crr_rows.append({'Section': c['CLO'], 'Detail': 'Attainment %', 'Value': f"{c['Average (%)']:.2f}"})
            
        pd.DataFrame(crr_rows).to_excel(writer, sheet_name='CRR (Audit Report)', index=False)
        
    return output.getvalue()

def calculate_gpa(grade):
    """Rough mapping of Grade to Point."""
    grades = {'A+': 4.0, 'A': 4.0, 'A-': 3.67, 'B+': 3.33, 'B': 3.0, 'B-': 2.67, 'C+': 2.33, 'C': 2.0, 'D': 1.0, 'F': 0.0}
    return grades.get(str(grade).strip().upper(), 0.0)

# --- Main Logic ---

def process_raw_data(data_df, config_map, plo_mapping):
    """
    Calculates stats based on user-defined mapping.
    """
    student_results = []
    total_gpa = 0
    count_gpa = 0
    
    for _, row in data_df.iterrows():
        s_id = str(row.get("Student No.", "")).strip()
        s_name = str(row.get("Student Name", "")).strip()
        grade = str(row.get("Grade", "")).strip()
        
        if grade:
            total_gpa += calculate_gpa(grade)
            count_gpa += 1
        
        student_clos = {}
        
        for col_name, config in config_map.items():
            try:
                val = row.get(col_name)
                raw_score = pd.to_numeric(val, errors='coerce')
                if pd.isna(raw_score): raw_score = 0
                
                contrib = (raw_score / config['full']) * config['weight']
                clo = config['clo']
                
                if clo not in student_clos: student_clos[clo] = {'earned': 0, 'weight': 0}
                student_clos[clo]['earned'] += contrib
                student_clos[clo]['weight'] += config['weight']
            except: pass
            
        res = {'id': s_id, 'name': s_name, 'Grade': grade}
        total_earned = 0
        
        for clo, vals in student_clos.items():
            if vals['weight'] > 0:
                perc = (vals['earned'] / vals['weight']) * 100
                res[clo] = perc
                total_earned += vals['earned']
        
        res['Total'] = total_earned
        student_results.append(res)
        
    avg_gpa = total_gpa / count_gpa if count_gpa > 0 else 0
    return pd.DataFrame(student_results), avg_gpa

# --- UI START ---

with st.sidebar:
    st.header("UNIMY Analytics")
    mode = st.radio("Mode:", ["Upload CampusOne (Raw)", "Direct Entry"])
    st.caption("© 2026 UNIMY Programme Assessment")

st.title("📊 UNIMY CLO Analytics (Complete Suite)")

if mode == "Upload CampusOne (Raw)":
    st.markdown("**Step 1: Upload Raw Export**")
    uploaded_file = st.file_uploader("Upload 'UNIMY - Coursework Result.xlsx'", type=['xlsx'])
    
    if uploaded_file:
        data_df, info, potential_cols = parse_campusone_file(uploaded_file)
        
        if data_df is not None:
            st.success(f"Loaded: {info['code']} - {info['name']} ({len(data_df)} Students)")
            
            # --- Step 2: Configuration (Enhanced) ---
            st.markdown("### Step 2: Assessment & PLO Mapping")
            st.info("Map columns to CLOs, Categories (CA/FP/FE), and link CLOs to PLOs.")
            
            with st.form("mapping_form"):
                # 1. Assessment Config
                st.subheader("1. Assessment Configuration")
                configs = {}
                cols = st.columns(3)
                
                for i, col_name in enumerate(potential_cols):
                    with cols[i % 3]:
                        st.markdown(f"**{col_name}**")
                        use = st.checkbox(f"Include {col_name}", value=True, key=f"use_{i}")
                        if use:
                            clo = st.selectbox("CLO Tag", ["CLO 1", "CLO 2", "CLO 3", "CLO 4"], key=f"clo_{i}")
                            cat = st.selectbox("Category", ["CA", "FP", "FE"], key=f"cat_{i}")
                            weight = st.number_input("Weight (%)", value=20.0, key=f"w_{i}")
                            full = st.number_input("Full Marks", value=100.0, key=f"f_{i}")
                            configs[col_name] = {'clo': clo, 'cat': cat, 'weight': weight, 'full': full}
                
                st.divider()
                
                # 2. PLO Mapping
                st.subheader("2. CLO -> PLO Mapping")
                st.caption("Link each CLO to a Programme Learning Outcome.")
                c1, c2, c3, c4 = st.columns(4)
                plo_map = {}
                with c1: plo_map["CLO 1"] = st.selectbox("CLO 1 maps to:", [f"PLO {i}" for i in range(1, 13)])
                with c2: plo_map["CLO 2"] = st.selectbox("CLO 2 maps to:", [f"PLO {i}" for i in range(1, 13)])
                with c3: plo_map["CLO 3"] = st.selectbox("CLO 3 maps to:", [f"PLO {i}" for i in range(1, 13)])
                with c4: plo_map["CLO 4"] = st.selectbox("CLO 4 maps to:", [f"PLO {i}" for i in range(1, 13)])
                
                analyze_btn = st.form_submit_button("Analyze & Generate Full Report")
            
            # --- Step 3: Analysis ---
            if analyze_btn and configs:
                df_res, avg_gpa = process_raw_data(data_df, configs, plo_map)
                
                if not df_res.empty:
                    # Metrics
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Course", info['code'])
                    c2.metric("Lecturer", info['lecturer'])
                    c3.metric("Students", len(df_res))
                    pass_rate = (len(df_res[df_res['Total'] >= 50]) / len(df_res)) * 100
                    c4.metric("Pass Rate / GPA", f"{pass_rate:.1f}% / {avg_gpa:.2f}")
                    
                    # Tabs for Full Report
                    t1, t2, t3, t4, t5 = st.tabs(["Table 1 (Marks)", "Table 2 (CLO)", "Table 3 (PLO)", "CRR (Audit)", "Dashboard"])
                    
                    clo_cols = sorted([c for c in df_res.columns if "CLO" in c])
                    
                    # --- Table 1: Marks ---
                    with t1:
                        st.subheader("Table 1: Student Marks & CLO Scores")
                        st.dataframe(df_res.style.format("{:.1f}", subset=clo_cols+['Total']).map(lambda v: 'color: red;' if v < 50 else 'color: green;', subset=clo_cols+['Total']), use_container_width=True)
                        
                    # --- Table 2: CLO Analysis (Enhanced) ---
                    with t2:
                        st.subheader("Table 2: Analysis of CLO (CQI)")
                        clo_stats = []
                        for clo in clo_cols:
                            avg = df_res[clo].mean()
                            pass_cnt = len(df_res[df_res[clo] >= 50])
                            p_rate = (pass_cnt / len(df_res)) * 100
                            status = "YES" if avg >= 50 else "NO"
                            
                            # Smart Suggestion Logic
                            issue = ""
                            action = ""
                            audit_cat = ""
                            if status == "NO":
                                issue = "Low Attainment"
                                action = get_smart_recommendation(info['name'], 100-p_rate)
                                audit_cat = "Area 1: Syllabus/Content"
                                
                            clo_stats.append({
                                "CLO": clo,
                                "PLO": plo_map.get(clo, "-"),
                                "Overall %": avg,
                                "Pass/Fail": status,
                                "Issue": issue,
                                "Suggestion": action,
                                "Audit Category": audit_cat
                            })
                        
                        df_stats = pd.DataFrame(clo_stats)
                        
                        # Editable Table for Manual Override
                        edited_clo_stats = st.data_editor(df_stats, num_rows="fixed", use_container_width=True)
                        
                        fig, ax = plt.subplots(figsize=(8, 3))
                        ax.bar(edited_clo_stats['CLO'], edited_clo_stats['Overall %'], color=['#4CAF50' if x>=50 else '#F44336' for x in edited_clo_stats['Overall %']])
                        ax.axhline(50, color='black', linestyle='--')
                        st.pyplot(fig)

                    # --- Table 3: PLO Analysis ---
                    with t3:
                        st.subheader("Table 3: PLO Attainment Analysis")
                        # Map CLO scores to PLO
                        plo_data = []
                        for _, row in edited_clo_stats.iterrows():
                            plo_data.append({
                                "CLO": row['CLO'],
                                "Mapped PLO": row['PLO'],
                                "Score (%)": row['Overall %'],
                                "Status": "PASS" if row['Overall %'] >= 50 else "NO"
                            })
                        st.dataframe(pd.DataFrame(plo_data), use_container_width=True)

                    # --- CRR: Audit Report ---
                    with t4:
                        st.subheader("Course Review Report (CRR)")
                        st.markdown(f"""
                        **1. COURSE PERFORMANCE**
                        * **Pass Rate:** {pass_rate:.1f}%
                        * **Average GPA:** {avg_gpa:.2f}
                        
                        **2. CLO ATTAINMENT ANALYSIS**
                        """)
                        st.table(edited_clo_stats[['CLO', 'Overall %', 'Pass/Fail']])
                        
                        st.markdown("**3. CQI ACTION PLAN**")
                        cqi_plan = edited_clo_stats[edited_clo_stats['Pass/Fail'] == "NO"][['CLO', 'Audit Category', 'Suggestion']]
                        if not cqi_plan.empty:
                            st.table(cqi_plan)
                        else:
                            st.info("No CQI actions required (All CLOs Met KPI).")

                    # --- Dashboard ---
                    with t5:
                        st.subheader("Course Analytics Dashboard")
                        
                        # Grade Distribution
                        grades = df_res['Grade'].value_counts().sort_index()
                        fig2, ax2 = plt.subplots(figsize=(8, 4))
                        grades.plot(kind='bar', ax=ax2, color='#4A90E2')
                        ax2.set_title("Grade Distribution")
                        st.pyplot(fig2)
                        
                    # --- AUDIT DOWNLOAD ---
                    st.divider()
                    st.subheader("💾 Save Audit Evidence")
                    
                    excel_data = generate_evidence_excel(
                        info, 
                        df_res.to_dict('records'), 
                        edited_clo_stats.to_dict('records'),
                        plo_data,
                        {'pass_rate': pass_rate, 'gpa': avg_gpa}
                    )
                    
                    st.download_button(
                        label="Download Full Master Excel (Evidence)",
                        data=excel_data,
                        file_name=f"Processed_Master_{info['code']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        elif potential_cols: # Error parsing
            st.error(potential_cols) # Print error msg

# === DIRECT ENTRY MODE (Legacy) ===
elif mode == "Direct Entry":
    st.info("Use this if you don't have the CampusOne file.")
    with st.expander("Configuration", expanded=True):
        c_code = st.text_input("Course Code", "DMIM1012")
        st.write("Please select 'Upload CampusOne' mode for best results.")
