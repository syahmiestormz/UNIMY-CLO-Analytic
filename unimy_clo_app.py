import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import io

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

def process_single_course(uploaded_file):
    """
    Reads UNIMY Master Excel and calculates CLO/PLO stats.
    """
    try:
        xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        
        sheet_setup = next((s for s in xls if "Setup" in s), None)
        sheet_marks = next((s for s in xls if "Table 1" in s or "Marks" in s), None)
        
        if not (sheet_setup and sheet_marks):
            return None, "Missing 'Setup' or 'Table 1' sheets."

        # 1. PARSE SETUP (Assessment Configuration)
        df_setup = xls[sheet_setup]
        course_info = {"name": "Unknown", "code": "Unknown"}
        
        # Extract Course Info
        for r in range(min(15, len(df_setup))):
            row_str = df_setup.iloc[r].astype(str).str.cat(sep=' ')
            if "Course Name" in row_str:
                parts = str(df_setup.iloc[r, 1])
                if parts != "nan": course_info["name"] = parts
            if "Course Code" in row_str:
                parts = str(df_setup.iloc[r, 1])
                if parts != "nan": course_info["code"] = parts

        # Extract Assessments
        assessments = {}
        setup_header = find_header_row(df_setup, ["Assessment Name", "Weightage"])
        
        if setup_header != -1:
            df_setup.columns = df_setup.iloc[setup_header]
            df_setup = df_setup.iloc[setup_header+1:].reset_index(drop=True)
            
            for _, row in df_setup.iterrows():
                name = str(row.get("Assessment Name", "")).strip()
                if name and name.lower() != "nan":
                    try:
                        assessments[name] = {
                            "weight": float(row.get("Weightage (%)", 0)),
                            "full": float(row.get("Full Marks", 100)),
                            "clo": str(row.get("CLO Tag", "Unmapped")).strip()
                        }
                    except: pass
        
        if not assessments:
            return None, "No assessments found in Setup sheet."

        # 2. PARSE MARKS & CALCULATE
        df_marks = xls[sheet_marks]
        marks_header = find_header_row(df_marks, ["STUDENT ID", "STUDENT NAME"])
        
        if marks_header == -1:
            return None, "Could not find Student ID header in Table 1."
            
        # Fix Duplicate Headers
        raw_header = df_marks.iloc[marks_header].astype(str).tolist()
        seen = {}
        clean_header = []
        for h in raw_header:
            if h in seen: seen[h] += 1; clean_header.append(f"{h}_{seen[h]}")
            else: seen[h] = 0; clean_header.append(h)
            
        df_marks.columns = clean_header
        df_marks = df_marks.iloc[marks_header+1:].reset_index(drop=True)
        
        # Map Columns
        col_map = {}
        for col in df_marks.columns:
            for asm in assessments.keys():
                if asm.lower() == col.lower(): col_map[asm] = col
                elif asm.lower() in col.lower() and "total" not in col.lower(): col_map[asm] = col

        # Calculations
        student_results = []
        clo_buckets = {} # { 'CLO 1': {'total_earned': 0, 'total_weight': 0} }
        
        for _, row in df_marks.iterrows():
            s_id = str(row.get("STUDENT ID", "")).strip()
            if len(s_id) < 2 or s_id.lower() == "nan": continue
            
            student_clos = {}
            
            for asm_name, config in assessments.items():
                col_name = col_map.get(asm_name)
                if col_name:
                    try:
                        raw_score = pd.to_numeric(row[col_name], errors='coerce')
                        if pd.isna(raw_score): raw_score = 0
                        
                        # Calc Contribution
                        # (Raw / Full) * Weight
                        contrib = (raw_score / config['full']) * config['weight']
                        
                        clo = config['clo']
                        if clo not in student_clos: student_clos[clo] = {'earned': 0, 'weight': 0}
                        student_clos[clo]['earned'] += contrib
                        student_clos[clo]['weight'] += config['weight']
                        
                    except: pass
            
            # Finalize Student Scores
            res = {'id': s_id, 'name': row.get("STUDENT NAME", "")}
            total_earned = 0
            
            for clo, vals in student_clos.items():
                if vals['weight'] > 0:
                    # Normalize to 100%
                    perc = (vals['earned'] / vals['weight']) * 100
                    res[clo] = perc
                    total_earned += vals['earned'] # Weight is already factored in 'earned' calculation relative to subject
            
            res['Total'] = total_earned # Sum of weighted contributions = Final Mark
            student_results.append(res)

        df_results = pd.DataFrame(student_results)
        return df_results, course_info

    except Exception as e:
        return None, str(e)

# --- Main Interface ---

with st.sidebar:
    st.header("UNIMY Analytics")
    st.info("Upload one Subject Excel file to analyze CLO/PLO.")
    st.caption("© 2026 UNIMY Programme Assessment")

st.title("📊 UNIMY CLO Analytics")
st.markdown("**Subject-Level CLO & PLO Calculator**")

uploaded_file = st.file_uploader("Upload Master Excel (e.g. DMIM1012.xlsx)", type=['xlsx'])

if uploaded_file:
    df_res, info = process_single_course(uploaded_file)
    
    if df_res is not None and not df_res.empty:
        # Display Info
        c1, c2, c3 = st.columns(3)
        c1.metric("Course Code", info['code'])
        c2.metric("Total Students", len(df_res))
        
        # Calculate Pass Rate (Standard: >= 50)
        passed = df_res[df_res['Total'] >= 50]
        pass_rate = (len(passed) / len(df_res)) * 100
        c3.metric("Pass Rate", f"{pass_rate:.1f}%")
        
        # Identify CLO Columns
        clo_cols = sorted([c for c in df_res.columns if "CLO" in c])
        
        # --- TAB INTERFACE ---
        tab1, tab2, tab3 = st.tabs(["Student Results", "CLO Analysis", "Generate Report"])
        
        with tab1:
            st.subheader("Student CLO Attainment")
            # Styling: Red if < 50
            st.dataframe(
                df_res.style.format("{:.1f}", subset=clo_cols + ['Total'])
                .map(lambda v: 'color: red;' if v < 50 else 'color: green;', subset=clo_cols + ['Total']),
                use_container_width=True
            )
            
        with tab2:
            st.subheader("CLO Analysis (Table 2)")
            
            # Calculate CLO Stats
            clo_stats = []
            for clo in clo_cols:
                avg = df_res[clo].mean()
                # Pass rate for specific CLO
                pass_count = len(df_res[df_res[clo] >= 50])
                clo_pass_rate = (pass_count / len(df_res)) * 100
                
                clo_stats.append({
                    "CLO": clo,
                    "Average (%)": avg,
                    "KPI Status": "ACHIEVED" if avg >= 50 else "NOT MET",
                    "Student Pass Rate (%)": clo_pass_rate
                })
            
            df_clo_stats = pd.DataFrame(clo_stats)
            st.dataframe(df_clo_stats.style.format("{:.1f}", subset=["Average (%)", "Student Pass Rate (%)"]), use_container_width=True)
            
            # Bar Chart
            fig, ax = plt.subplots(figsize=(8, 4))
            colors = ['#4CAF50' if x >= 50 else '#F44336' for x in df_clo_stats["Average (%)"]]
            ax.bar(df_clo_stats["CLO"], df_clo_stats["Average (%)"], color=colors)
            ax.axhline(50, color='black', linestyle='--')
            ax.set_title("Average CLO Attainment")
            st.pyplot(fig)

        with tab3:
            st.subheader("📄 Report Generator")
            st.info("Copy this text for your ESPAR / Course Review Report.")
            
            # Auto-Generate Text
            weak_clos = [c['CLO'] for c in clo_stats if c['Average (%)'] < 50]
            
            # Executive Summary Draft
            report = f"""**COURSE PERFORMANCE SUMMARY**
Course: {info['code']} - {info['name']}
Pass Rate: {pass_rate:.1f}%

**CLO ANALYSIS**
"""
            for stat in clo_stats:
                status = "MET" if stat['Average (%)'] >= 50 else "NOT MET"
                report += f"- {stat['CLO']}: {stat['Average (%)']:.1f}% ({status})\n"
                
            report += "\n**CQI ACTION PLAN (Suggestions)**\n"
            
            if weak_clos:
                report += "| CLO | Issue | Action |\n|---|---|---|\n"
                for clo in weak_clos:
                    fail_rate = 100 - df_clo_stats[df_clo_stats['CLO'] == clo]['Student Pass Rate (%)'].values[0]
                    rec = get_smart_recommendation(info['name'], fail_rate)
                    report += f"| {clo} | High failure rate ({fail_rate:.1f}%) | {rec} |\n"
            else:
                report += "All CLOs achieved the KPI target of 50%. No critical interventions required."
                
            st.text_area("Generated Text", report, height=400)

    elif df_res is None:
        st.error(f"Error: {info}")
    else:
        st.warning("No student data found in Table 1.")

else:
    st.info("Upload your Excel file to begin.")
