def generate_evidence_excel(course_info, student_results, clo_stats, plo_stats, crr_data):
    """
    Generates an Excel file mimicking the 'Master Template' for audit evidence.
    Includes Setup, Table 1, Table 2, Table 3, and CRR sheets.
    """
    # Use openpyxl as it is already installed for reading
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
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
