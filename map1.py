import subprocess
import sys

# -------- AUTOMATIC CORE INSTALLER (BYPASSES REQUIREMENTS.TXT) --------
try:
    from rapidfuzz import fuzz, process
except ModuleNotFoundError:
    # Forces Streamlit's linux virtual server to download the dependencies
    subprocess.check_call([sys.executable, "-m", "pip", "install", "rapidfuzz", "openpyxl"])
    from rapidfuzz import fuzz, process

import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment

# Set page configuration
st.set_page_config(page_title="Fuzzy Text Mapper Pro", page_icon="🧩", layout="wide")

# Initialize session state for results persistence
if 'mapping_results' not in st.session_state:
    st.session_state.mapping_results = None
if 'file_status' not in st.session_state:
    st.session_state.file_status = None

st.title("String Mapping Dashboard")
st.markdown("Upload your Main File(s) and a Mapping Reference File to reconcile text values using exact and fuzzy matching algorithms.")

# --- FILE UPLOAD SECTION ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Load Main File(s)")
    st.info("Upload up to 13 files (.csv, .xlsx) - all with same header structure")
    main_files = st.file_uploader(
        "Upload Main Sheet(s)",
        type=["csv", "xlsx"],
        key="main",
        accept_multiple_files=True
    )
    
    # Validate file count
    if main_files and len(main_files) > 13:
        st.error(f"❌ You uploaded {len(main_files)} files. Maximum allowed is 13 files. Please remove {len(main_files) - 13} file(s).")
        main_files = main_files[:13]
        st.info(f" Trimmed to first 13 files.")

with col2:
    st.subheader("2. Load Mapping Reference File")
    st.info(" Upload only 1 mapping file (will be used for all main files)")
    mapping_file = st.file_uploader("Upload Mapping Patterns (.csv, .xlsx)", type=["csv", "xlsx"], key="map")

if main_files and mapping_file:
    # Read files based on extension safely
    @st.cache_data(ttl=3600)
    def load_uploaded_data(file_obj):
        if file_obj.name.endswith(".csv"):
            return pd.read_csv(file_obj)
        return pd.read_excel(file_obj)

    # Load mapping file once
    df_map = load_uploaded_data(mapping_file).copy()
    
    # Display file info
    st.success(f" Successfully loaded {len(main_files)} main file(s) and mapping reference file!")
    st.info(f" Mapping file size: {len(df_map):,} rows")
    
    # Show list of uploaded files
    with st.expander(" View Uploaded Files"):
        for idx, file in enumerate(main_files, 1):
            st.write(f"{idx}. **{file.name}**")

    # --- COLUMN MAPPER SELECTION (USER VISIBILITY) ---
    st.subheader(" Map Your Column Fields")
    
    col_sel1, col_sel2 = st.columns(2)
    
    with col_sel1:
        st.info("Identify the text/value column in your Main file(s):")
        
        # Get columns from first file
        first_file_df = load_uploaded_data(main_files[0])
        subject_col_name = st.selectbox(
            "Select Subject Column (Main File)",
            options=first_file_df.columns,
            key="subject_select"
        )
        
    with col_sel2:
        st.info("Identify matching pattern columns from your Map file:")
        pattern_col = st.selectbox("Pattern Column (Map File)", options=df_map.columns, index=0 if len(df_map.columns) > 0 else 0)

    # --- DYNAMIC OUTPUT COLUMN SELECTION ---
    st.subheader(" Select Output Mapping Columns")
    
    available_output_cols = [col for col in df_map.columns if col != pattern_col]
    
    if len(available_output_cols) > 0:
        selected_output_cols = st.multiselect(
            "Select which columns to map as output (you can select multiple)",
            options=available_output_cols,
            default=available_output_cols[:min(2, len(available_output_cols))]
        )
    else:
        st.warning("⚠️ No additional columns available in your mapping file for output!")
        selected_output_cols = []

    # --- CUSTOMIZABLE OUTPUT COLUMN NAMES ---
    if selected_output_cols:
        st.subheader(" Customize Output Column Names")
        
        custom_output_names = {}
        cols_per_row = 2
        cols = st.columns(cols_per_row)
        
        for idx, col in enumerate(selected_output_cols):
            with cols[idx % cols_per_row]:
                custom_name = st.text_input(
                    f"Custom name for '{col}'",
                    value=f"Mapped_{col}",
                    key=f"custom_{col}"
                )
                custom_output_names[col] = custom_name

        # --- MATCHING ALGORITHM SETTINGS ---
        st.subheader(" Matching Algorithm Settings")
        
        algo_col1, algo_col2, algo_col3 = st.columns(3)
        
        with algo_col1:
            fuzzy_threshold = st.slider(
                "Fuzzy Match Score Threshold",
                min_value=50,
                max_value=100,
                value=75,
                step=5,
                help="Lower = more lenient matching, Higher = stricter matching"
            )
        
        with algo_col2:
            exact_match_type = st.selectbox(
                "Exact Match Type",
                options=["Substring Match", "Exact Match", "Both"],
                index=0,
                help="Substring: 'apple' matches in 'pineapple' | Exact: must match exactly"
            )
        
        with algo_col3:
            case_sensitive = st.checkbox(
                "Case Sensitive Matching",
                value=False,
                help="If unchecked, 'Apple' and 'apple' are treated as the same"
            )

        # --- PROCESSING TRIGGER ---
        if st.button("Run Mapping Engine on All Files", type="primary"):
            
            # Process each file
            all_results = []
            file_status_info = []
            
            # Overall progress bar
            overall_progress_bar = st.progress(0)
            overall_status_text = st.empty()
            
            for file_idx, main_file in enumerate(main_files, 1):
                file_name = main_file.name
                overall_progress = (file_idx - 1) / len(main_files)
                overall_status_text.write(f"📄 **Processing File {file_idx}/{len(main_files)}: {file_name}**")
                
                with st.status(f"Processing {file_name}...", expanded=False) as status:
                    
                    try:
                        # Load main file
                        df_main = load_uploaded_data(main_file).copy()
                        df_main_original = df_main.copy()  # Keep original for merging later
                        status.write(f"✓ File loaded: {len(df_main):,} rows")
                        
                        # Validate column exists
                        if subject_col_name not in df_main.columns:
                            st.error(f"❌ Column '{subject_col_name}' not found in {file_name}. Skipping this file.")
                            file_status_info.append({
                                "file": file_name,
                                "status": "❌ Failed",
                                "rows": 0,
                                "reason": f"Column '{subject_col_name}' not found"
                            })
                            continue
                        
                        # Preprocessing
                        status.write(" Step 1/5: Cleaning and normalizing text keys...")
                        step_progress = st.progress(0.0)
                        
                        df_map_temp = df_map.copy()
                        df_map_temp['pattern_normalized'] = df_map_temp[pattern_col].astype(str).fillna('').str.strip()
                        df_main['subject_normalized'] = df_main[subject_col_name].astype(str).fillna('').str.strip()
                        
                        step_progress.progress(0.5)
                        
                        # Apply case sensitivity
                        if not case_sensitive:
                            df_map_temp['pattern_normalized'] = df_map_temp['pattern_normalized'].str.lower()
                            df_main['subject_normalized'] = df_main['subject_normalized'].str.lower()
                        
                        step_progress.progress(1.0)
                        step_progress.empty()
                        
                        # Hash Lookup generation for selected output columns
                        status.write(" Step 2/5: Building lookup dictionary...")
                        step_progress = st.progress(0.0)
                        
                        exact_match_dict = {}
                        for idx, row in df_map_temp.iterrows():
                            pattern_key = row['pattern_normalized']
                            output_values = {col: row[col] for col in selected_output_cols}
                            exact_match_dict[pattern_key] = output_values
                        
                        step_progress.progress(1.0)
                        step_progress.empty()

                        # Initialize results columns
                        for col in selected_output_cols:
                            df_main[f'mapped_{col}'] = None

                        # --- STEP 1: EXACT MATCHING ---
                        status.write(" Step 3/5: Executing Direct Fast-Lookup (Exact Matching)...")
                        step_progress = st.progress(0.0)
                        
                        patterns_items = list(exact_match_dict.items())
                        

                        def check_exact_match(subj):
                            """Check for exact or substring matches based on user selection"""
                            if pd.isna(subj) or subj == '':
                                return {col: None for col in selected_output_cols}
                            
                            subj_str = str(subj).strip()
                            if not case_sensitive:
                                subj_str = subj_str.lower()
                            
                            for patt, output_dict in patterns_items:
                                if patt == '' or patt is None:
                                    continue
                                
                                match_found = False
                                
                                if exact_match_type == "Substring Match":
                                    match_found = patt in subj_str
                                elif exact_match_type == "Exact Match":
                                    match_found = patt == subj_str
                                elif exact_match_type == "Both":
                                    match_found = (patt in subj_str) or (patt == subj_str)
                                
                                if match_found:
                                    return output_dict
                            
                            return {col: None for col in selected_output_cols}

                        # Apply matching with progress
                        exact_results = []
                        for idx, val in enumerate(df_main['subject_normalized']):
                            exact_results.append(check_exact_match(val))
                            if idx % max(1, len(df_main) // 100) == 0:
                                step_progress.progress(min(1.0, idx / len(df_main)))
                        
                        step_progress.progress(1.0)
                        step_progress.empty()
                        
                        for col in selected_output_cols:
                            df_main[f'mapped_{col}'] = [r.get(col) for r in exact_results]

                        # Count matches
                        exact_matches_per_col = {col: (df_main[f'mapped_{col}'].notna()).sum() for col in selected_output_cols}
                        unmatched_mask = df_main['mapped_' + selected_output_cols[0]].isna()
                        unmatched_count = unmatched_mask.sum()

                        status.write(f"✓ Exact Match Phase Closed")
                        for col, count in exact_matches_per_col.items():
                            status.write(f"  • {col}: {count:,} matches found")
                        status.write(f" Remainder to resolve via Fuzzy Distance calculations: {unmatched_count:,} rows.")

                        # --- STEP 2: RAPIDFUZZ FUZZY MATCHING ---
                        if unmatched_count > 0:
                            status.write(" Step 4/5: Spinning up RapidFuzz text vector alignment...")
                            
                            patterns_list = df_map_temp['pattern_normalized'].tolist()
                            unmatched_indices = df_main[unmatched_mask].index
                            
                            # Pre-extract values
                            subjects_unmatched = df_main.loc[unmatched_indices, 'subject_normalized'].values
                            
                            # Initialize fuzzy results
                            fuzzy_results_dict = {col: [] for col in selected_output_cols}

                            # Progress handling
                            fuzzy_progress = st.progress(0.0)
                            
                            for idx, subj in enumerate(subjects_unmatched):
                                best_match = process.extractOne(
                                    subj,
                                    patterns_list,
                                    scorer=fuzz.token_set_ratio,
                                    score_cutoff=fuzzy_threshold
                                )
                                
                                if best_match:
                                    m_pattern, score, _ = best_match
                                    matched_outputs = exact_match_dict[m_pattern]
                                    for col in selected_output_cols:
                                        fuzzy_results_dict[col].append(matched_outputs.get(col))
                                else:
                                    for col in selected_output_cols:
                                        fuzzy_results_dict[col].append("No Match")
                                
                                # Update progress occasionally
                                if idx % max(1, unmatched_count // 100) == 0 and unmatched_count > 0:
                                    fuzzy_progress.progress(min(1.0, float(idx / unmatched_count)))

                            fuzzy_progress.empty()
                            
                            # Write back fuzzy results
                            for col in selected_output_cols:
                                df_main.loc[unmatched_indices, f'mapped_{col}'] = fuzzy_results_dict[col]

                        # --- STEP 3: FINALIZING OUTPUT ---
                        status.write(" Step 5/5: Finalizing output...")
                        step_progress = st.progress(0.0)
                        
                        status.update(label=f"✅ {file_name} reconciled successfully!", state="complete")
                        step_progress.progress(1.0)
                        step_progress.empty()
                        
                        # --- BUILD FINAL OUTPUT ---
                        # Start with all main file columns
                        final_out = df_main_original.copy()
                        
                        # Add mapped columns
                        for col in selected_output_cols:
                            final_out[custom_output_names[col]] = df_main[f'mapped_{col}']
                        
                        all_results.append({
                            'file_name': file_name,
                            'dataframe': final_out,
                            'row_count': len(final_out)
                        })
                        
                        # Calculate statistics
                        matched_overall = (df_main['mapped_' + selected_output_cols[0]].notna() & (df_main['mapped_' + selected_output_cols[0]] != "No Match")).sum()
                        match_pct = (matched_overall / len(df_main) * 100) if len(df_main) > 0 else 0
                        
                        file_status_info.append({
                            "file": file_name,
                            "status": "✅ Success",
                            "rows": len(df_main),
                            "matched": matched_overall,
                            "match_pct": f"{match_pct:.1f}%"
                        })
                    
                    except Exception as e:
                        st.error(f"❌ Error processing {file_name}: {str(e)}")
                        file_status_info.append({
                            "file": file_name,
                            "status": "❌ Error",
                            "rows": 0,
                            "reason": str(e)
                        })
                
                # Update overall progress
                overall_progress = file_idx / len(main_files)
                overall_progress_bar.progress(overall_progress)

            overall_progress_bar.progress(1.0)
            overall_status_text.write(f"✅ **All {len(main_files)} file(s) processed!**")
            
            # Store results in session state
            st.session_state.mapping_results = all_results
            st.session_state.file_status = file_status_info

        # --- PERSIST & RENDER RESULTS FROM SESSION STATE ---
        if st.session_state.mapping_results:
            all_results = st.session_state.mapping_results
            file_status_info = st.session_state.file_status

            st.success(" Batch Processing Complete!")
            
            st.subheader(" Processing Summary")
            summary_df = pd.DataFrame(file_status_info)
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            
            # --- INDIVIDUAL FILE PREVIEWS & DOWNLOADS ---
            st.subheader(" Preview & Download Results")
            
            download_files = []
            
            for result in all_results:
                with st.expander(f" {result['file_name']} - {result['row_count']:,} rows"):
                    st.dataframe(result['dataframe'].head(15), use_container_width=True)
                    
                    # Individual file download
                    csv_buffer = io.StringIO()
                    result['dataframe'].to_csv(csv_buffer, index=False)
                    csv_data = csv_buffer.getvalue()
                    
                    st.download_button(
                        label=f"📥 Download {result['file_name'].split('.')[0]}_mapped.csv",
                        data=csv_data,
                        file_name=f"{result['file_name'].split('.')[0]}_mapped.csv",
                        mime="text/csv",
                        key=f"download_{result['file_name']}"
                    )
                    
                    download_files.append({
                        'file_name': f"{result['file_name'].split('.')[0]}_mapped.csv",
                        'data': csv_data
                    })
            
            # --- BATCH DOWNLOAD AS ZIP ---
            if len(all_results) > 1:
                st.subheader(" Batch Download")
                
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for file_info in download_files:
                        zip_file.writestr(file_info['file_name'], file_info['data'])
                
                zip_buffer.seek(0)
                
                st.download_button(
                    label=" Download All Files as ZIP",
                    data=zip_buffer.getvalue(),
                    file_name=f"mapped_data_batch_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            
            # --- SUMMARY DASHBOARD ---
            st.subheader("📊 Summary Dashboard")
            st.info("Select columns from your mapped output to generate summary statistics")
            
            # Get all columns from output dataframe
            output_columns = list(all_results[0]['dataframe'].columns) if all_results else []
            
            # Allow user to select columns for summary
            selected_summary_cols = st.multiselect(
                "Select columns for summary (combination counts)",
                options=output_columns,
                default=output_columns[:min(3, len(output_columns))],
                key="summary_cols_select"
            )
            
            if selected_summary_cols:
                if st.button("Generate Summary Dashboard", type="secondary"):
                    # Create combination statistics
                    combo_stats = []
                    
                    for result in all_results:
                        temp_df = result['dataframe'].copy()
                        
                        # Group by selected columns and count
                        grouped = temp_df.groupby(selected_summary_cols, dropna=False).size().reset_index(name='Count')
                        grouped['File'] = result['file_name']
                        combo_stats.append(grouped)
                    
                    if combo_stats:
                        summary_dashboard_df = pd.concat(combo_stats, ignore_index=True)
                        
                        st.dataframe(summary_dashboard_df, use_container_width=True, hide_index=True)
                        
                        # --- DOWNLOAD SUMMARY DASHBOARD AS EXCEL ---
                        def create_summary_excel():
                            excel_buffer = io.BytesIO()
                            
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                summary_dashboard_df.to_excel(writer, sheet_name='Combination Counts', index=False)
                                
                                # Style the header row
                                worksheet = writer.sheets['Combination Counts']
                                header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                                header_font = Font(bold=True, color="FFFFFF")
                                
                                for cell in worksheet[1]:
                                    cell.fill = header_fill
                                    cell.font = header_font
                                    cell.alignment = Alignment(horizontal="center", vertical="center")
                                
                                # Adjust column widths
                                for column in worksheet.columns:
                                    max_length = 0
                                    column_letter = column[0].column_letter
                                    for cell in column:
                                        try:
                                            if len(str(cell.value)) > max_length:
                                                max_length = len(str(cell.value))
                                        except:
                                            pass
                                    adjusted_width = (max_length + 2)
                                    worksheet.column_dimensions[column_letter].width = adjusted_width
                            
                            excel_buffer.seek(0)
                            return excel_buffer.getvalue()
                        
                        excel_data = create_summary_excel()
                        
                        st.download_button(
                            label=" Download Summary Dashboard (Excel)",
                            data=excel_data,
                            file_name=f"summary_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            
            # --- DETAILED STATISTICS ---
            st.subheader("📋 Detailed Settings Used")
            settings_col1, settings_col2 = st.columns(2)
            
            with settings_col1:
                st.info("**Matching Configuration:**")
                st.write(f"""
                - Fuzzy Threshold: **{fuzzy_threshold}%**
                - Match Type: **{exact_match_type}**
                - Case Sensitive: **{case_sensitive}**
                """)
            
            with settings_col2:
                st.info("**Column Mapping:**")
                st.write(f"""
                - Subject Column: **{subject_col_name}**
                - Pattern Column: **{pattern_col}**
                - Output Columns: **{', '.join(selected_output_cols)}**
                """)
    else:
        st.warning("Please select at least one output column from your mapping file.")