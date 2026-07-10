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
import re
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

# =====================================================================
# === TEMPLATE-PATTERN HELPERS (merged in from the batch script) ====
# =====================================================================
# Mapping patterns may contain {{placeholder}} tokens (e.g.
# "Your {{product}} renews on {{date}}"). These helpers pull out the
# literal text that comes BEFORE and AFTER the placeholder(s) so that
# matching can be done on the stable parts of the string instead of
# failing outright because the literal "{{...}}" text never appears
# in a real subject line. Patterns with no placeholder at all simply
# collapse to a single "before" fragment with an empty "after", which
# preserves the original plain substring/exact-match behavior.

def clean_text(text, case_sensitive=False):
    """Normalize text the same way the batch script does: strip
    asterisks, collapse everything down to alphanumerics + spaces,
    and squash whitespace. Honors the app's case-sensitivity toggle."""
    text = str(text)
    if not case_sensitive:
        text = text.lower()
        text = re.sub(r"\*+", " ", text)
        text = re.sub(r"[^a-z0-9 ]", " ", text)
    else:
        text = re.sub(r"\*+", " ", text)
        text = re.sub(r"[^a-zA-Z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_parts(pattern):
    """Split a mapping pattern on {{...}} placeholders and return the
    first literal chunk ('before') and last literal chunk ('after').
    Patterns without a placeholder return (whole_pattern, '')."""
    parts = re.split(r"\{\{.*?\}\}", str(pattern))
    parts = [p.strip() for p in parts if p.strip() != ""]

    if len(parts) == 1:
        return parts[0], ""
    elif len(parts) >= 2:
        return parts[0], parts[-1]
    else:
        return "", ""


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
        
        algo_col1, algo_col2, algo_col3, algo_col4 = st.columns(4)
        
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
                help="Substring: 'apple' matches in 'pineapple' | Exact: must match exactly. Note: patterns containing {{placeholder}} tokens always use whole-pattern regex matching, since an exact whole-string match isn't possible when part of the pattern is a variable."
            )
        
        with algo_col3:
            case_sensitive = st.checkbox(
                "Case Sensitive Matching",
                value=False,
                help="If unchecked, 'Apple' and 'apple' are treated as the same"
            )

        with algo_col4:
            min_fragment_len = st.slider(
                "Min. Literal Fragment Length",
                min_value=1,
                max_value=10,
                value=3,
                step=1,
                help="Patterns with {{placeholders}} are matched using the literal text around them. If every literal chunk of a pattern is shorter than this many characters, that pattern is too generic to match safely and is skipped for exact matching (it can still be reached via fuzzy matching)."
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
                        status.write(" Step 1/5: Cleaning and normalizing text keys (incl. {{placeholder}} extraction)...")
                        step_progress = st.progress(0.0)
                        
                        df_map_temp = df_map.copy()

                        # Subject text normalized with the shared cleaning
                        # function so pattern fragments and subject text are
                        # apples-to-apples (same alphanumeric-only cleaning).
                        df_main['subject_normalized'] = df_main[subject_col_name].astype(str).fillna('').apply(
                            lambda t: clean_text(t, case_sensitive)
                        )
                        
                        step_progress.progress(1.0)
                        step_progress.empty()
                        
                        # Build a compiled regex per pattern that keeps EVERY
                        # literal fragment (not just the first/last), requires
                        # them to appear in the correct order, and wraps each
                        # fragment in \b word boundaries so short fragments
                        # can't match inside unrelated words (e.g. "re" no
                        # longer matches inside "renewal", "reminder", etc).
                        # {{placeholder}} tokens become a ".*?" wildcard gap.
                        # Also build a deduped fuzzy candidate list for the
                        # fallback phase.
                        status.write(" Step 2/5: Building pattern regex lookup...")
                        step_progress = st.progress(0.0)

                        regex_flags = 0 if case_sensitive else re.IGNORECASE

                        patterns_items = []          # list of (compiled_regex, output_dict)
                        fuzzy_text_to_output = {}     # dedup map for fuzzy candidates
                        patterns_list_fuzzy = []
                        duplicate_count = 0
                        skipped_generic_count = 0

                        for idx, row in df_map_temp.iterrows():
                            raw_pattern = str(row[pattern_col])
                            output_values = {col: row[col] for col in selected_output_cols}

                            # Split, keeping the {{...}} tokens so we know
                            # exactly where each wildcard gap belongs.
                            segments = re.split(r"(\{\{.*?\}\})", raw_pattern)

                            regex_parts = []
                            literal_fragments = []
                            has_strong_fragment = False

                            for seg in segments:
                                if re.fullmatch(r"\{\{.*?\}\}", seg):
                                    regex_parts.append(r".*?")
                                else:
                                    cleaned = clean_text(seg, case_sensitive)
                                    if not cleaned:
                                        continue
                                    literal_fragments.append(cleaned)
                                    if len(cleaned) >= min_fragment_len:
                                        has_strong_fragment = True
                                    regex_parts.append(r"\b" + re.escape(cleaned) + r"\b")

                            if not literal_fragments:
                                # Pattern is 100% placeholder - nothing to
                                # anchor on, can't be matched safely at all.
                                skipped_generic_count += 1
                                continue

                            if not has_strong_fragment:
                                # Every literal chunk is shorter than the
                                # minimum fragment length - too generic to
                                # trust for exact matching (this is what was
                                # causing unrelated subjects to match).
                                skipped_generic_count += 1
                            else:
                                has_placeholder = "{{" in raw_pattern
                                if not has_placeholder and exact_match_type == "Exact Match":
                                    # Plain pattern (no {{placeholder}}) and
                                    # the user wants a true whole-string
                                    # match, not "appears somewhere".
                                    regex_str = r"^\s*" + re.escape(literal_fragments[0]) + r"\s*$"
                                else:
                                    # Substring/Both, or any pattern that has
                                    # a placeholder (whole-string exact match
                                    # isn't meaningful when part of it is a
                                    # variable) - use the word-boundary,
                                    # ordered-fragment regex built above.
                                    regex_str = "".join(regex_parts)
                                try:
                                    compiled = re.compile(regex_str, regex_flags)
                                    patterns_items.append((compiled, output_values))
                                except re.error:
                                    skipped_generic_count += 1

                            # Fuzzy candidate always built (even for
                            # generic/skipped patterns) since fuzzy scoring
                            # is similarity-based, not containment-based, so
                            # it's far less prone to the same false-positive
                            # problem.
                            fuzzy_text = " ".join(literal_fragments).strip()
                            if fuzzy_text:
                                if fuzzy_text not in fuzzy_text_to_output:
                                    fuzzy_text_to_output[fuzzy_text] = output_values
                                    patterns_list_fuzzy.append(fuzzy_text)
                                else:
                                    duplicate_count += 1

                        if duplicate_count > 0:
                            status.write(f"  ⚠️ {duplicate_count} duplicate pattern(s) detected — first occurrence used for fuzzy matching.")
                        if skipped_generic_count > 0:
                            status.write(f"  ⚠️ {skipped_generic_count} pattern(s) skipped for exact matching — literal text too short/generic (below {min_fragment_len} chars). Still reachable via fuzzy matching. Lower the 'Min. Literal Fragment Length' setting if you want these used for exact matching too.")
                        
                        step_progress.progress(1.0)
                        step_progress.empty()

                        # Initialize results columns
                        for col in selected_output_cols:
                            df_main[f'mapped_{col}'] = None

                        # --- STEP 1: EXACT / REGEX PLACEHOLDER MATCHING ---
                        status.write(" Step 3/5: Executing Direct Fast-Lookup (Full-Pattern Regex Matching)...")
                        step_progress = st.progress(0.0)

                        def check_exact_match(subj):
                            """Check each compiled pattern regex against the
                            subject, in mapping-row order (first match
                            wins). Word-boundaries + fragment order are
                            already baked into the compiled regex."""
                            if pd.isna(subj) or subj == '':
                                return {col: None for col in selected_output_cols}

                            subj_str = subj  # already cleaned upstream

                            for compiled, output_dict in patterns_items:
                                if compiled.search(subj_str):
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
                        # A row is only "unmatched" if ALL selected output
                        # columns came back empty - checking just the first
                        # column would wrongly flag legitimately-blank values
                        # in that column as no-match.
                        unmatched_mask = df_main[[f'mapped_{c}' for c in selected_output_cols]].isna().all(axis=1)
                        unmatched_count = unmatched_mask.sum()

                        status.write(f"✓ Exact Match Phase Closed")
                        for col, count in exact_matches_per_col.items():
                            status.write(f"  • {col}: {count:,} matches found")
                        status.write(f" Remainder to resolve via Fuzzy Distance calculations: {unmatched_count:,} rows.")

                        # --- STEP 2: RAPIDFUZZ FUZZY MATCHING ---
                        if unmatched_count > 0:
                            status.write(" Step 4/5: Spinning up RapidFuzz text vector alignment...")
                            
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
                                    patterns_list_fuzzy,
                                    scorer=fuzz.token_set_ratio,
                                    score_cutoff=fuzzy_threshold
                                )
                                
                                if best_match:
                                    m_pattern, score, _ = best_match
                                    matched_outputs = fuzzy_text_to_output[m_pattern]
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
            st.subheader("Detailed Settings Used")
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
