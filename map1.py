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

# Set page configuration
st.set_page_config(page_title="Fuzzy Text Mapper Pro", page_icon="🧩", layout="wide")

st.title("🧩 Automated String Mapping Dashboard")
st.markdown("Upload your Main File and a Mapping Reference File to reconcile text values using exact and fuzzy matching algorithms.")

# --- FILE UPLOAD SECTION ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Load Main File")
    main_file = st.file_uploader("Upload Main Sheet (.csv, .xlsx)", type=["csv", "xlsx"], key="main")

with col2:
    st.subheader("2. Load Mapping Reference File")
    mapping_file = st.file_uploader("Upload Mapping Patterns (.csv, .xlsx)", type=["csv", "xlsx"], key="map")

if main_file and mapping_file:
    # Read files based on extension safely
    @st.cache_data(ttl=3600)
    def load_uploaded_data(file_obj):
        if file_obj.name.endswith(".csv"):
            return pd.read_csv(file_obj)
        return pd.read_excel(file_obj)

    df_main = load_uploaded_data(main_file).copy()
    df_map = load_uploaded_data(mapping_file).copy()

    # --- COLUMN MAPPER SELECTION (USER VISIBILITY) ---
    st.success("🎉 Files uploaded successfully!")
    st.subheader("⚙️ Map Your Column Fields")
    
    col_sel1, col_sel2 = st.columns(2)
    
    with col_sel1:
        st.info("Identify the text/value column in your Main file:")
        subject_col = st.selectbox("Subject Column (Main File)", options=df_main.columns)
        
    with col_sel2:
        st.info("Identify matching pattern columns from your Map file:")
        pattern_col = st.selectbox("Pattern Column (Map File)", options=df_map.columns, index=0 if len(df_map.columns) > 0 else 0)

    # --- DYNAMIC OUTPUT COLUMN SELECTION ---
    st.subheader("📊 Select Output Mapping Columns")
    
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
        st.subheader("📋 Customize Output Column Names")
        
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
        st.subheader("⚙️ Matching Algorithm Settings")
        
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
        if st.button("Run Mapping Engine", type="primary"):
            
            with st.status("Initializing High Performance Compute Matrix...", expanded=True) as status:
                
                # Preprocessing using lightning fast pandas vector methods
                status.write("🧼 Cleaning and normalizing text keys...")
                
                # Convert everything to string and handle NaN values
                df_map['pattern_normalized'] = df_map[pattern_col].astype(str).fillna('').str.strip()
                df_main['subject_normalized'] = df_main[subject_col].astype(str).fillna('').str.strip()
                
                # Apply case sensitivity
                if not case_sensitive:
                    df_map['pattern_normalized'] = df_map['pattern_normalized'].str.lower()
                    df_main['subject_normalized'] = df_main['subject_normalized'].str.lower()
                
                # Hash Lookup generation for selected output columns
                exact_match_dict = {}
                for idx, row in df_map.iterrows():
                    pattern_key = row['pattern_normalized']
                    output_values = {col: row[col] for col in selected_output_cols}
                    exact_match_dict[pattern_key] = output_values

                # Initialize results columns
                for col in selected_output_cols:
                    df_main[f'mapped_{col}'] = None

                # --- STEP 1: EXACT MATCHING ---
                status.write("⚡ Step 1: Executing Direct Fast-Lookup (Exact Matching)...")
                
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

                # Apply matching
                exact_results = df_main['subject_normalized'].apply(check_exact_match)
                
                for col in selected_output_cols:
                    df_main[f'mapped_{col}'] = [r.get(col) for r in exact_results]

                # Count matches
                exact_matches_per_col = {col: (df_main[f'mapped_{col}'].notna()).sum() for col in selected_output_cols}
                unmatched_mask = df_main['mapped_' + selected_output_cols[0]].isna()
                unmatched_count = unmatched_mask.sum()

                status.write(f"✓ Exact Match Phase Closed")
                for col, count in exact_matches_per_col.items():
                    status.write(f"  • {col}: {count:,} matches found")
                status.write(f"⚠️ Remainder to resolve via Fuzzy Distance calculations: {unmatched_count:,} rows.")

                # --- STEP 2: RAPIDFUZZ FUZZY MATCHING ---
                if unmatched_count > 0:
                    status.write("🤖 Step 2: Spinning up RapidFuzz text vector alignment...")
                    
                    patterns_list = df_map['pattern_normalized'].tolist()
                    unmatched_indices = df_main[unmatched_mask].index
                    
                    # Pre-extract values
                    subjects_unmatched = df_main.loc[unmatched_indices, 'subject_normalized'].values
                    
                    # Initialize fuzzy results
                    fuzzy_results_dict = {col: [] for col in selected_output_cols}

                    # Progress handling
                    p_bar = st.progress(0.0)
                    
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
                            p_bar.progress(min(1.0, float(idx / unmatched_count)))

                    p_bar.empty()
                    
                    # Write back fuzzy results
                    for col in selected_output_cols:
                        df_main.loc[unmatched_indices, f'mapped_{col}'] = fuzzy_results_dict[col]

                status.update(label="✅ All data arrays reconciled successfully!", state="complete")

            # --- VIEW SUMMARY AND RESULTS ---
            st.success("📊 Run Execution Summary")
            
            tot = len(df_main)
            
            # Calculate match statistics per column
            stats_cols = st.columns(len(selected_output_cols) + 1)
            
            stats_cols[0].metric("Total Rows Processed", f"{tot:,}")
            
            for idx, col in enumerate(selected_output_cols):
                matched = (df_main[f'mapped_{col}'].notna() & (df_main[f'mapped_{col}'] != "No Match")).sum()
                no_match = (df_main[f'mapped_{col}'] == "No Match").sum()
                unmatched = df_main[f'mapped_{col}'].isna().sum()
                
                match_pct = (matched / tot * 100) if tot > 0 else 0
                stats_cols[idx + 1].metric(
                    f"Mapped: {col}",
                    f"{matched:,}",
                    f"{match_pct:.1f}%"
                )

            # Visualizing sample outputs
            st.subheader("👀 Preview Target Results Table")
            
            # Build final output with original and mapped columns
            final_out = df_main[[subject_col]].copy()
            
            for col in selected_output_cols:
                final_out[custom_output_names[col]] = df_main[f'mapped_{col}']
            
            st.dataframe(final_out.head(20), use_container_width=True)

            # File Export Streaming Logic
            csv_buffer = io.StringIO()
            final_out.to_csv(csv_buffer, index=False)
            csv_data = csv_buffer.getvalue()

            st.download_button(
                label="📥 Download Structured CSV Report",
                data=csv_data,
                file_name="reconciled_mapping_dataset.csv",
                mime="text/csv",
                use_container_width=True
            )
            
            # Show detailed statistics
            st.subheader("📈 Detailed Statistics")
            
            detail_col1, detail_col2 = st.columns(2)
            
            with detail_col1:
                st.info("**Match Summary:**")
                for col in selected_output_cols:
                    matched = (df_main[f'mapped_{col}'].notna() & (df_main[f'mapped_{col}'] != "No Match")).sum()
                    no_match = (df_main[f'mapped_{col}'] == "No Match").sum()
                    unmatched = df_main[f'mapped_{col}'].isna().sum()
                    
                    st.write(f"""
                    **{col}:**
                    - Matched: {matched:,}
                    - No Match (Fuzzy Failed): {no_match:,}
                    - Unmatched (Exact Failed): {unmatched:,}
                    """)
            
            with detail_col2:
                st.info("**Settings Used:**")
                st.write(f"""
                - Fuzzy Threshold: {fuzzy_threshold}%
                - Match Type: {exact_match_type}
                - Case Sensitive: {case_sensitive}
                - Pattern Column: {pattern_col}
                - Subject Column: {subject_col}
                """)
    else:
        st.warning("Please select at least one output column from your mapping file.")
