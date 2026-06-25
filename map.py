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
        st.info("Identify the text column in your Main file:")
        subject_col = st.selectbox("Subject Column (Main File)", options=df_main.columns)
        
    with col_sel2:
        st.info("Identify matching pattern columns from your Map file:")
        pattern_col = st.selectbox("Pattern Column (Map File)", options=df_map.columns, index=0 if len(df_map.columns) > 0 else 0)
        stage_col = st.selectbox("New Stage Output Column", options=df_map.columns, index=1 if len(df_map.columns) > 1 else 0)
        lob_col = st.selectbox("LOB Output Column", options=df_map.columns, index=2 if len(df_map.columns) > 2 else 0)

    # --- PROCESSING TRIGGER ---
    if st.button("Run Mapping Engine", type="primary"):
        
        with st.status("Initializing High Performance Compute Matrix...", expanded=True) as status:
            
            # Preprocessing using lightning fast pandas vector methods instead of manual loops
            status.write("🧼 Cleaning and normalizing text keys...")
            df_map['pattern_lower'] = df_map[pattern_col].astype(str).str.lower().str.strip()
            df_main['subject_lower'] = df_main[subject_col].astype(str).str.lower().str.strip()

            # Hash Lookup generation
            exact_match_dict = dict(zip(df_map['pattern_lower'], zip(df_map[stage_col], df_map[lob_col])))

            # Initialize results with proper dtype (object for strings)
            df_main['mapped_category'] = pd.Series(dtype='object')
            df_main['mapped_lob'] = pd.Series(dtype='object')
            df_main['match_type'] = pd.Series(dtype='object')

            # --- STEP 1: VECTORIZED EXACT MATCHING ---
            status.write("⚡ Step 1: Executing Direct Fast-Lookup (Exact Substring Hash Routing)...")
            
            # Compile logic safely for vectorized parsing
            patterns_items = list(exact_match_dict.items())
            
            def check_exact_match(subj):
                for patt, (stg, lb) in patterns_items:
                    if patt in subj:
                        return stg, lb, 'exact_match'
                return None, None, None

            # Use faster .apply() mapping route over row iterations
            exact_results = df_main['subject_lower'].apply(check_exact_match)
            
            df_main['mapped_category'] = [r[0] for r in exact_results]
            df_main['mapped_lob'] = [r[1] for r in exact_results]
            df_main['match_type'] = [r[2] for r in exact_results]

            exact_count = (df_main['match_type'] == 'exact_match').sum()
            unmatched_mask = df_main['mapped_category'].isna()
            unmatched_count = unmatched_mask.sum()

            status.write(f"✓ Exact Match Phase Closed: Found {exact_count:,} matches.")
            status.write(f"⚠️ Remainder to resolve via Fuzzy Distance calculations: {unmatched_count:,} rows.")

            # --- STEP 2: RAPIDFUZZ BLOCKING SYSTEM ---
            if unmatched_count > 0:
                status.write("🤖 Step 2: Spinning up RapidFuzz text vector alignment...")
                
                patterns_list = df_map['pattern_lower'].tolist()
                unmatched_indices = df_main[unmatched_mask].index
                
                # Pre-extract values to bypass slow dataframe pointer steps
                subjects_unmatched = df_main.loc[unmatched_indices, 'subject_lower'].values
                
                cat_outs = []
                lob_outs = []
                type_outs = []

                # Progress handling for visual tracking
                p_bar = st.progress(0.0)
                
                for idx, subj in enumerate(subjects_unmatched):
                    best_match = process.extractOne(
                        subj,
                        patterns_list,
                        scorer=fuzz.token_set_ratio,
                        score_cutoff=75
                    )
                    
                    if best_match:
                        m_pattern, score, _ = best_match
                        m_stage, m_lob = exact_match_dict[m_pattern]
                        cat_outs.append(m_stage)
                        lob_outs.append(m_lob)
                        type_outs.append(f"fuzzy_{int(score)}")
                    else:
                        cat_outs.append("No Match")
                        lob_outs.append("No Match")
                        type_outs.append("no_match")
                    
                    # Update status progress occasionally to prevent ui freezing 
                    if idx % 1000 == 0 and unmatched_count > 0:
                        p_bar.progress(float(idx / unmatched_count))

                p_bar.empty()
                
                # Write back remaining arrays using reset_index to avoid index mismatch
                df_main.loc[unmatched_indices, 'mapped_category'] = pd.Series(cat_outs, index=unmatched_indices)
                df_main.loc[unmatched_indices, 'mapped_lob'] = pd.Series(lob_outs, index=unmatched_indices)
                df_main.loc[unmatched_indices, 'match_type'] = pd.Series(type_outs, index=unmatched_indices)

            status.update(label="✅ All data arrays reconciled successfully!", state="complete")

        # --- VIEW SUMMARY AND RESULTS ---
        st.success("📊 Run Execution Summary")
        
        tot = len(df_main)
        ex_m = (df_main['match_type'] == 'exact_match').sum()
        fz_m = df_main['match_type'].str.contains('fuzzy_', na=False).sum()
        no_m = (df_main['match_type'] == 'no_match').sum()

        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.metric("Total Rows Processed", f"{tot:,}")
        m_col2.metric("Exact Matches", f"{ex_m:,}", f"{ex_m/tot*100:.1f}%")
        m_col3.metric("Fuzzy Matches Extracted", f"{fz_m:,}", f"{fz_m/tot*100:.1f}%")
        m_col4.metric("Unresolved Rows", f"{no_m:,}", f"-{no_m/tot*100:.1f}%", delta_color="inverse")

        # Visualizing sample outputs
        st.subheader("👀 Preview Target Results Table")
        final_out = df_main[[subject_col, 'mapped_category', 'mapped_lob', 'match_type']].copy()
        st.dataframe(final_out.head(15), use_container_width=True)

        # File Export Streaming Logic for high memory efficiency
        csv_buffer = io.StringIO()
        final_out.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()

        st.download_button(
            label="📥 Download Structured CSV Report",
            data=csv_data,
            file_name="mapped_dataset.csv",
            mime="text/csv",
            use_container_width=True
        )
