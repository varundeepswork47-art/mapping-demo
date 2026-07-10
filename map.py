import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Subject Line Mapper", page_icon="📬")
st.title("Subject Line Mapper")

# === CLEAN FUNCTION ===
def clean_text(text):
    text = str(text).lower()
    text = re.sub(r"\*+", " ", text)
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

# === EXTRACT BEFORE / AFTER ===
def extract_parts(pattern):
    parts = re.split(r"\{\{.*?\}\}", pattern)
    parts = [p.strip() for p in parts if p.strip() != ""]

    if len(parts) == 1:
        return parts[0], ""
    elif len(parts) >= 2:
        return parts[0], parts[-1]
    else:
        return "", ""

# === FILE UPLOADS ===
mapping_file = st.file_uploader("Upload Mapping File (pattern, lob, frequency)", type=["xlsx", "csv"])
main_file = st.file_uploader("Upload Main File (subject)", type=["xlsx", "csv"])

if mapping_file and main_file:
    mapping = pd.read_csv(mapping_file) if mapping_file.name.endswith(".csv") else pd.read_excel(mapping_file)
    main = pd.read_csv(main_file) if main_file.name.endswith(".csv") else pd.read_excel(main_file)

    mapping.columns = ["pattern", "lob", "frequency"]
    main.columns = ["subject"]

    if st.button("Run Mapping", type="primary"):
        # Apply extraction BEFORE cleaning
        mapping["before"], mapping["after"] = zip(*mapping["pattern"].apply(extract_parts))

        # Clean text
        mapping["before"] = mapping["before"].apply(clean_text)
        mapping["after"] = mapping["after"].apply(clean_text)
        main["subject_clean"] = main["subject"].apply(clean_text)

        # Initialize output
        main["LOB"] = "No Match"
        main["Frequency"] = "No Match"

        # === MATCHING ===
        progress_bar = st.progress(0)
        total = len(mapping)
        for i, (idx, row) in enumerate(mapping.iterrows()):
            before = row["before"]
            after = row["after"]

            if before and after:
                mask = main["subject_clean"].str.contains(before, regex=False) & \
                       main["subject_clean"].str.contains(after, regex=False)
            elif before:
                mask = main["subject_clean"].str.contains(before, regex=False)
            else:
                progress_bar.progress((i + 1) / total)
                continue

            new_match = mask & (main["LOB"] == "No Match")

            main.loc[new_match, "LOB"] = row["lob"]
            main.loc[new_match, "Frequency"] = row["frequency"]

            progress_bar.progress((i + 1) / total)

        st.success("✅ Done!")
        st.dataframe(main, use_container_width=True)

        # === DOWNLOAD ===
        output_buffer = io.BytesIO()
        main.to_excel(output_buffer, index=False)
        output_buffer.seek(0)

        st.download_button(
            label="📥 Download Output",
            data=output_buffer,
            file_name="final_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
