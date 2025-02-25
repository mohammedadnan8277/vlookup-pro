import streamlit as st
import pandas as pd

# Function to clean text (removes extra spaces, converts to lowercase)
def clean_text(value):
    if isinstance(value, str):
        return value.strip().lower()
    return value

# Streamlit UI
st.title("🔍 Exact Match VLOOKUP Pro ")  
st.markdown("<h5>Precision in Every Lookup! – By Adnan </h5>", unsafe_allow_html=True)
st.write("Upload your **Main Data** and **Lookup Table**, and this tool will fetch values only for exact matches!")

# Upload Files
main_file = st.file_uploader("📂 Step 1: Upload the Main Data File (Excel) – This is the dataset where you need to fetch values.", type=["xlsx"])
lookup_file = st.file_uploader("📂 Step 2: Upload the Lookup Table (Excel) – This file contains the reference values for matching.", type=["xlsx"])

if main_file and lookup_file:
    try:
        # Load Excel Data
        df_main = pd.read_excel(main_file)
        df_lookup = pd.read_excel(lookup_file)

        st.write("✅ Files Uploaded Successfully!")

        # Ensure both files contain data
        if df_main.empty or df_lookup.empty:
            st.error("❌ One or both uploaded files are empty. Please upload valid Excel files.")
        else:
            # Column Selection
            main_col = st.selectbox("🔍 Select the Matching Column from Main Data:", df_main.columns)
            lookup_col = st.selectbox("🔍 Select the Matching Column from Lookup Table:", df_lookup.columns)
            value_col = st.selectbox("🎯 Select the Column to Fetch from Lookup Table:", df_lookup.columns)

            if st.button("🚀 Fetch Exact Matches"):
                # Clean data
                df_main[main_col] = df_main[main_col].apply(clean_text)
                df_lookup[lookup_col] = df_lookup[lookup_col].apply(clean_text)

                # Create lookup dictionary
                lookup_dict = dict(zip(df_lookup[lookup_col], df_lookup[value_col]))

                # Fetch values only for exact matches
                df_main["Fetched Value"] = df_main[main_col].map(lookup_dict)

                # Display results
                st.write("### 🎯 Updated Data Preview")
                st.dataframe(df_main)

                # Save Fixed File
                fixed_filename = "exact_match_results.xlsx"
                df_main.to_excel(fixed_filename, index=False)

                # Download Button
                with open(fixed_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download Updated File",
                        data=file,
                        file_name=fixed_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                st.success("✅ Exact matches fetched successfully! Download the file above. 🚀")

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
