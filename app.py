import streamlit as st
import pandas as pd
import re
import io
import webbrowser

st.set_page_config(page_title="Team Voting Theme Extractor", layout="centered")

st.title("🏆 Team Awards Theme Extractor")
st.markdown("Upload your Excel file to extract award themes and paste them into Microsoft Forms.")

# File uploader
uploaded_file = st.file_uploader("📤 Upload your Excel (.xlsx) file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)
        st.success("✅ File uploaded successfully!")

        # Extract text from the first column
        column_data = df.iloc[:, 0].dropna().astype(str).tolist()

        # Regex to match lines that look like themes
        pattern = r'^[\W_]{0,2}.*?(Theme|Award).*?:\s*“.*?”'
        themes_with_details = [
            re.sub(r'\s*\n\s*', ' ', item).strip()
            for item in column_data
            if re.search(pattern, item)
        ]

        if themes_with_details:
            formatted_text = "\n\n".join(themes_with_details)
            st.subheader("🎯 Extracted Themes + Descriptions")
            st.text_area("👇 Copy this and paste into Microsoft Forms:", formatted_text, height=400)

            # Button to open Microsoft Forms
            if st.button("🚀 Open Microsoft Forms"):
                js = "window.open('https://forms.office.com')"  # New tab
                st.markdown(f"<script>{js}</script>", unsafe_allow_html=True)
        else:
            st.warning("⚠️ No themes found. Make sure the first column contains valid award theme entries.")

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")

st.markdown("---")
st.caption("Built with ❤️ using Streamlit. This app runs entirely in your browser.")
