import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="HackerRank Monthly Report Combiner", layout="centered")

st.title("📊 HackerRank Monthly Report Combiner")
st.markdown("Upload **weekly CSV or Excel files** to generate a single **monthly report**.")

uploaded_files = st.file_uploader("📁 Upload CSV or Excel files", type=["csv", "xlsx"], accept_multiple_files=True)

def read_file(file):
    try:
        if file.name.endswith(".csv"):
            return pd.read_csv(io.StringIO(file.getvalue().decode("utf-8")))
        elif file.name.endswith(".xlsx"):
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"❌ Error reading file {file.name}: {e}")
        return None

def standardize_column_names(df):
    df.columns = df.columns.str.strip().str.lower()
    return df

def find_identifier_column(columns):
    for possible_id in ['email', 'id', 'name', 'student id', 'roll no']:
        for col in columns:
            if possible_id in col:
                return col
    return None

if uploaded_files:
    weekly_data = []
    identifier_col = None

    for idx, file in enumerate(uploaded_files):
        df = read_file(file)
        if df is None:
            st.stop()

        df = standardize_column_names(df)

        if identifier_col is None:
            identifier_col = find_identifier_column(df.columns)
            if not identifier_col:
                st.error(f"❌ Could not find identifier (email, id, or name) in {file.name}")
                st.stop()

        score_col = df.select_dtypes(include='number').columns
        if len(score_col) == 0:
            st.error(f"❌ No numeric 'score' column found in {file.name}")
            st.stop()
        score_col = score_col[0]

        week_label = f"Week {idx+1}"
        temp_df = df[[identifier_col, score_col]].copy()
        temp_df.columns = [identifier_col, week_label]
        weekly_data.append(temp_df)

    merged_df = weekly_data[0]
    for df in weekly_data[1:]:
        merged_df = pd.merge(merged_df, df, on=identifier_col, how='outer')

    merged_df = merged_df.fillna("-")

    st.success("✅ Successfully combined files!")
    st.write(merged_df)

    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Monthly Report")
        return output.getvalue()

    st.download_button(
        label="📥 Download Combined Report",
        data=to_excel(merged_df),
        file_name="Monthly_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
