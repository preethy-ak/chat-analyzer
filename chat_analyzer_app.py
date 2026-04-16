import streamlit as st
import pandas as pd

st.set_page_config(page_title="Chat Analyzer", layout="wide")

st.title("📊 Chat Analyzer Dashboard")

# Upload file
uploaded_file = st.file_uploader("Upload your CSV/Excel file", type=["csv", "xlsx"])

if uploaded_file:

    # Read file
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        st.stop()

    # ✅ Normalize column names
    df.columns = df.columns.str.strip().str.upper()

    st.write("### 🔍 Columns detected:")
    st.write(df.columns.tolist())

    # ✅ Required columns
    REQUIRED_COLS = ['STORE_CODE', 'CUSTOMER_NAME', 'SENTIMENT']

    # ✅ Create missing columns instead of failing
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = ""

    # ✅ Clean data (VERY IMPORTANT)
    df['STORE_CODE'] = df['STORE_CODE'].fillna("UNKNOWN").astype(str)
    df['CUSTOMER_NAME'] = df['CUSTOMER_NAME'].fillna("UNKNOWN").astype(str)
    df['SENTIMENT'] = df['SENTIMENT'].fillna("").astype(str).str.lower().str.strip()

    # Optional: normalize sentiment variations
    df['SENTIMENT'] = df['SENTIMENT'].replace({
        'anger': 'angry',
        'angry ': 'angry',
        'angry.': 'angry'
    })

    # ✅ Safe grouping function
    def store_summary(group):
        try:
            angry_df = group[group['SENTIMENT'] == 'angry']

            angry_names = (
                angry_df['CUSTOMER_NAME']
                .dropna()
                .unique()
                .tolist()
            )

            return pd.Series({
                'TOTAL_CHATS': int(len(group)),
                'ANGRY_CUSTOMERS': int(len(angry_names)),
                'ANGRY_NAMES': ", ".join(angry_names) if angry_names else "-"
            })

        except Exception as e:
            # Fail-safe (never crash app)
            return pd.Series({
                'TOTAL_CHATS': int(len(group)),
                'ANGRY_CUSTOMERS': 0,
                'ANGRY_NAMES': "ERROR"
            })

    # ✅ Group safely (no apply crash)
    try:
        summary_df = (
            df.groupby('STORE_CODE', dropna=False)
              .apply(store_summary)
              .reset_index()
        )
    except Exception as e:
        st.error(f"❌ Processing error: {e}")
        st.stop()

    # ✅ Rename columns for UI
    summary_df = summary_df.rename(columns={
        'STORE_CODE': '🏬 Store Code',
        'TOTAL_CHATS': '💬 Total Chats',
        'ANGRY_CUSTOMERS': '😡 Angry Customers',
        'ANGRY_NAMES': '⚠️ Angry Customer Names'
    })

    # ✅ Display
    st.write("### 📈 Store Summary")
    st.dataframe(summary_df, use_container_width=True)

    # ✅ Download
    csv = summary_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Summary",
        data=csv,
        file_name="chat_summary.csv",
        mime="text/csv"
    )

else:
    st.info("👆 Please upload a file to begin.")
