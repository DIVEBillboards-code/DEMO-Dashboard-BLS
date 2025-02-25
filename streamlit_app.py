import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Survey Visualizer", layout="wide")

# Custom CSS
st.markdown("""
<style>
    .main { padding: 2rem; }
    .stTabs [data-baseweb="tab-list"] { gap: 2rem; }
    .stTabs [data-baseweb="tab"] {
        height: 50px; background-color: #f0f2f6; border-radius: 5px 5px 0px 0px; padding: 1rem;
    }
    .stTabs [aria-selected="true"] { background-color: #ffffff; border-bottom: 2px solid #4e89ae; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Survey Data Visualizer")
st.write("Upload your survey Excel file to generate tailored visualizations.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

# Simplified column type detection
def get_numeric_columns(df):
    return [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col]) and col != "User ID"]

def get_categorical_columns(df):
    return [col for col in df.columns if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_categorical_dtype(df[col])]

# Define ordinal sequences
ordinal_sequences = {
    "Brand image": ["Muy negativa", "Negativa", "Neutra", "Positiva", "Muy positiva"],
    "Frequency": ["Nunca", "Una vez al mes", "Pocas veces al mes", "Una vez a la semana", "Varias veces a la semana"],
    "Ad recall": ["No", "Sí, una vez", "Sí, varias veces"]
}

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        if not sheet_names:
            st.error("No sheets found in the Excel file.")
        else:
            selected_sheet = st.selectbox("Select a sheet:", sheet_names)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

            # Debug: Show raw data
            st.write("### Raw Data Preview")
            st.dataframe(df.head())

            # Identify column types
            numeric_cols = get_numeric_columns(df)
            categorical_cols = get_categorical_columns(df)

            # Debug: Display detected columns
            st.write("### Detected Column Types")
            st.write("Numeric Columns:", numeric_cols)
            st.write("Categorical Columns:", categorical_cols)

            # Handle ordinal variables
            ordinal_cols = []
            for col in numeric_cols[:]:  # Copy to avoid modifying list during iteration
                series = df[col].dropna().astype(float)
                is_ordinal = (series.apply(lambda x: x.is_integer()).all() and 
                              series.min() >= 0 and series.max() <= 10 and series.nunique() <= 11)
                if is_ordinal:
                    df[col] = pd.Categorical(df[col], categories=range(0, 11), ordered=True)
                    numeric_cols.remove(col)
                    categorical_cols.append(col)
                    ordinal_cols.append(col)

            for col in categorical_cols[:]:
                unique_vals = df[col].dropna().unique()
                for seq_name, seq in ordinal_sequences.items():
                    if set(unique_vals).issubset(seq):
                        df[col] = pd.Categorical(df[col], categories=seq, ordered=True)
                        ordinal_cols.append(col)
                        break

            # Debug: Updated column types after ordinal handling
            st.write("### Updated Column Types After Ordinal Detection")
            st.write("Numeric Columns:", numeric_cols)
            st.write("Categorical Columns:", categorical_cols)
            st.write("Ordinal Columns:", ordinal_cols)

            # Sidebar Filters
            st.sidebar.header("Filters")
            filter_cols = [col for col in df.columns if df[col].nunique() <= 20 and col != "User ID"]
            filters = {}
            for col in filter_cols:
                unique_vals = df[col].dropna().unique()
                selected_vals = st.sidebar.multiselect(f"Select {col}", unique_vals, default=unique_vals)
                filters[col] = selected_vals

            # Apply filters
            filtered_df = df.copy()
            for col, vals in filters.items():
                filtered_df = filtered_df[filtered_df[col].isin(vals)]

            # Weight column selection
            weight_col = st.sidebar.selectbox("Select weight column (optional):", 
                                              ["None"] + list(df.columns), index=0)

            # Data Overview
            st.subheader("Data Overview")
            col1, col2, col3 = st.columns(3)
            col1.metric("Rows", filtered_df.shape[0])
            col2.metric("Columns", filtered_df.shape[1])
            col3.metric("Missing Values", filtered_df.isna().sum().sum())
            with st.expander("View Filtered Data", expanded=False):
                st.dataframe(filtered_df)

            # Tabs
            tab1, tab2, tab3, tab4 = st.tabs(["📈 Distribution", "🔄 Correlation", "📊 Categorical", "📅 Time Series"])

            # Distribution Tab
            with tab1:
                st.subheader("Distribution Analysis")
                all_cols = numeric_cols + categorical_cols
                st.write("Available columns for distribution:", all_cols)
                if all_cols:
                    selected_col = st.selectbox("Select a column for distribution:", all_cols, key="dist_col")
                    if selected_col in numeric_cols:
                        st.write(f"Plotting histogram for numeric column: {selected_col}")
                        fig_hist = px.histogram(filtered_df, x=selected_col, title=f"Histogram of {selected_col}")
                        st.plotly_chart(fig_hist, use_container_width=True)
                        fig_box = px.box(filtered_df, y=selected_col, title=f"Box Plot of {selected_col}")
                        st.plotly_chart(fig_box, use_container_width=True)
                    else:
                        st.write(f"Plotting bar chart for categorical/ordinal column: {selected_col}")
                        if weight_col != "None" and weight_col in filtered_df.columns:
                            counts = filtered_df.groupby(selected_col, observed=True)[weight_col].sum().reset_index()
                            counts.columns = [selected_col, "Weighted Count"]
                        else:
                            counts = filtered_df[selected_col].value_counts().reset_index()
                            counts.columns = [selected_col, "Count"]
                        if selected_col in ordinal_cols:
                            order_key = [k for k, v in ordinal_sequences.items() if selected_col in k]
                            if order_key:
                                order = ordinal_sequences[order_key[0]]
                                counts[selected_col] = pd.Categorical(counts[selected_col], categories=order, ordered=True)
                                counts = counts.sort_values(selected_col)
                        fig = px.bar(counts, x=selected_col, y=counts.columns[1], title=f"Distribution of {selected_col}")
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("No columns available for distribution analysis.")

            # Correlation Tab
            with tab2:
                st.subheader("Correlation Analysis")
                ordinal_cat_cols = [col for col in categorical_cols if pd.api.types.is_categorical_dtype(df[col]) and df[col].cat.ordered]
                corr_cols = numeric_cols + ordinal_cat_cols
                st.write("Columns available for correlation:", corr_cols)
                if len(corr_cols) >= 2:
                    selected_corr_cols = st.multiselect("Select columns for correlation:", corr_cols, default=corr_cols[:min(5, len(corr_cols))])
                    if len(selected_corr_cols) >= 2:
                        corr_method = st.radio("Correlation method:", ["Pearson", "Spearman"])
                        df_corr = filtered_df[selected_corr_cols].copy()
                        for col in selected_corr_cols:
                            if col in ordinal_cat_cols:
                                df_corr[col] = filtered_df[col].cat.codes
                        corr_matrix = df_corr.corr(method=corr_method.lower())
                        fig, ax = plt.subplots(figsize=(10, 8))
                        sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0, ax=ax)
                        st.pyplot(fig)
                    else:
                        st.info("Select at least 2 columns for correlation.")
                else:
                    st.warning("Need at least 2 columns for correlation analysis.")

            # Categorical Tab
            with tab3:
                st.subheader("Categorical Analysis")
                st.write("Categorical columns available:", categorical_cols)
                if categorical_cols:
                    selected_cat_col = st.selectbox("Select a categorical column:", categorical_cols, key="cat_col")
                    plot_type = st.radio("Plot type:", ["Bar", "Pie"])
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        counts = filtered_df.groupby(selected_cat_col, observed=True)[weight_col].sum().reset_index()
                        counts.columns = [selected_cat_col, "Weighted Count"]
                    else:
                        counts = filtered_df[selected_cat_col].value_counts().reset_index()
                        counts.columns = [selected_cat_col, "Count"]
                    if selected_cat_col in ordinal_cols:
                        order_key = [k for k, v in ordinal_sequences.items() if selected_cat_col in k]
                        if order_key:
                            order = ordinal_sequences[order_key[0]]
                            counts[selected_cat_col] = pd.Categorical(counts[selected_cat_col], categories=order, ordered=True)
                            counts = counts.sort_values(selected_cat_col)
                    if plot_type == "Bar":
                        fig = px.bar(counts, x=selected_cat_col, y=counts.columns[1], title=f"{selected_cat_col}")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        fig = px.pie(counts, names=selected_cat_col, values=counts.columns[1], title=f"{selected_cat_col}")
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("No categorical columns available.")

            # Time Series Tab
            with tab4:
                st.subheader("Time Series Analysis")
                st.info("No datetime columns detected in your data.")

            # Download Options
            st.subheader("Download Options")
            download_format = st.radio("Select download format:", ["CSV", "Excel", "JSON"])
            if st.button("Download Processed Data"):
                if download_format == "CSV":
                    csv = filtered_df.to_csv(index=False)
                    b64 = BytesIO(csv.encode())
                    st.download_button(label="Download CSV", data=b64, file_name=f"{selected_sheet}.csv", mime="text/csv")
                elif download_format == "Excel":
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        filtered_df.to_excel(writer, sheet_name=selected_sheet, index=False)
                    st.download_button(label="Download Excel", data=output, file_name=f"{selected_sheet}.xlsx", 
                                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    json_str = filtered_df.to_json(orient='records')
                    b64 = BytesIO(json_str.encode())
                    st.download_button(label="Download JSON", data=b64, file_name=f"{selected_sheet}.json", mime="application/json")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("👈 Upload an Excel file to begin.")
    st.subheader("Sample Visualizations:")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Distribution**")
        st.image("https://via.placeholder.com/400x300?text=Histogram+and+Bar+Charts", use_column_width=True)
    with col2:
        st.markdown("**Categorical**")
        st.image("https://via.placeholder.com/400x300?text=Pie+and+Bar+Charts", use_column_width=True)
    st.markdown("""
    ### How to use:
    1. Upload your survey Excel file
    2. Apply filters and select weights
    3. Explore visualizations
    4. Download your data
    """)

st.markdown("---\nCreated with Streamlit • Survey Visualizer App")
