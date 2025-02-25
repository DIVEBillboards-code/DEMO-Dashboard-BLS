import streamlit as st
import pandas as pd
import plotly.express as px
import scipy.stats as stats
from io import BytesIO
from fpdf import FPDF
import plotly.io as pio
import os

# Page configuration
st.set_page_config(page_title="Excel Data Visualizer", layout="wide")

# Custom CSS for better UX
st.markdown("""
<style>
    .main { padding: 2rem; background-color: #f9f9f9; }
    .stTabs [data-baseweb="tab-list"] { gap: 1.5rem; }
    .stTabs [data-baseweb="tab"] {
        height: 50px; background-color: #e6f1fa; border-radius: 8px 8px 0 0; padding: 1rem;
        font-weight: bold; color: #333;
    }
    .stTabs [aria-selected="true"] { background-color: #ffffff; border-bottom: 3px solid #0078d4; color: #0078d4; }
    .sidebar .sidebar-content { background-color: #f0f2f6; padding: 1rem; }
    h1 { color: #0078d4; font-family: 'Segoe UI', sans-serif; }
    h2, h3 { color: #333; font-family: 'Segoe UI', sans-serif; }
    .stButton>button { background-color: #0078d4; color: white; border-radius: 5px; }
    .stButton>button:hover { background-color: #005a9e; }
</style>
""", unsafe_allow_html=True)

# Title and description
st.title("📊 Excel Data Visualizer")
st.write("Explore your data with interactive visualizations and statistical insights!")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file here", type=["xlsx", "xls"], help="Supports .xlsx and .xls formats.")

# Function to detect column types
def detect_survey_columns(df):
    numeric_cols = []
    categorical_cols = []
    ordinal_cols = []

    for col in df.columns:
        if "id" in col.lower() or df[col].nunique() > 0.5 * len(df):
            continue

        if pd.api.types.is_numeric_dtype(df[col]):
            series = df[col].dropna().astype(float)
            if series.nunique() > 20 and not series.apply(lambda x: x.is_integer()).all():
                numeric_cols.append(col)
            elif series.nunique() <= 10 or (series.min() >= 0 and series.max() <= 10 and series.apply(lambda x: x.is_integer()).all()):
                df[col] = pd.Categorical(df[col], categories=sorted(series.unique()), ordered=True)
                ordinal_cols.append(col)
            else:
                numeric_cols.append(col)

        elif pd.api.types.is_object_dtype(df[col]):
            unique_vals = df[col].nunique()
            sample_vals = df[col].dropna().unique()
            ordinal_indicators = ["muy", "negativa", "positiva", "nunca", "siempre", "frecuente", "low", "medium", "high", "yes", "no"]
            if unique_vals <= 10 and any(ind.lower() in " ".join(str(val).lower() for val in sample_vals) for ind in ordinal_indicators):
                df[col] = pd.Categorical(df[col], categories=sample_vals, ordered=True)
                ordinal_cols.append(col)
            else:
                categorical_cols.append(col)

    brand_image_col = '[Brand image] This is an advertisement for Cetaphil. What image does it give you of Cetaphil?'
    if brand_image_col in categorical_cols:
        categorical_cols.remove(brand_image_col)
        ordinal_cols.append(brand_image_col)
        df[brand_image_col] = pd.Categorical(df[brand_image_col], 
                                            categories=["Negative", "Neutral", "Positive"], ordered=True)

    attribution_col = '[Attribution] In your opinion, this ad is for:'
    if attribution_col in ordinal_cols:
        ordinal_cols.remove(attribution_col)
        categorical_cols.append(attribution_col)

    return numeric_cols, categorical_cols, ordinal_cols

# Function to create PDF report
def create_pdf(filtered_df, numeric_cols, categorical_cols, ordinal_cols, figures):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Title
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Excel Data Visualizer Report", ln=True, align="C")

    # Data Overview
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Data Overview", ln=True)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"Rows: {filtered_df.shape[0]}", ln=True)
    pdf.cell(0, 6, f"Columns: {filtered_df.shape[1]}", ln=True)
    pdf.cell(0, 6, f"Missing Values: {filtered_df.isna().sum().sum()}", ln=True)

    # Column Types
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Column Types", ln=True)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"Numeric (Continuous): {', '.join(numeric_cols)}", ln=True)
    pdf.cell(0, 6, f"Categorical (Unordered): {', '.join(categorical_cols)}", ln=True)
    pdf.cell(0, 6, f"Ordinal (Survey-like): {', '.join(ordinal_cols)}", ln=True)

    # Data Table
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Filtered Data Preview", ln=True)
    pdf.set_font("Arial", "", 8)
    col_width = pdf.w / (len(filtered_df.columns) + 1)
    row_height = 6

    for col in filtered_df.columns:
        pdf.cell(col_width, row_height, str(col), border=1)
    pdf.ln(row_height)

    for i, row in filtered_df.head(10).iterrows():
        for value in row:
            pdf.cell(col_width, row_height, str(value)[:20], border=1)
        pdf.ln(row_height)

    # Add Graphs
    pdf.set_font("Arial", "B", 12)
    for fig_name, fig in figures.items():
        pdf.add_page()
        pdf.cell(0, 10, f"Visualization: {fig_name}", ln=True)
        img_path = f"{fig_name.replace(' ', '_')}.png"
        pio.write_image(fig, img_path, width=800, height=600)
        pdf.image(img_path, x=10, y=pdf.get_y() + 5, w=190)
        os.remove(img_path)  # Clean up temporary file

    pdf_output = BytesIO()
    pdf_output.write(pdf.output(dest='S').encode('latin1'))
    pdf_output.seek(0)
    return pdf_output

# Main app logic
if uploaded_file:
    with st.spinner("Processing your file..."):
        try:
            df = pd.read_excel(uploaded_file)
            st.success("File uploaded successfully!")

            # Sidebar controls
            with st.sidebar:
                st.header("Controls")
                with st.expander("Filters", expanded=True):
                    filterable_cols = [col for col in df.columns if df[col].nunique() <= 20]
                    filters = {}
                    for col in filterable_cols:
                        unique_vals = df[col].dropna().unique()
                        selected_vals = st.multiselect(f"{col}", unique_vals, default=unique_vals, key=f"filter_{col}")
                        filters[col] = selected_vals
                    if st.button("Reset Filters"):
                        st.rerun()

                numeric_cols, categorical_cols, ordinal_cols = detect_survey_columns(df)
                weight_options = ["None"] + numeric_cols
                weight_col = st.selectbox("Apply weights (optional)", weight_options, index=0, 
                                        help="Choose a numeric column to weight the data.")

            # Apply filters
            filtered_df = df.copy()
            for col, vals in filters.items():
                filtered_df = filtered_df[filtered_df[col].isin(vals)]
            if filtered_df.empty:
                st.warning("Filters resulted in no data. Showing full dataset instead.")
                filtered_df = df.copy()

            # Data overview
            st.subheader("Data Overview", anchor="overview")
            col1, col2, col3 = st.columns(3)
            col1.metric("Rows", filtered_df.shape[0])
            col2.metric("Columns", filtered_df.shape[1])
            col3.metric("Missing Values", filtered_df.isna().sum().sum())

            with st.expander("View Data Preview"):
                st.dataframe(filtered_df.head())
            with st.expander("Column Types"):
                st.write(f"**Numeric (Continuous):** {numeric_cols}")
                st.write(f"**Categorical (Unordered):** {categorical_cols}")
                st.write(f"**Ordinal (Survey-like):** {ordinal_cols}")

            # Dictionary to store figures for PDF
            figures = {}

            # Tabs
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Overview", "📈 Insights", "🔄 Explore", "🌟 Profiles", "📉 Stats"])

            # Placeholder for existing tabs (replace with your visualizations)
            with tab1:
                st.subheader("Overview")
                # Example: Add your pie charts, bar charts, etc.
                pass

            with tab2:
                st.subheader("Insights")
                # Example: Add your bar charts, heatmaps, etc.
                pass

            with tab3:
                st.subheader("Explore")
                # Example: Add your box plots, scatter plots, etc.
                pass

            with tab4:
                st.subheader("Profiles")
                # Example: Add your radar charts, etc.
                pass

            # Stats Tab
            with tab5:
                st.subheader("Statistical Insights")

                # Correlation Matrix
                with st.expander("Correlation Matrix", expanded=False):
                    st.write("**Correlation Matrix**: Visualizes the strength and direction of relationships between numeric variables. Values range from -1 (strong negative) to 1 (strong positive).")
                    numeric_cols_selected = st.multiselect("Select numeric columns for correlation", numeric_cols, default=numeric_cols, key="corr_cols")
                    if st.button("Generate Correlation Matrix", key="corr_button"):
                        if len(numeric_cols_selected) < 2:
                            st.warning("Please select at least two numeric columns.")
                        else:
                            corr_matrix = filtered_df[numeric_cols_selected].corr()
                            fig_corr = px.imshow(corr_matrix, text_auto=True, aspect="auto", 
                                                 title="Correlation Matrix", color_continuous_scale="RdBu", 
                                                 template="plotly_white")
                            st.plotly_chart(fig_corr, use_container_width=True)
                            figures["Correlation Matrix"] = fig_corr

                # Summary Statistics
                with st.expander("Summary Statistics", expanded=False):
                    st.write("**Summary Statistics**: Provides key metrics like mean, median, and standard deviation for selected columns.")
                    summary_cols = st.multiselect("Select columns for summary statistics", 
                                                  numeric_cols + categorical_cols + ordinal_cols, 
                                                  default=numeric_cols, key="summary_cols")
                    if st.button("Generate Summary Statistics", key="summary_button"):
                        if not summary_cols:
                            st.warning("Please select at least one column.")
                        else:
                            summary = filtered_df[summary_cols].describe(include='all')
                            st.dataframe(summary)

                # Statistical Tests
                with st.expander("Statistical Tests", expanded=False):
                    st.write("Perform basic statistical tests to uncover relationships in your data.")

                    # Chi-square Test
                    st.write("**Chi-square Test**: Determines if there is a significant association between two categorical variables.")
                    cat_col1 = st.selectbox("Select first categorical column", categorical_cols, key="chi_cat1")
                    cat_col2 = st.selectbox("Select second categorical column", 
                                            [col for col in categorical_cols if col != cat_col1], key="chi_cat2")
                    if st.button("Run Chi-square Test", key="chi_button"):
                        if cat_col1 and cat_col2:
                            contingency_table = pd.crosstab(filtered_df[cat_col1], filtered_df[cat_col2])
                            chi2, p, dof, expected = stats.chi2_contingency(contingency_table)
                            st.write(f"Chi-square statistic: {chi2:.2f}")
                            st.write(f"P-value: {p:.4f}")
                            if p < 0.05:
                                st.write("Interpretation: There is a significant association between the two variables (p < 0.05).")
                            else:
                                st.write("Interpretation: There is no significant association between the two variables (p >= 0.05).")
                        else:
                            st.warning("Please select two categorical columns.")

                    # T-test
                    st.write("**T-test**: Compares the means of a numeric variable between two groups to see if they are significantly different.")
                    numeric_col = st.selectbox("Select numeric column", numeric_cols, key="ttest_num")
                    group_col = st.selectbox("Select categorical column with two groups", 
                                             [col for col in categorical_cols if filtered_df[col].nunique() == 2], key="ttest_group")
                    if st.button("Run T-test", key="ttest_button"):
                        if numeric_col and group_col:
                            group1 = filtered_df[filtered_df[group_col] == filtered_df[group_col].unique()[0]][numeric_col]
                            group2 = filtered_df[filtered_df[group_col] == filtered_df[group_col].unique()[1]][numeric_col]
                            t_stat, p_val = stats.ttest_ind(group1, group2, equal_var=False)  # Welch's t-test
                            st.write(f"T-statistic: {t_stat:.2f}")
                            st.write(f"P-value: {p_val:.4f}")
                            if p_val < 0.05:
                                st.write("Interpretation: There is a significant difference in means between the two groups (p < 0.05).")
                            else:
                                st.write("Interpretation: There is no significant difference in means between the two groups (p >= 0.05).")
                        else:
                            st.warning("Please select a numeric column and a categorical column with exactly two unique values.")

            # Download Section
            st.subheader("Download Your Data")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Download as CSV", help="Save the filtered data as a CSV file"):
                    csv = filtered_df.to_csv(index=False)
                    st.download_button(label="Download CSV", data=csv, file_name="processed_data.csv", mime="text/csv")
            with col2:
                if st.button("Download as PDF", help="Save data and graphs as a PDF"):
                    with st.spinner("Generating PDF with graphs..."):
                        pdf_output = create_pdf(filtered_df, numeric_cols, categorical_cols, ordinal_cols, figures)
                        st.download_button(label="Download PDF", data=pdf_output, file_name="processed_data_with_graphs.pdf", mime="application/pdf")

        except Exception as e:
            st.error(f"Error processing file: {e}")
else:
    st.info("👈 Upload an Excel file to get started!")
    with st.expander("How to Use This Tool"):
        st.markdown("""
        - **Upload**: Drop your Excel file (.xlsx or .xls) to analyze.
        - **Filter**: Refine your data in the sidebar.
        - **Explore Tabs**:
          - **Overview**: Distributions and breakdowns.
          - **Insights**: Comparisons and heatmaps.
          - **Explore**: Relationships and histograms.
          - **Profiles**: Radar charts for multi-variable analysis.
          - **Stats**: Statistical insights like correlation, summary stats, and tests.
        - **Download**: Save as CSV or PDF with graphs!
        """)

# Footer
st.markdown("---\nCreated with ❤️ for DIVE")
