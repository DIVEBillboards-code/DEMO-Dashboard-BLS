import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Excel Data Visualizer", layout="wide")

# Custom CSS for layout
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

st.title("📊 Excel Data Visualizer")
st.write("Upload any Excel file to automatically detect and visualize survey-like questions.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

def detect_survey_columns(df):
    """Detect columns that resemble survey questions."""
    numeric_cols = []  # Continuous numeric data
    categorical_cols = []  # Unordered categorical data
    ordinal_cols = []  # Ordered categorical or numeric scales
    
    for col in df.columns:
        # Skip columns that are likely identifiers (e.g., "ID" in name or too many unique values)
        if "id" in col.lower() or df[col].nunique() > 0.5 * len(df):
            continue
        
        # Numeric columns
        if pd.api.types.is_numeric_dtype(df[col]):
            series = df[col].dropna().astype(float)
            # Check for ordinal scales (e.g., 0-10, integers, limited unique values)
            if (series.apply(lambda x: x.is_integer()).all() and 
                series.min() >= 0 and series.max() <= 10 and series.nunique() <= 11):
                df[col] = pd.Categorical(df[col], categories=range(int(series.min()), int(series.max()) + 1), ordered=True)
                ordinal_cols.append(col)
            else:
                numeric_cols.append(col)
        
        # Categorical columns (text or object types)
        elif pd.api.types.is_object_dtype(df[col]):
            unique_vals = df[col].nunique()
            # Heuristic: <= 10 unique values might be a survey response (e.g., Yes/No, ratings)
            if unique_vals <= 10:
                # Check if it resembles an ordinal scale (e.g., "Low", "Medium", "High")
                sample_vals = df[col].dropna().unique()
                ordinal_indicators = ["muy", "negativa", "positiva", "nunca", "siempre", "frecuente", "low", "medium", "high"]
                if any(ind.lower() in " ".join(str(val).lower() for val in sample_vals) for ind in ordinal_indicators):
                    df[col] = pd.Categorical(df[col], categories=sample_vals, ordered=True)
                    ordinal_cols.append(col)
                else:
                    categorical_cols.append(col)
            else:
                categorical_cols.append(col)  # Treat as categorical if too many unique values

    return numeric_cols, categorical_cols, ordinal_cols

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.write("### Data Preview")
        st.dataframe(df.head())

        # Detect survey-like columns
        numeric_cols, categorical_cols, ordinal_cols = detect_survey_columns(df)
        st.write("### Detected Column Types")
        st.write(f"Numeric (Continuous): {numeric_cols}")
        st.write(f"Categorical (Unordered): {categorical_cols}")
        st.write(f"Ordinal (Survey-like): {ordinal_cols}")

        # Sidebar Filters
        st.sidebar.header("Filters")
        filters = {}
        filterable_cols = [col for col in df.columns if df[col].nunique() <= 20]
        for col in filterable_cols:
            unique_vals = df[col].dropna().unique()
            selected_vals = st.sidebar.multiselect(f"Filter {col}", unique_vals, default=unique_vals)
            filters[col] = selected_vals

        # Apply filters
        filtered_df = df.copy()
        for col, vals in filters.items():
            filtered_df = filtered_df[filtered_df[col].isin(vals)]

        # Weight column selection (for numeric columns only)
        weight_options = ["None"] + numeric_cols
        weight_col = st.sidebar.selectbox("Apply weights (optional):", weight_options, index=0)

        # Data Overview
        st.subheader("Data Overview")
        col1, col2, col3 = st.columns(3)
        col1.metric("Rows", filtered_df.shape[0])
        col2.metric("Columns", filtered_df.shape[1])
        col3.metric("Missing Values", filtered_df.isna().sum().sum())

        # Visualization Tabs
        tab1, tab2, tab3 = st.tabs(["📊 Distributions", "📈 Comparisons", "🔄 Relationships"])

        # Distributions Tab
        with tab1:
            st.subheader("Distributions")
            col1, col2 = st.columns(2)

            with col1:
                # Numeric Distribution
                if numeric_cols:
                    num_col = st.selectbox("Select a numeric column:", numeric_cols, key="num_dist")
                    fig_num = px.histogram(filtered_df, x=num_col, title=f"Distribution of {num_col}", 
                                         nbins=min(50, filtered_df[num_col].nunique()))
                    st.plotly_chart(fig_num, use_container_width=True)
                else:
                    st.info("No continuous numeric columns detected.")

            with col2:
                # Survey-like (Ordinal/Categorical) Distribution
                survey_cols = ordinal_cols + categorical_cols
                if survey_cols:
                    survey_col = st.selectbox("Select a survey question:", survey_cols, key="survey_dist")
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        counts = filtered_df.groupby(survey_col, observed=True)[weight_col].sum().reset_index()
                        counts.columns = [survey_col, "Weighted Count"]
                    else:
                        counts = filtered_df[survey_col].value_counts().reset_index()
                        counts.columns = [survey_col, "Count"]
                    if survey_col in ordinal_cols:
                        counts[survey_col] = pd.Categorical(counts[survey_col], 
                                                          categories=filtered_df[survey_col].cat.categories, ordered=True)
                        counts = counts.sort_values(survey_col)
                    fig_survey = px.bar(counts, x=survey_col, y=counts.columns[1], title=f"Distribution of {survey_col}")
                    st.plotly_chart(fig_survey, use_container_width=True)
                else:
                    st.info("No survey-like columns detected.")

        # Comparisons Tab
        with tab2:
            st.subheader("Comparisons")
            col1, col2 = st.columns(2)

            with col1:
                # Survey vs Numeric
                survey_cols = ordinal_cols + categorical_cols
                if survey_cols and numeric_cols:
                    survey_x = st.selectbox("Select survey question (X):", survey_cols, key="comp_survey")
                    num_y = st.selectbox("Select numeric (Y):", numeric_cols, key="comp_num")
                    fig_comp = px.box(filtered_df, x=survey_x, y=num_y, title=f"{num_y} by {survey_x}")
                    st.plotly_chart(fig_comp, use_container_width=True)
                else:
                    st.info("Need both survey-like and numeric columns for comparison.")

            with col2:
                # Survey vs Survey
                if len(survey_cols) >= 2:
                    survey_x2 = st.selectbox("Select survey question (X):", survey_cols, key="comp_survey_x")
                    survey_y2 = st.selectbox("Select survey question (Y):", 
                                           [col for col in survey_cols if col != survey_x2], key="comp_survey_y")
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        cross_tab = filtered_df.groupby([survey_x2, survey_y2], observed=True)[weight_col].sum().reset_index()
                        cross_tab.columns = [survey_x2, survey_y2, "Weighted Count"]
                    else:
                        cross_tab = pd.crosstab(filtered_df[survey_x2], filtered_df[survey_y2]).reset_index()
                        cross_tab = cross_tab.melt(id_vars=[survey_x2], var_name=survey_y2, value_name="Count")
                    fig_cross = px.bar(cross_tab, x=survey_x2, y="Count", color=survey_y2, 
                                      title=f"{survey_x2} vs {survey_y2}", barmode="group")
                    st.plotly_chart(fig_cross, use_container_width=True)
                else:
                    st.info("Need at least two survey-like columns for comparison.")

        # Relationships Tab
        with tab3:
            st.subheader("Relationships")
            col1, col2 = st.columns(2)

            with col1:
                # Numeric vs Numeric
                if len(numeric_cols) >= 2:
                    num_x = st.selectbox("Select numeric (X):", numeric_cols, key="rel_num_x")
                    num_y = st.selectbox("Select numeric (Y):", [col for col in numeric_cols if col != num_x], key="rel_num_y")
                    fig_scatter = px.scatter(filtered_df, x=num_x, y=num_y, title=f"{num_x} vs {num_y}", trendline="ols")
                    st.plotly_chart(fig_scatter, use_container_width=True)
                else:
                    st.info("Need at least two numeric columns for scatter plot.")

            with col2:
                # Survey vs Numeric/Survey
                survey_cols = ordinal_cols + categorical_cols
                if survey_cols and (numeric_cols or len(survey_cols) >= 2):
                    survey_x = st.selectbox("Select survey question (X):", survey_cols, key="rel_survey_x")
                    y_options = numeric_cols + [col for col in survey_cols if col != survey_x]
                    y_col = st.selectbox("Select Y (numeric or survey):", y_options, key="rel_y")
                    if y_col in numeric_cols:
                        fig_rel = px.box(filtered_df, x=survey_x, y=y_col, title=f"{y_col} by {survey_x}")
                    else:
                        fig_rel = px.box(filtered_df, x=survey_x, y=filtered_df[y_col].cat.codes if y_col in ordinal_cols else y_col, 
                                       title=f"{y_col} (codes if ordinal) by {survey_x}")
                    st.plotly_chart(fig_rel, use_container_width=True)
                else:
                    st.info("Need survey-like and numeric/survey columns for relationship analysis.")

        # Download Option
        st.subheader("Download Data")
        if st.button("Download as CSV"):
            csv = filtered_df.to_csv(index=False)
            st.download_button(label="Download CSV", data=csv, file_name="processed_data.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("👈 Please upload an Excel file to see visualizations.")
    st.markdown("""
    ### How It Works:
    - Automatically detects survey-like questions (e.g., ratings, yes/no, categories).
    - Visualizes distributions, comparisons, and relationships.
    - Supports filtering and weighting.
    """)

st.markdown("---\nCreated with Streamlit • Universal Excel Visualizer")
