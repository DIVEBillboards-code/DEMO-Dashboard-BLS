import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Excel Data Visualizer", layout="wide")

# Enhanced CSS for better UX
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

st.title("📊 Excel Data Visualizer")
st.write("Explore your data with interactive visualizations!")

# File Uploader
uploaded_file = st.file_uploader("Upload your Excel file here", type=["xlsx", "xls"], help="Supports .xlsx and .xls formats.")

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

if uploaded_file:
    with st.spinner("Processing your file..."):
        try:
            df = pd.read_excel(uploaded_file)
            st.success("File uploaded successfully!")
            
            # Sidebar Controls
            with st.sidebar:
                st.header("Controls")
                
                # Filters
                with st.expander("Filters", expanded=True):
                    filterable_cols = [col for col in df.columns if df[col].nunique() <= 20]
                    filters = {}
                    for col in filterable_cols:
                        unique_vals = df[col].dropna().unique()
                        selected_vals = st.multiselect(f"{col}", unique_vals, default=unique_vals, key=f"filter_{col}")
                        filters[col] = selected_vals
                    if st.button("Reset Filters"):
                        st.rerun()

                # Weight Selection
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

            # Main Content
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

            # Tabs
            tab1, tab2, tab3 = st.tabs(["📊 Overview", "📈 Insights", "🔄 Explore"])

            # Overview Tab
            with tab1:
                st.subheader("Overview")
                col1, col2 = st.columns(2)

                with col1:
                    if numeric_cols:
                        num_col = st.selectbox("Numeric Data", numeric_cols, key="num_overview")
                        fig_num = px.histogram(filtered_df, x=num_col, title=f"{num_col} Distribution",
                                             template="plotly_white", color_discrete_sequence=["#0078d4"],
                                             nbins=min(50, filtered_df[num_col].nunique()))
                        fig_num.update_layout(xaxis_title=num_col, yaxis_title="Count", font=dict(size=12))
                        st.plotly_chart(fig_num, use_container_width=True, key="overview_num_chart")
                    else:
                        st.info("No numeric columns available.")

                with col2:
                    survey_cols = ordinal_cols + categorical_cols
                    if survey_cols:
                        survey_col = st.selectbox("Survey Responses", survey_cols, key="survey_overview")
                        if weight_col != "None" and weight_col in filtered_df.columns:
                            counts = filtered_df.groupby(survey_col, observed=True)[weight_col].sum().reset_index()
                            counts.columns = [survey_col, "Weighted Count"]
                        else:
                            counts = filtered_df[survey_col].value_counts().reset_index()
                            counts.columns = [survey_col, "Count"]
                        if survey_col in ordinal_cols and filtered_df[survey_col].dtype.name == "category":
                            counts[survey_col] = pd.Categorical(counts[survey_col], 
                                                              categories=filtered_df[survey_col].cat.categories, ordered=True)
                            counts = counts.sort_values(survey_col)
                        fig_survey = px.bar(counts, x=survey_col, y=counts.columns[1], title=f"{survey_col}",
                                          template="plotly_white", color_discrete_sequence=["#00cc96"])
                        fig_survey.update_layout(xaxis=dict(tickangle=45), font=dict(size=12))
                        st.plotly_chart(fig_survey, use_container_width=True, key="overview_survey_chart")
                    else:
                        st.info("No survey-like columns available.")

            # Insights Tab
            with tab2:
                st.subheader("Insights")
                col1, col2 = st.columns(2)

                with col1:
                    survey_cols = ordinal_cols + categorical_cols
                    if survey_cols and numeric_cols:
                        survey_x = st.selectbox("Survey Question (X)", survey_cols, key="comp_survey")
                        num_y = st.selectbox("Numeric (Y)", numeric_cols, key="comp_num")
                        fig_comp = px.box(filtered_df, x=survey_x, y=num_y, title=f"{num_y} by {survey_x}",
                                        template="plotly_white", color_discrete_sequence=["#ff5733"])
                        fig_comp.update_layout(xaxis=dict(tickangle=45), font=dict(size=12))
                        st.plotly_chart(fig_comp, use_container_width=True, key="insights_box_chart")
                    else:
                        st.info("Need survey and numeric columns for insights.")

                with col2:
                    if len(survey_cols) >= 2:
                        survey_x2 = st.selectbox("Survey Question (X-axis)", survey_cols, key="comp_survey_x")
                        survey_y2 = st.selectbox("Survey Question (Color)", 
                                               [col for col in survey_cols if col != survey_x2], key="comp_survey_y")
                        if weight_col != "None" and weight_col in filtered_df.columns:
                            cross_tab = filtered_df.groupby([survey_x2, survey_y2], observed=True)[weight_col].sum().reset_index()
                            cross_tab.columns = [survey_x2, survey_y2, "Weighted Count"]
                        else:
                            cross_tab = pd.crosstab(filtered_df[survey_x2], filtered_df[survey_y2]).reset_index()
                            cross_tab = cross_tab.melt(id_vars=[survey_x2], var_name=survey_y2, value_name="Count")
                        fig_cross = px.bar(cross_tab, x=survey_x2, y="Count", color=survey_y2, 
                                         title=f"{survey_x2} vs {survey_y2}", barmode="group",
                                         template="plotly_white", color_discrete_sequence=px.colors.qualitative.Pastel)
                        fig_cross.update_layout(xaxis=dict(tickangle=45), font=dict(size=12))
                        st.plotly_chart(fig_cross, use_container_width=True, key="insights_bar_chart")
                    else:
                        st.info("Need at least two survey columns for comparison.")

            # Explore Tab
            with tab3:
                st.subheader("Explore Relationships")
                col1, col2 = st.columns(2)

                with col1:
                    if len(numeric_cols) >= 2:
                        num_x = st.selectbox("Numeric X", numeric_cols, key="rel_num_x")
                        num_y = st.selectbox("Numeric Y", [col for col in numeric_cols if col != num_x], key="rel_num_y")
                        fig_scatter = px.scatter(filtered_df, x=num_x, y=num_y, title=f"{num_x} vs {num_y}", 
                                               trendline="ols", template="plotly_white", 
                                               color_discrete_sequence=["#636efa"])
                        fig_scatter.update_layout(font=dict(size=12))
                        st.plotly_chart(fig_scatter, use_container_width=True, key="explore_scatter_chart")
                    else:
                        st.info("Need at least two numeric columns.")

                with col2:
                    survey_cols = ordinal_cols + categorical_cols
                    if survey_cols and (numeric_cols or len(survey_cols) >= 2):
                        survey_x = st.selectbox("Survey Question", survey_cols, key="rel_survey_x")
                        y_options = numeric_cols + [col for col in survey_cols if col != survey_x]
                        y_col = st.selectbox("Y-Axis (Numeric or Survey)", y_options, key="rel_y")
                        if y_col in numeric_cols:
                            fig_rel = px.box(filtered_df, x=survey_x, y=y_col, title=f"{y_col} by {survey_x}",
                                           template="plotly_white", color_discrete_sequence=["#ab63fa"])
                        else:
                            y_data = filtered_df[y_col].cat.codes if y_col in ordinal_cols and filtered_df[y_col].dtype.name == "category" else filtered_df[y_col]
                            fig_rel = px.box(filtered_df, x=survey_x, y=y_data, 
                                           title=f"{y_col} {'(codes)' if y_col in ordinal_cols else ''} by {survey_x}",
                                           template="plotly_white", color_discrete_sequence=["#ab63fa"])
                        fig_rel.update_layout(xaxis=dict(tickangle=45), font=dict(size=12))
                        st.plotly_chart(fig_rel, use_container_width=True, key="explore_box_chart")
                    else:
                        st.info("Need survey and numeric/survey columns.")

            # Download
            st.subheader("Download Your Data")
            if st.button("Download as CSV", help="Save the filtered data as a CSV file"):
                csv = filtered_df.to_csv(index=False)
                st.download_button(label="Download CSV", data=csv, file_name="processed_data.csv", mime="text/csv")

        except Exception as e:
            st.error(f"Error processing file: {e}")
else:
    st.info("👈 Upload an Excel file to get started!")
    with st.expander("How to Use This Tool"):
        st.markdown("""
        - **Upload**: Drop your Excel file (.xlsx or .xls) to analyze.
        - **Filter**: Use the sidebar to refine your data.
        - **Explore**: Check out the tabs:
          - **Overview**: See distributions of your data.
          - **Insights**: Compare variables for deeper understanding.
          - **Explore**: Discover relationships between columns.
        - **Download**: Save your filtered data anytime!
        """)

st.markdown("---\nCreated with ❤️ for DIVE")
