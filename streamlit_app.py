import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from fpdf import FPDF
import plotly.io as pio
import os
import numpy as np

st.set_page_config(page_title="Excel Data Visualizer", layout="wide")

# Custom CSS (unchanged)
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

    # Adjust for the specific columns in this dataset
    brand_image_col = '[Brand image] Este es un anuncio de Coca Cola. ¿Qué imagen te da de Coca Cola?'
    if brand_image_col in categorical_cols:
        categorical_cols.remove(brand_image_col)
        ordinal_cols.append(brand_image_col)
        df[brand_image_col] = pd.Categorical(df[brand_image_col], 
                                            categories=["Muy negativa", "Negativa", "Neutra", "Positiva", "Muy positiva"], ordered=True)

    attribution_col = '[Attribution] Según tu opinión, este anuncio es para:'
    if attribution_col in ordinal_cols:
        ordinal_cols.remove(attribution_col)
        categorical_cols.append(attribution_col)

    return numeric_cols, categorical_cols, ordinal_cols

def calculate_impact_score(df, ad_recall_col, kpi_col):
    # Step 1: Define control and exposed groups
    df['Group'] = df[ad_recall_col].apply(lambda x: 'Exposed' if x in ['Sí, una vez', 'Sí, varias veces'] else 'Control')

    # Step 2: Calculate the average KPI for the control group
    control_avg = df[df['Group'] == 'Control'][kpi_col].mean()

    # Step 3: Calculate uplift for the exposed group (x_i = KPI value - control group average)
    exposed_df = df[df['Group'] == 'Exposed'].copy()
    exposed_df['Uplift'] = exposed_df[kpi_col] - control_avg

    # Step 4: Calculate benchmarks (top 25% and bottom 10% of uplifts)
    uplifts = exposed_df['Uplift'].dropna()
    if uplifts.empty:
        return None, "No data available for the exposed group."
    
    top_25_threshold = uplifts.quantile(0.75)
    bottom_10_threshold = uplifts.quantile(0.10)
    
    x_i_top25 = uplifts[uplifts >= top_25_threshold].mean()  # Average of top 25%
    x_i_flop10 = uplifts[uplifts <= bottom_10_threshold].mean()  # Average of bottom 10%

    # Avoid division by zero
    denominator = x_i_top25 - x_i_flop10
    if denominator == 0:
        return None, "Denominator (x_i_top25 - x_i_flop10) is zero, cannot calculate IS."

    # Step 5: Calculate the IS for each respondent in the exposed group
    exposed_df['IS_Component'] = (exposed_df['Uplift'] - x_i_flop10) / denominator

    # Step 6: Sum the IS components and divide by the number of KPIs (n=1), then multiply by 100
    n = 1  # Since we're evaluating only one KPI
    IS = (exposed_df['IS_Component'].sum() / n) * 100

    return IS, None

def create_pdf(filtered_df, numeric_cols, categorical_cols, ordinal_cols, figures, impact_score=None, impact_score_error=None):
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

    # Group Breakdown
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Group Breakdown", ln=True)
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"Control Group: {len(filtered_df[filtered_df['Group'] == 'Control'])} respondents", ln=True)
    pdf.cell(0, 6, f"Exposed Group: {len(filtered_df[filtered_df['Group'] == 'Exposed'])} respondents", ln=True)

    # Impact Score
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Impact Score", ln=True)
    pdf.set_font("Arial", "", 10)
    if impact_score is not None:
        pdf.cell(0, 6, f"Impact Score (IS): {impact_score:.2f}", ln=True)
    else:
        pdf.cell(0, 6, f"Impact Score (IS): Not calculated ({impact_score_error})", ln=True)

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

if uploaded_file:
    with st.spinner("Processing your file..."):
        try:
            df = pd.read_excel(uploaded_file)
            st.success("File uploaded successfully!")

            # Define control and exposed groups
            ad_recall_col = '[Ad recall] ¿Recuerda haber visto este anuncio en un cartel digital?'
            df['Group'] = df[ad_recall_col].apply(lambda x: 'Exposed' if x in ['Sí, una vez', 'Sí, varias veces'] else 'Control')

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

            filtered_df = df.copy()
            for col, vals in filters.items():
                filtered_df = filtered_df[filtered_df[col].isin(vals)]
            if filtered_df.empty:
                st.warning("Filters resulted in no data. Showing full dataset instead.")
                filtered_df = df.copy()

            st.subheader("Data Overview", anchor="overview")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Rows", filtered_df.shape[0])
            col2.metric("Columns", filtered_df.shape[1])
            col3.metric("Missing Values", filtered_df.isna().sum().sum())
            col4.metric("Exposed Group", len(filtered_df[filtered_df['Group'] == 'Exposed']))

            with st.expander("View Data Preview"):
                st.dataframe(filtered_df.head())
            with st.expander("Column Types"):
                st.write(f"**Numeric (Continuous):** {numeric_cols}")
                st.write(f"**Categorical (Unordered):** {categorical_cols}")
                st.write(f"**Ordinal (Survey-like):** {ordinal_cols}")

            # Calculate Impact Score for the Consideration KPI
            kpi_col = '[Consideration] ¿En el futuro considerarías comprar Coca Cola?'
            impact_score, impact_score_error = calculate_impact_score(filtered_df, ad_recall_col, kpi_col)

            st.subheader("Impact Score Analysis")
            if impact_score is not None:
                st.metric("Impact Score (IS)", f"{impact_score:.2f}")
            else:
                st.error(f"Could not calculate Impact Score: {impact_score_error}")

            # Store figures for PDF
            figures = {}

            tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Overview", "📈 Insights", "🔄 Explore", "🌟 Profiles", "📉 Impact Analysis"])

            with tab1:
                st.subheader("Overview")
                col1, col2 = st.columns(2)

                with col1:
                    if categorical_cols:
                        cat_col = st.selectbox("Categorical Data", categorical_cols, key="cat_overview")
                        counts = filtered_df[cat_col].value_counts().reset_index()
                        counts.columns = [cat_col, "Count"]
                        fig_pie = px.pie(counts, names=cat_col, values="Count", title=f"{cat_col} Breakdown",
                                       template="plotly_white", color_discrete_sequence=px.colors.qualitative.Pastel)
                        fig_pie.update_layout(font=dict(size=12))
                        st.plotly_chart(fig_pie, use_container_width=True, key="overview_pie_chart")
                        figures["Overview Pie Chart"] = fig_pie
                    else:
                        st.info("No categorical columns available.")

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
                        fig_bar = px.bar(counts, x=survey_col, y=counts.columns[1], title=f"{survey_col}",
                                       template="plotly_white", color_discrete_sequence=["#00cc96"])
                        fig_bar.update_layout(xaxis=dict(tickangle=45), font=dict(size=12))
                        st.plotly_chart(fig_bar, use_container_width=True, key="overview_bar_chart")
                        figures["Overview Bar Chart"] = fig_bar
                    else:
                        st.info("No survey-like columns available.")

            with tab2:
                st.subheader("Insights")
                col1, col2 = st.columns(2)

                with col1:
                    survey_cols = ordinal_cols + categorical_cols
                    if survey_cols and numeric_cols:
                        survey_x = st.selectbox("Survey Question (X)", survey_cols, key="comp_survey")
                        num_y = st.selectbox("Numeric (Y)", numeric_cols, key="comp_num")
                        fig_box = px.box(filtered_df, x=survey_x, y=num_y, title=f"{num_y} by {survey_x}",
                                       template="plotly_white", color_discrete_sequence=["#ff5733"])
                        fig_box.update_layout(xaxis=dict(tickangle=45), font=dict(size=12))
                        st.plotly_chart(fig_box, use_container_width=True, key="insights_box_chart")
                        figures["Insights Box Plot"] = fig_box
                    else:
                        st.info("Need survey and numeric columns for insights.")

                with col2:
                    if len(survey_cols) >= 2:
                        survey_x2 = st.selectbox("Survey Question (X-axis)", survey_cols, key="comp_survey_x")
                        survey_y2 = st.selectbox("Survey Question (Y-axis)", 
                                               [col for col in survey_cols if col != survey_x2], key="comp_survey_y")
                        cross_tab = pd.crosstab(filtered_df[survey_x2], filtered_df[survey_y2])
                        fig_heatmap = px.imshow(cross_tab, text_auto=True, aspect="auto",
                                              title=f"{survey_x2} vs {survey_y2}", 
                                              color_continuous_scale="Blues", template="plotly_white")
                        fig_heatmap.update_layout(xaxis=dict(tickangle=45), font=dict(size=12))
                        st.plotly_chart(fig_heatmap, use_container_width=True, key="insights_heatmap_chart")
                        figures["Insights Heatmap"] = fig_heatmap
                    else:
                        st.info("Need at least two survey columns for heatmap.")

            with tab3:
                st.subheader("Explore Relationships")
                col1, col2 = st.columns(2)

                with col1:
                    if numeric_cols:
                        num_col = st.selectbox("Numeric Data", numeric_cols, key="num_explore")
                        fig_hist = px.histogram(filtered_df, x=num_col, title=f"{num_col} Distribution",
                                              template="plotly_white", color_discrete_sequence=["#0078d4"],
                                              nbins=min(50, filtered_df[num_col].nunique()))
                        fig_hist.update_layout(font=dict(size=12))
                        st.plotly_chart(fig_hist, use_container_width=True, key="explore_hist_chart")
                        figures["Explore Histogram"] = fig_hist
                    else:
                        st.info("No numeric columns available.")

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
                        figures["Explore Box Plot"] = fig_rel
                    else:
                        st.info("Need survey and numeric/survey columns.")

            with tab4:
                st.subheader("Profiles")
                survey_cols = numeric_cols + ordinal_cols
                if len(survey_cols) >= 2 and categorical_cols:
                    group_col = st.selectbox("Group By", categorical_cols, key="radar_group")
                    radar_cols = st.multiselect("Select Variables (2+)", survey_cols, 
                                              default=survey_cols[:min(3, len(survey_cols))], 
                                              key="radar_vars")
                    if len(radar_cols) >= 2:
                        radar_df = filtered_df.copy()
                        for col in radar_cols:
                            if col in ordinal_cols and radar_df[col].dtype.name == "category":
                                radar_df[col] = radar_df[col].cat.codes
                        agg_data = radar_df.groupby(group_col)[radar_cols].mean().reset_index()
                        fig_radar = go.Figure()
                        for i, row in agg_data.iterrows():
                            values = [row[col] for col in radar_cols]
                            fig_radar.add_trace(go.Scatterpolar(
                                r=values + [values[0]],
                                theta=radar_cols + [radar_cols[0]],
                                fill='toself',
                                name=row[group_col],
                                line=dict(color=px.colors.qualitative.Pastel[i % len(px.colors.qualitative.Pastel)])
                            ))
                        max_val = agg_data[radar_cols].max().max()
                        fig_radar.update_layout(
                            polar=dict(radialaxis=dict(visible=True, range=[0, max(10, max_val)])),
                            showlegend=True, template="plotly_white", font=dict(size=12)
                        )
                        st.plotly_chart(fig_radar, use_container_width=True, key="profiles_radar_chart")
                        figures["Profiles Radar Chart"] = fig_radar
                    else:
                        st.info("Select at least 2 numeric or ordinal variables for radar chart.")
                else:
                    st.info("Need at least 2 numeric/ordinal columns and 1 categorical column for radar charts.")

            with tab5:
                st.subheader("Impact Analysis: Control vs Exposed Groups")
                col1, col2 = st.columns(2)

                with col1:
                    # Box plot comparing Consideration between Control and Exposed groups
                    fig_box_impact = px.box(filtered_df, x='Group', y=kpi_col,
                                            title=f"{kpi_col} by Group",
                                            template="plotly_white", color='Group',
                                            color_discrete_sequence=["#ff5733", "#00cc96"])
                    fig_box_impact.update_layout(xaxis=dict(tickangle=45), font=dict(size=12))
                    st.plotly_chart(fig_box_impact, use_container_width=True, key="impact_box_chart")
                    figures["Impact Box Plot"] = fig_box_impact

                with col2:
                    # Histogram of Consideration scores for both groups
                    fig_hist_impact = px.histogram(filtered_df, x=kpi_col, color='Group',
                                                  title=f"{kpi_col} Distribution by Group",
                                                  template="plotly_white",
                                                  color_discrete_sequence=["#ff5733", "#00cc96"],
                                                  nbins=min(50, filtered_df[kpi_col].nunique()))
                    fig_hist_impact.update_layout(font=dict(size=12), barmode='overlay')
                    fig_hist_impact.update_traces(opacity=0.75)
                    st.plotly_chart(fig_hist_impact, use_container_width=True, key="impact_hist_chart")
                    figures["Impact Histogram"] = fig_hist_impact

            st.subheader("Download Your Data")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Download as CSV", help="Save the filtered data as a CSV file"):
                    csv = filtered_df.to_csv(index=False)
                    st.download_button(label="Download CSV", data=csv, file_name="processed_data.csv", mime="text/csv")
            with col2:
                if st.button("Download as PDF", help="Save data and graphs as a PDF"):
                    with st.spinner("Generating PDF with graphs..."):
                        pdf_output = create_pdf(filtered_df, numeric_cols, categorical_cols, ordinal_cols, figures, impact_score, impact_score_error)
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
          - **Impact Analysis**: Compare control vs exposed groups and view Impact Score.
        - **Download**: Save as CSV or PDF with graphs!
        """)

st.markdown("---\nCreated with ❤️ for DIVE")
