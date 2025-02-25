import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Survey Visualizer", layout="wide")

# Custom CSS for better layout
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

st.title("📊 Coca-Cola Survey Visualizer")
st.write("Upload your survey Excel file to explore visualizations of the data.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

# Rename columns for simplicity and match your data
column_mapping = {
    "[Profiling] ¿Qué edad tienes?": "Age",
    "[Profiling] Eres...": "Gender",
    "[Ad recall] ¿Recuerda haber visto este anuncio en un cartel digital?": "Ad Recall",
    "[Interest] ¿Te interesa este anuncio?": "Interest",
    "[Attribution] Según tu opinión, este anuncio es para": "Attribution",
    "[Brand image] Este es un anuncio de Coca Cola. ¿Qué imagen te da de Coca Cola?": "Brand Image",
    "[Consideration] ¿En el futuro considerarías comprar Coca Cola?": "Consideration",
    "¿Que tan seguido tomas Coca Cola?": "Frequency",
    "Adjustment Weight": "Weight",
    "User ID": "User ID",
    "Area": "Area"
}

# Define ordinal orders
ordinal_orders = {
    "Brand Image": ["Muy negativa", "Negativa", "Neutra", "Positiva", "Muy positiva"],
    "Frequency": ["Nunca", "Una vez al mes", "Pocas veces al mes", "Una vez a la semana", "Varias veces a la semana"],
    "Ad Recall": ["No", "Sí, una vez", "Sí, varias veces"],
    "Interest": list(range(0, 11)),
    "Consideration": list(range(0, 11))
}

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Rename columns for easier handling
        df = df.rename(columns=column_mapping)
        st.write("### Data Preview")
        st.dataframe(df.head())

        # Convert ordinal columns to categorical with order
        for col, order in ordinal_orders.items():
            if col in df.columns:
                df[col] = pd.Categorical(df[col], categories=order, ordered=True)

        # Sidebar Filters
        st.sidebar.header("Filters")
        filters = {}
        for col in ["Age", "Gender", "Ad Recall", "Attribution", "Area"]:
            if col in df.columns:
                unique_vals = df[col].dropna().unique()
                selected_vals = st.sidebar.multiselect(f"Filter {col}", unique_vals, default=unique_vals)
                filters[col] = selected_vals

        # Apply filters
        filtered_df = df.copy()
        for col, vals in filters.items():
            filtered_df = filtered_df[filtered_df[col].isin(vals)]

        # Weight option
        weight_col = st.sidebar.selectbox("Apply weights (optional):", ["None", "Weight"], index=0)

        # Data Overview
        st.subheader("Data Overview")
        col1, col2, col3 = st.columns(3)
        col1.metric("Rows", filtered_df.shape[0])
        col2.metric("Columns", filtered_df.shape[1])
        col3.metric("Missing Values", filtered_df.isna().sum().sum())

        # Visualization Tabs
        tab1, tab2, tab3 = st.tabs(["📊 Demographics", "📈 Perceptions", "🔄 Relationships"])

        # Demographics Tab
        with tab1:
            st.subheader("Demographics and Ad Recall")
            col1, col2 = st.columns(2)
            
            with col1:
                # Age Distribution
                if "Age" in filtered_df.columns:
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        age_counts = filtered_df.groupby("Age", observed=True)[weight_col].sum().reset_index()
                        age_counts.columns = ["Age", "Weighted Count"]
                    else:
                        age_counts = filtered_df["Age"].value_counts().reset_index()
                        age_counts.columns = ["Age", "Count"]
                    fig_age = px.bar(age_counts, x="Age", y=age_counts.columns[1], title="Age Distribution")
                    st.plotly_chart(fig_age, use_container_width=True)
                
            with col2:
                # Gender Distribution
                if "Gender" in filtered_df.columns:
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        gender_counts = filtered_df.groupby("Gender", observed=True)[weight_col].sum().reset_index()
                        gender_counts.columns = ["Gender", "Weighted Count"]
                    else:
                        gender_counts = filtered_df["Gender"].value_counts().reset_index()
                        gender_counts.columns = ["Gender", "Count"]
                    fig_gender = px.pie(gender_counts, names="Gender", values=gender_counts.columns[1], title="Gender Distribution")
                    st.plotly_chart(fig_gender, use_container_width=True)

            # Ad Recall
            if "Ad Recall" in filtered_df.columns:
                if weight_col != "None" and weight_col in filtered_df.columns:
                    recall_counts = filtered_df.groupby("Ad Recall", observed=True)[weight_col].sum().reset_index()
                    recall_counts.columns = ["Ad Recall", "Weighted Count"]
                else:
                    recall_counts = filtered_df["Ad Recall"].value_counts().reset_index()
                    recall_counts.columns = ["Ad Recall", "Count"]
                fig_recall = px.bar(recall_counts, x="Ad Recall", y=recall_counts.columns[1], title="Ad Recall Distribution")
                st.plotly_chart(fig_recall, use_container_width=True)

        # Perceptions Tab
        with tab2:
            st.subheader("Brand Perceptions and Behavior")
            col1, col2 = st.columns(2)
            
            with col1:
                # Interest Distribution
                if "Interest" in filtered_df.columns:
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        interest_counts = filtered_df.groupby("Interest", observed=True)[weight_col].sum().reset_index()
                        interest_counts.columns = ["Interest", "Weighted Count"]
                    else:
                        interest_counts = filtered_df["Interest"].value_counts().reset_index()
                        interest_counts.columns = ["Interest", "Count"]
                    fig_interest = px.bar(interest_counts, x="Interest", y=interest_counts.columns[1], title="Interest in Ad (0-10)")
                    st.plotly_chart(fig_interest, use_container_width=True)
                
                # Brand Image
                if "Brand Image" in filtered_df.columns:
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        image_counts = filtered_df.groupby("Brand Image", observed=True)[weight_col].sum().reset_index()
                        image_counts.columns = ["Brand Image", "Weighted Count"]
                    else:
                        image_counts = filtered_df["Brand Image"].value_counts().reset_index()
                        image_counts.columns = ["Brand Image", "Count"]
                    fig_image = px.bar(image_counts, x="Brand Image", y=image_counts.columns[1], title="Brand Image Perception")
                    st.plotly_chart(fig_image, use_container_width=True)

            with col2:
                # Consideration
                if "Consideration" in filtered_df.columns:
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        consid_counts = filtered_df.groupby("Consideration", observed=True)[weight_col].sum().reset_index()
                        consid_counts.columns = ["Consideration", "Weighted Count"]
                    else:
                        consid_counts = filtered_df["Consideration"].value_counts().reset_index()
                        consid_counts.columns = ["Consideration", "Count"]
                    fig_consid = px.bar(consid_counts, x="Consideration", y=consid_counts.columns[1], title="Consideration to Buy (0-10)")
                    st.plotly_chart(fig_consid, use_container_width=True)
                
                # Frequency
                if "Frequency" in filtered_df.columns:
                    if weight_col != "None" and weight_col in filtered_df.columns:
                        freq_counts = filtered_df.groupby("Frequency", observed=True)[weight_col].sum().reset_index()
                        freq_counts.columns = ["Frequency", "Weighted Count"]
                    else:
                        freq_counts = filtered_df["Frequency"].value_counts().reset_index()
                        freq_counts.columns = ["Frequency", "Count"]
                    fig_freq = px.bar(freq_counts, x="Frequency", y=freq_counts.columns[1], title="Consumption Frequency")
                    st.plotly_chart(fig_freq, use_container_width=True)

        # Relationships Tab
        with tab3:
            st.subheader("Relationships Between Variables")
            col1, col2 = st.columns(2)
            
            with col1:
                # Interest vs Consideration
                if "Interest" in filtered_df.columns and "Consideration" in filtered_df.columns:
                    fig_ic = px.scatter(filtered_df, x="Interest", y="Consideration", 
                                      title="Interest vs Consideration", trendline="ols")
                    st.plotly_chart(fig_ic, use_container_width=True)

            with col2:
                # Brand Image vs Frequency
                if "Brand Image" in filtered_df.columns and "Frequency" in filtered_df.columns:
                    fig_bf = px.box(filtered_df, x="Brand Image", y="Frequency", 
                                   title="Brand Image vs Consumption Frequency")
                    st.plotly_chart(fig_bf, use_container_width=True)

        # Download Option
        st.subheader("Download Data")
        if st.button("Download as CSV"):
            csv = filtered_df.to_csv(index=False)
            st.download_button(label="Download CSV", data=csv, file_name="survey_data.csv", mime="text/csv")

    except Exception as e:
        st.error(f"Error processing file: {e}")
else:
    st.info("👈 Please upload your Excel file to see visualizations.")
    st.markdown("""
    ### Expected Visualizations:
    - **Demographics**: Age, Gender, Ad Recall distributions
    - **Perceptions**: Interest, Brand Image, Consideration, and Frequency
    - **Relationships**: Interest vs Consideration, Brand Image vs Frequency
    """)

st.markdown("---\nCreated with Streamlit • Coca-Cola Survey Visualizer")
