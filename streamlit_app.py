import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Excel Visualizer", layout="wide")

# Custom CSS to improve the appearance
st.markdown("""
<style>
    .main {
        padding: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 2rem;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 5px 5px 0px 0px;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff;
        border-bottom: 2px solid #4e89ae;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Excel File Visualizer")
st.write("Upload your Excel file to automatically generate visualizations from your data.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

def get_numeric_columns(df):
    return df.select_dtypes(include=['number']).columns.tolist()

def get_categorical_columns(df):
    return df.select_dtypes(include=['object', 'category']).columns.tolist()

def get_datetime_columns(df):
    return df.select_dtypes(include=['datetime']).columns.tolist()

if uploaded_file:
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        if not sheet_names:
            st.error("No sheets found in the Excel file.")
        else:
            # Select sheet
            selected_sheet = st.selectbox("Select a sheet:", sheet_names)
            
            # Read the selected sheet
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
            
            # Show basic info about the data
            st.subheader("Data Overview")
            col1, col2, col3 = st.columns(3)
            col1.metric("Rows", df.shape[0])
            col2.metric("Columns", df.shape[1])
            col3.metric("Missing Values", df.isna().sum().sum())
            
            # Display the dataframe
            with st.expander("View Data", expanded=False):
                st.dataframe(df)
            
            # Identify column types
            numeric_cols = get_numeric_columns(df)
            categorical_cols = get_categorical_columns(df)
            datetime_cols = get_datetime_columns(df)
            
            # Create tabs for different visualization types
            tab1, tab2, tab3, tab4 = st.tabs(["📈 Distribution", "🔄 Correlation", "📊 Categorical", "📅 Time Series"])
            
            with tab1:
                st.subheader("Distribution Analysis")
                
                if numeric_cols:
                    selected_num_col = st.selectbox("Select a numeric column for distribution analysis:", numeric_cols)
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader(f"Histogram of {selected_num_col}")
                        fig = px.histogram(df, x=selected_num_col, marginal="box", 
                                        title=f"Distribution of {selected_num_col}")
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
                        st.subheader(f"Box Plot of {selected_num_col}")
                        fig = px.box(df, y=selected_num_col, 
                                    title=f"Box Plot of {selected_num_col}")
                        st.plotly_chart(fig, use_container_width=True)
                    
                    # Basic statistics
                    st.subheader(f"Statistics for {selected_num_col}")
                    stats = df[selected_num_col].describe()
                    st.dataframe(stats)
                else:
                    st.info("No numeric columns found for distribution analysis.")
            
            with tab2:
                st.subheader("Correlation Analysis")
                
                if len(numeric_cols) >= 2:
                    st.write("Select columns for correlation analysis:")
                    
                    # Allow user to select columns for correlation analysis
                    selected_corr_cols = st.multiselect(
                        "Select numeric columns (at least 2):",
                        numeric_cols,
                        default=numeric_cols[:min(5, len(numeric_cols))]
                    )
                    
                    if len(selected_corr_cols) >= 2:
                        # Calculate correlation matrix
                        corr_matrix = df[selected_corr_cols].corr()
                        
                        # Plot correlation heatmap
                        fig, ax = plt.subplots(figsize=(10, 8))
                        sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0, ax=ax)
                        plt.title('Correlation Matrix')
                        st.pyplot(fig)
                        
                        # Scatter plot for selected pairs
                        st.subheader("Scatter Plot Analysis")
                        x_col = st.selectbox("Select X-axis column:", selected_corr_cols)
                        y_col = st.selectbox("Select Y-axis column:", 
                                            [col for col in selected_corr_cols if col != x_col],
                                            index=min(1, len(selected_corr_cols) - 1))
                        
                        color_col = None
                        if categorical_cols:
                            color_col = st.selectbox("Color by (optional):", 
                                                    ["None"] + categorical_cols)
                        
                        if color_col == "None":
                            fig = px.scatter(df, x=x_col, y=y_col, 
                                            title=f"{x_col} vs {y_col}",
                                            trendline="ols")
                        else:
                            fig = px.scatter(df, x=x_col, y=y_col, color=color_col,
                                            title=f"{x_col} vs {y_col} by {color_col}")
                        
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.info("Please select at least 2 columns for correlation analysis.")
                else:
                    st.info("Need at least 2 numeric columns for correlation analysis.")
            
            with tab3:
                st.subheader("Categorical Analysis")
                
                if categorical_cols:
                    selected_cat_col = st.selectbox("Select a categorical column:", categorical_cols)
                    
                    # Count plot
                    value_counts = df[selected_cat_col].value_counts().reset_index()
                    value_counts.columns = [selected_cat_col, 'Count']
                    
                    # Sort by count
                    value_counts = value_counts.sort_values('Count', ascending=False)
                    
                    # Limit to top 20 categories if there are many
                    if len(value_counts) > 20:
                        st.info(f"Showing top 20 categories out of {len(value_counts)}")
                        value_counts = value_counts.head(20)
                    
                    fig = px.bar(value_counts, x=selected_cat_col, y='Count', 
                                title=f"Count of {selected_cat_col}",
                                text='Count')
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # If there are numeric columns, allow relationship visualization
                    if numeric_cols:
                        st.subheader(f"{selected_cat_col} vs Numeric Variables")
                        selected_num_col = st.selectbox("Select a numeric column:", numeric_cols)
                        
                        plot_type = st.radio("Select plot type:", ["Box Plot", "Bar Chart (Mean)"])
                        
                        if plot_type == "Box Plot":
                            fig = px.box(df, x=selected_cat_col, y=selected_num_col,
                                        title=f"{selected_num_col} by {selected_cat_col}")
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            # Calculate mean for each category
                            agg_df = df.groupby(selected_cat_col)[selected_num_col].mean().reset_index()
                            agg_df = agg_df.sort_values(selected_num_col, ascending=False)
                            
                            # Limit to top 20 categories if there are many
                            if len(agg_df) > 20:
                                st.info(f"Showing top 20 categories out of {len(agg_df)}")
                                agg_df = agg_df.head(20)
                            
                            fig = px.bar(agg_df, x=selected_cat_col, y=selected_num_col,
                                        title=f"Mean {selected_num_col} by {selected_cat_col}")
                            st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No categorical columns found for analysis.")
            
            with tab4:
                st.subheader("Time Series Analysis")
                
                # Check for datetime columns
                if datetime_cols:
                    date_col = st.selectbox("Select date column:", datetime_cols)
                    
                    if numeric_cols:
                        selected_ts_col = st.selectbox("Select value to plot over time:", numeric_cols)
                        
                        # Resample options
                        resample_options = {
                            "Original data": None,
                            "Daily": 'D', 
                            "Weekly": 'W', 
                            "Monthly": 'M',
                            "Quarterly": 'Q',
                            "Yearly": 'Y'
                        }
                        
                        resample_period = st.selectbox("Resample period:", list(resample_options.keys()))
                        
                        # Create time series plot
                        df_copy = df.copy()
                        df_copy[date_col] = pd.to_datetime(df_copy[date_col], errors='coerce')
                        df_copy = df_copy.dropna(subset=[date_col])
                        
                        # Sort by date
                        df_copy = df_copy.sort_values(by=date_col)
                        
                        # Resample if needed
                        if resample_options[resample_period]:
                            # Set date as index
                            df_copy = df_copy.set_index(date_col)
                            # Resample and take mean
                            df_resampled = df_copy[selected_ts_col].resample(resample_options[resample_period]).mean().reset_index()
                            fig = px.line(df_resampled, x=date_col, y=selected_ts_col,
                                        title=f"{selected_ts_col} over Time ({resample_period})")
                        else:
                            fig = px.line(df_copy, x=date_col, y=selected_ts_col,
                                        title=f"{selected_ts_col} over Time")
                        
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # Moving average
                        st.subheader("Moving Average")
                        window_size = st.slider("Select window size for moving average:", 
                                                min_value=2, 
                                                max_value=min(30, len(df_copy) // 2),
                                                value=5)
                        
                        if len(df_copy) > window_size:
                            if resample_options[resample_period]:
                                # Use resampled data
                                df_ma = df_resampled.copy()
                                df_ma['Moving Average'] = df_ma[selected_ts_col].rolling(window=window_size).mean()
                            else:
                                # Use original data
                                df_ma = df_copy.copy()
                                df_ma['Moving Average'] = df_ma[selected_ts_col].rolling(window=window_size).mean()
                            
                            fig = px.line(df_ma, x=date_col, y=[selected_ts_col, 'Moving Average'],
                                        title=f"{selected_ts_col} with {window_size}-period Moving Average")
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            st.info("Not enough data points for moving average calculation.")
                    else:
                        st.info("No numeric columns found for time series analysis.")
                
                elif numeric_cols and categorical_cols:
                    # If no datetime columns, suggest creating a time series from categorical + numeric
                    st.info("No datetime columns found. You can try creating a time series from categorical data.")
                    
                    potential_time_col = st.selectbox("Select a categorical column that might represent time periods:",
                                                    categorical_cols)
                    selected_ts_col = st.selectbox("Select value to plot:", numeric_cols)
                    
                    # Group by the categorical column
                    df_grouped = df.groupby(potential_time_col)[selected_ts_col].mean().reset_index()
                    
                    # Try to convert categorical to orderable format
                    try:
                        # First, check if it can be converted to datetime
                        df_grouped['temp_date'] = pd.to_datetime(df_grouped[potential_time_col], errors='coerce')
                        if not df_grouped['temp_date'].isna().all():
                            # Can be converted to date
                            df_grouped = df_grouped.sort_values('temp_date')
                            fig = px.line(df_grouped, x=potential_time_col, y=selected_ts_col,
                                        title=f"{selected_ts_col} by {potential_time_col}")
                            st.plotly_chart(fig, use_container_width=True)
                        else:
                            # Try numeric conversion
                            df_grouped['temp_num'] = pd.to_numeric(df_grouped[potential_time_col], errors='coerce')
                            if not df_grouped['temp_num'].isna().all():
                                # Can be converted to number
                                df_grouped = df_grouped.sort_values('temp_num')
                                fig = px.line(df_grouped, x=potential_time_col, y=selected_ts_col,
                                            title=f"{selected_ts_col} by {potential_time_col}")
                                st.plotly_chart(fig, use_container_width=True)
                            else:
                                # Use as is, no particular order
                                fig = px.line(df_grouped, x=potential_time_col, y=selected_ts_col,
                                            title=f"{selected_ts_col} by {potential_time_col}")
                                st.plotly_chart(fig, use_container_width=True)
                                st.info("Note: The x-axis order may not represent chronological time.")
                    except:
                        fig = px.line(df_grouped, x=potential_time_col, y=selected_ts_col,
                                    title=f"{selected_ts_col} by {potential_time_col}")
                        st.plotly_chart(fig, use_container_width=True)
                        st.info("Note: The x-axis order may not represent chronological time.")
                else:
                    st.info("No suitable columns found for time series analysis.")
            
            # Add download functionality
            st.subheader("Download Options")
            
            download_format = st.radio("Select download format:", ["CSV", "Excel", "JSON"])
            
            if st.button("Download Processed Data"):
                if download_format == "CSV":
                    csv = df.to_csv(index=False)
                    b64 = BytesIO()
                    b64.write(csv.encode())
                    b64.seek(0)
                    st.download_button(
                        label="Download CSV",
                        data=b64,
                        file_name=f"{selected_sheet}.csv",
                        mime="text/csv"
                    )
                elif download_format == "Excel":
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, sheet_name=selected_sheet, index=False)
                    output.seek(0)
                    st.download_button(
                        label="Download Excel",
                        data=output,
                        file_name=f"{selected_sheet}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:  # JSON
                    json_str = df.to_json(orient='records')
                    b64 = BytesIO()
                    b64.write(json_str.encode())
                    b64.seek(0)
                    st.download_button(
                        label="Download JSON",
                        data=b64,
                        file_name=f"{selected_sheet}.json",
                        mime="application/json"
                    )
                    
    except Exception as e:
        st.error(f"Error processing the file: {e}")
else:
    # Show sample images when no file is uploaded
    st.info("👈 Please upload an Excel file to begin visualization.")
    
    # Display sample visualizations
    st.subheader("Sample Visualizations You'll Get:")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Distribution Analysis**")
        st.image("https://via.placeholder.com/400x300?text=Histogram+and+Box+Plots", use_column_width=True)
        
        st.markdown("**Correlation Analysis**")
        st.image("https://via.placeholder.com/400x300?text=Correlation+Heatmap", use_column_width=True)
    
    with col2:
        st.markdown("**Categorical Analysis**")
        st.image("https://via.placeholder.com/400x300?text=Bar+Charts", use_column_width=True)
        
        st.markdown("**Time Series Analysis**")
        st.image("https://via.placeholder.com/400x300?text=Time+Series+Charts", use_column_width=True)
    
    st.markdown("""
    ### How to use:
    1. Upload any Excel file using the file uploader
    2. Select a sheet from your Excel file
    3. Explore different visualizations through the tabs
    4. Download your processed data in various formats
    
    This app automatically detects numeric, categorical, and date columns to provide relevant visualizations.
    """)

# Add footer
st.markdown("""
---
Created with Streamlit • Excel Visualizer App
""")
