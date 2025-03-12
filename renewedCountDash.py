import streamlit as st
from simple_salesforce import Salesforce
import os
from dotenv import load_dotenv
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd

# Load environment variables from .env file
load_dotenv()

# Function to connect to Salesforce and run SOQL queries
def connect_to_salesforce():
    """Connect to Salesforce and execute SOQL queries."""
    try:
        # Salesforce connection using environment variables
        sf = Salesforce(
            username=os.getenv("SF_USERNAME_PRO"),
            password=os.getenv("SF_PASSWORD_PRO"),
            security_token=os.getenv("SF_SECURITY_TOKEN_PRO"),
        )

        # Query to get the count of Renewals
        renewal_query = """
            SELECT COUNT(Id)
            FROM Opportunity
            WHERE New_Business_or_Renewal__c IN ('Personal Lines - Renewal', 'Commercial Lines - Renewal')
        """
        renewal_results = sf.query(renewal_query)
        renewal_count = renewal_results['records'][0]['expr0']

        # Query to get the total count of Opportunities
        total_query = """
            SELECT COUNT(Id)
            FROM Opportunity
            WHERE New_Business_or_Renewal__c IN ('Personal Lines - New Business', 'Personal Lines - Renewal', 
                                                 'Commercial Lines - New Business', 'Commercial Lines - Renewal')
        """
        total_results = sf.query(total_query)
        total_count = total_results['records'][0]['expr0']

        return int(renewal_count), int(total_count)
    except Exception as e:
        st.error(f"Error connecting to Salesforce: {str(e)}")
        return 0, 0

# Streamlit UI
st.title("Renewal Opportunity Dashboard")

# Sidebar for user interaction
st.sidebar.header("Dashboard Options")
show_data = st.sidebar.checkbox("Show Raw Data", value=False)

# Main content
renewal_count, total_count = connect_to_salesforce()
if total_count > 0:
    renewal_percentage = (renewal_count / total_count) * 100
    st.subheader("Opportunity Summary")
    st.write(f"**Renewal Opportunities:** {renewal_count}")
    st.write(f"**Total Opportunities:** {total_count}")
    st.write(f"**Renewal Percentage:** {renewal_percentage:.2f}%")
else:
    st.warning("No opportunities found.")

# Visualization
if total_count > 0:
    # Prepare data for charts
    data = {
        "Category": ["Renewals", "Other Opportunities"],
        "Count": [renewal_count, total_count - renewal_count],
    }
    df = pd.DataFrame(data)

    # Sidebar chart selection
    chart_type = st.sidebar.selectbox(
        "Select Chart Type",
        options=[
            "Pie Chart",
            "Bar Chart",
            "Scatter Plot",
            "Line Chart",
            "Area Chart",
            "Histogram",
            "Box Plot",
            "Violin Plot",
            "Funnel Chart",
            "Density Heatmap",
            "Radar Chart",
            "Sunburst Chart",
            "Treemap",
            "Parallel Coordinates",
            "Bubble Chart",
        ],
    )

    # Generate selected chart
    st.subheader("Visualization")
    if chart_type == "Pie Chart":
        fig = px.pie(
            df,
            names="Category",
            values="Count",
            title="Renewals vs Other Opportunities",
        )
    elif chart_type == "Bar Chart":
        fig = px.bar(
            df,
            x="Category",
            y="Count",
            title="Renewals vs Other Opportunities",
            text="Count",
        )
    elif chart_type == "Scatter Plot":
        fig = px.scatter(
            df,
            x="Category",
            y="Count",
            title="Scatter Plot of Opportunities",
            size="Count",
            color="Category",
        )
    elif chart_type == "Line Chart":
        fig = px.line(
            df,
            x="Category",
            y="Count",
            title="Line Chart of Opportunities",
        )
    elif chart_type == "Area Chart":
        fig = px.area(
            df,
            x="Category",
            y="Count",
            title="Area Chart of Opportunities",
        )
    elif chart_type == "Histogram":
        fig = px.histogram(
            df,
            x="Count",
            title="Histogram of Opportunity Counts",
        )
    elif chart_type == "Box Plot":
        fig = px.box(
            df,
            x="Category",
            y="Count",
            title="Box Plot of Opportunities",
        )
    elif chart_type == "Violin Plot":
        fig = px.violin(
            df,
            x="Category",
            y="Count",
            title="Violin Plot of Opportunities",
        )
    elif chart_type == "Funnel Chart":
        fig = px.funnel(
            df,
            x="Count",
            y="Category",
            title="Funnel Chart of Opportunities",
        )
    elif chart_type == "Density Heatmap":
        fig = px.density_heatmap(
            df,
            x="Category",
            y="Count",
            title="Density Heatmap of Opportunities",
        )
    elif chart_type == "Radar Chart":
        fig = go.Figure(
            data=go.Scatterpolar(
                r=df["Count"],
                theta=df["Category"],
                fill="toself",
            )
        )
        fig.update_layout(title="Radar Chart of Opportunities")
    elif chart_type == "Sunburst Chart":
        fig = px.sunburst(
            df,
            path=["Category"],
            values="Count",
            title="Sunburst Chart of Opportunities",
        )
    elif chart_type == "Treemap":
        fig = px.treemap(
            df,
            path=["Category"],
            values="Count",
            title="Treemap of Opportunities",
        )
    elif chart_type == "Parallel Coordinates":
        df["Category_Encoded"] = df["Category"].factorize()[0]
        fig = px.parallel_coordinates(
            df,
            dimensions=["Category_Encoded", "Count"],
            title="Parallel Coordinates of Opportunities",
        )
    elif chart_type == "Bubble Chart":
        fig = px.scatter(
            df,
            x="Category",
            y="Count",
            size="Count",
            color="Category",
            title="Bubble Chart of Opportunities",
        )

    # Display the selected chart
    st.plotly_chart(fig)

# Optionally show raw data
if show_data and total_count > 0:
    st.subheader("Raw Data")
    st.dataframe(df)
