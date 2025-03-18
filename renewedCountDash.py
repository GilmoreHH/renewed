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

        # Query to get the count of renewal opportunities that were eligible and renewed
        # Assuming there's a field that tracks if an opportunity was renewed (StageName = 'Closed Won' or similar)
        eligible_renewed_query = """
            SELECT COUNT(Id)
            FROM Opportunity
            WHERE New_Business_or_Renewal__c IN ('Personal Lines - Renewal', 'Commercial Lines - Renewal')
            AND StageName = 'Closed Won'
            AND CloseDate = LAST_N_DAYS:7 
        """
        eligible_renewed_results = sf.query(eligible_renewed_query)
        eligible_renewed_count = eligible_renewed_results['records'][0]['expr0']

        # Query to get the total count of eligible renewal opportunities
        eligible_renewal_query = """
            SELECT COUNT(Id)
            FROM Opportunity
            WHERE New_Business_or_Renewal__c IN ('Personal Lines - Renewal', 'Commercial Lines - Renewal')
            AND CloseDate = LAST_N_DAYS:7 
        """
        eligible_renewal_results = sf.query(eligible_renewal_query)
        eligible_renewal_count = eligible_renewal_results['records'][0]['expr0']

        # Get data for weekly trend analysis (last 8 weeks)
        # Get all opportunities first, then process them on the client side
        all_opps_query = """
            SELECT 
                Id, 
                CloseDate, 
                StageName
            FROM Opportunity
            WHERE New_Business_or_Renewal__c IN ('Personal Lines - Renewal', 'Commercial Lines - Renewal')
            AND CloseDate = LAST_N_DAYS:56
        """
        all_opps_results = sf.query_all(all_opps_query)
        # We don't need this line anymore as we're using a different approach
        # Process results into a weekly data structure
        import datetime
        
        # Create dictionaries to track counts by week
        eligible_by_week = {}
        renewed_by_week = {}
        
        # Process all opportunities
        for record in all_opps_results['records']:
            # Convert CloseDate string to date object
            close_date = datetime.datetime.strptime(record['CloseDate'], '%Y-%m-%d').date()
            # Get ISO week number (1-53)
            year, week_num, _ = close_date.isocalendar()
            week_key = f"{year}-W{week_num:02d}"
            
            # Count eligible opportunities
            eligible_by_week[week_key] = eligible_by_week.get(week_key, 0) + 1
            
            # Count renewed opportunities (closed won)
            if record['StageName'] == 'Closed Won':
                renewed_by_week[week_key] = renewed_by_week.get(week_key, 0) + 1
        
        # Combine the data
        weekly_data = []
        for week in sorted(eligible_by_week.keys()):
            eligible_count = eligible_by_week[week]
            renewed_count = renewed_by_week.get(week, 0)  # Default to 0 if no renewals that week
            weekly_data.append({
                'Week': week,
                'EligibleCount': eligible_count,
                'RenewedCount': renewed_count,
                'RenewalRate': (renewed_count / eligible_count * 100) if eligible_count > 0 else 0
            })
        
        weekly_df = pd.DataFrame(weekly_data)

        return int(eligible_renewed_count), int(eligible_renewal_count), weekly_df
    except Exception as e:
        st.error(f"Error connecting to Salesforce: {str(e)}")
        return 0, 0, pd.DataFrame()

# Streamlit UI
st.title("Renewal Opportunity Dashboard")

# Sidebar for user interaction
st.sidebar.header("Dashboard Options")
show_data = st.sidebar.checkbox("Show Raw Data", value=False)
time_period = st.sidebar.selectbox(
    "Select Time Period",
    options=["Last 7 Days", "Last 30 Days", "Last Quarter"],
    index=0
)

# Main content
eligible_renewed_count, eligible_renewal_count, weekly_df = connect_to_salesforce()
if eligible_renewal_count > 0:
    renewal_percentage = (eligible_renewed_count / eligible_renewal_count) * 100
    st.subheader("Renewal Opportunity Summary")
    st.write(f"**Eligible Renewal Opportunities:** {eligible_renewal_count}")
    st.write(f"**Successfully Renewed Opportunities:** {eligible_renewed_count}")
    st.write(f"**Renewal Success Rate:** {renewal_percentage:.2f}%")
else:
    st.warning("No renewal opportunities found.")

# Visualization
if eligible_renewal_count > 0:
    # Prepare data for charts
    data = {
        "Category": ["Renewed", "Not Renewed"],
        "Count": [eligible_renewed_count, eligible_renewal_count - eligible_renewed_count],
    }
    df = pd.DataFrame(data)

    # Sidebar chart selection
    chart_type = st.sidebar.selectbox(
        "Select Chart Type",
        options=[
            "Pie Chart",
            "Bar Chart",
            "Line Chart (Weekly Trend)",
            "Area Chart",
            "Gauge Chart",
            "Treemap",
        ],
    )

    # Generate selected chart
    st.subheader("Visualization")
    if chart_type == "Pie Chart":
        fig = px.pie(
            df,
            names="Category",
            values="Count",
            title="Renewed vs Not Renewed Opportunities",
            color="Category",
            color_discrete_map={"Renewed": "#2ecc71", "Not Renewed": "#e74c3c"}
        )
    elif chart_type == "Bar Chart":
        fig = px.bar(
            df,
            x="Category",
            y="Count",
            title="Renewed vs Not Renewed Opportunities",
            text="Count",
            color="Category",
            color_discrete_map={"Renewed": "#2ecc71", "Not Renewed": "#e74c3c"}
        )
    elif chart_type == "Line Chart (Weekly Trend)":
        if not weekly_df.empty:
            fig = px.line(
                weekly_df,
                x="Week",
                y="RenewalRate",
                title="Weekly Renewal Rate Trend (%)",
                markers=True,
            )
            fig.update_layout(yaxis_title="Renewal Rate (%)")
        else:
            st.warning("Weekly trend data not available.")
            fig = go.Figure()
    elif chart_type == "Area Chart":
        if not weekly_df.empty:
            fig = px.area(
                weekly_df,
                x="Week",
                y=["RenewedCount", "EligibleCount"],
                title="Weekly Renewal Opportunities",
            )
            fig.update_layout(yaxis_title="Count")
        else:
            st.warning("Weekly trend data not available.")
            fig = go.Figure()
    elif chart_type == "Gauge Chart":
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=renewal_percentage,
            title={"text": "Renewal Success Rate"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": "#1f77b4"},
                "steps": [
                    {"range": [0, 50], "color": "#e74c3c"},
                    {"range": [50, 75], "color": "#f39c12"},
                    {"range": [75, 100], "color": "#2ecc71"}
                ],
            }
        ))
    elif chart_type == "Treemap":
        fig = px.treemap(
            df,
            path=["Category"],
            values="Count",
            title="Treemap of Renewal Opportunities",
            color="Category",
            color_discrete_map={"Renewed": "#2ecc71", "Not Renewed": "#e74c3c"}
        )

    # Display the selected chart
    st.plotly_chart(fig)
    
    # Weekly trend table
    if not weekly_df.empty:
        st.subheader("Weekly Renewal Trend")
        weekly_display = weekly_df.copy()
        weekly_display['RenewalRate'] = weekly_display['RenewalRate'].round(2).astype(str) + '%'
        st.dataframe(weekly_display)

# Optionally show raw data
if show_data and eligible_renewal_count > 0:
    st.subheader("Raw Data")
    st.dataframe(df)
