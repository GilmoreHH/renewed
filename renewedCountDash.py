import streamlit as st
from simple_salesforce import Salesforce
import os
from dotenv import load_dotenv
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np
import datetime

# Load environment variables from .env file
load_dotenv()

# Set page configuration with an icon for Renewals Count Dashboard
st.set_page_config(
    page_title="Renewals Count",
    page_icon="🔄",  # Icon representing renewals (e.g., a refresh/renewal symbol)
    layout="wide",  # Use a wide layout for the app
)


# Define all stages and their probabilities
def get_stage_metadata():
    """Return a dictionary of stages with their probabilities and order."""
    stages = {
        "New": {"probability": 10, "order": 1, "category": "Open"},
        "Information Gathering": {"probability": 20, "order": 2, "category": "Open"},
        "Rating": {"probability": 30, "order": 3, "category": "Open"},
        "Proposal Generation": {"probability": 40, "order": 4, "category": "Open"},
        "Decision Pending": {"probability": 50, "order": 5, "category": "Open"},
        "Pre-Bind Review": {"probability": 75, "order": 6, "category": "Open"},
        "Quote to Bind": {"probability": 85, "order": 7, "category": "Open"},
        "Binding": {"probability": 85, "order": 8, "category": "Open"},
        "Billing": {"probability": 95, "order": 9, "category": "Open"},
        "Post-Binding": {"probability": 98, "order": 10, "category": "Open"},
        "Closed Won": {"probability": 100, "order": 11, "category": "Closed/Won"},
        "Closed Lost": {"probability": 0, "order": 12, "category": "Closed/Lost"}
    }
    return stages

# Define loss reasons
def get_loss_reasons():
    """Return a list of all possible loss reasons."""
    return [
        "4 Point Issues",
        "AOR",
        "Choosing to Self-Insure",
        "Flood not required - for flood only",
        "I was lazy",
        "No Buyers Information Received",
        "No Inspections Received",
        "No Market",
        "No Response from Buyer",
        "Paid off Mortgage",
        "Rate",
        "Rate / Went with another agency",
        "Rate - No Updated Inspections Received",
        "Sale Fell Through",
        "Service / went with another agency",
        "Sold",
        "Unknown",
        "Property Damaged Or Lost"
    ]

# Function to get ISO week calendar data for 2025
def get_iso_week_calendar():
    """Return a DataFrame with ISO week numbers and date ranges for 2025."""
    weeks = []
    for week_num in range(1, 53):
        # Create entry for each week
        weeks.append({
            "Week": f"Week {week_num:02d}",
            "From": "",
            "To": ""
        })
    
    # Populate with the specific dates provided
    iso_weeks = {
        1: {"from": "Dec 30, 2024", "to": "Jan 5, 2025"},
        2: {"from": "Jan 6, 2025", "to": "Jan 12, 2025"},
        3: {"from": "Jan 13, 2025", "to": "Jan 19, 2025"},
        4: {"from": "Jan 20, 2025", "to": "Jan 26, 2025"},
        5: {"from": "Jan 27, 2025", "to": "Feb 2, 2025"},
        6: {"from": "Feb 3, 2025", "to": "Feb 9, 2025"},
        7: {"from": "Feb 10, 2025", "to": "Feb 16, 2025"},
        8: {"from": "Feb 17, 2025", "to": "Feb 23, 2025"},
        9: {"from": "Feb 24, 2025", "to": "Mar 2, 2025"},
        10: {"from": "Mar 3, 2025", "to": "Mar 9, 2025"},
        11: {"from": "Mar 10, 2025", "to": "Mar 16, 2025"},
        12: {"from": "Mar 17, 2025", "to": "Mar 23, 2025"},
        13: {"from": "Mar 24, 2025", "to": "Mar 30, 2025"},
        14: {"from": "Mar 31, 2025", "to": "Apr 6, 2025"},
        15: {"from": "Apr 7, 2025", "to": "Apr 13, 2025"},
        16: {"from": "Apr 14, 2025", "to": "Apr 20, 2025"},
        17: {"from": "Apr 21, 2025", "to": "Apr 27, 2025"},
        18: {"from": "Apr 28, 2025", "to": "May 4, 2025"},
        19: {"from": "May 5, 2025", "to": "May 11, 2025"},
        20: {"from": "May 12, 2025", "to": "May 18, 2025"},
        21: {"from": "May 19, 2025", "to": "May 25, 2025"},
        22: {"from": "May 26, 2025", "to": "Jun 1, 2025"},
        23: {"from": "Jun 2, 2025", "to": "Jun 8, 2025"},
        24: {"from": "Jun 9, 2025", "to": "Jun 15, 2025"},
        25: {"from": "Jun 16, 2025", "to": "Jun 22, 2025"},
        26: {"from": "Jun 23, 2025", "to": "Jun 29, 2025"},
        27: {"from": "Jun 30, 2025", "to": "Jul 6, 2025"},
        28: {"from": "Jul 7, 2025", "to": "Jul 13, 2025"},
        29: {"from": "Jul 14, 2025", "to": "Jul 20, 2025"},
        30: {"from": "Jul 21, 2025", "to": "Jul 27, 2025"},
        31: {"from": "Jul 28, 2025", "to": "Aug 3, 2025"},
        32: {"from": "Aug 4, 2025", "to": "Aug 10, 2025"},
        33: {"from": "Aug 11, 2025", "to": "Aug 17, 2025"},
        34: {"from": "Aug 18, 2025", "to": "Aug 24, 2025"},
        35: {"from": "Aug 25, 2025", "to": "Aug 31, 2025"},
        36: {"from": "Sep 1, 2025", "to": "Sep 7, 2025"},
        37: {"from": "Sep 8, 2025", "to": "Sep 14, 2025"},
        38: {"from": "Sep 15, 2025", "to": "Sep 21, 2025"},
        39: {"from": "Sep 22, 2025", "to": "Sep 28, 2025"},
        40: {"from": "Sep 29, 2025", "to": "Oct 5, 2025"},
        41: {"from": "Oct 6, 2025", "to": "Oct 12, 2025"},
        42: {"from": "Oct 13, 2025", "to": "Oct 19, 2025"},
        43: {"from": "Oct 20, 2025", "to": "Oct 26, 2025"},
        44: {"from": "Oct 27, 2025", "to": "Nov 2, 2025"},
        45: {"from": "Nov 3, 2025", "to": "Nov 9, 2025"},
        46: {"from": "Nov 10, 2025", "to": "Nov 16, 2025"},
        47: {"from": "Nov 17, 2025", "to": "Nov 23, 2025"},
        48: {"from": "Nov 24, 2025", "to": "Nov 30, 2025"},
        49: {"from": "Dec 1, 2025", "to": "Dec 7, 2025"},
        50: {"from": "Dec 8, 2025", "to": "Dec 14, 2025"},
        51: {"from": "Dec 15, 2025", "to": "Dec 21, 2025"},
        52: {"from": "Dec 22, 2025", "to": "Dec 28, 2025"},
    }
    
    for i, week in enumerate(weeks):
        week_num = i + 1
        if week_num in iso_weeks:
            week["From"] = iso_weeks[week_num]["from"]
            week["To"] = iso_weeks[week_num]["to"]
    
    return pd.DataFrame(weeks)

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

        # Get stage metadata
        stage_metadata = get_stage_metadata()
        
        # Query to get the count of renewal opportunities by stage
        stage_query = """
            SELECT StageName, COUNT(Id) oppCount
            FROM Opportunity
            WHERE New_Business_or_Renewal__c IN ('Personal Lines - Renewal', 'Commercial Lines - Renewal')
            AND CloseDate = LAST_N_DAYS:7
            GROUP BY StageName
        """
        stage_results = sf.query_all(stage_query)
        
        # Process stage results into a DataFrame
        stage_data = []
        for record in stage_results['records']:
            stage_name = record['StageName']
            count = record['oppCount']
            
            # Get stage metadata or use defaults if not found
            metadata = stage_metadata.get(stage_name, {"probability": 0, "order": 99, "category": "Unknown"})
            
            stage_data.append({
                'StageName': stage_name,
                'Count': count,
                'Probability': metadata["probability"],
                'Order': metadata["order"],
                'Category': metadata["category"]
            })
        
        # Create DataFrame and ensure all stages are represented
        stage_df = pd.DataFrame(stage_data)
        
        # Add any missing stages with count 0
        existing_stages = set(stage_df['StageName'])
        for stage_name, metadata in stage_metadata.items():
            if stage_name not in existing_stages:
                stage_df = pd.concat([stage_df, pd.DataFrame([{
                    'StageName': stage_name,
                    'Count': 0,
                    'Probability': metadata["probability"],
                    'Order': metadata["order"],
                    'Category': metadata["category"]
                }])], ignore_index=True)
                
        # Sort by the defined order
        stage_df = stage_df.sort_values('Order')
        
        # Get all possible loss reasons
        all_loss_reasons = get_loss_reasons()
        
        # Query to get reasons for closed lost opportunities
        loss_reason_query = """
            SELECT Loss_Reason__c, COUNT(Id) reasonCount
            FROM Opportunity
            WHERE New_Business_or_Renewal__c IN ('Personal Lines - Renewal', 'Commercial Lines - Renewal')
            AND StageName = 'Closed Lost'
            AND CloseDate = LAST_N_DAYS:7
            GROUP BY Loss_Reason__c
            ORDER BY COUNT(Id) DESC
        """
        loss_reason_results = sf.query_all(loss_reason_query)
        
        # Process loss reason results into a DataFrame
        loss_reason_data = []
        for record in loss_reason_results['records']:
            reason = record['Loss_Reason__c'] if record['Loss_Reason__c'] else 'Not Specified'
            loss_reason_data.append({
                'Loss_Reason': reason,
                'Count': record['reasonCount']
            })
        
        # Create DataFrame and ensure all reasons are represented
        loss_reason_df = pd.DataFrame(loss_reason_data)
        
        # Add any missing reasons with count 0
        existing_reasons = set(loss_reason_df['Loss_Reason'])
        for reason in all_loss_reasons:
            if reason not in existing_reasons:
                loss_reason_df = pd.concat([loss_reason_df, pd.DataFrame([{
                    'Loss_Reason': reason,
                    'Count': 0
                }])], ignore_index=True)
                
        # Sort by count
        loss_reason_df = loss_reason_df.sort_values('Count', ascending=False)
        
        # Get data for weekly trend analysis (last 8 weeks)
        all_opps_query = """
            SELECT 
                Id, 
                CloseDate, 
                StageName,
                Loss_Reason__c
            FROM Opportunity
            WHERE New_Business_or_Renewal__c IN ('Personal Lines - Renewal', 'Commercial Lines - Renewal')
            AND CloseDate = LAST_N_DAYS:91
        """
        all_opps_results = sf.query_all(all_opps_query)
        
        # Get stage metadata
        stage_metadata = get_stage_metadata()
        
        # Process opportunities by week
        import datetime
        
        # Create dictionaries to track counts by week
        eligible_by_week = {}
        stage_counts_by_week = {stage: {} for stage in stage_metadata.keys()}
        
        # Process all opportunities
        for record in all_opps_results['records']:
            # Convert CloseDate string to date object
            close_date = datetime.datetime.strptime(record['CloseDate'], '%Y-%m-%d').date()
            # Get ISO week number (1-53)
            year, week_num, _ = close_date.isocalendar()
            week_key = f"{year}-W{week_num:02d}"
            
            # Count eligible opportunities
            eligible_by_week[week_key] = eligible_by_week.get(week_key, 0) + 1
            
            # Count by stage
            stage = record['StageName']
            if stage not in stage_counts_by_week:
                stage_counts_by_week[stage] = {}
            
            stage_counts_by_week[stage][week_key] = stage_counts_by_week[stage].get(week_key, 0) + 1
        
        # Combine the data
        weekly_data = []
        for week in sorted(eligible_by_week.keys()):
            eligible_count = eligible_by_week[week]
            
            week_data = {
                'Week': week,
                'TotalOpportunities': eligible_count,
            }
            
            # Add counts for each stage
            for stage, counts_by_week in stage_counts_by_week.items():
                stage_count = counts_by_week.get(week, 0)
                # Clean stage name for column
                column_name = ''.join(c if c.isalnum() else '_' for c in stage)
                week_data[column_name] = stage_count
            
            # Add special calculated columns
            closed_won_count = stage_counts_by_week.get('Closed Won', {}).get(week, 0)
            closed_lost_count = stage_counts_by_week.get('Closed Lost', {}).get(week, 0)
            
            # Calculate other stages (all except Closed Won and Closed Lost)
            other_stages_count = eligible_count - closed_won_count - closed_lost_count
            
            week_data['ClosedWon'] = closed_won_count
            week_data['ClosedLost'] = closed_lost_count
            week_data['OtherStages'] = other_stages_count
            week_data['WinRate'] = (closed_won_count / eligible_count * 100) if eligible_count > 0 else 0
            
            weekly_data.append(week_data)
        
        weekly_df = pd.DataFrame(weekly_data)

        # Calculate summary statistics
        total_opportunities = stage_df['Count'].sum() if not stage_df.empty else 0
        closed_won_count = stage_df[stage_df['StageName'] == 'Closed Won']['Count'].sum() if not stage_df.empty else 0
        closed_lost_count = stage_df[stage_df['StageName'] == 'Closed Lost']['Count'].sum() if not stage_df.empty else 0
        other_stages_count = total_opportunities - closed_won_count - closed_lost_count
        
        return stage_df, loss_reason_df, weekly_df, total_opportunities, closed_won_count, closed_lost_count, other_stages_count
    
    except Exception as e:
        st.error(f"Error connecting to Salesforce: {str(e)}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), 0, 0, 0, 0

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

opportunity_type = st.sidebar.selectbox(
    "Opportunity Type",
    options=["All Renewals", "Personal Lines - Renewal", "Commercial Lines - Renewal"],
    index=0
)

# Sidebar ISO Week Calendar
show_iso_calendar = st.sidebar.checkbox("Show ISO Week Calendar", value=False)

# Main content
stage_df, loss_reason_df, weekly_df, total_opportunities, closed_won_count, closed_lost_count, other_stages_count = connect_to_salesforce()

# Renewal Scorecard
st.subheader("Renewal Opportunity Scorecard")
col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric("Total Opportunities", total_opportunities)
with col2:
    st.metric("Closed Won (Bound, Paid)", closed_won_count)
with col3:
    st.metric("Closed Lost", closed_lost_count)
with col4:
    win_rate = (closed_won_count / total_opportunities * 100) if total_opportunities > 0 else 0
    st.metric("Win Rate", f"{win_rate:.2f}%")

# ISO Week Calendar expander
if show_iso_calendar:
    with st.expander("ISO Week Calendar 2025", expanded=True):
        iso_calendar_df = get_iso_week_calendar()
        
        # Format the calendar in a 13x4 grid (13 weeks per quarter, 4 quarters per year)
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.subheader("Q1 (Weeks 1-13)")
            st.dataframe(iso_calendar_df.iloc[0:13], hide_index=True)
        
        with col2:
            st.subheader("Q2 (Weeks 14-26)")
            st.dataframe(iso_calendar_df.iloc[13:26], hide_index=True)
        
        with col3:
            st.subheader("Q3 (Weeks 27-39)")
            st.dataframe(iso_calendar_df.iloc[26:39], hide_index=True)
        
        with col4:
            st.subheader("Q4 (Weeks 40-52)")
            st.dataframe(iso_calendar_df.iloc[39:52], hide_index=True)

# Visualization selector
chart_type = st.sidebar.selectbox(
    "Select Chart Type",
    options=[
        "Status Distribution",
        "Closed Won vs Closed Lost",
        "Pipeline by Stage",
        "Weekly Trend",
        "Loss Reasons",
        "Win Rate Gauge",
        "Stage Progression"
    ],
)

# Generate selected chart
st.subheader("Visualization")
if chart_type == "Status Distribution":
    # Group by category
    category_data = stage_df.groupby('Category')['Count'].sum().reset_index()
    
    fig = px.pie(
        category_data,
        names="Category",
        values="Count",
        title="Renewal Opportunity Status Distribution",
        color="Category",
        color_discrete_map={
            "Closed/Won": "#2ecc71", 
            "Closed/Lost": "#e74c3c",
            "Open": "#f39c12"
        }
    )
    st.plotly_chart(fig)

elif chart_type == "Closed Won vs Closed Lost":
    # Filter for just Closed Won and Closed Lost
    closed_data = stage_df[stage_df['StageName'].isin(['Closed Won', 'Closed Lost'])]
    
    fig = px.bar(
        closed_data,
        x="StageName",
        y="Count",
        title="Closed Won vs Closed Lost Opportunities",
        text="Count",
        color="StageName",
        color_discrete_map={
            "Closed Won": "#2ecc71", 
            "Closed Lost": "#e74c3c"
        }
    )
    st.plotly_chart(fig)
    
elif chart_type == "Pipeline by Stage":
    # Create a bar chart showing all stages
    fig = px.bar(
        stage_df,
        x="StageName",
        y="Count",
        title="Pipeline by Stage",
        text="Count",
        color="Probability",
        color_continuous_scale=px.colors.sequential.Viridis
    )
    fig.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig)
    
    # Add a table with stage details
    st.subheader("Stage Details")
    details_df = stage_df[['StageName', 'Count', 'Probability', 'Category']].copy()
    details_df = details_df.sort_values('Count', ascending=False)
    st.dataframe(details_df)

elif chart_type == "Weekly Trend":
    if not weekly_df.empty:
        # Add a reference to the ISO week calendar
        st.info("👆 Enable 'Show ISO Week Calendar' in the sidebar to see the date ranges for each ISO week.")
        
        fig = px.line(
            weekly_df,
            x="Week",
            y=["ClosedWon", "ClosedLost", "OtherStages"],
            title="Weekly Renewal Opportunity Trend",
            markers=True,
        )
        fig.update_layout(yaxis_title="Count")
        st.plotly_chart(fig)
        
        # Also show win rate trend
        fig2 = px.line(
            weekly_df,
            x="Week",
            y="WinRate",
            title="Weekly Win Rate Trend (%)",
            markers=True,
        )
        fig2.update_layout(yaxis_title="Win Rate (%)")
        st.plotly_chart(fig2)
    else:
        st.warning("Weekly trend data not available.")

elif chart_type == "Loss Reasons":
    if not loss_reason_df.empty:
        # Filter out zero counts
        non_zero_reasons = loss_reason_df[loss_reason_df['Count'] > 0]
        
        if not non_zero_reasons.empty:
            # Create horizontal bar chart for loss reasons
            fig = px.bar(
                non_zero_reasons.head(10),  # Show top 10 reasons
                y="Loss_Reason",
                x="Count",
                title="Top Loss Reasons by Frequency",
                text="Count",
                orientation='h',
                color_discrete_sequence=["#e74c3c"]
            )
            fig.update_layout(yaxis_title="", xaxis_title="Count")
            st.plotly_chart(fig)
        
        # Show complete table of loss reasons
        st.subheader("All Loss Reasons")
        st.dataframe(loss_reason_df)
    else:
        st.warning("No loss reason data available.")

elif chart_type == "Win Rate Gauge":
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=win_rate,
        title={"text": "Win Rate (%)"},
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
    st.plotly_chart(fig)
    
    # Add a table with win rate by stage
    if not stage_df.empty:
        st.subheader("Conversion Analysis")
        
        # Calculate conversion rates between stages
        stage_df_sorted = stage_df.sort_values('Order')
        stages_list = stage_df_sorted['StageName'].tolist()
        
        conversion_data = []
        for i in range(len(stages_list) - 1):
            current_stage = stages_list[i]
            next_stage = stages_list[i + 1]
            
            current_count = stage_df_sorted[stage_df_sorted['StageName'] == current_stage]['Count'].iloc[0]
            next_count = stage_df_sorted[stage_df_sorted['StageName'] == next_stage]['Count'].iloc[0]
            
            conversion_rate = (next_count / current_count * 100) if current_count > 0 else 0
            
            conversion_data.append({
                'From Stage': current_stage,
                'To Stage': next_stage,
                'From Count': current_count,
                'To Count': next_count,
                'Conversion Rate': f"{conversion_rate:.2f}%",
            })
        
        conversion_df = pd.DataFrame(conversion_data)
        st.dataframe(conversion_df)

elif chart_type == "Stage Progression":
    # Create a funnel chart showing progression through stages
    if not stage_df.empty:
        # Sort by the defined order
        funnel_data = stage_df.sort_values('Order')
        
        fig = go.Figure(go.Funnel(
            y=funnel_data['StageName'],
            x=funnel_data['Count'],
            textinfo="value+percent initial",
            marker={"color": funnel_data['Probability'], "colorscale": "Viridis"}
        ))
        
        fig.update_layout(
            title="Sales Pipeline Funnel",
            margin=dict(l=150)
        )
        
        st.plotly_chart(fig)

# Show detailed opportunities by stage
if not stage_df.empty:
    st.subheader("Opportunities by Stage")
    # Sort stages with Closed Won and Closed Lost first, then others alphabetically
    def stage_sorter(stage_name):
        if stage_name == "Closed Won":
            return "1"
        elif stage_name == "Closed Lost":
            return "2"
        else:
            return "3" + stage_name
    
    stage_df['SortKey'] = stage_df['StageName'].apply(stage_sorter)
    stage_df = stage_df.sort_values('SortKey').drop('SortKey', axis=1)
    st.dataframe(stage_df)

# Show loss reasons table
if not loss_reason_df.empty:
    st.subheader("Loss Reasons Analysis")
    st.dataframe(loss_reason_df)

# Weekly trend table
if not weekly_df.empty:
    st.subheader("Weekly Renewal Trend")
    display_df = weekly_df.copy()
    display_df['WinRate'] = display_df['WinRate'].round(2).astype(str) + '%'
    st.dataframe(display_df)
    
    # Add reminder about ISO week calendar
    if not show_iso_calendar:
        st.info("📅 Enable 'Show ISO Week Calendar' in the sidebar to see the date ranges for each ISO week.")

# Optionally show raw data
if show_data:
    st.subheader("Raw Data")
    st.write("Stage Data:")
    st.dataframe(stage_df)
    
    st.write("Loss Reason Data:")
    st.dataframe(loss_reason_df)
    
    st.write("Weekly Trend Data:")
    st.dataframe(weekly_df)
    
    st.write("ISO Week Calendar:")
    st.dataframe(get_iso_week_calendar())
