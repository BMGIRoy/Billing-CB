import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import numpy as np
from datetime import datetime

def create_time_series_chart(df):
    """
    Create a time series chart showing monthly T Amt and N Amt
    """
    # Group by month and calculate sum of T Amt and N Amt
    monthly_data = df.groupby('Year-Month').agg({
        'T Amt': 'sum',
        'N Amt': 'sum'
    }).reset_index()
    
    # Ensure data is sorted chronologically
    monthly_data['Date'] = pd.to_datetime(monthly_data['Year-Month'] + '-01')
    monthly_data = monthly_data.sort_values('Date')
    
    # Create figure with two y-axes
    fig = go.Figure()
    
    # Add T Amt line
    fig.add_trace(go.Scatter(
        x=monthly_data['Date'], 
        y=monthly_data['T Amt'],
        mode='lines+markers',
        name='Total Amount',
        line=dict(color='royalblue', width=3),
        marker=dict(size=8)
    ))
    
    # Add N Amt line
    fig.add_trace(go.Scatter(
        x=monthly_data['Date'], 
        y=monthly_data['N Amt'],
        mode='lines+markers',
        name='Net Amount',
        line=dict(color='firebrick', width=3),
        marker=dict(size=8)
    ))
    
    # Format the chart
    fig.update_layout(
        title='Monthly Financial Trends',
        xaxis_title='Month',
        yaxis_title='Amount ($)',
        legend_title='Metric',
        hovermode='x unified',
        height=500
    )
    
    # Format x-axis to show month and year
    fig.update_xaxes(
        tickformat='%b %Y',
        tickmode='auto',
        nticks=12
    )
    
    return fig

def create_hierarchy_chart(df):
    """
    Create a hierarchical visualization (treemap) of business performance
    """
    # Group by hierarchy and calculate sum of T Amt
    hierarchy_data = df.groupby(['Business Head', 'Consultant', 'Client']).agg({
        'T Amt': 'sum'
    }).reset_index()
    
    # Create treemap
    fig = px.treemap(
        hierarchy_data,
        path=['Business Head', 'Consultant', 'Client'],
        values='T Amt',
        color='T Amt',
        color_continuous_scale='Blues',
        title='Business Hierarchy Performance (Total Amount)'
    )
    
    # Format the chart
    fig.update_layout(
        height=600,
        margin=dict(l=0, r=0, t=30, b=0)
    )
    
    # Format hover text
    fig.update_traces(
        hovertemplate='<b>%{label}</b><br>Total Amount: $%{value:,.2f}'
    )
    
    return fig

def create_comparison_chart(df):
    """
    Create a scatter plot comparing T Amt vs N Amt by consultant
    """
    # Group by consultant and calculate sum of T Amt and N Amt
    comparison_data = df.groupby('Consultant').agg({
        'T Amt': 'sum',
        'N Amt': 'sum',
        'Client': 'nunique'  # Number of unique clients per consultant
    }).reset_index()
    
    # Create scatter plot
    fig = px.scatter(
        comparison_data,
        x='T Amt',
        y='N Amt',
        size='Client',  # Bubble size based on number of clients
        hover_name='Consultant',
        text='Consultant',
        title='T Amt vs N Amt Comparison by Consultant',
        labels={
            'T Amt': 'Total Amount ($)',
            'N Amt': 'Net Amount ($)',
            'Client': 'Number of Clients'
        },
        size_max=30,
        color='T Amt',
        color_continuous_scale='Viridis'
    )
    
    # Draw diagonal reference line (T Amt = N Amt)
    max_val = max(comparison_data['T Amt'].max(), comparison_data['N Amt'].max())
    fig.add_trace(go.Scatter(
        x=[0, max_val],
        y=[0, max_val],
        mode='lines',
        line=dict(color='grey', dash='dash'),
        name='T Amt = N Amt',
        hoverinfo='skip'
    ))
    
    # Format the chart
    fig.update_layout(
        height=600,
        hovermode='closest'
    )
    
    # Reduce text size to avoid overlap
    fig.update_traces(
        textposition='top center',
        textfont=dict(size=10)
    )
    
    return fig

def create_quarterly_chart(df):
    """
    Create a quarterly analysis bar chart for fiscal quarters
    """
    # Group by fiscal year and quarter to calculate sum of T Amt and N Amt
    quarterly_data = df.groupby(['Fiscal Year', 'Fiscal Quarter']).agg({
        'T Amt': 'sum',
        'N Amt': 'sum'
    }).reset_index()
    
    # Create a combined period column for proper ordering (e.g., "FY 2022-23 Q1")
    quarterly_data['Period'] = quarterly_data['Fiscal Year'] + ' ' + quarterly_data['Fiscal Quarter']
    
    # Define proper quarter order
    quarter_order = ['Q1', 'Q2', 'Q3', 'Q4']
    
    # Sort data by fiscal year and quarter
    quarterly_data['Quarter_Num'] = quarterly_data['Fiscal Quarter'].apply(lambda q: quarter_order.index(q))
    quarterly_data = quarterly_data.sort_values(['Fiscal Year', 'Quarter_Num'])
    
    # Create a grouped bar chart
    fig = go.Figure()
    
    # Add T Amt bars
    fig.add_trace(go.Bar(
        x=quarterly_data['Period'],
        y=quarterly_data['T Amt'],
        name='Total Amount',
        marker_color='royalblue'
    ))
    
    # Add N Amt bars
    fig.add_trace(go.Bar(
        x=quarterly_data['Period'],
        y=quarterly_data['N Amt'],
        name='Net Amount',
        marker_color='firebrick'
    ))
    
    # Format the chart
    fig.update_layout(
        title='Quarterly Financial Performance',
        xaxis_title='Fiscal Quarter',
        yaxis_title='Amount ($)',
        legend_title='Metric',
        barmode='group',
        height=500
    )
    
    return fig

def create_annual_chart(df):
    """
    Create an annual financial trends chart
    """
    # Group by fiscal year to calculate sum of T Amt and N Amt
    annual_data = df.groupby('Fiscal Year').agg({
        'T Amt': 'sum',
        'N Amt': 'sum'
    }).reset_index()
    
    # Sort data by fiscal year
    annual_data = annual_data.sort_values('Fiscal Year')
    
    # Create a grouped bar chart
    fig = go.Figure()
    
    # Add T Amt bars
    fig.add_trace(go.Bar(
        x=annual_data['Fiscal Year'],
        y=annual_data['T Amt'],
        name='Total Amount',
        marker_color='royalblue'
    ))
    
    # Add N Amt bars
    fig.add_trace(go.Bar(
        x=annual_data['Fiscal Year'],
        y=annual_data['N Amt'],
        name='Net Amount',
        marker_color='firebrick'
    ))
    
    # Add line showing percentage difference
    annual_data['Diff_Percent'] = (annual_data['T Amt'] - annual_data['N Amt']) / annual_data['T Amt'] * 100
    
    fig.add_trace(go.Scatter(
        x=annual_data['Fiscal Year'],
        y=annual_data['Diff_Percent'],
        mode='lines+markers',
        name='Difference (%)',
        line=dict(color='green', width=3),
        marker=dict(size=8),
        yaxis='y2'
    ))
    
    # Format the chart with dual y-axis
    fig.update_layout(
        title='Annual Financial Trends',
        xaxis_title='Fiscal Year',
        yaxis_title='Amount ($)',
        yaxis2=dict(
            title='Difference (%)',
            titlefont=dict(color='green'),
            tickfont=dict(color='green'),
            overlaying='y',
            side='right'
        ),
        legend_title='Metric',
        barmode='group',
        height=500
    )
    
    return fig

def create_consultant_performance_chart(df):
    """
    Create a chart showing consultant performance
    """
    # Group by consultant to calculate various performance metrics
    consultant_data = df.groupby('Consultant').agg({
        'T Amt': 'sum',
        'N Amt': 'sum',
        'Client': 'nunique',
        'Date': 'count'  # Number of billing entries as a proxy for activity
    }).reset_index()
    
    # Calculate the average amount per billing
    consultant_data['Avg_Billing'] = consultant_data['T Amt'] / consultant_data['Date']
    
    # Sort consultants by total amount
    consultant_data = consultant_data.sort_values('T Amt', ascending=False)
    
    # Take top 10 consultants by total amount
    top_consultants = consultant_data.head(10)
    
    # Create a bar chart with secondary y-axis
    fig = go.Figure()
    
    # Add T Amt bars
    fig.add_trace(go.Bar(
        x=top_consultants['Consultant'],
        y=top_consultants['T Amt'],
        name='Total Amount',
        marker_color='royalblue'
    ))
    
    # Add N Amt bars
    fig.add_trace(go.Bar(
        x=top_consultants['Consultant'],
        y=top_consultants['N Amt'],
        name='Net Amount',
        marker_color='firebrick'
    ))
    
    # Add number of clients line
    fig.add_trace(go.Scatter(
        x=top_consultants['Consultant'],
        y=top_consultants['Client'],
        mode='lines+markers',
        name='Number of Clients',
        line=dict(color='green', width=3),
        marker=dict(size=8),
        yaxis='y2'
    ))
    
    # Format the chart with dual y-axis
    fig.update_layout(
        title='Top 10 Consultant Performance',
        xaxis_title='Consultant',
        yaxis_title='Amount ($)',
        yaxis2=dict(
            title='Number of Clients',
            titlefont=dict(color='green'),
            tickfont=dict(color='green'),
            overlaying='y',
            side='right',
            range=[0, max(top_consultants['Client']) * 1.2]
        ),
        legend_title='Metric',
        height=600,
        xaxis=dict(
            tickangle=45
        )
    )
    
    return fig
