# Connector Channel Sales Performance Dashboard

A simulated MIS dashboard tracking the lead-to-disbursal sales funnel across regions, 
loan products, and connector segments, built to mirror how Sales Excellence/MIS teams 
in BFSI track connector channel performance.

## Problem Statement
Banking connector channels generate leads across thousands of connectors, but tracking 
conversion (lead → approval → disbursal) across regions, products, and connector 
segments requires structured reporting and automation to flag underperformance early.

## Approach
- **Data**: Simulated dataset of 5,000 leads across 60 connectors, 6 regions, 
  4 loan products, and 4 connector segments
- **Excel layer**: Pivot Tables for region/product/segment-wise funnel analysis, 
  XLOOKUP + SUMIFS/COUNTIF-based connector scorecard, conditional formatting to 
  flag low-converting connectors
- **Python layer**: Automated funnel metric calculation (pandas) and auto-generated 
  PowerPoint business review deck (python-pptx)

## Tools
Excel (Pivot Tables, XLOOKUP, SUMIFS, COUNTIF, Conditional Formatting), 
Python (pandas, python-pptx)

## Key Insight
Chennai region shows the lowest conversion rate at 28.1%, driven primarily by 
underperformance in the Aggregator connector segment (39.2% conversion vs. 
~60%+ for other segments). This points to a targeted root cause — Aggregator-sourced 
leads in Chennai specifically — rather than a broad regional or channel-wide issue.

## How to Run
1. `pip install pandas python-pptx`
2. `python generate_report.py`
3. Generates `Connector_Channel_Weekly_Report.pptx` with funnel charts and insight slide

## Screenshots
![Region Funnel](screenshots/region_funnel.png)
![Connector Scorecard](screenshots/connector_scorecard.png)
![Generated Report](screenshots/ppt_slide.png)
