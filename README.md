# Shipment Summary Analyzer

This Streamlit app takes the output Excel from the **MS Shipment Analyzer** app and produces a clean summary report.

## Features
- Excludes shipments with 2 stops or no milestone received
- Generates 3 pivot-style summaries:
  - Out-of-order vs. milestone status
  - Origin-level summary with % out of order
  - Carrier-level summary with % out of order
- Exports a downloadable Excel with two sheets:
  - `Filtered_Data`
  - `Summary_Analysis`

## How to Run Locally
1. Clone this repo or create a new Streamlit app.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
