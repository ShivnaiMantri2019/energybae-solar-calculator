# Energybae Solar Load Calculator

## Overview
AI-powered tool that reads MSEDCL electricity bills
and generates Solar Load Calculator Excel report automatically.

## What It Does
1. User uploads MSEDCL electricity bill (JPG/PNG/PDF)
2. AI reads and extracts all key data from the bill
3. Solar system size is calculated automatically
4. Formatted Excel report is downloaded instantly

## Tools Used
- Google Colab
- OpenRouter API (Free)
- NVIDIA Nemotron Vision Language Model
- openpyxl (Excel generation)
- Pillow (image processing)

## How To Run
1. Open Google Colab
2. Run Cell 1 to install libraries
3. Run Cell 2 to upload your bill
4. Run Cell 3 to set your API key
5. Run remaining cells one by one
6. Excel report downloads automatically

## Solar Calculation Logic
- Avg Monthly Units = Sum of 12 months / 12
- Solar Capacity (kW) = Avg daily units / 4.5 x 1.25
- No. of Panels = Solar kW x 1000 / 450W
- Annual Savings = Avg units x 12 x Rs 7.5
- Payback Period = Installed cost / Annual savings
- CO2 Offset = Avg units x 12 x 0.82 kg

## Excel Report Structure
- Section 1 — Consumer Details
- Section 2 — Meter Reading
- Section 3 — 12-Month Consumption History
- Section 4 — Solar System Recommendation

## What I Would Improve Next
- Add Streamlit web interface
- Batch processing for multiple bills
- Better solar formulas with roof and shading data
- Auto email Excel report to customer

## About
Built for Energybae — Empowering People with Renewable Energy
www.energybae.in | energybae.co@gmail.com | +91 9112233120
