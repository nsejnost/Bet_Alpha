# CBB Market vs Model

## Overview
A Flask web application that compares College Basketball market odds against predictions from KenPom, Haslametrics, and Barttorvik models. It fetches live odds from the Odds API, runs model comparisons, and displays totals and spreads comparison tables in the browser.

## Recent Changes
- 2026-02-11: Initial Replit setup — installed Python 3.12, Flask, numpy, pandas, openpyxl. Configured workflow and deployment.

## Project Architecture
- `app.py` — Flask web server (entry point), serves HTML and `/api/data` endpoint
- `main.py` — Core data pipeline: loads model data from Excel, fetches odds from API, computes comparisons
- `templates/index.html` — Single-page frontend with tabbed Totals/Spreads tables
- `CBB_Model_Inputs.xlsx` — Input workbook with KenPom, Haslametrics, and Barttorvik data
- `requirements.txt` — Python dependencies

## Environment Variables
- `ODDS_API_KEY` (required) — API key for the-odds-api.com to fetch live market data

## How to Run
- The app runs via `python app.py` on port 5000
- Set `ODDS_API_KEY` secret before using the data features
