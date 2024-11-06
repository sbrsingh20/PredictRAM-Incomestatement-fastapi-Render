import pandas as pd
import os
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware

# Load the data from Excel files
inflation_data = pd.read_excel('Inflation_event_stock_analysis_resultsOct.xlsx')
income_data = pd.read_excel('Inflation_IncomeStatement_correlation_results.xlsx')
interest_rate_data = pd.read_excel('interestrate_event_stock_analysis_resultsOct.xlsx')
interest_rate_income_data = pd.read_excel('interestrate_IncomeStatement_correlation_results.xlsx')

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # or "*" to allow all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allow all HTTP methods
    allow_headers=["*"],  # Allow all headers
)

# Root endpoint
@app.get("/")
async def read_root():
    return {"message": "Welcome to the Stock Analysis API! Use the /stock-details endpoint."}

class StockRequest(BaseModel):
    stock_symbol: str
    event_type: str
    expected_rate: float
    method: str

@app.post("/stock-details/")
async def get_stock_details(request: StockRequest):
    stock_symbol = request.stock_symbol
    event_type = request.event_type
    expected_rate = request.expected_rate
    method = request.method

    if event_type == 'Inflation':
        event_row = inflation_data[inflation_data['Symbol'] == stock_symbol]
        income_row = income_data[income_data['Stock Name'] == stock_symbol]
    else:  # Interest Rate
        event_row = interest_rate_data[interest_rate_data['Symbol'] == stock_symbol]
        income_row = interest_rate_income_data[interest_rate_income_data['Stock Name'] == stock_symbol]

    if event_row.empty or income_row.empty:
        raise HTTPException(status_code=404, detail="Stock symbol not found")

    event_details = event_row.iloc[0]
    income_details = income_row.iloc[0]

    projections = generate_projections(event_details, income_details, expected_rate, event_type, method)
    interpretation = interpret_data(event_details, income_details, event_type)

    return {
        "stock_symbol": stock_symbol,
        "event_type": event_type,
        "projections": projections,
        "interpretation": interpretation
    }


def generate_projections(event_details, income_details, expected_rate, event_type, method):
    latest_event_value = pd.to_numeric(income_details.get('Latest Event Value', 0), errors='coerce')
    projections = []

    if 'Latest Close Price' in event_details.index:
        latest_close_price = pd.to_numeric(event_details['Latest Close Price'], errors='coerce')

        if method == 'Dynamic':
            rate_change = expected_rate - latest_event_value
            price_change = event_details['Event Coefficient'] * rate_change
            projected_price = latest_close_price + price_change
            change = price_change
            explanation = "Dynamic calculation considers the event coefficient and rate change."
        else:  # Simple
            price_change = latest_close_price * (expected_rate / 100)
            projected_price = latest_close_price + price_change
            change = expected_rate
            explanation = "Simple calculation uses the expected rate directly."

        projections.append({
            'Parameter': 'Projected Stock Price',
            'Current Value': latest_close_price,
            'Projected Value': projected_price,
            'Change': change
        })
    else:
        raise HTTPException(status_code=404, detail="Stock Price data not available in event details.")

    for column in income_details.index:
        if column != 'Stock Name':
            current_value = pd.to_numeric(income_details[column], errors='coerce')
            if pd.notna(current_value):
                if method == 'Dynamic':
                    if column in event_details.index:
                        correlation_factor = event_details[column] if column in event_details.index else 0
                        projected_value = current_value + (current_value * correlation_factor * (expected_rate - latest_event_value) / 100)
                    else:
                        projected_value = current_value * (1 + (expected_rate - latest_event_value) / 100)
                    change = projected_value - current_value
                else:  # Simple
                    projected_value = current_value * (1 + expected_rate / 100)
                    change = expected_rate

                projections.append({
                    'Parameter': column,
                    'Current Value': current_value,
                    'Projected Value': projected_value,
                    'Change': change
                })

    new_columns = [
        'June 2024 Total Revenue/Income',
        'June 2024 Total Operating Expense',
        'June 2024 Operating Income/Profit',
        'June 2024 EBITDA',
        'June 2024 EBIT',
        'June 2024 Income/Profit Before Tax',
        'June 2024 Net Income From Continuing Operation',
        'June 2024 Net Income',
        'June 2024 Net Income Applicable to Common Share',
        'June 2024 EPS (Earning Per Share)'
    ]

    for col in new_columns:
        if col in income_details.index:
            current_value = pd.to_numeric(income_details[col], errors='coerce')
            if pd.notna(current_value):
                if method == 'Dynamic':
                    projected_value = current_value * (1 + (expected_rate - latest_event_value) / 100)
                else:  # Simple
                    projected_value = current_value * (1 + expected_rate / 100)

                projections.append({
                    'Parameter': col,
                    'Current Value': current_value,
                    'Projected Value': projected_value,
                    'Change': expected_rate
                })

    return projections

def interpret_data(event_details, income_details, event_type):
    interpretation = {}
    
    if event_type == 'Inflation':
        if 'Event Coefficient' in event_details.index:
            if event_details['Event Coefficient'] < -1:
                interpretation['Inflation'] = "Stock price decreases significantly. Increase portfolio risk."
            elif event_details['Event Coefficient'] > 1:
                interpretation['Inflation'] = "Stock price increases, benefiting from inflation."
    else:  # Interest Rate
        if 'Event Coefficient' in event_details.index:
            if event_details['Event Coefficient'] < -1:
                interpretation['Interest Rate'] = "Stock price decreases significantly. Increase portfolio risk."
            elif event_details['Event Coefficient'] > 1:
                interpretation['Interest Rate'] = "Stock price increases, benefiting from interest hikes."
    
    if 'Average Operating Margin' in income_details.index:
        average_operating_margin = income_details['Average Operating Margin']
        if average_operating_margin > 0.2:
            interpretation['Income Statement'] = "High Operating Margin: Indicates strong management effectiveness."
        elif average_operating_margin < 0.1:
            interpretation['Income Statement'] = "Low Operating Margin: Reflects risk in profitability."

    return interpretation

