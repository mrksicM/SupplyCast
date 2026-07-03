import pandas as pd
import numpy as np
import math
from prophet import Prophet
from statsmodels.tsa.arima.model import ARIMA
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, mean_absolute_percentage_error
from openpyxl import Workbook

# --- Load storage data ---
storage = pd.read_excel("HistoryGenerator/Export/Storage.xlsx", sheet_name=None)
df = pd.concat(storage.values())
df["TrueDemand"] = df["Consumed"] + df.get("Shortage", 0)

items = df["Item Code"].unique()

# --- Create Excel workbook ---
wb = Workbook()
ws = wb.active
ws.title = "Orders"

# Header row
ws.append([
    "Week","Product Code","Product Name",
    "Predicted Qty","Accuracy %","Possible Error (units)",
    "Current Stock","User Adjustment","Final Order"
])

# Determine next week number
next_week = df["Date"].max().isocalendar().week
next_yearweek = df["Date"].max().strftime("%Y-%U")

for code in items:
    item_df = df[df["Item Code"]==code].copy()
    product_name = item_df["Item Name"].iloc[0] if "Item Name" in item_df.columns else code
    
    # Training series
    train = item_df[["Date","TrueDemand"]].rename(columns={"Date":"ds","TrueDemand":"y"})
    
    # --- Prophet ---
    prophet_model = Prophet(yearly_seasonality=True, weekly_seasonality=True)
    prophet_model.fit(train)
    future = prophet_model.make_future_dataframe(periods=7)  # only next week
    forecast_prophet = prophet_model.predict(future)
    
    y_true = train["y"].values
    y_pred_prophet = forecast_prophet.loc[:len(y_true)-1,"yhat"].values
    mape_prophet = mean_absolute_percentage_error(y_true, y_pred_prophet)
    
    # --- ARIMA ---
    series = train["y"].reset_index(drop=True)
    try:
        arima_model = ARIMA(series, order=(2,1,2))
        arima_fit = arima_model.fit()
        forecast_arima = arima_fit.forecast(steps=7)
        y_pred_arima = arima_fit.fittedvalues
        y_true_arima = series.iloc[y_pred_arima.index]
        mape_arima = mean_absolute_percentage_error(y_true_arima, y_pred_arima)
    except Exception:
        forecast_arima = pd.Series([np.nan]*7)
        mape_arima = np.nan
    
    # --- Random Forest ---
    rf = RandomForestRegressor(n_estimators=100, random_state=42)
    item_df["dow"] = item_df["Date"].dt.dayofweek
    item_df["month"] = item_df["Date"].dt.month
    X = item_df[["dow","month"]]
    y = item_df["TrueDemand"]
    rf.fit(X,y)
    y_pred_rf = rf.predict(X)
    mape_rf = mean_absolute_percentage_error(y, y_pred_rf)
    
    # --- Choose best accuracy ---
    mape_scores = {"Prophet": mape_prophet, "ARIMA": mape_arima, "RandomForest": mape_rf}
    best_model = min(mape_scores, key=mape_scores.get)
    best_mape = mape_scores[best_model]
    best_accuracy = round((1 - best_mape) * 100, 1)
    
    # --- Forecast next week total ---
    dates = pd.to_datetime(future.iloc[len(train):]["ds"].values)
    forecast_df = pd.DataFrame({
        "Date": dates,
        "Prophet": forecast_prophet.iloc[len(train):]["yhat"].values,
        "ARIMA": forecast_arima.values,
        "RF": [rf.predict(pd.DataFrame([[d.dayofweek, d.month]], columns=["dow","month"]))[0] for d in dates]
    })
    forecast_df["Ensemble"] = forecast_df[["Prophet","ARIMA","RF"]].mean(axis=1)
    forecast_df = forecast_df.round(0)
    forecast_df["YearWeek"] = forecast_df["Date"].dt.strftime("%Y-%U")
    
    next_week_orders = forecast_df.groupby("YearWeek")[["Ensemble"]].sum().reset_index()
    next_week_qty = int(next_week_orders.loc[next_week_orders["YearWeek"]==next_yearweek,"Ensemble"].values[0]) if next_yearweek in next_week_orders["YearWeek"].values else 0
    
    # --- Current stock (last End Stock) ---
    current_stock = max(int(item_df.sort_values("Date").iloc[-1]["End Stock"]), 0)
    
    # --- Export row ---
    ws.append([
        next_week,
        code,
        product_name,
        next_week_qty,
        best_accuracy,
        int(next_week_qty * best_mape),
        current_stock,
        "",  # User Adjustment
        ""   # Final Order
    ])

# Save workbook
wb.save("Orders.xlsx")
