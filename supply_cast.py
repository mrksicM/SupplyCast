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

# --- Use true demand (Consumed + Shortage) ---
df["TrueDemand"] = df["Consumed"] + df.get("Shortage", 0)

# --- Get all item codes ---
items = df["Item Code"].unique()

# --- Create Excel workbook ---
wb = Workbook()
wb.remove(wb.active)

for code in items:
    item_df = df[df["Item Code"]==code].copy()
    product_name = item_df["Item Name"].iloc[0] if "Item Name" in item_df.columns else code
    
    # Training series: True demand
    train = item_df[["Date","TrueDemand"]].rename(columns={"Date":"ds","TrueDemand":"y"})
    
    # --- Prophet ---
    prophet_model = Prophet(yearly_seasonality=True, weekly_seasonality=True)
    prophet_model.fit(train)
    future = prophet_model.make_future_dataframe(periods=30)
    forecast_prophet = prophet_model.predict(future)
    
    # Accuracy
    y_true = train["y"].values
    y_pred_prophet = forecast_prophet.loc[:len(y_true)-1,"yhat"].values
    rmse_prophet = math.sqrt(mean_squared_error(y_true, y_pred_prophet))
    mape_prophet = mean_absolute_percentage_error(y_true, y_pred_prophet)
    
    # --- ARIMA ---
    series = train["y"].reset_index(drop=True)
    try:
        arima_model = ARIMA(series, order=(2,1,2))
        arima_fit = arima_model.fit()
        forecast_arima = arima_fit.forecast(steps=30)
        y_pred_arima = arima_fit.fittedvalues
        y_true_arima = series.iloc[y_pred_arima.index]
        rmse_arima = math.sqrt(mean_squared_error(y_true_arima, y_pred_arima))
        mape_arima = mean_absolute_percentage_error(y_true_arima, y_pred_arima)
    except Exception:
        forecast_arima = pd.Series([np.nan]*30)
        rmse_arima, mape_arima = np.nan, np.nan
    
    # --- Random Forest ---
    rf = RandomForestRegressor(n_estimators=100, random_state=42)
    item_df["dow"] = item_df["Date"].dt.dayofweek
    item_df["month"] = item_df["Date"].dt.month
    X = item_df[["dow","month"]]
    y = item_df["TrueDemand"]
    rf.fit(X,y)
    y_pred_rf = rf.predict(X)
    rmse_rf = math.sqrt(mean_squared_error(y, y_pred_rf))
    mape_rf = mean_absolute_percentage_error(y, y_pred_rf)
    
    # --- Export to Excel ---
    ws = wb.create_sheet(title=code)
    ws.append(["Model","RMSE","MAPE"])
    ws.append(["Prophet", round(rmse_prophet,2), round(mape_prophet,4)])
    ws.append(["ARIMA", round(rmse_arima,2) if not np.isnan(rmse_arima) else "N/A",
               round(mape_arima,4) if not np.isnan(mape_arima) else "N/A"])
    ws.append(["RandomForest", round(rmse_rf,2), round(mape_rf,4)])
    
    ws.append([])
    ws.append(["Week","Prophet Forecast","ARIMA Forecast","RF Forecast","Ensemble"])
    
    # --- Weekly aggregation ---
    dates = pd.to_datetime(future.iloc[len(train):]["ds"].values)
    forecast_df = pd.DataFrame({
        "Date": dates,
        "Prophet": forecast_prophet.iloc[len(train):]["yhat"].values,
        "ARIMA": forecast_arima.values,
        "RF": [rf.predict(pd.DataFrame([[d.dayofweek, d.month]], columns=["dow","month"]))[0] for d in dates]
    })
    forecast_df["Ensemble"] = forecast_df[["Prophet","ARIMA","RF"]].mean(axis=1)
    forecast_df = forecast_df.round(0)
    forecast_df["Week"] = forecast_df["Date"].dt.isocalendar().week
    
    weekly = forecast_df.groupby("Week")[["Prophet","ARIMA","RF","Ensemble"]].sum().reset_index()
    
    for _, row in weekly.iterrows():
        ws.append([
            int(row["Week"]),
            int(row["Prophet"]),
            int(row["ARIMA"]) if not np.isnan(row["ARIMA"]) else None,
            int(row["RF"]),
            int(row["Ensemble"])
        ])

# Save workbook
wb.save("Forecasts_weekly.xlsx")
