import pandas as pd
import openpyxl, pprint, os, sys
from datetime import datetime, timedelta
from openpyxl.styles import Alignment
from dateutil.relativedelta import relativedelta

def get_target_order(products, orders):
    # 対象の注文IDを対話形式で入力
    target_id = input("注文IDを入力して下さい")

    # 情報を元に見積書を作成
    # 対象の取引データを抽出
    target_orders = orders[orders["order_id"] == target_id]

    # データを結合
    merged = target_orders.merge(products, on="product_id", how="left")

    # 金額列の小計
    merged["subtotal"]  = merged["unit_price"] * merged["quantity"]

    # 消費税
    merged["VAT"] = merged["subtotal"] * merged["tax_rate"] / 100
    merged["VAT"] = merged["VAT"].round(0).astype(int)

    # 合計
    merged["total_amount"] = merged["subtotal"] + merged["VAT"]
    merged["total_amount"] = merged["total_amount"].astype(int)

    return merged, target_id
