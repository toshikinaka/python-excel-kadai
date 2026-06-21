import pandas as pd

df = pd.DataFrame({
    "日付": ["2023-05-17", "2023-05-18", "2023-05-19", "2023-05-20", "2023-05-21"],
    "社員名": ["山田", "佐藤", "鈴木", "田中", "高橋"],
    "売上": [100, 200, 150, 300, 250],
    "部門": ["メーカー", "代理店", "メーカー", "商社", "代理店"]
})

average_sales = df["売上"].mean()

df["平均売上"] = average_sales

def performance(sales):
    if sales >= average_sales + 50:
        return "A"
    elif sales >= average_sales:
        return "B"
    else:
        return "C"

df["業績ランク"] = df["売上"].apply(performance)

df.to_excel("業績.xlsx", index=False)

print("業績.xlsxを作成しました")