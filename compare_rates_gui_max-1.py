import pandas as pd
import tkinter as tk
from tkinter import filedialog

def prepare_df(df, input_cols, rename_cols):
    df = df[input_cols].drop_duplicates()
    df.columns = rename_cols
    return df

def main():
    root = tk.Tk()
    root.withdraw()

    print("🔎 Вибери файл постачальника 1 (Code, Destination, Price, Increment)")
    df1 = pd.read_excel(filedialog.askopenfilename())

    print("🔎 Вибери файл постачальника 2 (Code, Destination, Price, Increment)")
    df2 = pd.read_excel(filedialog.askopenfilename())

    print("🔎 Вибери файл постачальника 3 (Code, Destination, Price, Increment)")
    df3 = pd.read_excel(filedialog.askopenfilename())

    print("🔎 Вибери файл постачальника 4 (Code, Destination, Price, Increment)")
    df4 = pd.read_excel(filedialog.askopenfilename())

    print("🔎 Вибери файл постачальника 5 (Code, Destination, Price, Increment)")
    df5 = pd.read_excel(filedialog.askopenfilename())

    # Підготовка
    df1 = prepare_df(df1, ['Code', 'Destination', 'Price', 'Increment'],
                     ['Code', 'Destination Pxn', 'Pxn Price', 'Pxn Increment'])
    df2 = prepare_df(df2, ['Code', 'Destination', 'Price', 'Increment'],
                     ['Code', 'Destination SkyTel', 'SkyTel Price', 'SkyTel Increment'])
    df3 = prepare_df(df3, ['Code', 'Destination', 'Price', 'Increment'],
                     ['Code', 'Destination SmartNet', 'SmartNet Price', 'SmartNet Increment'])
    df4 = prepare_df(df4, ['Code', 'Destination', 'Price', 'Increment'],
                     ['Code', 'Destination SVM', 'SVM Price', 'SVM Increment'])
    df5 = prepare_df(df5, ['Code', 'Destination', 'Price', 'Increment'],
                     ['Code', 'Destination RingHD', 'RingHD Price', 'RingHD Increment'])

    # Злиття
    df_all = pd.merge(df1, df2, on='Code', how='outer')
    df_all = pd.merge(df_all, df3, on='Code', how='outer')
    df_all = pd.merge(df_all, df4, on='Code', how='outer')
    df_all = pd.merge(df_all, df5, on='Code', how='outer')

    # Перетворення цін
    price_cols = ['Pxn Price', 'SkyTel Price', 'SmartNet Price', 'SVM Price', 'RingHD Price']
    for col in price_cols:
        df_all[col] = pd.to_numeric(df_all[col], errors='coerce')

    # Вибір максимальної ціни < 1
    def find_best(row):
        providers = ['Pxn', 'SkyTel', 'SmartNet', 'SVM', 'RingHD']
        prices = {}
        increments = {}

        for p in providers:
            price = row.get(f'{p} Price')
            increment = row.get(f'{p} Increment')

            if pd.notna(price):
                try:
                    price = float(price)
                    if price < 1.0:
                        prices[p] = price
                        increments[p] = increment
                except:
                    continue

        if not prices:
            return pd.Series([None, None, None])

        best_provider = max(prices, key=prices.get)
        best_price = prices[best_provider]
        best_increment = increments.get(best_provider, None)

        return pd.Series([best_price, best_provider, best_increment])

    # Застосування
    df_all[['Best Price', 'Best Provider', 'Best Increment']] = df_all.apply(find_best, axis=1)

    # Перегляд
    print("\n🔍 Перші рядки зведеного звіту:")
    print(df_all.head())

    # Збереження
    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        df_all.to_excel(save_path, index=False)
        print(f"✅ Файл збережено: {save_path}")
    else:
        print("❌ Збереження скасовано.")

if __name__ == '__main__':
    main()
