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

    # Вибір 5 файлів постачальників
    print("🔎 Вибери файл постачальника 1 (Code, Destination, Price, Increment)")
    file1 = filedialog.askopenfilename()
    df1 = pd.read_excel(file1)

    print("🔎 Вибери файл постачальника 2 (Code, Destination, Price, Increment)")
    file2 = filedialog.askopenfilename()
    df2 = pd.read_excel(file2)

    print("🔎 Вибери файл постачальника 3 (Code, Destination, Price, Increment)")
    file3 = filedialog.askopenfilename()
    df3 = pd.read_excel(file3)

    print("🔎 Вибери файл постачальника 4 (Code, Destination, Price, Increment)")
    file4 = filedialog.askopenfilename()
    df4 = pd.read_excel(file4)

    print("🔎 Вибери файл постачальника 5 (Code, Destination, Price, Increment)")
    file5 = filedialog.askopenfilename()
    df5 = pd.read_excel(file5)

    # Підготовка даних
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

    # Об'єднання всіх
    df_all = pd.merge(df1, df2, on='Code', how='outer')
    df_all = pd.merge(df_all, df3, on='Code', how='outer')
    df_all = pd.merge(df_all, df4, on='Code', how='outer')
    df_all = pd.merge(df_all, df5, on='Code', how='outer')

    # Конвертація цін у числовий тип
    price_cols = ['Pxn Price', 'SkyTel Price', 'SmartNet Price', 'SVM Price', 'RingHD Price']
    for col in price_cols:
        df_all[col] = pd.to_numeric(df_all[col], errors='coerce')

    # 🧠 Найкраща ціна, постачальник, increment
    def find_best(row):
        providers = ['Pxn', 'SkyTel', 'SmartNet', 'SVM', 'RingHD']
        prices = {}
        increments = {}

        for p in providers:
            price = row.get(f'{p} Price')
            increment = row.get(f'{p} Increment')

            if pd.notna(price):
                try:
                    prices[p] = float(price)
                    increments[p] = increment
                except:
                    continue

        if not prices:
            return pd.Series([None, None, None])  # якщо немає цін

        # best_provider = max(prices, key=prices.get)
        best_provider = min(prices, key=prices.get)
        best_price = prices[best_provider]
        best_increment = increments.get(best_provider, None)

        return pd.Series([best_price, best_provider, best_increment])

    # ⬇️ Додаємо нові колонки
    df_all[['Best Price', 'Best Provider', 'Best Increment']] = df_all.apply(find_best, axis=1)

    print("\n🔍 Перші рядки зведеного звіту:")
    print(df_all.head())

    # Збереження
    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        df_all.to_excel(save_path, index=False)
        print(f"✅ Файл успішно збережено: {save_path}")
    else:
        print("❌ Збереження скасовано.")

if __name__ == '__main__':
    main()
