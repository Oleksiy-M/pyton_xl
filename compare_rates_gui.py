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
    file1 = filedialog.askopenfilename()
    df1 = pd.read_excel(file1)
    print(f"Available columns: {list(df1.columns)}")

    print("🔎 Вибери файл постачальника 2 (Code, Destination, Price, Increment)")
    file2 = filedialog.askopenfilename()
    df2 = pd.read_excel(file2)
    print(f"Available columns: {list(df2.columns)}")

    print("🔎 Вибери файл постачальника 3 (Code, Destination, Price, Increment)")
    file3 = filedialog.askopenfilename()
    df3 = pd.read_excel(file3)
    print(f"Available columns: {list(df3.columns)}")

    print("🔎 Вибери файл постачальника 4 (Code, Destination, Price, Increment)")
    file4 = filedialog.askopenfilename()
    df4 = pd.read_excel(file4)
    print(f"Available columns: {list(df4.columns)}")

    print("🔎 Вибери файл постачальника 5 (Code, Destination, Price, Increment)")
    file5 = filedialog.askopenfilename()
    df5 = pd.read_excel(file5)
    print(f"Available columns: {list(df5.columns)}")

    df1 = prepare_df(
        df1,
        input_cols=['Code', 'Destination', 'Price', 'Increment'],
        rename_cols=['Code', 'Destination Pxn', 'Pxn Price', 'Pxn Increment']
    )
    df2 = prepare_df(
        df2,
        input_cols=['Code', 'Destination', 'Price', 'Increment'],
        rename_cols=['Code', 'Destination SkyTel', 'SkyTel Price', 'SkyTel Increment']
    )
    df3 = prepare_df(
        df3,
        input_cols=['Code', 'Destination', 'Price', 'Increment'],
        rename_cols=['Code', 'Destination SmartNet', 'SmartNet Price', 'SmartNet Increment']
    )
    df4 = prepare_df(
        df4,
        input_cols=['Code', 'Destination', 'Price', 'Increment'],
        rename_cols=['Code', 'Destination Forex', 'Forex Price', 'Forex Increment']
    )
    df5 = prepare_df(
        df5,
        input_cols=['Code', 'Destination', 'Price', 'Increment'],
        rename_cols=['Code', 'Destination RingHD', 'RingHD Price', 'RingHD Increment']
    )

    df_all = pd.merge(df1, df2, on='Code', how='outer')
    df_all = pd.merge(df_all, df3, on='Code', how='outer')
    df_all = pd.merge(df_all, df4, on='Code', how='outer')
    df_all = pd.merge(df_all, df5, on='Code', how='outer')

    for col in ['Pxn Price', 'SkyTel Price', 'SmartNet Price', 'Forex Price', 'RingHD Price']:
        df_all[col] = pd.to_numeric(df_all[col], errors='coerce')

    def find_best(row):
        prices = {
            'Pxn': row.get('Pxn Price', float('inf')),
            'SkyTel': row.get('SkyTel Price', float('inf')),
            'SmartNet': row.get('SmartNet Price', float('inf')),
            'Forex': row.get('Forex Price', float('inf')),
            'RingHD': row.get('RingHD Price', float('inf')),
        }
        prices = {k: (v if pd.notna(v) else float('inf')) for k, v in prices.items()}
        best_provider = min(prices, key=prices.get)
        best_price = prices[best_provider]
        if best_price == float('inf'):
            return pd.Series([None, None])
        return pd.Series([best_price, best_provider])

    df_all[['Best Price', 'Best Provider']] = df_all.apply(find_best, axis=1)

    print("\nОб'єднаний датафрейм:")
    print(df_all.head())

    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        df_all.to_excel(save_path, index=False)
        print(f"Файл успішно збережено: {save_path}")
    else:
        print("Файл не збережено.")

if __name__ == '__main__':
    main()
