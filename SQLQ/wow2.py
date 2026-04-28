import pandas as pd

# ====== 核心分配邏輯（FIFO + 追蹤版）======
def allocate(df, col_name):
    data = df[['月份', col_name]].copy()

    demand = []
    supply = []

    # ===== 分類需求 / 供給 =====
    for _, row in data.iterrows():
        val = row[col_name]

        if pd.isna(val):
            continue

        val = float(val)

        if val > 0:
            demand.append([row['月份'], val])
        elif val < 0:
            supply.append([row['月份'], -val])  # 轉正

    result = []

    i = 0
    j = 0

    # ===== FIFO 配對 =====
    while i < len(demand) and j < len(supply):
        d_month, d_val = demand[i]
        s_month, s_val = supply[j]

        alloc = min(d_val, s_val)

        result.append({
            '需求月份': d_month,
            '供給月份': s_month,
            '分配量': alloc,
            '欄位': col_name
        })

        demand[i][1] -= alloc
        supply[j][1] -= alloc

        if demand[i][1] == 0:
            i += 1
        if supply[j][1] == 0:
            j += 1

    return pd.DataFrame(result)


# ====== 主程式 ======
def main():
    input_file = 'input.xlsx'
    output_file = 'output.xlsx'

    df = pd.read_excel(input_file, skiprows=0)

    # ===== 清理欄位 =====
    df.columns = df.columns.str.strip()

    if '列標號' in df.columns:
        df.rename(columns={'列標號': '月份'}, inplace=True)

    # ===== 找所有數值欄位 =====
    columns = [c for c in df.columns if c.startswith('_')]

    # ===== 強制轉數字 =====
    for c in columns:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    print("欄位名稱：", df.columns.tolist())

    all_results = []

    # ===== 合併所有欄位結果（重點改這裡）=====
    for col in columns:
        print(f'處理中: {col}')

        result = allocate(df, col)
        result['欄位'] = col

        all_results.append(result)

    # ===== 合併成一張表 =====
    final_df = pd.concat(all_results, ignore_index=True)

    # ===== 輸出單一 sheet =====
    final_df.to_excel(output_file, index=False)

    print('完成！輸出檔案：', output_file)


# ====== 執行 ======
if __name__ == '__main__':
    main()
