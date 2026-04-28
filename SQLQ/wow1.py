import pandas as pd

# ====== 核心分配邏輯（FIFO）======
def allocate(df, col_name):
    data = df[['月份', col_name]].copy()

    demand = []
    supply = []

    for _, row in data.iterrows():
        val = row[col_name]

        if pd.isna(val):
            continue

        val = float(val)  # ✅ 強制轉數字（關鍵）

        if val > 0:
            demand.append([row['月份'], val])
        elif val < 0:
            supply.append([row['月份'], -val])  # 轉正

    result = []

    i = 0
    j = 0

    while i < len(demand) and j < len(supply):
        d_month, d_val = demand[i]
        s_month, s_val = supply[j]

        alloc = min(d_val, s_val)

        result.append({
            '需求月份': d_month,
            '供給月份': s_month,
            '分配量': alloc
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

    # ✅ 清理欄位名稱
    df.columns = df.columns.str.strip()

    # ✅ 改欄位名稱
    if '列標號' in df.columns:
        df.rename(columns={'列標號': '月份'}, inplace=True)

    # ✅ 🔥 重點：所有 _ 開頭欄位轉數字（避免 string bug）
    columns = [c for c in df.columns if c.startswith('_')]

    for c in columns:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    print("欄位名稱：", df.columns.tolist())

    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    all_results = []

    for col in columns:
        print(f'處理中: {col}')

        result = allocate(df, col)
        result['欄位'] = col

        result.to_excel(writer, sheet_name=col[:31], index=False)  # Excel限制31字

        all_results.append(result)

    final_df = pd.concat(all_results, ignore_index=True)
    final_df.to_excel(writer, sheet_name='總表', index=False)

    writer.close()

    print('完成！輸出檔案：', output_file)


# ====== 執行 ======
if __name__ == '__main__':
    main()
