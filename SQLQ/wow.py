import pandas as pd

# ====== allocate 函數（你原本那個）======
def allocate(df, col_name):
    data = df[['月份', col_name]].copy()

    demand = []
    supply = []

    for _, row in data.iterrows():
        val = row[col_name]

        if pd.isna(val):
            continue

        val = float(val)

        if val > 0:
            demand.append([row['月份'], val])
        elif val < 0:
            supply.append([row['月份'], -val])

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

    df.columns = df.columns.str.strip()

    if '列標號' in df.columns:
        df.rename(columns={'列標號': '月份'}, inplace=True)

    columns = [c for c in df.columns if c.startswith('_')]

    for c in columns:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    all_results = []

    for col in columns:
        print(f'處理中: {col}')

        result = allocate(df, col)
        result['欄位'] = col

        all_results.append(result)

    final_df = pd.concat(all_results, ignore_index=True)

    summary_df = (
        final_df
        .groupby('欄位', as_index=False)['分配量']
        .sum()
        .rename(columns={'分配量': '總分配量'})
    )

    print(summary_df)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, sheet_name='明細', index=False)
        summary_df.to_excel(writer, sheet_name='總結', index=False)

    print('完成！輸出檔案：', output_file)


if __name__ == '__main__':
    main()
