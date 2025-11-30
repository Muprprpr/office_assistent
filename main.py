import pandas as pd
import matplotlib.pyplot as plt
import os
import argparse
import sys
import re
from tqdm import tqdm

def clean_series(series, remove_zero=True):
    series = pd.to_numeric(series, errors='coerce')
    if remove_zero:
        series = series[series != 0]
    return series.dropna()

def parse_clean_args(clean_args):
    if not clean_args:
        return set()
    args_str = " ".join(clean_args).lower()
    if 'both' in args_str:
        return {'x', 'y'}
    tokens = re.split(r'\s+', args_str.strip())
    result = set()
    for t in tokens:
        if t == 'x':
            result.add('x')
        elif t == 'y':
            result.add('y')
    return result

def group_bar_data(x_series, y_series, group_size):
    """
    将 (x, y) 按每 group_size 行分组，返回聚合后的 x_labels 和 y_means
    """
    if group_size <= 1:
        return x_series, y_series

    df = pd.DataFrame({'x': x_series, 'y': y_series}).reset_index(drop=True)

    # 如果 x 是数值型，先排序（更合理）
    try:
        df['x_num'] = pd.to_numeric(df['x'], errors='coerce')
        if df['x_num'].notna().all():
            df = df.sort_values('x_num').reset_index(drop=True)
            numeric_x = True
        else:
            numeric_x = False
    except:
        numeric_x = False

    # 分组
    grouped = []
    n = len(df)
    num_groups = (n + group_size - 1) // group_size  # 向上取整

    with tqdm(total=num_groups, desc="分组聚合", bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}]") as pbar:
        for i in range(0, n, group_size):
            group = df.iloc[i:i+group_size]
            y_mean = group['y'].mean()
            if numeric_x:
                x_vals = pd.to_numeric(group['x'], errors='coerce')
                x_min = x_vals.min()
                x_max = x_vals.max()
                if x_min == x_max:
                    x_label = str(x_min)
                else:
                    x_label = f"{x_min:.2f}~{x_max:.2f}"
            else:
                # 类别型：取第一个或拼接
                if len(group) == 1:
                    x_label = str(group['x'].iloc[0])
                else:
                    x_label = f"{group['x'].iloc[0]}...{group['x'].iloc[-1]}"
            grouped.append((x_label, y_mean))
            pbar.update(1)

    x_labels, y_means = zip(*grouped)
    return list(x_labels), list(y_means)

def reject_outliers_by_percentile(x_series, y_series, ratio=1.0):
    if ratio <= 0:
        return x_series, y_series
    df_temp = pd.DataFrame({'x': x_series, 'y': y_series}).dropna()
    q_low = ratio / 100.0
    q_high = 1 - q_low

    with tqdm(total=100, desc="剔除极值", leave=False) as pbar:
        x_low = df_temp['x'].quantile(q_low)
        x_high = df_temp['x'].quantile(q_high)
        y_low = df_temp['y'].quantile(q_low)
        y_high = df_temp['y'].quantile(q_high)
        pbar.update(50)

        mask = (
            (df_temp['x'] >= x_low) & (df_temp['x'] <= x_high) &
            (df_temp['y'] >= y_low) & (df_temp['y'] <= y_high)
        )
        filtered = df_temp[mask]
        pbar.update(50)
    return filtered['x'], filtered['y']

def main():
    parser = argparse.ArgumentParser(description="从Excel绘制图表，支持分组柱状图与进度反馈。")
    parser.add_argument("--path", "-p", required=True, help="Excel文件路径")
    parser.add_argument("--x", required=True, help="横轴列名")
    parser.add_argument("--y", required=True, help="纵轴列名")
    parser.add_argument("--plot-type", choices=["scatter", "bar"], default="scatter")
    parser.add_argument("--clean", nargs="*", help="清洗列: x, y, both")
    parser.add_argument("--reject", choices=["y", "n"], default="n")
    parser.add_argument("--reject-ratio", type=float, default=1.0)
    parser.add_argument("--group-size", type=int, help="柱状图分组大小（每N行合并为1柱）")

    args = parser.parse_args()

    excel_path = args.path.strip('"').strip("'")
    if not os.path.isfile(excel_path):
        print(f"❌ 文件不存在: {excel_path}", file=sys.stderr)
        sys.exit(1)

    # ============ 阶段 1: 加载数据 ============
    print("⏳ 正在加载Excel文件...")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"❌ 读取失败: {e}", file=sys.stderr)
        sys.exit(1)
    print(f"✅ 成功加载 {len(df)} 行数据")

    x_col, y_col = args.x, args.y
    if x_col not in df.columns or y_col not in df.columns:
        print(f"❌ 列缺失。可用: {list(df.columns)}", file=sys.stderr)
        sys.exit(1)

    clean_set = parse_clean_args(args.clean)
    x_data = df[x_col].copy()
    y_data = df[y_col].copy()

    # ============ 阶段 2: 数据清洗 ============
    with tqdm(total=100, desc="数据清洗", leave=False) as pbar:
        valid_idx = df.index
        pbar.update(20)
        if 'x' in clean_set:
            cleaned_x = clean_series(x_data, remove_zero=True)
            valid_idx = valid_idx.intersection(cleaned_x.index)
            pbar.update(30)
        if 'y' in clean_set:
            cleaned_y = clean_series(y_data, remove_zero=True)
            valid_idx = valid_idx.intersection(cleaned_y.index)
            pbar.update(30)
        x_data = x_data.loc[valid_idx]
        y_data = y_data.loc[valid_idx]
        pbar.update(20)

    # ============ 阶段 3: 类型转换 ============
    if args.plot_type == "scatter":
        x_num = pd.to_numeric(x_data, errors='coerce')
        y_num = pd.to_numeric(y_data, errors='coerce')
        valid = x_num.notna() & y_num.notna()
        x_final = x_num[valid]
        y_final = y_num[valid]
    else:  # bar
        y_num = pd.to_numeric(y_data, errors='coerce')
        valid = y_num.notna()
        x_final = x_data.loc[valid]  # 保留原始类型（可能为str或num）
        y_final = y_num[valid]

    if len(x_final) == 0:
        print("❌ 无有效数据", file=sys.stderr)
        sys.exit(1)

    # ============ 阶段 4: 剔除极值 ============
    if args.reject == 'y':
        if args.plot_type == "scatter":
            x_final, y_final = reject_outliers_by_percentile(x_final, y_final, args.reject_ratio)
        else:
            df_bar = pd.DataFrame({'x': x_final, 'y': y_final}).dropna()
            q_low = args.reject_ratio / 100.0
            q_high = 1 - q_low
            y_low = df_bar['y'].quantile(q_low)
            y_high = df_bar['y'].quantile(q_high)
            mask = (df_bar['y'] >= y_low) & (df_bar['y'] <= y_high)
            df_bar = df_bar[mask]
            x_final = df_bar['x']
            y_final = df_bar['y']

    if len(x_final) == 0:
        print("❌ 剔除后无数据", file=sys.stderr)
        sys.exit(1)

    # ============ 阶段 5: 分组聚合（仅 bar 图） ============
    if args.plot_type == "bar" and args.group_size is not None and args.group_size > 1:
        if len(x_final) <= args.group_size:
            print(f"⚠️ 数据量 ({len(x_final)}) ≤ group-size ({args.group_size})，跳过分组。")
        else:
            x_final, y_final = group_bar_data(x_final, y_final, args.group_size)
            print(f"📊 分组完成：{len(x_final)} 个柱子（原 {len(y_final)*args.group_size} 行）")

    # ============ 阶段 6: 绘图 ============
    print("🎨 正在绘图...")
    plt.figure(figsize=(min(12, max(8, len(x_final) * 0.3)), 6))  # 自适应宽度
    if args.plot_type == "scatter":
        plt.scatter(x_final, y_final, alpha=0.6, edgecolors='w', linewidth=0.5)
        plt.xlabel(x_col)
        plt.ylabel(y_col)
        plt.title(f'Scatter Plot: {x_col} vs {y_col}')
    else:
        plt.bar(range(len(x_final)), y_final, color='skyblue', edgecolor='black')
        plt.xlabel(x_col)
        plt.ylabel(y_col)
        plt.title(f'Bar Chart (Grouped): {x_col} vs {y_col}')
        plt.xticks(range(len(x_final)), x_final, rotation=45, ha='right')

    plt.grid(True, linestyle='--', alpha=0.5)
    plt.tight_layout()

    # 保存
    output_dir = os.path.dirname(excel_path)
    base_name = f"{x_col}-{y_col}-{args.plot_type}"
    if args.group_size:
        base_name += f"-group{args.group_size}"
    safe_name = "".join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in base_name) + ".png"
    output_path = os.path.join(output_dir, safe_name)
    plt.savefig(output_path, dpi=150)
    plt.close()

    print(f"✅ 完成！图像已保存至:\n{output_path}")
    print(f"📊 最终数据点数: {len(x_final)}")

if __name__ == "__main__":
    main()