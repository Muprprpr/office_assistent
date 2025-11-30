📊 Excel 图表生成工具
一款支持散点图/柱状图绘制、数据清洗、异常值剔除、分组聚合的命令行工具，专为处理含大量零值、空值或冗余数据的 Excel 文件设计。内置进度反馈，适合科研、工程与业务数据分析场景。

✨ 核心功能

功能 说明
------ ------
双图表类型 散点图（scatter）用于数值-数值关系；柱状图（bar）用于类别-数值展示
灵活数据清洗 可选择清洗 x 列、y 列、两者或不清洗（自动剔除 0 和空值）
异常值剔除 按百分比自动剔除最大/最小值（如各剔除 1%）
柱状图分组 对海量数据按行数分组聚合（如每 100 行合并为 1 个柱子），提升可读性
进度可视化 使用 tqdm 显示加载、清洗、分组等阶段进度，避免“假死”感
路径兼容性 自动处理带引号的文件路径（Windows 拖拽友好）
一键打包 支持 PyInstaller 打包为独立 .exe，无需 Python 环境

📦 安装依赖

bash
pip install pandas matplotlib openpyxl tqdm
💡 说明：
openpyxl 用于读取 .xlsx 文件（若需支持 .xls，额外安装 xlrd）
tqdm 提供进度条
所有依赖均可被 PyInstaller 自动打包

🚀 快速使用
基础命令结构

bash
python plot_excel.py --path <文件路径> --x <横轴列名> --y <纵轴列名> [其他选项]
示例 1：绘制散点图（清洗 x/y，剔除极值）

bash
python plot_excel.py \
--path "C:\data.xlsx" \
--x Temperature \
--y Pressure \
--plot-type scatter \
--clean both \
--reject y \
--reject-ratio 1.0
示例 2：绘制分组柱状图（每 500 行合并）

bash
python plot_excel.py \
--path sales.xlsx \
--x Date \
--y Revenue \
--plot-type bar \
--clean y \
--group-size 500
示例 3：不清洗，直接绘图

bash
python plot_excel.py --path data.xlsx --x Category --y Count --plot-type bar

⚙️ 参数详解

参数 简写 必填 默认值 说明
------ ------ ------ -------- ------
--path -p ✅ 是 — Excel 文件路径（支持 "C:\file.xlsx" 引号包裹）
--x — ✅ 是 — 横轴列名（区分大小写）
--y — ✅ 是 — 纵轴列名
--plot-type — 否 scatter 图表类型：scatter（散点图）或 bar（柱状图）
--clean — 否 不清洗 清洗哪些列：<br>• x → 仅清洗横轴<br>• y → 仅清洗纵轴<br>• both 或 x y → 清洗两者
--reject — 否 n 是否剔除最大/最小值：y / n
--reject-ratio — 否 1.0 每侧剔除的百分比（如 1.0 = 各剔除 1%，共 2%）
--group-size — 否 — 仅柱状图有效：每 N 行合并为 1 个柱子（y 取平均）
🔔 注意：
--clean 接受多种写法：--clean x、--clean x y、--clean both
--group-size 仅在 --plot-type bar 时生效
若 --reject y 但未指定 --reject-ratio，默认使用 1.0

📁 输出结果
图像保存位置：与输入 Excel 文件同目录
文件命名规则：

{x列名}-{y列名}-{图表类型}[-group{N}].png

示例：
Temperature-Pressure-scatter.png
Date-Revenue-bar-group500.png

🧪 错误处理与提示

场景 系统行为
------ --------
文件不存在 报错并退出
列名不存在 列出所有可用列名供参考
清洗后无数据 提示“无有效数据”并退出
分组大小 ≥ 数据量 跳过分组并警告
路径含引号 自动去除首尾 " 或 '

📦 打包为独立可执行文件（.exe）

bash
pip install pyinstaller
pyinstaller --onefile plot_excel.py

生成文件：dist/plot_excel.exe

用户使用方式（无需 Python）：

cmd
plot_excel.exe --path "C:\data.xlsx" --x A --y B --plot-type bar --group-size 100

📝 使用建议
大数据集：优先使用 --group-size 避免柱状图过于密集
时间序列：若 x 为时间戳，建议先在 Excel 中排序，再分组
异常值敏感场景：设置 --reject y --reject-ratio 0.5 剔除极端噪声
调试模式：先用小样本测试参数组合，再处理全量数据

🛠️ 开发与扩展
聚合函数扩展：当前分组使用 mean()，可修改 group_bar_data 函数支持 sum/median
GUI 版本：可基于此脚本封装 tkinter 界面
多图批量生成：循环调用本脚本或扩展为支持多列组合

📜 许可证

MIT License — 免费用于个人及商业项目。

💬 作者：Muprprpr
📫 联系邮箱：muprprpr@163.com
📅 最后更新：2025年11月30日
✉️ 反馈建议：欢迎提交 Issue 或 PR！

✅ 现在，只需一条命令，即可从混乱的 Excel 数据中生成清晰、专业的图表！