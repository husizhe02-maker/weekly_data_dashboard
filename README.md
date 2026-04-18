# 周报数据可视化看板

一个基于 Streamlit 的本地可视化看板，用于上传包含 **视频维度数据** 和 **账号维度数据** 两张工作表的 Excel 文件，并自动生成周报分析页面。

当前版本基于最新版 `app.py` 生成，代码使用了 `streamlit`、`pandas`、`numpy`、`plotly.express`、`plotly.graph_objects` 和 `plotly.subplots.make_subplots`，并通过 `pd.read_excel` 读取 Excel 工作簿。fileciteturn1file0

## 功能概览

应用包含 3 个标签页：

### Tab1 整体概览
- 周次选择器
- KPI 卡片：总播放量、总新增粉丝、总点赞数、活跃账号数
- 最近最多 4 周的播放量与新增粉丝双轴趋势图
- 当前周账号维度 / 视频维度底层明细展开查看

相关逻辑见 `render_tab_overview()` 与 `main()`。fileciteturn1file1 fileciteturn1file2

### Tab2 账号维度
- 账号汇总表：周播放量、日播、发帖量、条均及各自环比
- 当前周账号横截面表现：
  - 账号播放量占比饼图
  - 播放量 vs 新增粉丝散点图
  - 播放量 vs 条均散点图
- 最多最近 5 周的多账号播放量趋势图

相关逻辑见 `render_tab_accounts()`。fileciteturn1file3 fileciteturn1file4

### Tab3 视频维度
- 周次多选 + 账号多选筛选器
- Top10 视频横向柱状图
- 互动率气泡图
- 完播率 vs 播放量散点图
- 话题标签聚合表格
- 各时长区间条均播放量柱状图
- 解析出标签的视频明细

相关逻辑见 `render_tab_videos()`。fileciteturn1file4

## 环境要求

建议：
- Python 3.10 或 3.11
- pip 最新版本

## 安装依赖

在项目目录运行：

```bash
pip install -r requirements.txt
```

## 运行方式

确保目录下存在你的 `app.py` 后，运行：

```bash
streamlit run app.py
```

启动后，在浏览器中打开本地地址，即可上传 Excel 文件使用看板。

## Excel 文件要求

上传的 Excel 文件必须包含以下两个工作表：
- `视频维度数据`
- `账号维度数据`

代码会在加载时校验工作表名称和关键字段；如果缺失，会在页面报错提示。fileciteturn1file0

## 主要依赖说明

`requirements.txt` 中保留了当前代码真正需要的核心依赖：

- `streamlit`：构建交互式 Web 看板
- `pandas`：数据读取、清洗、聚合
- `numpy`：数值计算与条件逻辑
- `plotly`：绘制饼图、折线图、柱状图、散点图、双轴图
- `openpyxl`：支持 `pandas.read_excel()` 读取 `.xlsx` 文件

## 推荐目录结构

```text
project/
├── app.py
├── requirements.txt
└── README.md
```

如果你的数据样例也想放在项目里，可以扩展为：

```text
project/
├── app.py
├── requirements.txt
├── README.md
└── data/
    └── 周报数据_脱敏版.xlsx
```

## 常见问题

### 1. 依赖安装失败
可以先升级 pip：

```bash
python -m pip install --upgrade pip
```

然后重新安装依赖。

### 2. Excel 上传后提示缺少字段
请检查：
- 工作表名称是否完全一致
- 列名是否与原始模板一致
- 是否有隐藏空格或被手动改名

### 3. 趋势图或表格为空
通常是因为：
- 当前周次没有数据
- 多选筛选器没有选中任何账号
- Excel 中日期字段格式异常，导致周次未被正确解析

## 后续可扩展方向

- 导出 PNG / PDF 周报截图
- 导出汇总 Excel
- 增加 KPI 卡片同比/环比说明
- 增加异常检测和预警提示
- 增加更细粒度的视频标签分析
