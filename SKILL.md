---
name: "可转债安全性评估+生成日期"
description: "从理杏仁导出可转债数据，按三个安全性指标计算评级，生成格式化Excel并上传至腾讯文档。适用于可转债安全性评估的定期更新场景。"
agent_created: true
---

# 可转债安全性评估 + 生成日期

## 触发条件

当用户要求「更新可转债安全性评估」、「做可转债安全性表格」、「跑可转债数据」或类似表述时触发本 Skill。

## 前置条件

1. 用户需提供**理杏仁登录凭据**（账号 + 密码）
2. 用户需**连接腾讯文档 MCP 连接器**（左侧边栏 → 连接管理 → 腾讯文档 → 连接授权）

## 完整工作流

### 步骤 1: 浏览器自动化导出数据

使用 Playwright（已安装于工作目录 venv）登录理杏仁并导出 CSV：

```bash
# 浏览器环境已就绪，Playwright 安装在 venv 中
# browser-user 命令路径: <workspace>/venv/Scripts/browser-use.exe
```

**操作路径：**

| 表格 | 路径 | 说明 |
|------|------|------|
| 表格一 | 筛选器 → 我的公司筛选 → 可转债安全性 → 开始筛选 → 导出CSV（单位壹） | 公司财务数据，317+条 |
| 表格二 | 债券 → 转股排名表格 → 导出CSV | 债券行情数据，320条 |

**注意：** 
- 筛选器配置页 URL：`https://www.lixinger.com/analytics/screener/company-fundamental/cn?screener-id=652be67894b68bd3b6bad12e`
- 导出CSV时需先点击「导出CSV」下拉按钮，选择「壹」，再点击「确定」
- CSV 编码为 UTF-8-SIG，数值带有 `=` 前缀需要清洗

### 步骤 2: 合并数据并计算安全性

使用 `scripts/generate_excel.py` 脚本处理两个 CSV：

```bash
python scripts/generate_excel.py
```

该脚本自动完成：
- 读取两个 CSV（可转债安全性.csv + 转股排名表格.csv）
- 清洗数值（去除 `=` 前缀）
- 计算三个安全性指标
- 生成 `可转债安全性_最终版.xlsx`

**安全性指标：**

| 指标 | 公式 | 达标条件 |
|------|------|----------|
| 利息保障倍数 | `EBIT / 利息支出(财务费用)` | >= 7 |
| 偿债能力 | `(货币资金 + 交易性金融资产) >= 流动负债 或 有息负债` | 满足其一 |
| 杠杆率 | `总负债 / 总市值` | <= 1.5 |

**评级标准：**

| 符合指标数 | 评级 |
|:----------:|:----:|
| 3 | **安全** 🟢 |
| 2 | **低风险** 🟢 |
| 1 | **中风险** 🟡 |
| 0 | **高风险** 🟠 |

### 步骤 3: 生成最终格式化 Excel

运行 `scripts/generate_excel.py` 后会生成 `可转债安全性_最终版.xlsx`，包含：
- 14 列：债券代码、债券名称、正股名称、PE-TTM、PB、股息率、最近转股更新日、最新债券价格、涨跌幅、双低、最近转股溢价率、最近转股价、最近转股价值、**安全性**
- 320 只可转债完整数据
- 行颜色（openpyxl 格式）：

| 评级 | 背景色 |
|------|:------:|
| 安全 | `00B050` 标准绿 |
| 低风险 | `92D050` 标准浅绿 |
| 中风险 | `FFFF00` 标准黄 |
| 高风险 | `FFC000` 标准橙 |
| 表头 | `00A3F5` 天蓝 + 黑色加粗字体 |

### 步骤 4: 上传至腾讯文档

使用 Tencent Docs MCP 工具连接并写入数据：

1. **获取文档信息：**
   - 目标文档 ID（使用 `manage.query_file_info` 或直接使用已知 file_id）
   - 获取 sheet 信息（`sheet.get_sheet_info`）
   - 读取原文档列结构（`sheet.get_cell_data`）确定表头

2. **写入数据：**
   - 清除旧数据（`sheet.clear_range_cells`，rows 2-321, cols 0-13）
   - 批量写入新数据（`sheet.set_range_value`，每行14列，共320行）
   - 注意：表头在 0-based row 0，数据从 row 1 开始

3. **应用样式：**
   - 表头样式：`set_cell_style` start_row=0, end_row=0, bg_color=FF00A3F5, font_color=FF000000, bold=true, horizontal_align=center
   - 数据行样式：逐行设置 bg_color（需要 226+ 次调用），或者直接使用导入方式（pre_import → curl 上传 → async_import）保留 openpyxl 样式

4. **导入方式（推荐，可保留样式）：**
   ```python
   # 1. pre_import → 获取 upload_url
   # 2. curl -X PUT -T excel_file.xlsx upload_url
   # 3. async_import
   # 4. import_progress 轮询至 100%
   ```

### 步骤 5: 文件名格式

输出文件命名规则：`可转债安全性_YYYY-MM-DD.xlsx`（使用当前日期）

## 参考资源

- `scripts/generate_excel.py` — 核心数据处理与Excel生成脚本（已包含列映射、指标计算、格式化）
- 腾讯文档 MCP 工具：`manage.pre_import` / `manage.async_import` / `manage.import_progress` / `sheet.set_cell_style` / `sheet.set_range_value` / `sheet.clear_range_cells`

## 注意事项

1. 理杏仁筛选器 ID `652be67894b68bd3b6bad12e` 是用户已配置的「可转债安全性」筛选条件
2. 安全性指标需要从原始财务数据（列11=货币资金, 12=交易性金融资产, 13=有息负债, 14=EBIT, 15=市值, 16=总负债, 17=流动负债, 8=利息支出）自行计算，而非使用筛选器自带的计算结果列
3. 腾讯文档 row 索引为 0-based：row 0 = 表头，row 1-320 = 数据
4. 320 只债券通过 正股名称 ↔ 公司名称 匹配安全性评级
5. 导出日期记录在表格说明行（可选）
