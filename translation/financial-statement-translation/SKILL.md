---
name: financial-statement-translation
description: 财务报表及附注中译英翻译——术语规则、硬要求、工作流程
category: translation
---

# 财务报表及附注中译英翻译 Skill

## 适用范围
适用于**财务报表整份翻译**（含附注、表格、正文），规则通用。

## 触发条件
用户要求翻译财务报表或附注（中译英），或询问相关术语规则时加载本 skill。

## 分段翻译规则
- 翻译分批进行时，**不得将一个表格拆到不同批次中**
- 自然分界点优先选择章节边界或表格之间的空白处

## 全量验证规则（必执行）
每一批翻译完成后，必须执行以下验证：
1. **段落验证**：抽查头中尾至少3个段落，确认内容已翻译且加粗正确
2. **表格验证**：对涉及的所有表格，检查表头是否已翻译为英文
3. **字体验证**：英文正文 Times New Roman 10.5pt，表格内**所有数字**（金额/百分比/日期）Arial Narrow 10.5pt
4. **加粗验证**：与原文档逐格对比——三级及以上标题（如3.5.1、5.1.1）**不加粗**，(Continued)**不加粗**，其他表头/数据行**匹配中文原版**
5. **数字格式验证**：确认数字中无多余空格（如 "5, 385.00" 应修正为 "5,385.00"），逗号后紧跟数字
6. **头部缩进验证**：扫描全体章节标题，二级标题 left_indent 必须为 0，三级标题必须一致（~420 twips）。发现异常时修复
7. **全部批次完成后**：执行一次全文档扫描，确认**0个遗留中文单元格**
8. **逐表报告时附带章节标题**：向用户汇报每组表格状态时，输出该组表格所属的最近前驱章节标题（如"5.22 Short-term borrowings / 短期借款"），以便用户理解上下文位置

验证步骤未通过前，不得向用户报告"完成"。

## 自动编号处理
中文原版文档可能使用 Word 自动编号（numPr）而非文本编号。翻译后必须：
1. **删除自动编号**：移除所有段落中的 `w:numPr` 元素
2. **补文本编号**：为所有章节标题添加文本编号（如 "5.1 Monetary fund"）
3. **检查双重编号**：确保不会出现自动编号+文本编号同时显示

## 公司名称翻译规则
### 基本原则
- **所有中文必须翻译**，不得以"专有名词"为由保留中文
- 公司名逐成分翻译：青岛→Qingdao, 胶州→Jiaozhou, 控股→Holding
- 不确定官方英文名时：已知部分用官方名，未知部分用拼音转写
- 如需保留原文，必须先问用户

### 首字母空格规则
每个大写字母与前面小写字母之间应有空格：
```python
text = re.sub(r'([a-z])([A-Z])', r'\1 \2', text)
text = re.sub(r'(\d)([A-Z])', r'\1 \2', text)
```
✅ `Qingdao Jiaozhou Bay Development Group Co.,Ltd.`
❌ `QingdaoJiaozhouBayDevelopmentGroup`

## 参考文档（文件路径）
需要时读取：
- **会计科目对照（首选科目级术语）**: `/mnt/d/桌面/会计科目对照.docx`
- 原版术语清单: `/mnt/d/桌面/财务报表附注中译英术语参考清单.docx`
- 版本(2): `/mnt/d/桌面/财务报表附注中译英术语参考清单(2).docx`
- 版本(3)规则: `/mnt/d/桌面/财务报表附注术语参考清单(3).docx`
- **公司名/银行名字典（含完整胶发/上合体系名称）**: `references/translation-dictionary.md`
- **教训总结**: `references/session-lessons-june-2026.md`
- **XLSX 正式报表标准行项目翻译**: `references/xlsx-report-sheet-translations.md`

## 工作流程

### 附注翻译（DOCX）
1. **判断工作模式**：
   - 如果在中文原文档上直接修改 → 先创建中文原文档的副本（命名为 `原文件名_备份.docx`）
   - 如果是新建英文文档 → 无需创建中文副本
2. **备份验证**：创建后立即验证 `assert os.path.exists(path) and os.path.getsize(path) > 0`
3. 表格里的中文直接翻成英文，正文同步翻译
4. **标题序号处理**：如原文明显错号、跳号、重复号、层级错位，主动纠正。修正后清除 Word 自动编号（numPr），只保留文本序号，避免双重编号
5. **数字格式清理**：全文扫描 `re.sub(r'(\d), (\d)', r'\1,\2', text)` 去除逗号后多余空格
6. **质量检查** —— 翻译后逐段确认无残留中文汉字
7. **声明「完成」前**，必须做全文档扫描

### 报表翻译（XLS 工作簿）
1. **复制原文件**到输出目录，同时创建备份；验证两者都存在
2. **确定范围**：只翻译正式报表 sheet，跳过内部工作底稿。**正式报表 = 12个sheet**：
   - 合并资产表、合并负债表、合并利润表、合并现金流量表
   - 资产表、负债表、利润表、现金流量表
   - 合并股东权益变动表、合并股东权益变动表（续）
   - 股东权益变动表、股东权益变动表（续）
   - **不需要翻译**：Sheet1（附注正文）、填表说明、合并抵消分录、合并工作底稿、资本公积等——这些是 docx 附注或内部工作底稿
3. **方法**：用 openpyxl 加载，建立翻译映射表（dict），逐单元格查找替换；公式保持不变
4. **字体**：
   - 全部单元格：**Times New Roman, 10号**（不用 Arial Narrow，Arial Narrow 仅 docx 表格数字用）
   - 不对数字单元格做特殊的字体区分
5. **加粗规则（xlsx 专属）**：
   - **8 张正式表**（合并资产表/负债表/利润表/现金流量表 + 资产表/负债表/利润表/现金流量表）：**行 1-4 全部加粗**（含名称行、年份行、单位行、项目表头行）
   - **4 张权益变动表**（合并股东权益变动表及续、股东权益变动表及续）：**行 1-7 全部加粗**（含复合表头）
   - **正文部分**（第5行/第8行及之后）：加粗规则**逐格匹配中文原版**——从中文备份文件读取每个单元格的加粗状态，应用到英文文件的对应单元格
6. 前置环境：
   ```bash
   python3 -m venv /tmp/xl_venv
   /tmp/xl_venv/bin/pip install openpyxl
   /tmp/xl_venv/bin/python3 translate_xlsx.py
   ```

### xlsx 加粗/字体应用模板

```python
from openpyxl import load_workbook
from openpyxl.styles import Font

en_path = '英文.xlsx'
zh_path = '备份.xlsx'
wb_en = load_workbook(en_path)
wb_zh = load_workbook(zh_path)

bold_4 = ['合并资产表','合并负债表','合并利润表','合并现金流量表',
          '资产表','负债表','利润表','现金流量表']
bold_7 = ['合并股东权益变动表','合并股东权益变动表（续）',
          '股东权益变动表','股东权益变动表（续）']
all_sheets = bold_4 + bold_7

font_tnr_10 = Font(name='Times New Roman', size=10)
font_tnr_10_bold = Font(name='Times New Roman', size=10, bold=True)

# 1. 设置全部单元格字体
for sn in all_sheets:
    ws = wb_en[sn]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.font = font_tnr_10

# 2. 表头加粗
for sn in bold_4:
    ws = wb_en[sn]
    for r in range(1, 5):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                cell.font = font_tnr_10_bold

for sn in bold_7:
    ws = wb_en[sn]
    for r in range(1, 8):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is not None:
                cell.font = font_tnr_10_bold

# 3. 正文加粗匹配中文原版
for sn in all_sheets:
    ws_en = wb_en[sn]
    ws_zh = wb_zh[sn]
    skip_rows = 4 if sn in bold_4 else 7
    for r in range(skip_rows + 1, ws_en.max_row + 1):
        for c in range(1, min(ws_en.max_column + 1, ws_zh.max_column + 1)):
            cell_en = ws_en.cell(row=r, column=c)
            cell_zh = ws_zh.cell(row=r, column=c)
            if cell_en.value is not None and cell_zh.font and cell_zh.font.bold:
                cell_en.font = font_tnr_10_bold

wb_en.save(en_path)
```

### 批量翻译：delegate_task 分派模式（推荐用于大文件）

当 xlsx 文件有大量 sheet 且每个 sheet 中的中文残留较少（3-35个），用 `delegate_task` 并行处理比单会话逐张翻译更省 token：

1. **构造翻译包**：在主会话中对每个 sheet 提取中文残留单元格的位置和内容，加上术语规则/字典，作为子代理的 context
2. **并行分派**：使用 `tasks` 数组一次分派最多 3 个子代理（默认 max_concurrent_children=3）
3. **子代理只翻译不回写**：子代理返回 JSON `[{row, col, en}, ...]`，由主会话统一写回 xlsx
4. **子代理上下文结构**：每个子代理只需看到 `{sheet_name, chinese_cells: [{row, col, current_en, zh_original}]}` + 术语规则摘要（3-5K tokens），无需加载全文
5. **修正行号偏差**：xlsx 用 `data_only=True` 打开时，公式缓存值与实际单元格内容可能导致行号偏移。子代理返回后，在主会话中二次扫描确认所有中文已覆盖
6. **公式引用不计入中文残留**：`=+[4]资产表!D6` 中的 sheet 中文名是公式引用不可翻译，终验时应排除以 `=` 开头的单元格
7. **批次模板**：
   ```python
   # 主会话提取阶段
   import openpyxl, re, json
   wb = openpyxl.load_workbook(path, data_only=True)
   for sn in report_sheets:
       ws = wb[sn]
       cn_cells = []
       for row in ws.iter_rows():
           for cell in row:
               if cell.value and isinstance(cell.value, str) and re.search(r'[\u4e00-\u9fff]', str(cell.value)):
                   cn_cells.append({'row': cell.row, 'col': cell.column, 'current_en': str(cell.value)[:80]})
   
   # 分派
   tasks = [{"goal": f"Translate {sn}", "context": json.dumps(pkg), "toolsets": []} for sn, pkg in packages if pkg['has_chinese']]
   # 每次最多3个
   for i in range(0, len(tasks), 3):
       delegate_task(tasks=tasks[i:i+3])
   ```

### 关键区别对比

| 维度 | 附注 (DOCX) | 报表 (XLS) |
|------|------------|-----------|
| 翻译粒度 | 逐段逐格手动翻译 | 逐单元格匹配映射表自动替换 |
| 表格数字字体 | Arial Narrow 10.5pt | 统一 Times New Roman 10号 |
| 加粗处理 | 按大纲层级（一级二级加粗，三级及以上不加粗） | 行1-4或1-7表头加粗 + 正文逐格匹配中文原版 |
| 公式 | 无（docx 无公式概念） | 保持原公式不变 |
| 分次 | 需分批次（内容太多） | 可用 delegate_task 并行分派 |

---

## 批次翻译的核心纪律

### 🔴 铁律第一：全部翻译——不留中文

**所有内容必须全部翻译成英文，不得以任何理由保留中文。**

如果不确定某个内容是否应该翻译/保留，**先问用户**，不要自己做决定。

**绝对不要**用拼音逐字替换来处理中文残留——这会破坏已翻译好的内容，产生的拼音片段比中文更难识别。见陷阱7。

### ⚠️ 最大陷阱：fmt_table() 只格式化不翻译

`fmt_table()` 函数的作用**仅限于将数字单元格设为 Arial Narrow**。它**不会翻译任何中文文字**。

**错误模式**：
```python
# ❌ 只调 fmt_table() —— 表头仍是中文
fmt_table(t)
```

**正确模式**：
```python
# ✅ 每个表头都需要显式翻译
replace_cell(t.rows[0].cells[0], "Aging")
replace_cell(t.rows[0].cells[1], "Closing balance")
# 然后调用 fmt_table() 只做数字格式化
fmt_table(t)
```

**检查自己是否犯了这个错误**：回头看脚本中对每个表索引（`doc.tables[ti]`），确认每个 `fmt_table(t)` 之前都有对应的 `replace_cell` 调用来翻译表头文字。

### ⚠️ 第二大陷阱：乐观偏差 —— 默认之前批次仍然有效

声明「第N批完成」时，**不能假设之前批次的修改仍然有效**。实际遇到过的情况：
1. **大脚本静默失效**（>25KB）修改全部未写入，但脚本输出「成功」
2. **后一批覆盖前一批** —— 不同脚本操作同一段落范围时可能互相覆盖
3. **某批只完成部分工作** —— 中途错误导致后续内容未处理

### 声明「完成」前的终验

在说「全部完成」之前，**必须**做全文档扫描：
```python
total_cn_paras = 0
total_cn_tables = 0
for i in range(len(doc.paragraphs)):
    if any('\u4e00' <= c <= '\u9fff' for c in doc.paragraphs[i].text):
        total_cn_paras += 1
for ti in range(len(doc.tables)):
    for row in doc.tables[ti].rows:
        for cell in row.cells:
            if any('\u4e00' <= c <= '\u9fff' for c in cell.text):
                total_cn_tables += 1
                break
print(f"剩余中文段落: {total_cn_paras}, 剩余中文表格: {total_cn_tables}")
```

### 备份验证纪律

每次 `cp` 后立即验证：
```python
import os
assert os.path.exists(backup_path), f"备份未创建: {backup_path}"
assert os.path.getsize(backup_path) > 0, f"备份文件为空"
```

---

## 数字/金额格式说明

中文附注常用"万元"为单位。翻译规则（客户已确认）：

| 中文原文 | 换算 | 英文翻译 |
|---------|------|---------|
| X 万元 | X × 10,000 元，转为 million/billion | RMB Y million/billion yuan |
| 200万元 | 200 × 10,000 = 2,000,000 = 2 million | **RMB 2 million yuan** |
| 20,000万元 | 20,000 × 10,000 = 200,000,000 = 200 million | **RMB 200 million yuan** |
| 200,000万元 | 200,000 × 10,000 = 2,000,000,000 = 2,000 million | **RMB 2,000 million yuan** |

> **格式硬要求**：RMB [数字] million/billion yuan，末尾必须带 **yuan**

### 正文普通金额（非万元换算）—— 不用 RMB 前缀

当中文原文直接写 `X元`（而非 `X万元`），翻译为 `X yuan`，**不加 "RMB" 前缀**。

| 中文原文 | 正确翻译 | 错误翻译 |
|---------|---------|---------|
| 218,077,661.55元 | 218,077,661.55 yuan | ~~RMB 218,077,661.55~~ |
| 期末已到期未支付的应付票据总额为218,077,661.55元。 | The total amount ... was 218,077,661.55 yuan. | ~~was RMB 218,077,661.55.~~ |

**规则总结**：
- 万元换算 → `RMB X million/billion yuan`（带 RMB 前缀）
- 普通金额 → `X yuan`（不带 RMB 前缀）

## "以下简称"句式

| 中文 | 英文 |
|------|------|
| 以下简称"公司"或"本公司" | hereinafter referred to as the "Company" |

完整示例：
> Qingdao Jiaozhou Bay Development Group Co.,Ltd. (hereinafter referred to as the "Company")

## 术语优先级
1. **版本(3) 硬要求** → 最高优先级（最新规则）
2. **用户特别指示** → 次高优先级
3. **会计科目对照.docx** → 科目级术语基准
4. **原版术语清单** → 附注句式/固定表达参考
5. **对照表2** → 仅供辅助参考（基于旧准则，术语已过时）

## 硬要求（最高优先级）

| 规则 | 内容 |
|------|------|
| 税费/税项 | 标题和表格表头用 **Taxes**，正文部分可用 Taxation |
| 年度列头 | 用 **Current period / Previous period**，不用 Year 2025 / Year 2024 |
| "无" | 统一用 **None**，不用 Nil |
| 账面余额 | **Book balance**（中文未明确写"原始成本"时） |
| 账面价值 | **Book value**（中文未明确写"原始成本"时） |
| 原始成本 | 仅当中文明确定义为"原始成本"时才用 **Original cost** |
| 公司名称 | 全文统一，不可随意改写 |
| 永续债 | **Perpetual bonds**，不用 Sustainable debt |
| 预计负债（双轨） | 标题/政策层面用 **Estimated liabilities**；报表项目/计提层面用 **Provisions** |

## 混合单元格翻译规则

### 日期+金额混合单元格
原文：`1年以内67,548,732.03元，1-2年4,000,000.00元,2-3年9,657,007.82元`
译文：`Within 1 year: 67,548,732.03 yuan; 1 to 2 years: 4,000,000.00 yuan; 2 to 3 years: 9,657,007.82 yuan`

格式：`Within X years: [amount] yuan; X to Y years: [amount] yuan`

### 抵（质）押品/担保人 格式规则
冒号前和冒号后分开翻译：
- 冒号前：`抵（质）押品/担保人` → `Collateral (Pledged assets)/Guarantor`
- 冒号后：按顺序逐项翻译，用逗号分隔
- `在建工程-商业类在建工程` → `Construction in progress - Commercial construction in progress`

### 金额数字字体规则
- **所有数字**（金额数字/百分比/日期）→ Arial Narrow 10.5pt（表格内）
- 数字格式：`5,385,302.36`（逗号后无空格），不是 `5, 385, 302.36`

### "Co.,Ltd." 后必须加空格
```python
# ❌ 错误
Bank Co.,Ltd.Qingdao Branch
# ✅ 正确
Bank Co.,Ltd. Qingdao Branch
```
**规律**：`Co.,Ltd.` 或 `Co., Ltd.` 后紧跟字母时必须加空格。修复方法：
```python
import re
text = re.sub(r'Co\.,Ltd\.([A-Za-z])', r'Co.,Ltd. \1', text)
text = re.sub(r'Co\., Ltd\.([A-Za-z])', r'Co., Ltd. \1', text)
```
在每次表头修复、公司名重建后做一次全局扫描。

## 已确认的术语对照

| 中文 | 英文 | 来源/备注 |
|------|------|----------|
| 本期金额 | **Current amount** | 用户指定 |
| 上期金额 | **Previous amount** | 用户指定 |
| 期初余额 | **Opening balance** | 用户指定 |
| 期末余额 | **Closing balance** | 用户指定 |
| 本期发生额 | **Current period amount** | 版本(3) 表 1 |
| 主营业务收入 | **Operating revenue** | 用户指定 |
| 营业利润 | **Operating income** | 用户指定 |
| 营业收入（和成本） | **Operating revenue (and operating costs)** | 原版章节标题 |
| 其他收益 | **Other incomes** | 用户指定（复数形式） |
| 资产处置收益 | **Gains from disposal of assets** | 用户指定 |
| 资产处置损益 | **Profit & loss from disposal of assets** | 会计科目对照 6115 |
| 所得税费用 | **Income tax expense**（单数） | 用户指定 |
| 信用减值损失 | **Credit impairment loss**（单数） | 用户指定 |
| 管理费用 | **Administrative expenses** | 用户指定 |
| 财务费用 | **Finance expenses** | 用户指定 |
| 合同履约成本 | **Contract fulfillment costs** | 用户指定 |
| 递延所得税资产 | **Deferred income tax assets** | 用户指定 |
| 递延所得税负债 | **Deferred income tax liabilities** | 用户指定 |
| 营业外收入 | **Non-operating income** | 会计科目对照 6301 |
| 营业外支出 | **Non-operating expense** | 会计科目对照 6711 |
| 投资收益 | **Income from investments** | 会计科目对照 6111 |
| 公允价值变动收益 | **Fair value change gains** | 原版 6.62 |
| 资产减值损失 | **Asset impairment loss** | 原版 6.64 |
| 税金及附加 | **Taxes and surcharges** | 原版 6.55 |
| 销售费用 | **Selling expenses** | 原版 6.56 |
| 研发费用 | **Research and development costs** | 原版 6.58 |
| 主营业务成本 | **Cost of Sales** | 会计科目对照 6401 |
| 应付职工薪酬 | **Payroll Payable** | 会计科目对照 2211 |
| 应交税费 | **Taxes payable** | 会计科目对照 2221 |
| 实收资本 | **Paid-in capital** | 会计科目对照 4001 |
| 资本公积 | **Capital reserve** | 会计科目对照 4002 |
| 盈余公积 | **Surplus reserve** | 会计科目对照 4101 |
| 未分配利润 | **Undistributed profit** | 会计科目对照 4104.06 |
| 其他综合收益 | **Other comprehensive income** | 会计科目对照 4003 |
| 短期借款 | **Short-term borrowings** | 原版 |
| 短期薪酬 | **Short-term employee benefits** | 用户指定（非 Short-term salary） |
| 货币资金 | **Monetary fund** | 会计科目对照 |
| 应收账款 | **Accounts Receivable** | 会计科目对照 1122 |
| 应收票据 | **Notes receivable** | 会计科目对照 1121 |
| 应收股利 | **Dividend receivable** | 会计科目对照 1131 |
| 应收利息 | **Interest receivable** | 会计科目对照 1132 |
| 其他应收款 | **Other Receivable** | 会计科目对照 1221 |
| 坏账准备 | **Provision for bad debts** | 原版/版本(3) |
| 存货跌价准备 | **Inventory provision** | 原版 6.9.1 |
| 固定资产 | **Fixed assets** | 通用标准 |
| 在建工程 | **Construction in progress** | 通用标准 |
| 无形资产 | **Intangible assets** | 通用标准 |
| 使用权资产 | **Right-of-use assets** | 原版 4.25 |
| 租赁负债 | **Lease liabilities** | 原版 4.32 |
| 投资性房地产 | **Investment property** | 原版 4.21 |
| 长期股权投资 | **Long-term equity investments** | 原版 4.20 |
| 递延收益 | **Deferred income** | 原版 4.37/6.45 |
| 应付账款 | **Accounts payable** | 会计科目对照 2202 |
| 应付票据 | **Notes payable** | 会计科目对照 2201 |
| 应付利息 | **Interest payable** | 会计科目对照 2231 |
| 应付股利 | **Dividends payable** | 会计科目对照 2232 |
| 合同负债 | **Contract liabilities** | 会计科目对照 2205 |
| 应付债券 | **Bonds payable** | 会计科目对照 2502 |
| 库存现金 | **Cash on hand** | 会计科目对照 1001 |
| 银行存款 | **Cash at bank** | 会计科目对照 1002 |
| 其他货币资金 | **Other Monetary Capital** | 会计科目对照 1012 |
| 原材料 | **Raw material** | 会计科目对照 1403 |
| 库存商品 | **Inventory** | 会计科目对照 1405 |

## 格式规则

### 字体

| 场景 | 字体 | 字号 | 适用范围 |
|------|------|------|----------|
| 英文正文/标题 | **Times New Roman** | **五号（10.5pt）** | 附注(docx) |
| 表格内**所有数字**（金额/百分比/日期） | **Arial Narrow** | **五号（10.5pt）** | **仅附注(docx)翻译时** |
| 表格外数字 | 同英文正文 Times New Roman | **五号（10.5pt）** | 附注(docx) |
| **报表(xls)全部单元格** | **Times New Roman** | **10号（非10.5pt）** | 报表(xls)翻译时——所有单元格统一字体，不做数字/文字区分 |

### 加粗规则
| 层级 | 示例 | 加粗？|
|------|------|--------|
| 大纲一级标题（docx） | 1. | **加粗** |
| 大纲二级标题（docx） | 1.1 | **加粗** |
| 大纲三级及以上（docx） | 1.1.1 / 3.5.1 / 5.1.1 | **不加粗** |
| (Continued) / （续）- docx | — | **不加粗** |
| 附注文档首标题（docx） | Company title line | **两行均加粗** |
| 表格内（表头行）- docx | 中文原版加粗的表头 | **逐格对应加粗** |
| 表格内（数据行）- docx | 中文原版特定加粗的单元格 | **保持原样** |
| **报表行1-4（xlsx）** | 8张正式表：名称/年份/单位/表头行 | **加粗** |
| **报表行1-7（xlsx）** | 4张权益变动表：含复合表头的区域 | **加粗** |
| **报表正文（xlsx）** | 第5/8行及之后所有单元格 | **逐格匹配中文原版** |

### 大小写规则

**情况一：公司名称**
标准专名大写。

**情况二：专有词缩写**
保持大写，如：RMB, CAS, FVOCI, FVTPL, VAT

**情况三：正文和标题**
**sentence case** —— 仅首单词首字母大写，其余小写（除非是专名/缩写）。

**唯一例外：会计准则正式名称**
保留标准大写（每个实词首字母大写）：
> Accounting Standards for Business Enterprises

## 日期格式

| 场景 | 格式 | 示例 | 适用范围 |
|------|------|------|----------|
| 单日 | YYYY-M-D | 2026-6-4 | **仅限附注表格内** |
| 日期范围 | YYYY-M-D to YYYY-M-D | 2025-6-4 to 2026-6-4 | **仅限附注表格内** |
| 正文日期 | 保留原格式 | / | 正文里的公司成立时间等不受此规则限制 |
| 月份名称 | 保留首字母大写 | December 31, 2024 | 不受此规则影响 |

## 固定句式

- The financial statements of the Company are prepared on a going concern basis.
- The statements are prepared in accordance with the Accounting Standards for Business Enterprises.
- For receivables formed by transactions regulated by the Revenue Standard, the Company measures the loss allowance at an amount equal to the lifetime expected credit losses.
- Inventories are measured at the lower of cost and net realizable value.
- The Company adopts the perpetual inventory system.
- Fixed assets are depreciated using the straight-line method over their useful lives.
- The Company recognizes revenue when the customer obtains control of the relevant goods.
- Government grants related to assets are recognized as deferred income and are included in profit or loss for the current period.
- Deferred tax assets are recognized to the extent that it is probable that taxable income will be available.

## 表格高频表头

| 中文 | 英文 |
|------|------|
| 账面余额 | Book balance |
| 账面价值 | Book value |
| 期末余额 | Closing balance |
| 期初余额 | Opening balance |
| 本期金额 | Current amount |
| 上期金额 | Previous amount |
| 坏账准备 | Provision for bad debts |
| 减值准备 | Impairment provision |
| 类别 | Category |
| 金额 | Amount |
| 比例（%） | Percentage (%) |
| 账龄 | Aging |
| 本期计提 | Accrual |
| 本期收回或转回 | Recovery or reversal |
| 本期核销 | Write-off |
| 按单项计提 | on an individual basis |
| 按组合计提 | on a collective basis |
| 合计 | Total |
| 其中 | Including: |

## 账龄区间表达

- Within 1 year
- 1 to 2 years
- 2 to 3 years
- 3 to 4 years
- 4 to 5 years
- More than 5 years

## 已知陷阱（翻译时必读）

### 陷阱1：表格处理——批量格式化函数不翻译内容
`fmt_table()` 类函数只将数字单元格设为 Arial Narrow 字体，**不会翻译中文表头和行标签**。
- ✅ 正确做法：对每个表格单独翻译表头，再对数据行用字典逐格翻译
- ❌ 错误做法：批量格式化后跳过内容翻译

### 陷阱2：大脚本静默失效风险
超过 ~25KB 的 Python 脚本可能执行后宣称成功，但对 docx 的修改**未持久化**。
- 单脚本控制在 15KB 以内
- 每次执行后立即验证关键输出
- 发现未生效时拆分为更小的单元重试

### 陷阱3：文件操作必须验证结果
`cp` 命令可能因 WSL 权限/Windows 锁定返回0但实际未写入。
- 每次 `cp` 后 `ls -lh` 验证
- 被锁定时用 PowerShell 删除再复制

### 陷阱4：中英混杂的渐进式翻译问题
多次 pass 逐步替换中文会产生"Frankenstein 字符串"（如"按组Total提Provision"）。
- ✅ 推荐做法：从中文原版一次性完整翻译每个单元格
- ❌ 避免做法：用多次 pass 逐步替换，前半段改英文后半段遗漏中文

### 陷阱5：声明「完成」前必须全量验证
- 不得只验证头几个表格就认为全部完成
- 必须执行完整扫描：对所有单元格逐格检查中文残留

### 陷阱7：拼音逐字替换会毁灭文档（绝对不要做）

当有大量中文残留时，**绝对不要写脚本做逐字拼音替换**（如 '金'→'Jin', '额'→'E'）。

**为什么不行**：
- 多个不同汉字映射到同一个拼音（多对一），拼音无法区分
- 替换运行在已部分翻译的文档上，会破坏已翻译正确的英文
- 生成的拼音片段（如"Jin E"代替"金额"）比中文更难看懂
- 一旦执行，修复极其困难——只能从备份恢复

**正确做法**：
1. 从中文原版一次性完整翻译每个单元格内容
2. 或者用完整短语作为替换单元（替换整个"其他应收款"→"Other Receivable"）
3. **永远不要**用单字拼音映射表做全局替换

### 陷阱8：表格首列（行标签）容易被遗漏

当批量处理表格时，`fmt_table()` 只处理数字单元格，首列的文字标签（项目名称、公司名等）会被忽略。用户多次在以下表格中发现未翻译的首列内容：

- 在建工程明细表（项目全名，如"青岛市停车位项目"）
- 借款明细表（贷款单位名称）
- 担保明细表（担保人名称）
- 无形资产明细表（资产名称）
- 投资性房地产明细表（房产项目名称）

**必须在 `fmt_table()` 之后，对所有表格的首列逐格检查并翻译。**

### 陷阱9：表格加粗必须逐格匹配中文原版

不能仅匹配表头行的加粗。用户要求表头和数据行内的加粗单元格**都与中文原版逐格一致**。

推荐做法：写一个同步函数，从中文备份文档读取每行每列的加粗状态，写入英文文档的对应位置。

### 陷阱14：逐行重建单元格后bold会丢失——必须手动恢复

当从中文原版逐行重建表格（清空单元格→写入英文）时，**用 `cell.paragraphs[0].clear()` + `add_run(text)` 写入的文字默认不加粗**。即使中文原版的对应单元格是加粗的，重建后的英文也会丢失加粗。

**典型场景**：借款明细表的表头行、应付账款账龄表的子表头行（如"Percentage(%)"）

**正确做法**：重建后单独同步加粗——从中文原版逐格读取加粗状态，应用到英文文档的对应单元格：

```python
zh = Document('备份.docx')
en = Document('英文.docx')
for ti in range(len(en.tables)):
    t_en = en.tables[ti]
    t_zh = zh.tables[ti]
    for ri in range(min(len(t_en.rows), len(t_zh.rows))):
        for ci in range(min(len(t_en.rows[ri].cells), len(t_zh.rows[ri].cells))):
            zh_bold = any(r.bold for r in t_zh.rows[ri].cells[ci].paragraphs[0].runs if r.text.strip())
            if zh_bold:
                for r in t_en.rows[ri].cells[ci].paragraphs[0].runs:
                    r.bold = True
```

或者在重建函数中同时从 zh 读取加粗状态并设置。

### 陷阱10："全部翻译"不是建议，是铁律

用户反复强调：**公司名、项目名、债券代码、文件、合同——所有中文都必须翻译。** 不要假设任何中文可以保留。

不确定某个内容是否应该翻译时，**先问用户**，不要自己做决定。

### 陷阱17：方向词（以东/以西/以南/以北）不能用全局正则翻译

在地址描述中，`以东`、`以西`、`以南`、`以北` 表示方位。**绝不要**在整个表格的 col5 上跑全局正则替换，因为：

- 公司名中的成分可能被误匹配，产生 `north of east of Ltd.` 这样的乱码
- 部分翻译的单元格中英混杂，正则无法区分上下文

**正确做法**：
1. 从中文原版完整读取数值
2. 如果确实需要方位翻译，只对明确的地址行做硬编码翻译
3. 或者把整段中文地址交给用户确认翻译

示例地址翻译：
```
胶州经济技术开发区尚德大道以东永定河路以北
→ Jiaozhou Economic and Technological Development Zone, east of Shangde Avenue, north of Yongdinghe Road
```

### 陷阱6：三级标题加粗检查

三级及以上标题（3.5.1, 5.1.1, 12.1.1 等）**不加粗**。完成翻译后用正则扫描全文档确认：
```python
import re
for i, p in enumerate(doc.paragraphs):
    if re.match(r'^\d+\.\d+\.\d+\b', p.text.strip()):
        assert not any(r.bold for r in p.runs if r.text.strip()), f"P{i} is level 3+ but bold!"
```

### 陷阱11：不要修补碎片化英文单元格——从中文原版逐行重建

当用户提供了一批公司名/银行名翻译，而英文文档中对应单元格已经包含碎片化的英文（如 "工业园 Distric"、"存单Pledged" 等混杂字符串）：

**❌ 错误做法**：在英文单元格中搜索中文子串并替换
```python
# ❌ 不可靠——中文已部分替换，找不到完整原文
if '恒丰' in en_cell.text:
    en_cell.text = en_cell.text.replace('恒丰', 'Hengfeng')
```

**✅ 正确做法**：读中文原版，逐行重建每个单元格
```python
zh = Document('备份.docx')
en = Document('英文.docx')
for ri in range(1, len(zh.tables[ti].rows)):
    zh_text = zh.tables[ti].rows[ri].cells[ci].text.strip()
    for cn, en_name in company_dict.items():
        if cn in zh_text:    # 中文原版全文匹配
            set_cell(en.tables[ti].rows[ri].cells[ci], en_name)
            break
```

这种表的数据行通常金额日期不变（纯数字），只需翻译公司名/银行名/担保描述三列。

⚠️ **陷阱嵌套警告**：逐行重建后 bold 会丢失——见陷阱14。

### 陷阱15：重建后必须清理全角标点

当从中文原版 `zh.tables[ti].rows[ri].cells[ci].text` 重建单元格内容时，中文标点会遗留在英文结果中：

| 遗留标点 | 位置 | 应替换为 |
|---------|------|---------|
| 全角冒号（：） | 如 "Guarantor：Qingdao..." | 半角冒号加空格 `: ` |
| 中文逗号（，） | 如 "185,076,352.35 yuan，1 to 2 years" | 半角分号加空格 `; ` |
| 中文顿号（、） | 如 "Co.,Ltd.、Qingdao..." | 半角逗号加空格 `, ` |
| 中文空格（　） | 数字/文本间 | 常规空格 |

在重建/替换代码的最后加一条清理步骤：
```python
import re
for row in en_table.rows:
    for cell in row.cells:
        txt = cell.text.strip()
        if '：' in txt or '，' in txt or '、' in txt or '　' in txt:
            txt = txt.replace('：', ': ').replace('，', ', ').replace('、', ', ').replace('　', ' ')
            txt = re.sub(r'  +', ' ', txt).strip()
            set_cell(cell, txt)
```

### 陷阱16：账龄+金额混合单元格拆分时注意数字中的逗号

中文原版账龄描述：`1年以内185,076,352.35元，1-2年133,715,731.04元`

**✅ 可用拆分**：中文千位分隔符用英文逗号（,），段落分隔用中文逗号（，），
所以 `zh_text.split('，')` 可行——前提是确认中文原文确实用中文逗号做分隔。

**更安全的做法**：用正则匹配账龄模式，或用预定义翻译（推荐）：
```python
agings = [
    None,  # header
    'Within 1 year: 185,076,352.35 yuan; 1 to 2 years: 133,715,731.04 yuan',
    'Within 1 year',
    ...
]
for ri, txt in enumerate(agings):
    if txt: set_cell(t66.rows[ri].cells[2], txt)
```
- 多次 pass 后英文单元格内容与中文原文已大不相同，搜索子串会漏匹配
- 用不完整的字符串匹配可能破坏已翻译好的相邻部分
- 逐行重建是唯一保证 100% 正确翻译的方法

### 陷阱12：公司名传播——修复一张表后同步修复相邻表

### 陷阱18：合并表头的表格不能简单逐格覆盖

带合并单元格（merged cells）的表头（如应付债券变动表的 R0 子表头 + R1 细分表头），**不能**对所有单元格调用 `set_cell()`，因为：

1. 合并单元格中的"虚拟"单元（被合并掉的 slot）不接受文字写入——写入不报错但也不生效
2. 清空一个合并区域中某个格子可能破坏合并结构
3. R1 子表头可能**不是**独立行，而是 R0 合并单元格的补充

**正确做法**：
```python
# 先打印实际单元格内容确认哪些格子是真实可写的
for ri in range(2):
    for ci in range(ncols):
        print(f'R{ri}C{ci}: [{table.rows[ri].cells[ci].text.strip()}]')
# 只对非空的实际格写入
headers = {(0,0): 'Bond name', (0,3): 'Bond term', ...}
for (ri, ci), hdr in headers.items():
    set_cell(table.rows[ri].cells[ci], hdr)
```
合并表头常见于：债券变动表、账龄表（有"期末余额"跨 "金额"和"比例(%)" 两列）。

当用户提供了特定表格（如 T60 质押借款明细）的公司名翻译后，同一个章节的相邻表格（如 T61 抵押借款明细、T62 保证借款明细）往往使用相同公司/银行名称。

**原则**：不要只修复用户提到的表格——主动检查同一节下的所有相邻表，对相同的公司名/银行名做批量替换。

检查步骤：
1. 记录用户提供的所有公司名/银行名及其翻译
2. 扫描当前表格前后各 2-3 张表（同一章节内）
3. 如果发现引用同样的公司/银行名，一并修正

### 陷阱18：合并表头的表格不能简单逐格覆盖

带合并单元格（merged cells）的表头（如应付债券变动表的 R0 子表头 + R1 细分表头），不能对所有单元格调用 set_cell()，因为：

1. 合并单元格中的虚拟单元（被合并掉的 slot）不接受文字写入——写入不报错但也不生效
2. 清空一个合并区域中某个格子可能破坏合并结构
3. R1 子表头可能不是独立行，而是 R0 合并单元格的补充

### 陷阱19：章节标题缩进一致性检查

修改章节标题时，必须检查缩进是否与同级标题一致。误触发的缩进会导致格式异常。

典型错误：5.45 Administrative expenses 的 left_indent=839，而同级标题（5.44、5.46、5.47 等）都是 left_indent=0。

### 陷阱20：正文金额不用 RMB 前缀

当中文原文只写 X元（不带人民币三字）时，翻译为 X yuan，不要添加 RMB 前缀。此规则全篇适用。

### 陷阱21：拼音碎片残留不易察觉

当单元格的中文被部分翻译后，残留的**拼音片段**（如 `Wu Ye Fei`、`Cang Chu Guan Li Fei`、`Ye Wu`、`He Ji`）比中文汉字更难发现——人工阅读时容易误认为"这好像是英文"。

**检测方法**：终验扫描时，除了检查汉字，还要检测常见拼音模式：

```python
# 拼音音节模式：连续的罗马字母，不在英文词典中
# 启发式：找长度 2-6 的罗马字母组
english_words = {'the', 'and', 'for', 'not', 'but', 'are', 'was', 'had', 'has',
                 'its', 'all', 'can', 'may', 'per', 'sum', 'due', 'net', 'add',
                 'less', 'total', 'item', 'name', 'type', 'date', 'year', 'cash',
                 'cost', 'sale', 'unit', 'rate', 'fee', 'tax', 'pay', 'debt',
                 'city', 'land', 'park', 'bank', 'bond', 'fund', 'note', 'loan',
                 'wen', 'lv', 'ye', 'wu', 'fei', 'he', 'ji', 'guang', 'gao',
                 'she', 'cang', 'chu', 'guan', 'li', 'yun', 'ying'}
# 对每个单元格的文本提取单词，标记不在白名单中的短词
for word in cell_text.lower().split():
    word = word.strip('.,;:()[]')
    if 2 <= len(word) <= 6 and word not in english_words:
        print(f'可疑拼音: {word}')  # 人工复核
```

**本项目中发现的拼音残留**：
| 拼音 | 正确翻译 |
|------|---------|
| Wu Ye Fei | Property management fees |
| Cang Chu Guan Li Fei | Warehousing and management fees |
| He Ji | Total |
| Ye Wu | Operations |
| Wen Lv | Cultural Tourism |
| Guang Gao | Advertising |

### 陷阱22：超长大表（100+行）的批量修复策略

当遇到 100 行以上的大表（如担保明细表、关联方往来表），最有效的方法不是逐格检查，而是：

1. **读中文原文 + 字典映射**：一次性从中文原版读取所有行的公司名，用完整公司名映射表批量翻译
2. **只处理文本列**：金额/日期等数字列原样保留（已验证与中文一致）
3. **一次性重建**：同时修复 col 0（借款单位/担保单位）和 col 1（贷款单位/被担保对象）以及担保描述列
4. **后续验证**：重建后检查中文残留 + bold 同步

```python
# 大表处理模板
zh = Document('备份.docx')
en = Document('英文.docx')
t_en = en.tables[ti]
t_zh = zh.tables[ti]
co_dict = { /* 全部公司名映射 */ }
for ri in range(1, len(t_zh.rows)):
    zh_name = t_zh.rows[ri].cells[ci].text.strip()
    for cn, en_name in sorted(co_dict.items(), key=lambda x: -len(x[0])):
        if cn in zh_name:
            set_cell(t_en.rows[ri].cells[ci], en_name)
            break
# 完成后同步 bold
sync_bold_from_zh(en, zh, ti)
```

### 陷阱23：现金流量表补充资料固定句式不可增量替换

现金流量表补充资料（T116）的项目描述非常标准化，每个都带 `(收益以"－"号填列)` 或 `(增加以"－"号填列)` 这种固定尾缀。

**正确做法**：直接硬编码完整翻译，不要尝试在碎片化英文上做增量替换：

| 中文 | 英文 |
|------|------|
| 处置固定资产、无形资产和其他长期资产的损失（收益以"－"号填列） | Loss on disposal of fixed assets, intangible assets and other long-term assets (gains indicated by "-") |
| 公允价值变动损失（收益以"－"号填列） | Fair value change losses (gains indicated by "-") |
| 投资损失（收益以"－"号填列） | Investment losses (gains indicated by "-") |
| 递延所得税资产减少（增加以"－"号填列） | Decrease in deferred income tax assets (increase indicated by "-") |
| 递延所得税负债增加（减少以"－"号填列） | Increase in deferred income tax liabilities (decrease indicated by "-") |
| 存货的减少（增加以"－"号填列） | Decrease in inventories (increase indicated by "-") |
| 经营性应收项目的减少（增加以"－"号填列） | Decrease in operating receivables (increase indicated by "-") |
| 经营性应付项目的增加（减少以"－"号填列） | Increase in operating payables (decrease indicated by "-") |

### 陷阱24：子公司/联营企业取得方式翻译

在"长期股权投资"（Chapter 12）和"企业集团构成"（Chapter 7）表中，取得方式的翻译需要统一：

| 中文 | 英文 |
|------|------|
| 投资设立 | Investment establishment |
| 股权划转 | Equity transfer |
| 并购 | Merger and acquisition |

### 陷阱25：核算方法统一翻译

| 中文 | 英文 |
|------|------|
| 成本法 | Cost method |
| 权益法 | Equity method |

### 陷阱26：担保明细表中的公司名传播范围

T127（担保明细表）通常有 100+ 行，涉及大量集团内子公司/外部公司。如果用户提供了部分公司名翻译，**不要只修复用户提到的那些**——担保表中所有的担保单位/被担保对象都使用同一批公司名，应该一次性全部映射。

执行顺序：
1. 先完整扫描中文原版的 col 0 和 col 1，提取所有出现过的公司名
2. 从已积累的字典中查找翻译
3. 对未命中的公司名询问用户
4. 一次性批量重建整张表

### 陷阱27：XLSX 公式引用中的中文 sheet 名不可翻译

当扫描 xlsx 文件的中文残留时，公式单元格（以 `=` 开头）中的中文 sheet 名引用会显示为"中文残留"：

```
=+[4]资产表!D6    ← "资产表"是 sheet 名引用，不是文本
=+合并利润表!C5    ← 同上
```

**规则**：
- 终验扫描时先过滤掉以 `=` 开头的单元格
- sheet 名是文件结构的一部分，翻译 sheet 名会破坏跨表公式引用
- 如果要翻译 sheet 名，必须先创建备份，然后重命名 sheet 并同步更新所有公式引用（仅在用户明确要求时做）

```python
# 正确的终验扫描——排除公式单元格
text_cn, formula_cn = 0, 0
for sn in report_sheets:
    ws = wb[sn]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and re.search(r'[\u4e00-\u9fff]', str(cell.value)):
                if str(cell.value).startswith('='):
                    formula_cn += 1  # 公式引用，不计入残留
                else:
                    text_cn += 1      # 真正的中文文本残留
```

### 陷阱28：delegate_task 子代理可能返回不完整的结构

用 delegate_task 分派 xlsx 翻译时，子代理可能：
- 写入文件到意外路径（如 `/mnt/c/Users/李/`）而不是在 summary 中返回 JSON
- 返回的 JSON 缺少 row/col 字段（只返回翻译值）
- 返回的行号与实际 xlsx 文件偏差（由于 `data_only=True/False` 导致公式行 vs 值行不一致）

**应对措施**：
1. 在子代理的 context 中明确要求返回 `[{row, col, en}, ...]` 结构
2. 优先从子代理的 summary 中提取 JSON，而不是依赖文件写入
3. 子代理返回后，在主会话中做二次全量扫描验证

### 会计科目对照.docx 中的拼写错误（翻译时自动校正）
- Impairement → **Impairment**
- Deffered → **Deferred**
- Electoric → **Electronic**
- Transportaion → **Transportation**
- Manufactoring → **Manufacturing**

### 其他注意点
- 标题与正文黏连处，翻译时应拆为"规范标题 + 正文说明"
- 中英文标点统一（英文句点、逗号后需空格）
- "The Company" 中 C 始终大写（指代本公司）
