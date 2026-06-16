# Section Numbering Correction

When the Chinese source document has incorrect/out-of-order section numbering, 
renumber the English translation sequentially.

## Mapping from Chinese Original to Corrected English

### Chapter 3 — Major Sections (unnumbered in source → assigned)

| Source Chinese heading | Old Chinese # | Corrected English # |
|---|---|---|
| 遵循企业会计准则的声明 | (none) | 3.1 |
| 会计期间 | (none) | 3.2 |
| 营业周期 | (none) | 3.3 |
| 记账本位币 | (none) | 3.4 |
| 企业合并 | (none) | 3.5 |
| 合并财务报表编制方法 | (none) | 3.6 |
| 合营安排 | (none) | 3.7 |
| 现金及现金等价物 | (none) | 3.8 |
| 外币业务 | (none) | 3.9 |
| 金融工具 | (none) | 3.10 |
| 应收票据 | (none) | 3.11 |
| 应收账款 | (none) | 3.12 |
| 其他应收款 | (none) | 3.13 ← was 3.14, no 3.13 exists |
| 存货 | (none) | 3.14 |
| 合同资产 | (none) | 3.15 |
| 长期应收款 | (none) | 3.16 ← was 3.20, gap 3.17-3.19 |
| 长期股权投资 | (none) | 3.17 |
| 投资性房地产 | (none) | 3.18 |
| 固定资产 | (none) | 3.19 |
| 在建工程 | (none) | 3.20 |
| 借款费用 | (none) | 3.21 |
| 使用权资产 | (none) | 3.22 |
| 无形资产 | (none) | 3.23 |
| 长期待摊费用 | (none) | 3.24 |
| 长期资产减值 | (none) | 3.25 |
| 合同负债 | (none) | 3.26 |
| 职工薪酬 | (none) | 3.27 |
| 租赁负债 | (none) | 3.28 |
| 应付债券 | (none) | 3.29 |
| 收入 | (none) | 3.30 ← was 3.36, gap 3.34-3.35 |
| 合同成本 | (none) | 3.31 |
| 政府补助 | (none) | 3.32 |
| 递延所得税 | (none) | 3.33 |
| 租赁 | (none) | 3.34 |
| 其他重要会计政策和会计估计 | (none) | 3.35 |
| 重要会计政策、会计估计的变更 | (none) | 3.36 |

### Sub-section renumbering

| Old | New |
|-----|-----|
| 3.14.1 / 3.14.2 | → 3.13.1 / 3.13.2 |
| 3.15.1 – 3.15.5 | → 3.14.1 – 3.14.5 |
| 3.16.1 / 3.16.2 | → 3.15.1 / 3.15.2 |
| 3.20.1 / 3.20.2 | → 3.16.1 / 3.16.2 |
| 3.21.1 – 3.21.3.3 | → 3.17.1 – 3.17.3.3 |
| 3.23.1 – 3.23.4 | → 3.19.1 – 3.19.4 |
| 3.25.1 – 3.25.4 | → 3.21.1 – 3.21.4 |
| 3.27.1 – 3.27.3 | → 3.23.1 – 3.23.3 |
| 3.36.1 / 3.36.2 / 3.36.2.x | → 3.30.1 / 3.30.2 / 3.30.2.x |
| 3.39.1 – 3.39.4 | → 3.33.1 – 3.33.4 |
| 3.40.1 – 3.40.2.2 | → 3.34.1 – 3.34.2.2 |
| 3.41.1 | → 3.35.1 |
| 3.42.1 / 3.42.2 | → 3.36.1 / 3.36.2 |

## Implementation

```python
# Remove auto-numbering before adding text numbers
from docx.oxml.ns import qn
for p in doc.paragraphs:
    pPr = p._element.find(qn('w:pPr'))
    if pPr is not None:
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None:
            pPr.remove(numPr)

# Update cross-references in body text
xref_map = {
    'Note 3.21': 'Note 3.17',
    'Note 3.14': 'Note 3.13',
    'Note 3.20': 'Note 3.16',
    'Note 3.36': 'Note 3.30',
    'Note 3.39': 'Note 3.33',
    'Note 3.40': 'Note 3.34',
}
```

## Bold Rule for Headings

| Level | Example | Bold? |
|---|---|---|
| 1 | 3. | Yes |
| 2 | 3.1 | Yes |
| 3+ | 3.5.1, 3.10.1.1 | No |
