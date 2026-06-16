# Frankenstein 字符串恢复方案

## 问题描述

当对同一单元格进行多次 pass 的部分中文替换后，单元格内容变成中英混杂状态（如"按组Total提Provision for bad debts"），后续 pass 无法进一步清理。

## 恢复技术

```python
# 1. 同时加载英文版和中文原版备份
en_doc = Document(EN_PATH)
zh_doc = Document(ZH_BACKUP_PATH)

# 2. 对每个仍有中文的单元格
for ti in range(min(len(en_doc.tables), len(zh_doc.tables))):
    for ri in range(min(len(en_doc.tables[ti].rows), len(zh_doc.tables[ti].rows))):
        for ci in range(len(en_doc.tables[ti].rows[ri].cells)):
            cell = en_doc.tables[ti].rows[ri].cells[ci]
            if not has_cn(cell.text):
                continue
            
            # 关键：从原版备份获取干净的原始中文
            zh_text = zh_doc.tables[ti].rows[ri].cells[ci].text.strip()
            
            # 用完整字典一次性翻译
            new_text = zh_text
            for cn, en in sorted(DICT.items(), key=lambda x: -len(x[0])):
                if cn in new_text:
                    new_text = new_text.replace(cn, en)
            
            # 替换单元格内容（覆盖混杂状态）
            replace_cell(cell, new_text)
```

## 注意事项

- 必须同时持有英文版文档和中文原版备份
- 字典条目必须按长度降序应用（最长匹配优先）
- 该方法仅适用于 docx（python-docx）；xlsx 类似但需用 openpyxl
