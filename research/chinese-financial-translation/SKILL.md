---
name: chinese-financial-translation
description: Chinese-to-English translation of financial statement footnotes (财务报表附注中译英). Handles terminology extraction from Word docs, multi-source conflict resolution, and application of user-defined hard rules.
category: research
---

# Chinese-to-English Financial Statement Translation

## When to load this skill
Load this whenever the user asks to translate Chinese financial statements, footnotes, or related accounting documents — especially when they provide reference/terminology Word documents.

## Workflow

### 1. Collect source documents
- Accept `.docx` files from the user (via attachment or Windows path)
- Extract text using `python-docx` in a venv (PEP 668 environments)
- Extract both **paragraphs** and **tables** — tables often contain the most authoritative fixed terminology

### 2. Identify multi-source conflicts
When the user provides multiple reference documents:
1. **Version prioritize**: newer documents override older ones
2. **Hard rules first**: explicit rules (e.g. "use Taxes, not Taxation") take priority over implicit preferences in older docs
3. **Flag conflicts, don't resolve unilaterally** — ask the user before deciding

### 3. Build the terminology reference
Save the consolidated terminology as a `references/` file under this skill, organized by:
- **Hard rules** (用户硬要求) — non-negotiable, from the most recent source
- **Terminology tables** — Chinese term → English translation
- **Pending conflicts** — items the user hasn't decided yet

### 4. Apply translation rules
When translating actual financial statements:

#### Hard rules (from version(3)):
- **Taxes**, not Taxation (标题/项目)
- **Current period / Previous period**, not Year 2025 / Year 2024
- **None**, not Nil
- **Book balance / Book value** (账面余额/账面价值体系)
- **Perpetual bonds**, not Sustainable debt
- Company names must be consistent throughout
- Title style: **sentence case** — only first word capitalized (e.g. `Book balance`, `Provision for bad debts`)
- Keep fixed abbreviations uppercase: **RMB, CAS, FVOCI, FVTPL**

#### Prioritization:
1. Hard rules (highest)
2. Latest reference document table terms
3. Earlier reference document table terms
4. General accounting English conventions (lowest)

#### Sentence case rule:
- First word only: `Notes receivable`, `Provision for bad debts`
- Fixed abbreviations keep uppercase: RMB, CAS, FVOCI, FVTPL
- Proper names keep original case
- Exception: table headers in existing English reference docs may use different casing — flag to user

### 5. Deliverable
Return the translated text inline (there is no attachment channel — the user reads your response in their terminal).

## Pitfalls
- ⚠️ **PEP 668 blocks system-wide pip install** — always create a venv first: `python3 -m venv /tmp/docx_venv && /tmp/docx_venv/bin/pip install python-docx`
- ⚠️ **Don't resolve conflicts without asking** — the user explicitly wants to decide. Flag each conflict with options.
- ⚠️ **Tables matter more than paragraphs** — table terminology in reference docs is the user's fixed corporate translation. Prefer table terms over paragraph terms.
- ⚠️ **Some documents use "draft" vs "bills" inconsistently** (e.g. bank acceptance draft vs bank acceptance bills) — flag for user resolution.
- ⚠️ **Estimated liabilities vs Provisions** — title-level uses Estimated liabilities; table/accrual-level uses Provisions. Don't blindly unify.
- ⚠️ **Single vs plural** varies across docs (Income tax expense vs Income tax expenses) — flag to user.
- ⚠️ **Reference docs may contain OCR/scan artifacts** — check for character anomalies (e.g. "Stage Ш" should be "Stage III").

## Reference files
- `references/terminology-extraction.md` — full extracted terminology from all 4 user-provided documents (unresolved conflicts included)
