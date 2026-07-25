# MCPTox Reproduction Slides Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the two reproduction placeholders in the 12-slide Persona-theme MCPTox deck with verified 50-sample and 300-sample experimental results.

**Architecture:** Recompute all displayed metrics from the supplied report/raw JSON into one audited data object, then import the existing PPTX with `@oai/artifact-tool` and edit slides 10–11 in place. Export a new PPTX, render every slide, and validate statistics, text residue, placeholders, and Git state.

**Tech Stack:** PowerShell, Node.js ES modules, `@oai/artifact-tool`, Git.

## Global Constraints

- Preserve the MCP-Persona master, layout, palette, typography, headers, and page numbering.
- Keep the final deck at exactly 12 slides; replace slides 10 and 11 rather than appending slides.
- Use valid-response denominators for ASR: exclude `Invalid`.
- Do not use unsupported significance, proof, or causal language.
- Keep charts editable and add `[Sources]` blocks to speaker notes.
- Output `D:\研究生\文献阅读\7.29组会\MCPTox_文献汇报_Persona主题_含复现实验.pptx`; do not overwrite the current deck.

---

### Task 1: Produce an Audited Experiment Dataset

**Files:**
- Create: `C:\Users\10701\Documents\7.29组会\.codex-ppt-build\reproduction-data.json`
- Read: `F:\idmDownlowd\inspect-evals-mcptox\results\comparison_50samples.md`
- Read: `F:\idmDownlowd\inspect-evals-mcptox\results\mimo-v2.5-300.json`
- Read: `F:\idmDownlowd\inspect-evals-mcptox\results\mimo-v2.5-pro-300.json`

**Interfaces:**
- Consumes: the 50-sample report and two arrays of 300 sample objects with `outcome`, `paradigm`, and `security_risk`.
- Produces: JSON with `small.overall`, `small.byParadigm`, `large.overall`, `large.byParadigm`, and `large.byRisk`.

- [ ] **Step 1: Write the metric assertions**

```powershell
$data = Get-Content -Raw 'C:\Users\10701\Documents\7.29组会\.codex-ppt-build\reproduction-data.json' | ConvertFrom-Json
if ($data.large.overall.mimo25.asr -ne 32.1) { throw 'MiMo v2.5 ASR mismatch' }
if ($data.large.overall.mimo25pro.asr -ne 37.1) { throw 'MiMo v2.5 Pro ASR mismatch' }
if ($data.large.byParadigm.P1.mimo25.asr -ne 43.8) { throw 'P1 MiMo ASR mismatch' }
if ($data.large.byParadigm.P3.mimo25pro.asr -ne 44.1) { throw 'P3 Pro ASR mismatch' }
```

- [ ] **Step 2: Run the assertions before generating the file**

Run the PowerShell block above.

Expected: FAIL because `reproduction-data.json` does not exist.

- [ ] **Step 3: Compute counts and valid-response ASR**

Create `reproduction-data.json` from raw outcomes using:

```text
valid = total - invalid
ASR = success / valid × 100
refusal_rate = refused / valid × 100
```

Include these verified overall counts:

```json
{
  "mimo25": {"total": 300, "success": 93, "invalid": 10, "valid": 290, "asr": 32.1},
  "mimo25pro": {"total": 300, "success": 106, "invalid": 14, "valid": 286, "asr": 37.1}
}
```

- [ ] **Step 4: Run the metric assertions**

Expected: PASS with no exception.

- [ ] **Step 5: Record the audit finding**

Add a `methodology` object stating that the 300-sample Markdown report used total samples for headline ASR and reversed the P1/P2 defense wording; all slide values come from raw JSON with Invalid excluded.

### Task 2: Build Slide 10 — 50-Sample Reproduction

**Files:**
- Create: `C:\Users\10701\Documents\7.29组会\.codex-ppt-build\add_reproduction_slides.mjs`
- Modify through export: `D:\研究生\文献阅读\7.29组会\MCPTox_文献汇报_Persona主题_含复现实验.pptx`

**Interfaces:**
- Consumes: `reproduction-data.json` and slide 10 of `MCPTox_文献汇报_Persona主题.pptx`.
- Produces: slide 10 with an editable grouped bar chart and concise interpretation.

- [ ] **Step 1: Import the existing deck**

```js
const presentation = await PresentationFile.importPptx(
  await FileBlob.load("D:/研究生/文献阅读/7.29组会/MCPTox_文献汇报_Persona主题.pptx"),
);
const slide10 = presentation.slides.items[9];
```

- [ ] **Step 2: Replace the table and placeholder copy**

Set the table to:

```text
样本：50 / 1,312
模型：MiMo v2.5、DeepSeek Chat、Qwen 3.7 Plus
指标：ASR、Refusal Rate
评估：Inspect AI；目标模型兼任 Judge（compat mode）
```

- [ ] **Step 3: Add the editable result chart**

Use `slide10.charts.add("bar", ...)` or `"column"` with categories `Overall`, `P1`, `P2`, `P3` and series:

```js
[
  { name: "MiMo v2.5", values: [6.2, 12.5, 0.0, 15.4] },
  { name: "DeepSeek Chat", values: [28.0, 70.0, 22.2, 7.7] },
  { name: "Qwen 3.7 Plus", values: [8.0, 0.0, 3.7, 23.1] },
]
```

Label the y-axis `ASR (%)`, use a 0–80 range, and preserve the blue-gray palette with one restrained accent color.

- [ ] **Step 4: Replace the conclusion block**

Use audience-facing copy:

```text
DeepSeek Chat 的总体 ASR 最高，并在 P1 上达到 70.0%。
Qwen 3.7 Plus 对 P1 未出现成功攻击，但在 P3 上达到 23.1%。
主动拒绝率整体较低；分组样本量仅 10/27/13，结果应视为探索性证据。
```

- [ ] **Step 5: Add source notes**

```text
[Sources]
- inspect-evals-mcptox/results/comparison_50samples.md
- 50 samples from the 1,312-instance valid benchmark set; target model also served as judge.
```

### Task 3: Build Slide 11 — 300-Sample Extension

**Files:**
- Modify: `C:\Users\10701\Documents\7.29组会\.codex-ppt-build\add_reproduction_slides.mjs`
- Modify through export: `D:\研究生\文献阅读\7.29组会\MCPTox_文献汇报_Persona主题_含复现实验.pptx`

**Interfaces:**
- Consumes: `large.overall`, `large.byParadigm`, and selected `large.byRisk`.
- Produces: slide 11 with corrected ASR values and limitations.

- [ ] **Step 1: Replace the experiment header**

```text
300 样本 × 2 模型扩展实验
MiMo v2.5 与 MiMo v2.5 Pro；ASR 分母排除 Invalid 响应
```

- [ ] **Step 2: Add the editable paradigm chart**

Use categories `Overall`, `P1`, `P2`, `P3` and:

```js
[
  { name: "MiMo v2.5", values: [32.1, 43.8, 22.5, 36.7] },
  { name: "MiMo v2.5 Pro", values: [37.1, 33.3, 28.7, 44.1] },
]
```

Use a 0–50 y-axis and data labels to one decimal place.

- [ ] **Step 3: Replace conclusions and risk callouts**

```text
Pro 总体 ASR 更高，但差异随攻击范式变化：P1 更低，P2/P3 更高。
Credential Leakage：27.5% vs 44.2%；Service Disruption：42.9% vs 53.2%（有效响应口径）。
两模型拒绝率均低于 2%，说明主动安全响应仍然有限。
```

- [ ] **Step 4: Add limitations**

```text
限制：目标模型兼任 Judge，可能引入自评偏差；50 与 300 样本实验的模型集合不同，不能据此解释样本规模导致的性能变化。
```

- [ ] **Step 5: Add source notes**

```text
[Sources]
- inspect-evals-mcptox/results/mimo-v2.5-300.json
- inspect-evals-mcptox/results/mimo-v2.5-pro-300.json
- ASR recomputed as Success / (Total - Invalid).
```

### Task 4: Export, Render, Verify, and Version

**Files:**
- Create: `D:\研究生\文献阅读\7.29组会\MCPTox_文献汇报_Persona主题_含复现实验.pptx`
- Create under ignored build directory: rendered PNGs and layout JSON.

**Interfaces:**
- Consumes: the edited in-memory presentation.
- Produces: verified PPTX and a Git commit on `main`.

- [ ] **Step 1: Export the deck**

```js
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save("D:/研究生/文献阅读/7.29组会/MCPTox_文献汇报_Persona主题_含复现实验.pptx");
```

- [ ] **Step 2: Render all 12 slides**

Export every slide as PNG and layout JSON under:

```text
C:\Users\10701\Documents\7.29组会\.codex-ppt-build\reproduction-final-render
```

Expected: exactly 12 PNG files and 12 layout JSON files.

- [ ] **Step 3: Inspect slides 10 and 11 at full size**

Verify:

```text
no title wrapping
no clipped chart labels
no overlap between charts and conclusion blocks
ASR values match reproduction-data.json
limitations are legible
```

- [ ] **Step 4: Run package-level validation**

Check that the PPTX contains 12 slide XML files, no empty structural placeholders, and no remaining `待填写` text. Fail validation if any condition is unmet.

- [ ] **Step 5: Commit the new deck**

```powershell
git add -- 'MCPTox_文献汇报_Persona主题_含复现实验.pptx'
git commit -m 'Add MCPTox reproduction results'
```

- [ ] **Step 6: Push after verification**

```powershell
git push origin main
```

Expected: remote `main` advances to the new commit and the working tree is clean.
