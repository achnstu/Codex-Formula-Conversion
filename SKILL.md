---
name: Markdown裸变量转Word原生公式
description: >
  中文：把 Markdown 里的"裸文本变量/伪公式"（如 P_sc,k=0、a_k in [0,1]、g_k、c_j、
  alpha、eta_dc、fc_smoothness、ΔP_fc,k）包成规范 LaTeX，再用 pandoc 渲染成 Word
  原生公式（OMML），使变量在 .docx 里显示为真正的数学符号而非纯文本。默认保留 Word
  原生 OMML；当用户明确需要 MathType 格式时，先询问确认，再把 OMML 批量转为 MathType。
  涵盖公式块、公式说明段、表头、行内提及的全覆盖识别与转换，以及转换后的验证。
  触发词："裸变量转公式"、"转原生公式"、"公式变原生公式"、"OMML"、"pandoc公式"、
  "符号没渲染"、"美元符号没渲染"、"变量包LaTeX"、"$...$转Word"、"转MathType"、
  "自动转mathtype"、"OMML转MathType"。
---

# Markdown 裸变量 → Word 原生公式（OMML）

作者：集美大学 罗强

把 Markdown 底稿里散落的裸文本变量包成 LaTeX，再用 pandoc 转成 Word 原生公式。目标：`.docx` 里每个物理变量都是真正的数学符号（正体罗马、希腊字母、上下标、分数、根号），而不是 `P_sc,k` 或 `$F_{\mathrm{smooth}}$` 这样的纯文本。

默认输出 **Word 原生 OMML**。只有用户明确需要 MathType，且确认后，才运行 MathType 转换；未确认时不要自动转。

## 核心前提：为什么必须用 pandoc

- Node 的 `docx` 库管线**不渲染** `$...$`/`$$...$$`，公式只能以 Unicode 文本"看起来像"，`$F_{\mathrm{smooth}}$` 会原样显示美元符号和反斜杠。
- pandoc 把 LaTeX 公式转成 **Word 原生公式（OMML）**，`\mathrm{...}` 渲染成正体罗马，希腊字母、上下标、分数、求和、根号全部正确。
- 所以本 skill 的核心动作是：**(1) 把所有裸变量包成规范 LaTeX → (2) pandoc 转换为 OMML → (3) 可选、经确认后转 MathType**。

## 第一步：识别所有裸变量（全覆盖，别只看公式块）

裸变量藏在四类位置，**最易漏的是公式说明段**，不是公式块本身：

1. **公式块**内残留的 ASCII：`eta_dc`、`Delta t`、`alpha`、`sum_`、`<=`、`>=`、`in [0,1]`、`sqrt(...)`。
2. **公式说明段**："其中 X 为…""设 X 被划分为…""若 c_j 为…""连续动作 a_k in [0,1]""其中 N_k 为探索噪声"——这些句子里的 `P_L,k`、`a_k in`、`g_k`、`B_1,...,B_S`、`c_j`、`N_k`、`alpha` 都是裸变量。
3. **表头**：`H2 合计均值`、`fc_smoothness (W2)`、`SOC 终端误差`。
4. **行内提及**：正文中混写的 `ΔP_fc,k=P_fc,k-P_fc,k-1`、`E_b`、`c_H2`。

扫描命令（排除已正确的 `$...$` 内容、图片路径行）：

```bash
grep -nP '(?<!\$)\b([A-Za-z]+_\{?[a-z0-9,]+|eta_|Delta |alpha\b|gamma=|tau=|sqrt\(|fc_smoothness|H2 (raw|std|合计))' 底稿.md \
  | grep -v '!\[' | grep -v '图表文件夹'
```

逐行人工核对——区分"真裸变量"和"误报"（已在 `$...$` 内、或文末修订说明里作为字面引用的旧词）。

## 第二步：包成规范 LaTeX

### 编号公式（`$$...$$`）

- 编号**放公式内末尾**用 `\qquad (\text{N})`，**绝不用 `\tag{}`**（pandoc 2.9 不支持 `\tag`，会报错并把整块降级成裸 TeX，导致 `$`/`mathrm` 残留）。
- 例：

```latex
$$
P_{\mathrm{fc},k} = \mathrm{clip}\!\left( P_{\mathrm{fc},k}^{\mathrm{req}},\; P_{\mathrm{fc},k-1} - r_{\downarrow}\,\Delta t,\; P_{\mathrm{fc},k-1} + r_{\uparrow}\,\Delta t \right) \qquad (\text{4})
$$
```

### 行内变量（`$...$`）

公式说明段和表头里的裸变量逐个包成 `$...$`。例：

| 裸文本 | 规范 LaTeX |
|---|---|
| `P_sc,k=0` | `$P_{\mathrm{sc},k}=0$` |
| `a_k in [0,1]` | `$a_k \in [0,1]$` |
| `g_k` | `$g_k$` |
| `B_1,...,B_S` | `$B_1,\dots,B_S$` |
| `c_j` | `$c_j$` |
| `N_k` | `$N_k$` |
| `alpha` | `$\alpha$` |
| `ΔP_fc,k=P_fc,k-P_fc,k-1` | `$\Delta P_{\mathrm{fc},k}=P_{\mathrm{fc},k}-P_{\mathrm{fc},k-1}$` |
| `E_b`、`c_H2`、`η_fc` | `$E_{\mathrm{b}}$`、`$c_{\mathrm{H_2}}$`、`$\eta_{\mathrm{fc}}$` |

### LaTeX 书写规范

- 函数名/多字母算子用 `\mathrm{}`：`\mathrm{clip}`、`\mathrm{SOC}`、`\mathrm{req}`、`\max`、`\min`、`\log`。
- 下标里的多字母也要 `\mathrm`：`P_{\mathrm{fc},k}`、`\eta_{\mathrm{dc}}`、`m_{\mathrm{H_2}}^{\mathrm{corr}}`。
- 关系符/算子：`\le \ge \in \approx \sum \sqrt \nabla \times \cdot \frac{}{}`、`r_{\uparrow}`、`r_{\downarrow}`。
- 希腊字母：`\eta \Delta \gamma \tau \alpha \lambda \mu \pi \phi \sigma \epsilon \zeta \Phi`；上横线/帽 `\bar{\phi}`、`\hat{P}`。

### 替换怎么做最安全

大批量替换用**原子 Python**：读入 → 逐行精确替换（带 `assert old in line` 断言）→ 一次性写回。任一断言失败则整体不写，避免多次 Edit 工具的反复损坏窗口，也能立刻发现"这行已被前一批改过"。

## 第三步：pandoc 转换

图片中文相对路径经 `--resource-path` 常因编码解析失败。**最稳：cd 进 md 所在目录再转，让相对路径自然解析**：

```bash
cd /path/to/drafts_cn
pandoc 底稿.md -o /tmp/out.docx -f markdown -t docx \
  --resource-path=.:..:../图表文件夹:../../图表文件夹
```

若目标 docx 被 Word 占用（`permission denied` / 残留 `~$` 锁文件），先输出到 `/tmp` 再 `cp` 覆盖。

## 第四步：验证（每次转换后必跑）

```python
import docx, zipfile
d = docx.Document("out.docx")
xml = zipfile.ZipFile("out.docx").read("word/document.xml").decode()
print("OMML公式块:", xml.count("<m:oMath"))                 # 应 ≥ 编号公式数 + 行内公式数
print("$残留:", "$" in "".join(p.text for p in d.paragraphs)) # 应 False
print("mathrm残留:", "mathrm" in xml)                        # 应 False
```

文本层抽查裸变量是否清零（应全为 0）：

```bash
python3 -c "
import docx
t=''.join(p.text for p in docx.Document('out.docx').paragraphs)
for k in ['P_sc,k=0','a_k in','alpha 为','g_k ','c_j ','fc_smoothness','H2 合计','eta_']:
    print('OK' if k not in t else '⚠️残留', k)
"
```

注意：文末"修订说明/变更记录"里作为"旧→新"字面引用的旧词属正常，不算残留。

## 第五步：可选转 MathType（必须先问）

默认停止在 Word 原生 OMML。若用户说"需要 MathType"、"老师要求 MathType"、"转 MathType 格式"等，先问：

> 是否需要把 Word 原生公式（OMML）继续转换为 MathType 格式？确认后会生成一个新的 `_mathtype.docx`，原 OMML 文件保留。

只有用户明确确认后才执行。未确认、用户不需要、或环境不满足时，保留原生 OMML。

### MathType 转换前检查

Windows + Word + MathType + pywin32 都需要可用。先检测 MathType：

```powershell
cd "C:\Users\zwzhu\.myskills\skills\md-bare-var-to-native-equation"
py .\scripts\convert_omml_to_mathtype.py --check
```

若提示缺少 pywin32：

```powershell
py -m pip install pywin32
```

注意 Windows 中文路径：若输入文件路径含中文，先用 PowerShell 复制到 ASCII 路径再交给 Python：

```powershell
cd "C:\Users\zwzhu\Documents\Codex\..."
New-Item -ItemType Directory -Force -Path ".\work\mathtype" | Out-Null
Copy-Item -Path "C:\Users\zwzhu\Desktop\中文文件.docx" -Destination ".\work\mathtype\input.docx" -Force
```

Python 有时会把中文路径读成 `??`，导致误判文件不存在或复制出 0 字节文件。

### 执行转换

建议把输入输出写成绝对路径，避免 Word COM 找错文件：

```powershell
cd "C:\Users\zwzhu\.myskills\skills\md-bare-var-to-native-equation"
py .\scripts\convert_omml_to_mathtype.py "C:\path\to\out.docx" "C:\path\to\out_mathtype.docx"
```

转换脚本会打开 Word。不要在转换过程中手动操作 Word。转换完成后检查输出文件。

脚本调用 MathType 的 `MTCommand_ConvertEqns` 整篇转换宏，相当于手动打开"转换公式"窗口，勾选所有来源公式类型，并点击"转换"。它会按 Word 文档顺序批量转换所有 OMML 公式，不需要逐个公式点。

首次在新机器上使用时，先手动打开 MathType 的"转换公式"对话框，确认四个来源类型都勾选：

- `MathType 或 Equation Editor 公式`
- `Microsoft Word EQ 域`
- `MathType 转换文本公式`
- `Word 2007 及以上 (OMML) 公式`

目标选择：`MathType 公式 (OLE 对象)`。MathType 会记住该设置；之后脚本用 Enter 自动点"转换"和最后"确定"。

### MathType 转换后验证

确认三件事：

1. `_mathtype.docx` 已生成，原 OMML 文件未被覆盖。
2. Word 打开 `_mathtype.docx` 不报修复错误。
3. 随机双击 2-3 个公式，确认可由 MathType 编辑；若 MathType 没接管，交付原生 OMML 并说明自动转换失败原因。
4. 脚本输出里必须看到 `Output OLE objects > 0` 或 `Output Equation.DSMT4 markers > 0`，否则视为未真正转换，即使命令曾显示发送了 Alt+\。
5. 对批量转换，要求 `Initial OMML = Output OLE objects = Output Equation.DSMT4 markers`，且 `Remaining OMML = 0`。

## 反复踩的坑

- **`\tag{}` 不可用**（pandoc 2.9）→ 编号放公式内末尾 `\qquad (\text{N})`。
- **公式说明段的裸变量最易漏** → 单独扫一遍"其中/设/若 + 变量"句式。
- **挂载同步损坏**：本类工作区的 md 经 Edit 工具改后，挂载层会间歇性在文件**尾部插入 NUL 字节甚至截断尾部**。转换前用 `tr -cd '\000' < f | wc -c` 查真实 NUL 数（**不是 `grep -c`——那数的是行数会误读**）、`file` 查类型、`wc -l` + 看结尾 3 行确认无截断。检出 NUL 用 `tr -d '\000'` 原子 strip 后立即转换。
- **图片没嵌入**：pandoc 报 `PandocResourceNotFound` 且中文路径转义成 `\22270\34920...` → cd 进 md 目录解决，别只靠 `--resource-path`。
- **MathType 转换不要默认跑**：OMML 是稳定主输出；MathType 依赖 Word COM、MathType add-in、窗口焦点，失败率高。只有用户确认才转，且输出到新文件。
- **不要相信 SendKeys 成功提示**：`Alt+\` 发送成功不等于 MathType OLE 已生成。必须检查 docx 包内 `word/embeddings/` 或 `Equation.DSMT4` 标记。
- **批量转换用 MathType 整篇宏**：优先 `MTCommand_ConvertEqns`，不要逐个 OMath 发 `Alt+\`。成功标准：`Remaining OMML = 0`，且 `Output OLE objects` 数量等于原 OMML 数量。
- **公式变问号**：不要手写/拼接低配 OMML 做生产输入。MathType 对非标准 OMML 容错差，可能把不认识的节点转成 `?`。生产路线必须是规范 LaTeX → pandoc/Word 生成正规 OMML → MathType 转换。
- **0 字节 docx 不是转换问题**：先 `Get-Item` 看 `Length`。0 字节文件不能 ZIP 解析，也没有公式可转。
- **Word/WPS 占用文件**：转换前关闭目标 docx；输出到新文件。若文件被占用，先输出到临时 ASCII 路径，再复制回桌面。
- **开源前不要提交生成样例/缓存**：不要提交 `_mathtype.docx`、`~$*.docx`、`__pycache__/`、临时 `work/`、真实论文或用户文档。提交脚本、SKILL.md、最小 synthetic 测试即可。
