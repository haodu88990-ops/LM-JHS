# Vibe Coding 流程 — 保费测算工具

> 结合本项目 `代码脚本/` 的实际代码、架构、已知缺陷写的。不是通用指南，是这个项目跟 AI（Codex）配合的实操手册。

---

## 一、项目两条路径，先说清楚你用哪条

这个项目有两种完全不同的解析+测算方式，跟 AI 说需求前先定好用哪条：

| | 路径 A：内容驱动（自动检测） | 路径 B：格式驱动（YAML 定义） |
|---|---|---|
| **入口** | `web_app.py`（上传即解析） / `run.py` | `run_tests.py` / `generate_boundary.py` |
| **核心模块** | `src/rate_parser.py` | `src/rate_reader.py` + `formats/*.yaml` |
| **配置** | 无格式配置，全靠 `RateTableParser` 统计推断 | 需要 `formats/column.yaml` 或 `formats/grid.yaml` |
| **适用场景** | 随便丢个费率表进去试 | 同一产品反复测，格式固定 |
| **解析准确性** | 看运气（表头文本正则匹配 + 年龄列统计推断） | 你说了算（表头行号、区段范围全手动配） |

> ⚠ 已知缺陷：路径 A 的 `_find_first_age_sequence` 只找 ≤1 岁起的数据列（Bug #12），纯成人产品（18岁起）的费率表自动解析会扑空。遇到这种情况切路径 B 手动配格式。

---

## 二、数据流全链路（贴着你代码的）

```
 费率表 (.xlsx / .pdf)
     │
     ▼
 ┌─ 路径 A：parse_rate_table(file_path) → RateTableParser.parse()
 │    • _read_metadata()         ← 读"产品信息"Sheet
 │    • _find_first_age_sequence() ← 扫描各列找年龄序列（上限30列）
 │    • _detect_sections()       ← 按年龄断点拆分数据区段
 │    • _infer_header_hierarchy() ← 统计推断表头维度（交费期间/性别/保险期间）
 │    • _extract_section_data()  ← 逐行读费率（⚠ 空年龄行会截断，Bug #4）
 │
 └─ 路径 B：reader.read_all_sections(rate_file) → RateTableReader
       • 按 format_profile.sections 定义的区段读
       • _extract_column() / _extract_grid()
       • 依赖你在 column.yaml/grid.yaml 里配的行号
     │
     ▼
 write_boundary_xlsx(rows, tmp_boundary)  →  边界值汇总 Excel（10列标准格式）
     │
     ▼
 TestCaseLoader.load(tmp_boundary)  →  用例列表（每行产出 1~2 条用例）
     │
     ▼
 InsuranceAPIClient.login()  →  POST /user/login.html
     │
     ▼
 每个用例（并行，默认20个线程）：
   1. save_customer(age, gender)  →  POST /saveCustomer.html
   2. save_product(plan, period, pay_period, amount)  →  POST /saveProductExt.html
       │  （如果 failureReason 含金额限制关键词 → 自动调金额重试，最多10次）
       ▼
   比较 API 返回 fee 与 期望保费（费率 × 保额 / 1000），tolerance=0.01 判定 PASS/FAIL
     │
     ▼
 ReportGenerator.generate(results, output_path)  →  Excel 报告（明细 + 汇总统计）
```

---

## 三、跟 AI 说需求的套话模板

### 场景 1：发现 Bug

```
「web_app 上传 xlsx 后点解析，页面报错『费率表解析失败』，
  控制台输出 traceback 是 ...
  费率表是瑞泰鸿利传世致享版那个」
```

AI 会先去读 `rate_parser.py` 里解析流程，找到异常抛出的位置。你不用猜是哪行的问题。

### 场景 2：想改功能

```
「run_tests.py 现在登录失败会继续跑，结果全是 FAIL，
  改成登录失败就直接退出，提示用户检查账号配置」
```

这就是上面 Bug #9，AI 已经知道在哪改了。

### 场景 3：想加验证方式

```
「在 tester.py 的 _run_single 里加个判断：
  如果 api.save_product 返回 fee=0 且 code=200 但没有 failureReason，
  不要判 PASS，判 FAIL」
```

这是 Bug #3，AI 已经在你上次审查时看到这块代码了。

### 场景 4：想排查配置问题

```
「帮我看看 ruitai.yaml 的 duty "3" 的 default_value 为啥是数组不是字符串，
  跟 duty "2" 和 "4" 不一致，会不会有问题」
```

---

## 四、这个项目踩过的坑（给你和 AI 都省时间）

以下是在你的代码审查报告里确认过的实际问题，写在这里省得重复排查：

### 崩得最狠的

- **`run.py` 的 `case_count` 未定义**：Phase 2 直接 NameError 炸掉，走不到测算。这个变量 `parsed["case_count"]` 存在但没取出来。（Bug #1）
- **`boundary.py` 的 `read_grouped()` 不存在**：`--grouped` 参数一用就 AttributeError，功能半成品。（Bug #2）

### 结果不对但没崩的

- **`api_client.save_product` 把 fee=0 当成成功**：code=200 + fee=0 + 无 failureReason 时 `success=True`，报告里多条 PASS 但其实保费计算失败。（Bug #3）
- **`rate_parser._extract_section_data` 空年龄行截断**：某行年龄列为空但费率列有值，循环 break，后面整段数据丢失。（Bug #4）

### 上传和报告的问题

- **报告保存位置不可控**：前端用 `currentFile.path` 传路径，标准浏览器没有这个属性，报告写到工作目录的父目录去。（Bug #6）
- **`file://` 下载链接打不开**：主流浏览器禁止从 HTTP 页面开 file:// 链接，点了没反应。（Bug #7）

### 两个解析器打架的地方

- `rate_parser.py` 里 `_read_metadata` 从产品信息 Sheet 读 `data_type`，之后 `_extract_metadata` 用正则 `每\d+元.*保险费` 给覆盖了。如果你产品信息 Sheet 写了 `data_type: 1` 但表头有"保险费"三个字，算费方向会翻车。
- `rate_parser._find_first_age_sequence` 搜列上限 30 列、起始年龄 ≤1 岁，边界情况很可能认不出年龄列。

---

## 五、不同需求的对话范例

### 范例：修一个 Bug

```
你：
"web_app 上传费率表后点解析报错，控制台说"费率表解析失败""

AI：
（读 web_app.py 和 rate_parser.py 找到报错路径，定位问题）

你：
"好修吧"

AI：
（改代码，改完跑一个上传 → 解析的主流程，确认不报错）
```

### 范例：加一个新费率表格式

```
你：
"我要加一个中英爱永恒的费率表，但它的布局跟 column.yaml 和 grid.yaml 都不一样"

AI：
"能不能用 explore.py 跑一下那个费率表，让我看前10行的结构"

你：
（跑 `python explore.py 费率表.xlsx` 把输出贴过来）

AI：
"看懂了，它的表头在第 5-8 行，年龄在 B 列，但每 10 列一组方案。
  建议新建一个 formats/aiyongheng_custom.yaml，配 sections 带 column_start/column_end"
```

### 范例：排查一条 FAIL 的原因

```
你：
"跑完测算有 3 条 FAIL，我看报告里 failureReason 是空的"

AI：
"查一下 _uploads/_debug_failure.json，那里面存了最后一次失败的请求 payload 和响应 body"

你：
（贴内容）

AI：
"看到了，API 返回了 `code: 500, message: '费率计算异常'`。
  应该是 age_rate 传的参数不对，我看看 tester._run_single 里传了什么..."
```

---

## 六、验证清单（改完一样勾一样）

每次改完代码，按这个顺序跑通才算完事：

```
□ 上传 xlsx → 解析成功 → 预览区展示边界值
□ 输入 serialNo 点测算 → SSE 进度条走到 100%
□ 报告 Excel 能打开，明细 Sheet 有数据
□ 汇总 Sheet 各维度统计正确
□ 删掉上传文件再重新走一遍（确认无残留状态）
```

改的范围不同需要额外验证的：

| 改了 | 额外验证 |
|------|---------|
| `rate_parser.py` | 上传两种以上不同格式的 xlsx 看解析结果 |
| `rate_reader.py` | 用对应 format yaml 跑一次 `run_tests.py` |
| `api_client.py` | 看 `_debug_failure.json` 确认请求体正确 |
| `tester.py` | 跑 3 条用例停掉，看部分结果是否正常输出 |
| `reporter.py` | 打开生成的 xlsx 检查条件格式和冻结窗格 |
| `web_app.py` | 刷新页面重新上传，确认 SSE 能连接 |

---

## 七、协作原则（省你时间用的）

1. **先说清楚走路径 A 还是路径 B**，不然 AI 可能读了两个模块还没决定改哪个
2. **一次只说一件事**。修完一个 bug、确认通过了，再说下一个。并行三个需求 AI 上下文容易串
3. **给出业务判断**。比如算费方向、年龄范围、责任码值这些领域知识 AI 猜不准，你告诉它它就不用花时间翻代码猜了
4. **改完跑主流程**。AI 改了 `rate_parser.py` 之后不会自动触发上传页面帮你点，你要手动走一遍"上传 → 解析 → 测算 → 报告"确认没断

---

*本流程基于 `代码脚本/` 目录下的实际代码和已知缺陷编写，与代码审查报告（codex检查/保费测算工具代码审查报告_20260610.md）对照阅读效果更好。*
