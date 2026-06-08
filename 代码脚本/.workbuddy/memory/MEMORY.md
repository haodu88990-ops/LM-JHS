# 项目长期记忆 — 瑞泰鸿利传世（致享版）终身寿险 API 测试

## 产品信息
- 产品名称：瑞泰鸿利传世（致享版）终身寿险（分红型）
- productId: 991452, companyId: 100080

## API 系统
- 基地址：`https://kfzxtb.lmbaoxian.com:13080`
- 登录账号：15856990088 / dc483e80a7a0bd9ef71d8cf973673924
- 接口串联流程：login → age_rate(saveCustomer) → plan_rate(saveProductExt)
- **必须按顺序调用**，plan_rate 依赖 age_rate 返回的动态参数

## 关键传值格式（已验证正确）

### age_rate 接口（/broker/api/prospectus/saveCustomer.html）
| 参数 | 格式 | 说明 |
|:---|:---|:---|
| insurantSex | `"1"`=男, `"2"`=女 | 性别代码 |
| insurantAge | 整数 | 被保人年龄 |
| insurantOccLevel | 1 | 职业等级 |
| insurantSocialInsurance | `"1"` | 社保标志 |
| serialNo | 字符串 | 初始值 "1491455820292947968"，后续动态更新 |

**age_rate 响应关键提取**：
- `info.proposalPlanVOList[0].serialNo` → 更新 serialNo
- `info.proposalPlanVOList[0].proposalId` → 更新 proposalId（**不能硬编码**）
- `info.productExtVOList` 包含因子信息（ensurePlan 等）

### plan_rate 接口（/broker/api/prospectus/saveProductExt.html）
| 参数 | 格式 | 说明 |
|:---|:---|:---|
| ensurePlan | `"1"`=标准体, `"2"`=优选体 | **动态传入**，从用例的 ensurePlan 字段获取 |
| ensurePeriodCode | `"TO105"` | 终身 |
| payPeriodCode | `"1"/"3"/"5"/"10"/"15"/"20"/"30"` | 交费期间（数字字符串） |
| payModeCode | 交1年=`"1"`(一次交清), 其他=`"5"`(年交) | 交费方式 |
| serialNo | 从 age_rate 动态获取 | **不能硬编码** |
| proposalId | 从 age_rate 动态获取 | **不能硬编码** |
| dutyOptionList | 根据责任计划构建 | 见下方映射 |

### 责任计划 → dutyOptionList 映射
| 计划 | 包含的责任 code | 责任名称 |
|:---|:---|:---|
| 0 | [] | 仅身故保险金 |
| 1 | ["2"] | +全残保险金 |
| 2 | ["3"] | +身故额外关爱保险金 |
| 3 | ["3","2"] | +身故额外关爱+全残 |
| 4 | ["4"] | +意外身故额外保险金 |
| 5 | ["4","2"] | +意外身故额外+全残 |
| 6 | ["4","3"] | +意外身故额外+身故额外关爱 |
| 7 | ["4","3","2"] | +意外身故额外+身故额外关爱+全残 |

dutyOptionList 中每项结构：
```json
{
    "factorType": "dutyOption", "factorName": "责任名称",
    "isDisplay": "1", "order": 115218, "factorLevel": 1,
    "code": "2", "type": "dutyOption", "value": "2",
    "desc": "责任名称", "dataType": "5", "appFactorList": [],
    "defaultValue": "2", "appfactorCode": "2", "premium": ""
}
```
注：code=3 的 defaultValue 是数组 ["3"]，code=4 同理

## 费率验证公式
- 期望保费 = 保额 × 期望费率(‰) / 1000
- 与 API 返回的 fee 比对，允许 1% 误差（实际偏差 0%）

## 数据源
- **费率表原始文件**：`f:\workbuddy\test_Jhs\费率表.xlsx`
  - Sheet「标准体费率表」：行4=计划(0-7), 行5=保险期间, 行6=交费期间, 行7=性别, 行8-77=年龄数据
  - Sheet「优选体费率表」：同上结构，费率不同，承保年龄须满足 18~65 周岁
- **边界值汇总文件**（推荐直接用）：`f:\workbuddy\test_Jhs\瑞泰鸿利致享版边界值汇总.xlsx`（Sheet: 边界费率汇总）
  - 10列结构：保障方案 | ensurePlan | 责任计划 | 保险期间 | 交费期间(年) | 性别 | 最小年龄 | 最小年龄费率 | 最大年龄 | 最大年龄费率
  - 224行数据（2体型×8计划×7交费期间×2性别），展开为448条测试用例
  - ensurePlan: "1"=标准体, "2"=优选体
  - 生成脚本：`f:\workbuddy\test_Jhs\generate_boundary.py`

## batch_test.py 核心逻辑说明（最新版）

### 数据读取（load_test_cases）
- 从 `瑞泰鸿利致享版边界值汇总.xlsx` 的「边界费率汇总」Sheet 读取，从第4行开始
- 10列依次：col1=保障方案, col2=ensurePlan, col3=责任计划, col4=保险期间, col5=交费期间, col6=性别, col7=最小年龄, col8=最小年龄费率, col9=最大年龄, col10=最大年龄费率
- 每行展开为 2 条测试用例（最小年龄 + 最大年龄），共 448 条

### 接口调用流程（run_single_case）
1. `age_rate`（saveCustomer）：传入年龄、性别，动态提取 serialNo 和 proposalId
2. `plan_rate`（saveProductExt）：传入 ensurePlan（动态）、责任计划→dutyOptionList、交费期间等
3. 验证：`fee ≈ 保额 × 期望费率(‰) / 1000`，允许 1% 误差

### 报告生成（generate_report）
- 表头 24 列，包含：保障方案、ensurePlan、责任计划、责任描述、保险期间/Code、交费期间/Code、交费方式Code、性别/Code、年龄、年龄类型、保额、期望费率/保费、age_rate状态码/结果、plan_rate状态码/结果、API返回fee、failureReason、测试结论
- Sheet2「汇总统计」：按保障方案、责任计划、交费期间、性别、年龄类型分维度汇总
- 文件名格式：`API测试报告_YYYYMMDD_HHMMSS.xlsx`（每次运行自动生成不重复的文件名）

## 承保年龄范围
| 交费期间 | 最小年龄 | 最大年龄 |
|:---|:---|:---|
| 1年 | 0 | 69 |
| 3年 | 0 | 65 |
| 5年 | 0 | 62 |
| 10年 | 0 | 62 |
| 15年 | 0 | 60 |
| 20年 | 0 | 59 |
| 30年 | 0 | 50 |

## 踩坑经验
- `proposalId` 必须从 age_rate 响应动态获取，硬编码会导致"流水号编码未查到相关信息"
- `ensurePlan` 不能传 "0"，必须传 "1"（标准体）或 "2"（优选体），否则报"保费公式有问题"
- 优选体（ensurePlan="2"）年龄须满足 18~65 周岁，超出范围报"被保人年龄有误，投保优选体年龄须满足18～65周岁"；0岁/超65岁属于预期内FAIL，不是Bug
- age_rate 响应中 proposalId 是数字类型（如 40297），需要 str() 转换
- payModeCode 和 payPeriodCode 要匹配：1年交→payModeCode="1"(一次交清)，其他→payModeCode="5"(年交)
- 优选体费率表中，0岁对应的费率（col7/col9）若存在，其实是无效用例，API会拒绝；边界值汇总中的0岁优选体用例调用会FAIL（正常现象）
