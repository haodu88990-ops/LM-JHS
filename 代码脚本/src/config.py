"""
配置加载器 — 费率表格式与产品参数完全分离。

- FormatProfile: 纯 Excel 布局（与产品无关）
- ProductProfile: 纯产品参数（与费率表格式无关）
"""

import re
import yaml
from typing import Any, Optional
from dataclasses import dataclass, field


# ============================================================
# 费率表格式配置（纯 Excel 布局，不涉及任何产品）
# ============================================================

@dataclass
class RateLayoutConfig:
    """费率表布局参数"""
    # "column": 每列独立(计划,保险期间,交费期间,性别)各占一行表头
    # "grid":   行式网格，pay_period + gender 各占一行表头，age 在固定列
    layout_type: str = "column"

    # ---- column 布局 ----
    header_rows: dict = field(default_factory=lambda: {
        "plan": 4, "period": 5, "pay_period": 6, "gender": 7
    })

    # ---- grid 布局 ----
    pay_period_row: Optional[int] = None
    gender_row: Optional[int] = None

    # ---- 通用 ----
    data_start_row: int = 8
    data_end_row: Optional[int] = None     # None = 自动检测（到 max_row）
    age_column: int = 1
    rate_columns_start: int = 2
    rate_columns_end: Optional[int] = None  # None = 自动检测（到 max_column）
    period_override: Optional[str] = None   # 全局保险期间覆盖（如 "终身"）


@dataclass
class SheetSection:
    """
    费率表中的一个读取区段。

    同一物理 Sheet 可以拆成多个区段（如标准体在 cols 2-29，优选体在 cols 32-59）。
    """
    sheet: str                           # Sheet 名称
    label: str                           # 显示标签（如"标准体""方案一"）
    ensure_plan: str = "1"               # API 用的 ensurePlan 代码
    plan_override: Any = None            # 覆盖责任计划值
    column_start: Optional[int] = None   # 限定列范围
    column_end: Optional[int] = None
    plan_map: dict = field(default_factory=dict)  # 表头值→计划编号


@dataclass
class FormatProfile:
    """
    纯费率表格式定义 — 与任何产品无关。

    描述 Excel 的物理布局和要读取的 Sheet/区段。
    """
    format_name: str = ""
    format_type: str = "column"          # "column" | "grid"
    layout: RateLayoutConfig = field(default_factory=RateLayoutConfig)
    sections: list[SheetSection] = field(default_factory=list)

    @classmethod
    def from_yaml(cls, path: str) -> "FormatProfile":
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        if not isinstance(data, dict):
            raise ValueError(f"格式配置文件格式错误: {path}")

        layout_data = data.get("layout", {})
        layout = RateLayoutConfig(
            layout_type=data.get("format_type", layout_data.get("layout_type", "column")),
            header_rows=layout_data.get("header_rows", {
                "plan": 4, "period": 5, "pay_period": 6, "gender": 7
            }),
            pay_period_row=layout_data.get("pay_period_row"),
            gender_row=layout_data.get("gender_row"),
            data_start_row=layout_data.get("data_start_row", 8),
            data_end_row=layout_data.get("data_end_row"),
            age_column=layout_data.get("age_column", 1),
            rate_columns_start=layout_data.get("rate_columns_start", 2),
            rate_columns_end=layout_data.get("rate_columns_end"),
            period_override=layout_data.get("period_override"),
        )

        sections = []
        for s in data.get("sections", []):
            sections.append(SheetSection(
                sheet=s.get("sheet", ""),
                label=s.get("label", ""),
                ensure_plan=str(s.get("ensure_plan", "1")),
                plan_override=s.get("plan_override"),
                column_start=s.get("column_start"),
                column_end=s.get("column_end"),
                plan_map=s.get("plan_map", {}),
            ))

        # 如果没有定义 sections，用默认单 Sheet
        if not sections and data.get("sheet"):
            sections.append(SheetSection(
                sheet=data["sheet"],
                label=data.get("label", "标准体"),
                ensure_plan=str(data.get("ensure_plan", "1")),
                plan_override=data.get("plan_override"),
                column_start=data.get("column_start"),
                column_end=data.get("column_end"),
            ))

        return cls(
            format_name=data.get("format_name", ""),
            format_type=layout.layout_type,
            layout=layout,
            sections=sections,
        )


# ============================================================
# 产品配置（纯产品参数，不涉及费率表格式）
# ============================================================

@dataclass
class APIConfig:
    base_url: str = ""
    login_endpoint: str = "/broker/api/user/login.html"
    age_rate_endpoint: str = "/broker/api/prospectus/saveCustomer.html"
    plan_rate_endpoint: str = "/broker/api/prospectus/saveProductExt.html"
    verify_ssl: bool = False
    timeout: int = 30


@dataclass
class CredentialsConfig:
    account: str = ""
    password_md5: str = ""


@dataclass
class ProductIdConfig:
    product_id: str = ""
    company_id: str = ""
    # serial_no 和 proposal_id 由用户运行时提供，不存配置


@dataclass
class DefaultsConfig:
    insurant_id: int = 85320
    insurant_occ_level: int = 1
    insurant_social_insurance: str = "1"
    policy_holder_id: int = 85321
    policy_holder_age: int = 30
    policy_holder_sex: str = "1"
    dividend_draw_type: str = "2"
    request_type: str = "prospectus"


@dataclass
class MappingsConfig:
    gender: dict = field(default_factory=lambda: {"男": "1", "女": "2"})
    ensure_period: dict = field(default_factory=dict)
    pay_mode: dict = field(default_factory=lambda: {"single": "1", "annual": "5"})
    pay_period: Any = "direct"


@dataclass
class PlansConfig:
    type: str = "simple"
    duties: dict = field(default_factory=dict)
    plan_duties: dict = field(default_factory=dict)


@dataclass
class AgeLimitsConfig:
    preferred: Optional[dict] = None
    by_pay_period: dict = field(default_factory=dict)


@dataclass
class AmountConfig:
    fixed: Optional[int] = None
    min: int = 1000000
    max: int = 5000000
    step: int = 1000


@dataclass
class ThrottleConfig:
    interval: int = 5
    sleep: float = 0.3
    workers: int = 10


@dataclass
class TestConfig:
    amount: AmountConfig = field(default_factory=AmountConfig)
    tolerance: float = 0.01
    throttle: ThrottleConfig = field(default_factory=ThrottleConfig)


@dataclass
class OutputConfig:
    boundary_summary: str = "边界值汇总.xlsx"
    test_report: str = "API测试报告_{timestamp}.xlsx"


class ProductProfile:
    """
    产品配置档案 — 纯产品参数，与费率表格式无关。

    用法:
        profile = ProductProfile.from_yaml("products/my_product.yaml")
        client = InsuranceAPIClient(profile)
    """

    def __init__(self, data: dict):
        self._raw = data
        self.product_name: str = data.get("product_name", "未命名产品")

        self.api = APIConfig(**data.get("api", {}))
        self.credentials = CredentialsConfig(**data.get("credentials", {}))
        self.product = ProductIdConfig(**data.get("product", {}))
        self.defaults = DefaultsConfig(**data.get("defaults", {}))
        self.mappings = MappingsConfig(**data.get("mappings", {}))
        self.plans = PlansConfig(**data.get("plans", {}))
        self.age_limits = AgeLimitsConfig(**data.get("age_limits", {}))
        test_raw = data.get("test", {})
        amount_raw = test_raw.get("amount", {})
        amount = AmountConfig(**amount_raw) if isinstance(amount_raw, dict) else amount_raw
        throttle_raw = test_raw.get("throttle", {})
        throttle = ThrottleConfig(**throttle_raw) if isinstance(throttle_raw, dict) else throttle_raw
        self.test = TestConfig(
            amount=amount,
            tolerance=test_raw.get("tolerance", 0.01),
            throttle=throttle,
        )
        self.output = OutputConfig(**data.get("output", {}))

    @classmethod
    def from_yaml(cls, path: str) -> "ProductProfile":
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        if not isinstance(data, dict):
            raise ValueError(f"产品配置文件格式错误: {path}")
        return cls(data)

    @classmethod
    def from_dict(cls, data: dict) -> "ProductProfile":
        return cls(data)

    # ---- URL 属性 ----

    @property
    def login_url(self) -> str:
        return f"{self.api.base_url}{self.api.login_endpoint}"

    @property
    def age_rate_url(self) -> str:
        return f"{self.api.base_url}{self.api.age_rate_endpoint}"

    @property
    def plan_rate_url(self) -> str:
        return f"{self.api.base_url}{self.api.plan_rate_endpoint}"

    # ---- 代码映射 ----

    def get_gender_code(self, label: str) -> str:
        return self.mappings.gender.get(label, label)

    def get_ensure_period_code(self, label: str) -> str:
        """保险期间文本 → API 编码。先查映射表，再智能提取。

        - 岁满型「至...年满60周岁...」→ TO60
        - 年满型「30年」→ 30
        - 终身「终身」→ TO105（需映射表配置）
        """
        # 1) 直接映射
        mapped = self.mappings.ensure_period.get(label)
        if mapped:
            return mapped

        # 2) 岁满型：提取年龄数字 → TO{age}
        m = re.search(r'年满(\d+)\s*周?岁', label)
        if m:
            return f"TO{m.group(1)}"

        # 3) 年满型：纯数字+年 → 数字
        m = re.match(r'^(\d+)\s*年?$', label.strip())
        if m:
            return m.group(1)

        return label

    def get_pay_period_code(self, pay_period) -> str:
        mp = self.mappings.pay_period
        if mp == "direct":
            return str(pay_period)
        elif isinstance(mp, dict):
            return str(mp.get(str(pay_period), pay_period))
        return str(pay_period)

    def get_pay_mode_code(self, pay_period) -> str:
        pay = int(pay_period) if str(pay_period).isdigit() else pay_period
        if pay == 1 or str(pay) in ("1", "趸交"):
            return self.mappings.pay_mode.get("single", "1")
        return self.mappings.pay_mode.get("annual", "5")

    def normalize_pay_period(self, raw_value) -> Optional[int]:
        """将费率表中的交费期间原始值标准化为整数"""
        if raw_value is None:
            return None
        if isinstance(raw_value, (int, float)):
            return int(raw_value)
        s = str(raw_value).strip()
        if not s:
            return None
        try:
            return int(s)
        except ValueError:
            pass
        mp = self.mappings.pay_period
        if isinstance(mp, dict) and s in mp:
            return int(mp[s])
        if s in ("趸交", "一次性交清"):
            return 1
        m = re.match(r'^(\d+)年?$', s)
        if m:
            return int(m.group(1))
        return None

    def get_duties_for_plan(self, plan_index: int) -> list:
        if self.plans.type == "simple" or not self.plans.plan_duties:
            return []
        codes = self.plans.plan_duties.get(plan_index, [])
        if isinstance(codes, str):
            codes = [codes]
        result = []
        for code in codes:
            code_str = str(code)
            if code_str in self.plans.duties:
                d = self.plans.duties[code_str]
                item = {
                    "dutyCode": code_str,
                    "dutyName": d.get("name", ""),
                    "order": d.get("order", 0),
                    "defaultValue": d.get("default_value", code_str),
                }
                for k, v in d.items():
                    if k not in ("name", "order", "default_value"):
                        item[k] = v
                result.append(item)
        return result

    def plan_has_description(self, plan_index: int) -> str:
        if self.plans.type == "simple" or not self.plans.plan_duties:
            return "基础保障"
        codes = self.plans.plan_duties.get(plan_index, [])
        if isinstance(codes, str):
            codes = [codes]
        if not codes:
            return "仅身故保险金"
        names = []
        for code in codes:
            code_str = str(code)
            if code_str in self.plans.duties:
                names.append(self.plans.duties[code_str].get("name", code_str))
            else:
                names.append(f"责任{code_str}")
        return " + ".join(names)

    def get_age_limit(self, pay_period, ensure_plan: str = "1") -> Any:
        from .config import AgeLimitsConfig
        limit = type('AgeLimit', (), {"min": 0, "max": 100})()
        pay_key = str(pay_period)
        if pay_key in self.age_limits.by_pay_period:
            entry = self.age_limits.by_pay_period[pay_key]
            if isinstance(entry, dict):
                limit.min = entry.get("min", 0)
                limit.max = entry.get("max", 100)
            else:
                limit.max = int(entry)
        if ensure_plan == "2" and self.age_limits.preferred:
            pref = self.age_limits.preferred
            limit.min = max(limit.min, pref.get("min", 18))
            limit.max = min(limit.max, pref.get("max", 65))
        return limit

    def __repr__(self) -> str:
        return f"ProductProfile({self.product_name})"
