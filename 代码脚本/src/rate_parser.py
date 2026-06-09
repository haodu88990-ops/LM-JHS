"""
内容驱动的费率表解析器。

不从 Sheet 名或固定行列号出发，而是以「年龄列」为锚点，从数据内容
反推表格布局和维度映射。支持三种布局：

- title_grid : PDF 提取的多区段格式，性别/计划在标题文本中
- grid       : 行式网格，pay_period + gender 各占一行表头
- column     : 列式，每列独立维度组合（plan/period/pay_period/gender）
"""

import re
from typing import Optional, Any
from dataclasses import dataclass, field
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# ============================================================
# 维度值 → 标准化映射
# ============================================================

CN_PAY_MAP = {
    "趸交": 1, "一次性交清": 1,
    "三年": 3, "五年": 5, "六年": 6, "八年": 8, "十年": 10,
    "十二年": 12, "十五年": 15, "十八年": 18, "二十年": 20,
    "二十五年": 25, "三十年": 30,
}

CN_GENDER = {"男": "男", "女": "女", "M": "男", "F": "女", "m": "男", "f": "女"}


# ============================================================
# 数据结构
# ============================================================

@dataclass
class Region:
    """数据区段"""
    start_row: int          # 区段起始行（绝对行号）
    end_row: int            # 区段结束行（绝对行号）
    title_text: str = ""    # 标题行文本（如有）


@dataclass
class Layout:
    """检测到的布局结构"""
    layout_type: str = ""              # "title_grid" | "grid" | "column"
    age_col: int = 1                   # 年龄所在列
    rate_cols: list = field(default_factory=list)  # 费率列号列表
    data_start_row: int = 0            # 数据起始行（绝对）

    # column 布局：每列一个维度组合
    col_dims: dict = field(default_factory=dict)   # {col: {pay_period, gender, plan, period}}

    # grid / title_grid 布局
    pay_period_map: dict = field(default_factory=dict)  # {col: pay_period_int}
    gender: str = ""                   # grid: 从表头行读取, title_grid: 从标题提取
    plan_label: str = ""               # 保障方案标签
    period: str = "终身"               # 保险期间


# ============================================================
# 工具函数
# ============================================================

def _clean_int(val: Any) -> Optional[int]:
    if val is None:
        return None
    try:
        if isinstance(val, str):
            val = val.strip().rstrip("年")
            if not val:
                return None
        return int(float(val))
    except (ValueError, TypeError):
        return None


def _clean_float(val: Any) -> Optional[float]:
    if val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _norm_pay(val: Any) -> Optional[int]:
    """交费期间值 → 整数"""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return int(val)
    s = str(val).strip()
    if not s:
        return None
    try:
        return int(s)
    except ValueError:
        pass
    if s in CN_PAY_MAP:
        return CN_PAY_MAP[s]
    m = re.match(r'^(\d+)年?$', s)
    if m:
        return int(m.group(1))
    return None


def _is_pay_period(val: Any) -> bool:
    """判断值是否为交费期间"""
    return _norm_pay(val) is not None


def _is_gender(val: Any) -> bool:
    """判断值是否为性别"""
    if val is None:
        return False
    return str(val).strip() in CN_GENDER


def _is_period(val: Any) -> bool:
    """判断值是否为保险期间"""
    if val is None:
        return False
    s = str(val).strip()
    return "终身" in s or "至" in s or "定期" in s


def _classify_cell(val: Any) -> Optional[str]:
    """分类单元格的维度类型"""
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    if _is_pay_period(s):
        return "pay_period"
    if _is_gender(s):
        return "gender"
    if _is_period(s):
        return "period"
    if s.isdigit() and 0 <= int(s) <= 9:
        return "plan"
    return None


# ============================================================
# 解析器
# ============================================================

class RateTableParser:
    """内容驱动的费率表解析器。"""

    def parse(self, file_path: str) -> dict:
        """
        解析费率表文件，返回标准结果 dict。

        包含: format, product_name, data_type, fee_unit, fee_rule,
              dims, boundaries, case_count, boundary_count,
              _rows, rate_sheet, file_path
        """
        wb = load_workbook(file_path, data_only=True)
        sheets = wb.sheetnames

        # 1. 读元数据（如有"产品信息" Sheet）
        meta = self._read_metadata(wb)

        # 2. 找数据 Sheet
        data_sheets = self._find_data_sheets(wb)
        if not data_sheets and sheets:
            data_sheets = [sheets[-1]]

        # 3. 逐 Sheet 分段提取
        all_rows = []
        product_name = meta.get("product_name", "")

        for sn in data_sheets:
            ws = wb[sn]
            age_col = self._find_age_column(ws)
            if age_col is None:
                continue

            regions = self._segment_regions(ws, age_col)

            for region in regions:
                # 跳过 EM 加费 / 次标准体评点加费区段
                if self._is_em_surcharge(region.title_text):
                    continue

                layout = self._analyze_headers(ws, region, age_col)
                if layout is None:
                    continue
                rows = self._extract_region(ws, region, layout)
                all_rows.extend(rows)

                # 从标题提取产品名
                if not product_name and region.title_text:
                    pn = self._extract_product_name(region.title_text)
                    if pn:
                        product_name = pn

        wb.close()

        # 4. 组装结果
        return self._build_result(all_rows, product_name, meta, data_sheets,
                                  file_path)

    # ================================================================
    # 元数据
    # ================================================================

    def _read_metadata(self, wb) -> dict:
        """从'产品信息'/'产品说明' Sheet 读元数据"""
        meta = {"product_name": "", "data_type": "1", "fee_unit": 1000, "fee_rule": "四舍五入保留2位小数"}
        for name in ["产品信息", "产品说明"]:
            if name in wb.sheetnames:
                ws = wb[name]
                for row in range(1, ws.max_row + 1):
                    key = str(ws.cell(row=row, column=1).value or "").strip()
                    val = str(ws.cell(row=row, column=2).value or "").strip()
                    if key == "product_name":
                        meta["product_name"] = val
                    elif key == "data_type":
                        meta["data_type"] = val.split("_")[0] if "_" in val else val
                    elif key == "fee_unit":
                        try:
                            meta["fee_unit"] = int(float(val))
                        except ValueError:
                            pass
                    elif key == "fee_rule":
                        meta["fee_rule"] = val.split("_", 1)[1] if "_" in val else val
                break
        return meta

    # ================================================================
    # 数据 Sheet 发现
    # ================================================================

    def _find_data_sheets(self, wb) -> list[str]:
        """找含费率数据的 Sheet（排除纯元数据 Sheet）"""
        meta_names = {"产品信息", "产品说明", "info", "说明"}
        candidates = []
        for sn in wb.sheetnames:
            if sn in meta_names:
                continue
            ws = wb[sn]
            # 统计前 50 行的数值密度
            numeric = 0
            total = 0
            for row in range(1, min(ws.max_row + 1, 51)):
                for col in range(1, min(ws.max_column + 1, 20)):
                    total += 1
                    v = ws.cell(row=row, column=col).value
                    if v is not None:
                        try:
                            float(v)
                            numeric += 1
                        except (ValueError, TypeError):
                            pass
            if total > 0 and numeric / total > 0.05:
                candidates.append(sn)
        return candidates

    # ================================================================
    # 年龄列检测（核心锚点）
    # ================================================================

    def _find_age_column(self, ws: Worksheet) -> Optional[int]:
        """找年龄列 — 检测从 0 开始的连续整数序列。"""
        max_row = ws.max_row
        best_col = None
        best_len = 0

        for col in range(1, min(ws.max_column + 1, 6)):
            values = []
            rows = []
            for row in range(1, min(max_row + 1, 200)):
                v = _clean_int(ws.cell(row=row, column=col).value)
                values.append(v)
                rows.append(row)

            seq_len = self._measure_age_sequence(values)
            if seq_len > best_len:
                best_len = seq_len
                best_col = col

        return best_col if best_len >= 10 else None

    def _measure_age_sequence(self, values: list) -> int:
        """测量从 0 开始的最长连续递增序列。"""
        max_seq = 0
        in_seq = False
        expected = 0

        for v in values:
            if v is not None and v == expected:
                in_seq = True
                expected += 1
                max_seq = max(max_seq, expected)
            elif in_seq and v is not None and v == 0:
                # 重启序列
                expected = 1
                max_seq = max(max_seq, expected - 1)
            elif in_seq and v is None:
                pass  # 允许偶尔的空值，不中断
            else:
                if in_seq:
                    max_seq = max(max_seq, expected)
                in_seq = False
                expected = 0

        return max_seq

    # ================================================================
    # 区段分割
    # ================================================================

    def _segment_regions(self, ws: Worksheet, age_col: int) -> list[Region]:
        """将 Sheet 切分为数据区段。"""
        max_row = ws.max_row
        max_col = ws.max_column

        # 扫描所有行，标记区段边界
        boundaries = []  # [(row, type)]
        for row in range(1, max_row + 1):
            col1 = str(ws.cell(row=row, column=1).value or "").strip()
            # 检查 col 2+ 是否含交费期间（用于 grid 表头检测）
            pay_signals = 0
            for c in range(2, min(max_col + 1, 10)):
                if _is_pay_period(ws.cell(row=row, column=c).value):
                    pay_signals += 1

            # 标题行
            if ("费率表" in col1 or "加费表" in col1) and len(col1) > 15:
                boundaries.append((row, "title"))
            # 年龄表头行（且上方没有标题行）
            elif col1 in ("投保年龄", "年龄", "age", "Age", "AGE"):
                if not boundaries or boundaries[-1][0] < row - 3:
                    boundaries.append((row, "age_header"))
            # 交费期间表头行（grid 格式的首行）
            elif pay_signals >= 2 and col1 == "":
                prev_boundary = boundaries[-1][0] if boundaries else 0
                if row - prev_boundary > 5:  # 不是标题行紧邻的表头
                    boundaries.append((row, "pay_header"))

        # 从边界构建区段
        regions = []
        title_boundaries = [b for b in boundaries if b[1] in ("title", "pay_header")]

        if title_boundaries:
            for i, (brow, _) in enumerate(title_boundaries):
                end = (title_boundaries[i + 1][0] - 1
                       if i + 1 < len(title_boundaries) else max_row)
                title_text = str(ws.cell(row=brow, column=1).value or "")
                regions.append(Region(start_row=brow, end_row=end,
                                      title_text=title_text))
        else:
            # 整个 Sheet 只有一个区段
            # 找数据起始行
            data_start = 1
            for row in range(1, min(max_row + 1, 200)):
                age = _clean_int(ws.cell(row=row, column=age_col).value)
                if age is not None and age >= 0:
                    data_start = row
                    break
            regions.append(Region(start_row=1, end_row=max_row))

        return regions

    # ================================================================
    # 表头分析
    # ================================================================

    def _analyze_headers(self, ws: Worksheet, region: Region,
                         age_col: int) -> Optional[Layout]:
        """分析区段的表头结构，确定布局类型和维度映射。"""
        max_col = ws.max_column

        # 找数据起始行
        data_start = None
        for row in range(region.start_row, region.end_row + 1):
            age = _clean_int(ws.cell(row=row, column=age_col).value)
            if age is not None and age >= 0:
                data_start = row
                break
        if data_start is None:
            return None

        # 确定费率列范围
        rate_cols = []
        for col in range(age_col + 1, max_col + 1):
            rate = _clean_float(ws.cell(row=data_start, column=col).value)
            if rate is not None and rate > 0:
                rate_cols.append(col)
        if not rate_cols:
            return None

        # 提取表头行（region.start 到 data_start-1）
        header_rows_data = []  # [(row_num, [col_values])]
        for row in range(region.start_row, data_start):
            row_vals = {}
            for col in [age_col] + rate_cols:
                v = ws.cell(row=row, column=col).value
                if v is not None:
                    row_vals[col] = str(v).strip()
            if row_vals:
                header_rows_data.append((row, row_vals))

        # ---- 判断布局类型 ----

        # 检查是否有标题行含性别/计划信息
        title_gender = None
        title_plan = ""
        if region.title_text:
            flat = region.title_text.replace("\n", " ").replace("\r", " ")
            gm = re.search(r'[（(]\s*(男性|女性)\s+保险期间', flat)
            if gm:
                title_gender = "男" if gm.group(1) == "男性" else "女"
            pm = re.search(r'保险期间终身\s+(.+?)\s+每\d+元', flat)
            if pm:
                title_plan = pm.group(1).strip()
            # 次标准体
            if "次标准体" in flat:
                rm = re.search(r'加费评点(\d+)', flat)
                if rm:
                    title_plan = f"{title_plan} 评点{rm.group(1)}" if title_plan else f"次标准体 评点{rm.group(1)}"

        # 检查表头行中是否有直接的 pay_period → col 映射
        has_pay_row = False
        has_gender_row = False
        has_multi_dim_rows = False
        pay_from_header = {}  # {col: pay_period}

        for row_num, row_vals in header_rows_data:
            # 这一行包含多少种维度类型？
            dim_types = set()
            pay_count = 0
            gender_count = 0
            for col, val in row_vals.items():
                if col == age_col:
                    continue
                ct = _classify_cell(val)
                if ct:
                    dim_types.add(ct)
                if ct == "pay_period":
                    pay_count += 1
                    pp = _norm_pay(val)
                    if pp is not None:
                        pay_from_header[col] = pp
                elif ct == "gender":
                    gender_count += 1

            if pay_count >= 2:
                has_pay_row = True
            if gender_count >= 2:
                has_gender_row = True
            if len(dim_types) >= 2:
                has_multi_dim_rows = True

        # ---- 构造 Layout ----

        # Title grid: 标题含性别/计划, 表头有交费期间行
        if title_gender and has_pay_row:
            layout = Layout(
                layout_type="title_grid",
                age_col=age_col,
                rate_cols=rate_cols,
                data_start_row=data_start,
                pay_period_map=pay_from_header,
                gender=title_gender,
                plan_label=title_plan or region.title_text[:50],
                period="终身",
            )
            return layout

        # Grid: 表头有交费期间行 + 性别行（无标题性别）
        if has_pay_row and (has_gender_row or len(header_rows_data) <= 3):
            # 找性别行
            gender = ""
            gender_from_cols = {}
            for row_num, row_vals in header_rows_data:
                for col, val in row_vals.items():
                    if col == age_col:
                        continue
                    g = CN_GENDER.get(val)
                    if g:
                        gender_from_cols[col] = g
            # 如果男女都有，取第一个
            genders = list(gender_from_cols.values())
            gender = genders[0] if genders else (title_gender or "")

            layout = Layout(
                layout_type="grid",
                age_col=age_col,
                rate_cols=rate_cols,
                data_start_row=data_start,
                pay_period_map=pay_from_header,
                gender=gender,
                plan_label=title_plan or "",
                period="终身",
            )
            return layout

        # Column: 多行多维度表头
        if has_multi_dim_rows or len(header_rows_data) >= 3:
            # 每列独立维度
            col_dims = {}
            for col in rate_cols:
                col_dims[col] = {"pay_period": None, "gender": "",
                                 "plan": 0, "period": "终身"}
            for row_num, row_vals in header_rows_data:
                for col, val in row_vals.items():
                    if col == age_col or col not in col_dims:
                        continue
                    ct = _classify_cell(val)
                    if ct == "pay_period":
                        col_dims[col]["pay_period"] = _norm_pay(val)
                    elif ct == "gender":
                        col_dims[col]["gender"] = CN_GENDER.get(val, val)
                    elif ct == "period":
                        col_dims[col]["period"] = val
                    elif ct == "plan":
                        col_dims[col]["plan"] = int(val)

            # 过滤掉没有 pay_period 的列
            col_dims = {c: d for c, d in col_dims.items()
                        if d["pay_period"] is not None}

            layout = Layout(
                layout_type="column",
                age_col=age_col,
                rate_cols=list(col_dims.keys()),
                data_start_row=data_start,
                col_dims=col_dims,
                plan_label=title_plan or "",
                period="终身",
            )
            return layout

        # 兜底：尝试作为 title_grid 处理
        if pay_from_header:
            layout = Layout(
                layout_type="title_grid",
                age_col=age_col,
                rate_cols=rate_cols,
                data_start_row=data_start,
                pay_period_map=pay_from_header,
                gender=title_gender or "",
                plan_label=title_plan or region.title_text[:50],
                period="终身",
            )
            return layout

        return None

    # ================================================================
    # 数据提取
    # ================================================================

    def _extract_region(self, ws: Worksheet, region: Region,
                        layout: Layout) -> list[dict]:
        """从区段提取标准行字典。"""
        results = []

        if layout.layout_type == "column":
            results = self._extract_column_region(ws, region, layout)
        else:
            results = self._extract_grid_region(ws, region, layout)

        return results

    def _extract_grid_region(self, ws: Worksheet, region: Region,
                             layout: Layout) -> list[dict]:
        """提取 grid / title_grid 区段。"""
        # 初始化每个 (pay_period, gender) 的边界
        boundaries = {}
        for col, pp in layout.pay_period_map.items():
            key = (pp, layout.gender)
            boundaries[key] = {"col": col, "min_age": None, "min_rate": None,
                               "max_age": None, "max_rate": None}

        # 逐行读数据
        for row in range(layout.data_start_row, region.end_row + 1):
            age = _clean_int(ws.cell(row=row, column=layout.age_col).value)
            if age is None:
                break
            for (pp, gender), b in boundaries.items():
                rate = _clean_float(ws.cell(row=row, column=b["col"]).value)
                if rate is None:
                    continue
                if b["min_age"] is None:
                    b["min_age"] = age
                    b["min_rate"] = rate
                b["max_age"] = age
                b["max_rate"] = rate

        # 转为标准行
        label = layout.plan_label or "标准体"
        rows = []
        for (pp, gender), b in boundaries.items():
            if b["min_age"] is None:
                continue
            rows.append({
                "保障方案": label,
                "ensurePlan": "1",
                "责任计划": 0,
                "保险期间": layout.period,
                "交费期间": pp,
                "性别": gender,
                "最小年龄": b["min_age"],
                "最小年龄费率": b["min_rate"],
                "最大年龄": b["max_age"],
                "最大年龄费率": b["max_rate"],
            })
        return rows

    def _extract_column_region(self, ws: Worksheet, region: Region,
                               layout: Layout) -> list[dict]:
        """提取 column 区段。"""
        # 初始化边界
        boundaries = {}
        for col, dims in layout.col_dims.items():
            key = (col, dims["pay_period"], dims["gender"])
            boundaries[key] = {
                "col": col,
                "pay_period": dims["pay_period"],
                "gender": dims["gender"],
                "plan": dims.get("plan", 0),
                "period": dims.get("period", layout.period),
                "min_age": None, "min_rate": None,
                "max_age": None, "max_rate": None,
            }

        for row in range(layout.data_start_row, region.end_row + 1):
            age = _clean_int(ws.cell(row=row, column=layout.age_col).value)
            if age is None:
                break
            for key, b in boundaries.items():
                rate = _clean_float(ws.cell(row=row, column=b["col"]).value)
                if rate is None:
                    continue
                if b["min_age"] is None:
                    b["min_age"] = age
                    b["min_rate"] = rate
                b["max_age"] = age
                b["max_rate"] = rate

        label = layout.plan_label or "标准体"
        rows = []
        for key, b in boundaries.items():
            if b["min_age"] is None:
                continue
            rows.append({
                "保障方案": label,
                "ensurePlan": "1",
                "责任计划": b["plan"],
                "保险期间": b["period"],
                "交费期间": b["pay_period"],
                "性别": b["gender"],
                "最小年龄": b["min_age"],
                "最小年龄费率": b["min_rate"],
                "最大年龄": b["max_age"],
                "最大年龄费率": b["max_rate"],
            })
        return rows

    # ================================================================
    # 产品名提取
    # ================================================================

    def _is_em_surcharge(self, title_text: str) -> bool:
        """判断是否为 EM 加费 / 次标准体评点加费区段。"""
        if not title_text:
            return False
        # 次标准体评点加费、EM 加费等
        em_keywords = ["次标准体", "加费表", "EM加费", "EM 加费", "评点加费"]
        for kw in em_keywords:
            if kw in title_text:
                return True
        return False

    def _extract_product_name(self, title_text: str) -> str:
        """从标题文本提取产品名。"""
        flat = title_text.replace("\n", " ").replace("\r", " ")
        m = re.match(r'^(.+?)(?:费率表|次标准体评点加费表)', flat)
        if m:
            return m.group(1).strip()
        return ""

    # ================================================================
    # 结果组装
    # ================================================================

    def _build_result(self, all_rows: list[dict], product_name: str,
                      meta: dict, data_sheets: list[str],
                      file_path: str) -> dict:
        """组装标准返回 dict。"""
        # 维度汇总
        dims = {}
        for r in all_rows:
            for k in ["保障方案", "交费期间", "性别", "责任计划"]:
                dims.setdefault(k, set()).add(r.get(k, ""))

        # 边界值列表（前端用）
        boundaries = []
        for r in all_rows:
            boundaries.append({
                "label": r["保障方案"],
                "pay": r["交费期间"],
                "gender": r["性别"],
                "min_age": r["最小年龄"],
                "min_rate": r["最小年龄费率"],
                "max_age": r["最大年龄"],
                "max_rate": r["最大年龄费率"],
            })

        case_count = sum(2 if b["min_age"] != b["max_age"] else 1
                         for b in boundaries)

        data_type = meta.get("data_type", "1")
        data_type_label = "保费算保额" if data_type == "2" else "保额算保费"

        return {
            "format": "auto",
            "product_name": product_name or meta.get("product_name", "") or "(未标注)",
            "data_type": data_type_label,
            "fee_unit": meta.get("fee_unit", 1000),
            "fee_rule": meta.get("fee_rule", "四舍五入保留2位小数"),
            "dims": {k: sorted(v, key=str) for k, v in dims.items()},
            "boundaries": boundaries,
            "case_count": case_count,
            "boundary_count": len(boundaries),
            "_rows": all_rows,
            "rate_sheet": data_sheets[0] if data_sheets else "",
            "file_path": file_path,
        }


# ============================================================
# 工具函数：写边界值 Excel（供 BatchTester 读取）
# ============================================================

def write_boundary_xlsx(rows: list[dict], output_path: str, product_name: str = ""):
    """将标准行字典列表写入边界值汇总 Excel（TestCaseLoader 兼容格式）。"""
    from datetime import datetime as dt
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    HEADERS = [
        "保障方案", "ensurePlan", "责任计划", "保险期间", "交费期间（年）",
        "性别", "最小年龄（岁）", "最小年龄费率", "最大年龄（岁）", "最大年龄费率",
    ]
    COL_WIDTHS = [10, 12, 10, 10, 14, 8, 14, 14, 14, 14]
    ncols = len(HEADERS)

    wb = Workbook()
    ws = wb.active
    ws.title = "边界费率汇总"

    title = f"{product_name or '保险产品'} — 边界费率汇总表  生成时间：{dt.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value="保险期间：终身 | 金额单位：元（每1,000元保额）")
    c.font = Font(name="微软雅黑", size=9, color="666666")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 22

    hdr_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    hdr_font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    for col_idx, header in enumerate(HEADERS, 1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 22

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for i, row_data in enumerate(rows):
        row_num = 4 + i
        is_even = (i % 2 == 0)
        bg = "DEEAF1" if is_even else "FFFFFF"
        row_fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")

        values = [
            row_data.get("保障方案", ""),
            row_data.get("ensurePlan", "1"),
            row_data.get("责任计划", 0),
            row_data.get("保险期间", ""),
            row_data.get("交费期间", ""),
            row_data.get("性别", ""),
            row_data.get("最小年龄", ""),
            row_data.get("最小年龄费率", ""),
            row_data.get("最大年龄", ""),
            row_data.get("最大年龄费率", ""),
        ]

        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_num, column=col_idx, value=val)
            cell.font = Font(name="微软雅黑", size=10, color="1F4E79",
                            bold=(col_idx <= 2))
            cell.fill = row_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

    for col_idx, width in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(ncols)}{3 + len(rows)}"

    wb.save(output_path)
    print(f"✅ 边界值汇总已写入: {output_path}")
