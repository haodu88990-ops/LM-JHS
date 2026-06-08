#!/usr/bin/env python3
"""
保费测算 Web 界面。

启动:
    python web_app.py
    python web_app.py --port 8080

然后浏览器打开 http://localhost:5000
"""

import sys
import os
import io
import json
import tempfile
import traceback
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

from src.config import (
    FormatProfile, ProductProfile, RateLayoutConfig, SheetSection
)
from src.rate_reader import RateTableReader, RateTableMetadata
from src.boundary import BoundarySummaryGenerator
from src.tester import BatchTester
from src.reporter import ReportGenerator

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB

# 上传文件暂存
UPLOAD_DIR = Path(os.path.dirname(os.path.abspath(__file__))) / "_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


# ================================================================
# 默认产品配置
# ================================================================

DEFAULT_PRODUCT = {
    "product_name": "",
    "api": {
        "base_url": "https://kfzxtb.lmbaoxian.com:13080",
        "login_endpoint": "/broker/api/user/login.html",
        "age_rate_endpoint": "/broker/api/prospectus/saveCustomer.html",
        "plan_rate_endpoint": "/broker/api/prospectus/saveProductExt.html",
        "verify_ssl": False,
        "timeout": 30,
    },
    "credentials": {
        "account": "15856990088",
        "password_md5": "dc483e80a7a0bd9ef71d8cf973673924",
    },
    "product": {"product_id": "991452", "company_id": "100080"},
    "defaults": {
        "insurant_id": 85320, "insurant_occ_level": 1, "insurant_social_insurance": "1",
        "policy_holder_id": 85321, "policy_holder_age": 30, "policy_holder_sex": "1",
        "dividend_draw_type": "2", "request_type": "prospectus",
    },
    "mappings": {
        "gender": {"男": "1", "女": "2"},
        "ensure_period": {"终身": "TO105"},
        "pay_mode": {"single": "1", "annual": "5"},
        "pay_period": "direct",
    },
    "plans": {"type": "simple"},
    "age_limits": {},
    "test": {
        "amount": {"min": 1000000, "max": 5000000, "step": 1000},
        "tolerance": 0.01,
        "throttle": {"interval": 5, "sleep": 0.3},
    },
}


# ================================================================
# PDF → Excel 转换
# ================================================================

def pdf_to_excel(pdf_path: str) -> str:
    """从 PDF 提取费率表，保存为临时 Excel"""
    import pdfplumber
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "费率"

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and any(cell for cell in row):
                        ws.append([str(c).strip() if c else "" for c in row])

    out_path = os.path.join(UPLOAD_DIR, "_pdf_extracted.xlsx")
    wb.save(out_path)
    return out_path


# ================================================================
# 格式检测 + 解析
# ================================================================

def detect_format(file_path: str) -> str:
    from openpyxl import load_workbook
    wb = load_workbook(file_path, data_only=True)
    sheets = wb.sheetnames

    has_info = any(s in sheets for s in ["产品信息", "产品说明"])
    rate_sheet = None
    for s in sheets:
        if "费率" in s and "产品" not in s:
            rate_sheet = s
            break
    if not rate_sheet and sheets:
        rate_sheet = sheets[-1]

    if rate_sheet:
        ws = wb[rate_sheet]
        row3_first = str(ws.cell(row=3, column=1).value or "").strip().upper()
        row3_second = str(ws.cell(row=3, column=2).value or "").strip()
        if row3_first in ("PREMIUM", "保费") or "pay_period" in row3_second.lower():
            wb.close()
            return "grid"

    for s in sheets:
        if "费率表" in s:
            wb.close()
            return "column"

    wb.close()
    return "column"


def parse_rate_table(file_path: str) -> dict:
    """完整解析费率表，返回前端展示所需数据"""
    fmt_type = detect_format(file_path)

    if fmt_type == "grid":
        fmt = FormatProfile(
            format_name="auto", format_type="grid",
            layout=RateLayoutConfig(
                layout_type="grid",
                pay_period_row=3, gender_row=4,
                data_start_row=5, age_column=2, rate_columns_start=3,
            ),
            sections=[SheetSection(sheet="费率", label="标准体", ensure_plan="1", plan_override=0)],
        )
    else:
        fmt = FormatProfile(
            format_name="auto", format_type="column",
            layout=RateLayoutConfig(
                layout_type="column",
                header_rows={"plan": 4, "period": 5, "pay_period": 6, "gender": 7},
                data_start_row=8, age_column=1, rate_columns_start=2,
            ),
            sections=[SheetSection(sheet="标准体费率表", label="标准体", ensure_plan="1")],
        )

    reader = RateTableReader(fmt)
    meta = reader.read_metadata(file_path)
    product = ProductProfile.from_dict(DEFAULT_PRODUCT)
    if meta.product_name:
        product.product_name = meta.product_name
    reader = RateTableReader(fmt, product)
    rows = reader.read_all_sections(file_path)

    # 条件维度
    dims = {}
    for r in rows:
        for k in ["保障方案", "交费期间", "性别", "责任计划"]:
            dims.setdefault(k, set()).add(r.get(k, ""))

    # 边界值
    boundaries = []
    for r in rows:
        boundaries.append({
            "label": r["保障方案"],
            "pay": r["交费期间"],
            "gender": r["性别"],
            "min_age": r["最小年龄"],
            "min_rate": r["最小年龄费率"],
            "max_age": r["最大年龄"],
            "max_rate": r["最大年龄费率"],
        })

    case_count = sum(2 if b["min_age"] != b["max_age"] else 1 for b in boundaries)

    return {
        "format": fmt_type,
        "product_name": meta.product_name or "(未标注)",
        "data_type": meta.data_type_label,
        "fee_unit": meta.fee_unit,
        "fee_rule": meta.fee_rule,
        "dims": {k: sorted(v, key=str) for k, v in dims.items()},
        "boundaries": boundaries,
        "case_count": case_count,
        "boundary_count": len(boundaries),
    }


# ================================================================
# 路由
# ================================================================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    """上传文件 + 解析"""
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "请上传文件"}), 400

    filename = secure_filename(file.filename or "rate_table")
    ext = os.path.splitext(filename)[1].lower()

    # 保存上传文件
    file_path = os.path.join(UPLOAD_DIR, filename)
    file.save(file_path)

    # PDF → Excel
    if ext == ".pdf":
        try:
            file_path = pdf_to_excel(file_path)
        except Exception as e:
            return jsonify({"error": f"PDF 解析失败: {str(e)}"}), 400

    # 解析
    try:
        result = parse_rate_table(file_path)
        result["file_path"] = file_path
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": f"费率表解析失败: {str(e)}\n{traceback.format_exc()}"}), 400


@app.route("/api/run", methods=["POST"])
def run_test():
    """执行测算"""
    data = request.get_json()
    serial_no = (data.get("serial_no") or "").strip()
    file_path = data.get("file_path", "")

    if not serial_no:
        return jsonify({"error": "请输入 serialNo（方案序列号）"}), 400
    if not file_path or not os.path.exists(file_path):
        return jsonify({"error": "费率表文件丢失，请重新上传"}), 400

    try:
        # ---- 解析 ----
        parsed = parse_rate_table(file_path)
        fmt_type = parsed["format"]

        if fmt_type == "grid":
            fmt = FormatProfile(
                format_name="auto", format_type="grid",
                layout=RateLayoutConfig(
                    layout_type="grid",
                    pay_period_row=3, gender_row=4,
                    data_start_row=5, age_column=2, rate_columns_start=3,
                ),
                sections=[SheetSection(sheet="费率", label="标准体", ensure_plan="1", plan_override=0)],
            )
        else:
            fmt = FormatProfile(
                format_name="auto", format_type="column",
                layout=RateLayoutConfig(
                    layout_type="column",
                    header_rows={"plan": 4, "period": 5, "pay_period": 6, "gender": 7},
                    data_start_row=8, age_column=1, rate_columns_start=2,
                ),
                sections=[SheetSection(sheet="标准体费率表", label="标准体", ensure_plan="1")],
            )

        product = ProductProfile.from_dict(DEFAULT_PRODUCT)
        if parsed["product_name"] != "(未标注)":
            product.product_name = parsed["product_name"]

        # ---- 生成边界值 ----
        tmp_boundary = os.path.join(UPLOAD_DIR, "_boundary_tmp.xlsx")
        gen = BoundarySummaryGenerator(fmt, product)
        gen.generate(rate_file=file_path, output_file=tmp_boundary)

        # ---- 执行测试 ----
        tester = BatchTester(product, serial_no=serial_no, proposal_id="")
        results = tester.run(boundary_file=tmp_boundary)

        if not results:
            return jsonify({"error": "API 测算失败，请检查 serialNo 和网络连接"}), 500

        # ---- 生成报告 ----
        # 报告放到上传文件所在目录
        original_dir = os.path.dirname(os.path.abspath(
            data.get("original_path", UPLOAD_DIR)
        ))
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(original_dir, f"保费测算报告_{ts}.xlsx")

        reporter = ReportGenerator(product.product_name or "保险产品")
        reporter.generate(results, report_path)

        # 统计
        pass_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("PASS"))
        fail_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("FAIL"))
        error_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("ERROR"))

        # 清理
        if os.path.exists(tmp_boundary):
            os.remove(tmp_boundary)

        return jsonify({
            "success": True,
            "report_path": report_path,
            "report_name": os.path.basename(report_path),
            "stats": {
                "total": len(results),
                "pass": pass_count,
                "fail": fail_count,
                "error": error_count,
                "pass_rate": f"{100*pass_count/len(results):.1f}%" if results else "0%",
            },
        })

    except Exception as e:
        return jsonify({"error": f"测算异常: {str(e)}\n{traceback.format_exc()}"}), 500


# ================================================================
# 启动
# ================================================================

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=5000)
    args = parser.parse_args()

    print(f"""
╔══════════════════════════════════════════════╗
║       保费测算工具 - Web 界面                  ║
║                                              ║
║   打开浏览器访问: http://localhost:{args.port}     ║
║                                              ║
║   1. 上传 PDF / Excel 费率表                  ║
║   2. 输入 serialNo                           ║
║   3. 点击「开始测算」                          ║
║   4. 下载 Excel 报告                          ║
╚══════════════════════════════════════════════╝
""")
    app.run(host="0.0.0.0", port=args.port, debug=True)
