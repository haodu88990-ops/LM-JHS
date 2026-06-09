#!/usr/bin/env python3
"""
保费测算 Web 界面。

启动:
    python web_app.py
    python web_app.py --port 8080

然后浏览器打开 http://localhost:5050
（默认5050端口，避免与 macOS AirPlay 的5000端口冲突）
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

from src.config import ProductProfile
from src.rate_parser import RateTableParser, write_boundary_xlsx
from src.tester import BatchTester
from src.reporter import ReportGenerator

# 确保无论从哪个目录启动都能找到模板
BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB

# 上传文件暂存
UPLOAD_DIR = BASE_DIR.parent / "_uploads"
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
# 解析（内容驱动，自动适配格式）
# ================================================================


def parse_rate_table(file_path: str) -> dict:
    """解析费率表，自动检测布局并提取边界值。"""
    parser = RateTableParser()
    return parser.parse(file_path)







# ================================================================
# 路由
# ================================================================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})


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

        product = ProductProfile.from_dict(DEFAULT_PRODUCT)
        if parsed["product_name"] != "(未标注)":
            product.product_name = parsed["product_name"]

        tmp_boundary = os.path.join(UPLOAD_DIR, "_boundary_tmp.xlsx")

        # ---- 生成边界值 ----
        write_boundary_xlsx(parsed.get("_rows", []), tmp_boundary,
                             product.product_name or parsed["product_name"])

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
    parser.add_argument("--port", type=int, default=5050)
    parser.add_argument("--test", action="store_true", help="仅自检，不启动服务")
    args = parser.parse_args()

    if args.test:
        print("🔍 自检中...")
        print(f"   BASE_DIR: {BASE_DIR}")
        print(f"   templates: {BASE_DIR / 'templates'}")
        print(f"   index.html: {'✅' if (BASE_DIR / 'templates' / 'index.html').exists() else '❌ 缺失'}")
        print(f"   uploads dir: {'✅' if UPLOAD_DIR.exists() else '📁 创建中...'}")
        UPLOAD_DIR.mkdir(exist_ok=True)
        with app.test_client() as c:
            r = c.get('/')
            print(f"   GET /: {r.status_code}")
            r = c.get('/api/health')
            print(f"   GET /api/health: {r.status_code} {r.get_json()}")
        print("✅ 自检通过，可以启动: python web_app.py")
        import sys; sys.exit(0)

    print(f"""
╔══════════════════════════════════════════════╗
║       保费测算工具 - Web 界面                  ║
║                                              ║
║   浏览器打开: http://localhost:{args.port}           ║
║                                              ║
║   1. 上传 PDF / Excel 费率表                  ║
║   2. 输入 serialNo                           ║
║   3. 点击「开始测算」                          ║
║   4. 下载 Excel 报告                          ║
║                                              ║
║   按 Ctrl+C 退出                              ║
╚══════════════════════════════════════════════╝
""")
    try:
        app.run(host="0.0.0.0", port=args.port, debug=False)
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"❌ 端口 {args.port} 已被占用，换一个:")
            print(f"   python web_app.py --port {args.port + 1}")
        else:
            raise
