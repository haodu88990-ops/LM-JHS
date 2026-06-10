#!/usr/bin/env python3
"""
保费测算 Web 界面。

启动:
    python web_app.py
    python web_app.py --port 8080

然后浏览器打开 http://localhost:5050
（默认5050端口，避免与 macOS AirPlay 的5000端口冲突）
"""

import re
import sys
import os
import io
import json
import tempfile
import time
import uuid
import threading
import traceback
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

from src.config import ProductProfile
from src.rate_parser import RateTableParser, write_boundary_xlsx
from src.tester import BatchTester
from src.reporter import ReportGenerator

# 确保无论从哪个目录启动都能找到模板
BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB

# 错误处理：确保 API 返回 JSON（而非 HTML）
@app.errorhandler(RequestEntityTooLarge)
def handle_too_large(e):
    return jsonify({"error": f"文件过大，上限 {app.config['MAX_CONTENT_LENGTH'] // (1024*1024)}MB"}), 413

@app.errorhandler(500)
def handle_500(e):
    return jsonify({"error": "服务器内部错误，请查看终端日志"}), 500

# 上传文件暂存
UPLOAD_DIR = BASE_DIR.parent / "_uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

# 后台任务状态存储（用于 SSE 进度推送）
_tasks = {}          # task_id -> {"status", "current", "total", "message", "report_path", "error", "results"}
_tasks_lock = threading.Lock()


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
    "product": {"product_id": "", "company_id": ""},
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
        "throttle": {"interval": 5, "sleep": 0.3, "workers": 20},
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

    out_path = os.path.join(str(UPLOAD_DIR), "_pdf_extracted.xlsx")
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


@app.route("/api/ping", methods=["GET", "POST"])
def ping():
    """诊断接口：确认浏览器可正常访问"""
    return jsonify({
        "ok": True,
        "method": request.method,
        "content_type": str(request.content_type),
        "content_length": request.content_length,
    })


@app.route("/api/upload", methods=["POST"])
def upload():
    """上传文件 + 解析"""
    file = request.files.get("file")
    print(f"[UPLOAD] files={list(request.files.keys())} form={list(request.form.keys())} ct={request.content_type} cl={request.content_length} file={'OK' if file else 'NONE'}", flush=True)
    if not file:
        return jsonify({"error": "请上传文件"}), 400

    filename = secure_filename(file.filename or "rate_table")
    # 防止 secure_filename 对纯中文文件名返回空串，导致 file_path 变成目录
    if not filename or not os.path.splitext(filename)[1]:
        fallback_name = (file.filename or "rate_table.xlsx")
        fallback_ext = os.path.splitext(fallback_name)[1] or ".xlsx"
        filename = f"rate_table_{uuid.uuid4().hex[:8]}{fallback_ext}"
    ext = os.path.splitext(filename)[1].lower()

    # 保存上传文件 — 必须 try/except，否则异常会崩掉 threaded=True 下的请求线程，
    # 导致 TCP 连接被重置，浏览器收到 TypeError: Failed to fetch 而非错误响应
    file_path = os.path.join(str(UPLOAD_DIR), filename)
    try:
        file.save(file_path)
    except Exception as e:
        return jsonify({"error": f"文件保存失败: {str(e)}"}), 500

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
    """启动测算（后台执行，通过 SSE 推送进度）"""
    data = request.get_json()
    serial_no = (data.get("serial_no") or "").strip()
    file_path = data.get("file_path", "")

    if not serial_no:
        return jsonify({"error": "请输入 serialNo（方案序列号）"}), 400
    if not file_path or not os.path.exists(file_path):
        return jsonify({"error": "费率表文件丢失，请重新上传"}), 400

    task_id = uuid.uuid4().hex[:12]
    # 报告统一保存到 _uploads 目录（避免 original_path 被 File.path 传空导致路径错乱）
    report_dir = str(UPLOAD_DIR)

    with _tasks_lock:
        # 清理超过 10 分钟的过期任务（SSE 从未连接的情况）
        now = time.time()
        stale_ids = [
            tid for tid, t in _tasks.items()
            if now - t.get("created_at", 0) > 600
        ]
        for tid in stale_ids:
            del _tasks[tid]

        _tasks[task_id] = {
            "status": "running",
            "current": 0,
            "total": 0,
            "message": "🔑 登录中...",
            "report_path": "",
            "report_name": "",
            "stats": None,
            "error": None,
            "created_at": now,
        }

    def _run_background():
        try:
            # ---- 解析 ----
            parsed = parse_rate_table(file_path)

            # ---- 构建产品配置（合并用户输入） ----
            product_data = dict(DEFAULT_PRODUCT)
            if parsed["product_name"] != "(未标注)":
                product_data["product_name"] = parsed["product_name"]

            # 用户输入的 product_id / company_id（必填，无默认值）
            user_product_id = (data.get("product_id") or "").strip()
            user_company_id = (data.get("company_id") or "").strip()
            if user_product_id:
                product_data["product"]["product_id"] = user_product_id
            if user_company_id:
                product_data["product"]["company_id"] = user_company_id

            if not product_data["product"]["product_id"] or not product_data["product"]["company_id"]:
                with _tasks_lock:
                    _tasks[task_id]["status"] = "error"
                    _tasks[task_id]["error"] = "请填写 product_id 和 company_id（在「产品接口参数」区域）"
                return

            # 用户输入的保额范围
            amount_min = data.get("amount_min")
            amount_max = data.get("amount_max")
            amount_step = data.get("amount_step")
            if amount_min is not None:
                try:
                    product_data["test"]["amount"]["min"] = int(amount_min)
                except (ValueError, TypeError):
                    pass
            if amount_max is not None:
                try:
                    product_data["test"]["amount"]["max"] = int(amount_max)
                except (ValueError, TypeError):
                    pass
            if amount_step is not None:
                try:
                    product_data["test"]["amount"]["step"] = int(amount_step)
                except (ValueError, TypeError):
                    pass

            # 用户输入的责任码值映射
            duty_codes = data.get("duty_codes")  # {"责任名": "code", ...}
            pstruct = parsed.get("product_structure", {})
            if duty_codes and pstruct.get("has_optional_duties"):
                # 构建 duty_combination 模式
                duties = {"1": {"name": "基本保险责任", "order": 1, "default_value": "1"}}
                order = 2
                for duty_name, duty_code in duty_codes.items():
                    code_str = str(duty_code).strip()
                    if code_str and code_str != "1":
                        duties[code_str] = {
                            "name": duty_name,
                            "order": order,
                            "default_value": code_str,
                        }
                        order += 1

                # 构建 plan_duties: 根据方案名解析包含哪些责任
                plan_labels = pstruct.get("plan_labels", {})
                plan_duties = {}
                for label, plan_idx in plan_labels.items():
                    codes = []
                    parts = label.split("+")
                    for part in parts[1:]:  # 跳过基本责任
                        part = part.strip()
                        duty_name = re.sub(r'责任$', '', part)
                        if duty_name == "全部可选" or "全部可选" in duty_name:
                            # "全部可选" → 包含所有可选责任
                            codes.extend(
                                str(c).strip() for c in duty_codes.values()
                                if str(c).strip() and str(c).strip() != "1"
                            )
                        elif duty_name in duty_codes:
                            code = str(duty_codes[duty_name]).strip()
                            if code and code != "1":
                                codes.append(code)
                    plan_duties[plan_idx] = codes

                product_data["plans"] = {
                    "type": "duty_combination",
                    "duties": duties,
                    "plan_duties": plan_duties,
                }

            # 用户选择的承保方案（ensurePlan），覆盖解析器默认值
            user_ensure_plan = (data.get("ensure_plan") or "").strip()
            if user_ensure_plan in ("1", "2"):
                for row in parsed.get("_rows", []):
                    row["ensurePlan"] = user_ensure_plan

            # 用户输入的保险期间编码映射（覆盖自动提取）
            period_codes = data.get("period_codes")  # {"至...60周岁...": "TO60", ...}
            if period_codes and isinstance(period_codes, dict):
                for raw, code in period_codes.items():
                    code_str = str(code).strip()
                    if code_str:
                        product_data.setdefault("mappings", {}).setdefault("ensure_period", {})[raw] = code_str

            product = ProductProfile.from_dict(product_data)

            tmp_boundary = os.path.join(str(UPLOAD_DIR), f"_boundary_{task_id}.xlsx")

            # ---- 生成边界值 ----
            write_boundary_xlsx(parsed.get("_rows", []), tmp_boundary,
                                 product.product_name or parsed["product_name"])

            # ---- 进度回调 ----
            def on_progress(current, total, result=None):
                with _tasks_lock:
                    t = _tasks.get(task_id)
                    if t:
                        t["current"] = current
                        t["total"] = total
                        if result:
                            verdict = str(result.get("测试结论", "?"))
                            emoji = "✅" if verdict.startswith("PASS") else "❌" if verdict.startswith("FAIL") else "⚠️"
                            t["message"] = (f"{emoji} [{current}/{total}] {verdict} | "
                                            f"计划{result.get('责任计划','?')} {result.get('性别','?')} "
                                            f"{result.get('年龄','?')}岁")

            # ---- 执行测试 ----
            tester = BatchTester(product, serial_no=serial_no, proposal_id="")
            results = tester.run(boundary_file=tmp_boundary, progress_callback=on_progress)

            if not results:
                with _tasks_lock:
                    _tasks[task_id]["status"] = "error"
                    _tasks[task_id]["error"] = "API 测算失败，请检查 serialNo 和网络连接"
                return

            # ---- 生成报告 ----
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_path = os.path.join(report_dir, f"保费测算报告_{ts}.xlsx")

            reporter = ReportGenerator(product.product_name or "保险产品")
            reporter.generate(results, report_path)

            # 统计
            pass_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("PASS"))
            fail_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("FAIL"))
            error_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("ERROR"))

            # 清理
            if os.path.exists(tmp_boundary):
                os.remove(tmp_boundary)

            with _tasks_lock:
                _tasks[task_id]["status"] = "completed"
                _tasks[task_id]["message"] = "✅ 测算完成"
                _tasks[task_id]["report_path"] = report_path
                _tasks[task_id]["report_name"] = os.path.basename(report_path)
                _tasks[task_id]["stats"] = {
                    "total": len(results),
                    "pass": pass_count,
                    "fail": fail_count,
                    "error": error_count,
                    "pass_rate": f"{100*pass_count/len(results):.1f}%" if results else "0%",
                }

        except Exception as e:
            with _tasks_lock:
                _tasks[task_id]["status"] = "error"
                _tasks[task_id]["error"] = f"测算异常: {str(e)}\n{traceback.format_exc()}"

    thread = threading.Thread(target=_run_background, daemon=True)
    thread.start()

    return jsonify({"task_id": task_id})


@app.route("/api/progress/<task_id>")
def task_progress(task_id):
    """SSE 端点：推送任务进度"""
    def generate():
        # 首次检查任务是否存在
        with _tasks_lock:
            task = _tasks.get(task_id)
        if not task:
            yield f"event: task_error\ndata: {json.dumps({'message': '任务不存在'})}\n\n"
            return

        last_current = -1
        while True:
            with _tasks_lock:
                task = _tasks.get(task_id)
                if not task:
                    break

                status = task["status"]
                current = task["current"]
                total = task["total"]

                # 发送进度（仅在变化时）
                if status == "running" and current != last_current:
                    last_current = current
                    pct = round(current / total * 100, 1) if total > 0 else 0
                    yield f"event: progress\ndata: {json.dumps({'current': current, 'total': total, 'pct': pct, 'message': task['message']})}\n\n"

                if status == "completed":
                    yield f"event: complete\ndata: {json.dumps({'stats': task['stats'], 'report_path': task['report_path'], 'report_name': task['report_name'], 'message': task['message']})}\n\n"
                    with _tasks_lock:
                        _tasks.pop(task_id, None)
                    break

                if status == "error":
                    yield f"event: task_error\ndata: {json.dumps({'message': task['error']})}\n\n"
                    with _tasks_lock:
                        _tasks.pop(task_id, None)
                    break

            if status != "running":
                break
            time.sleep(0.3)

    return app.response_class(
        generate(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


@app.route("/api/download")
def download():
    """下载报告文件（替代浏览器禁用的 file:// 链接）"""
    task_id = request.args.get("task_id", "").strip()
    if not task_id:
        return jsonify({"error": "缺少 task_id 参数"}), 400

    with _tasks_lock:
        task = _tasks.get(task_id)
    if not task:
        return jsonify({"error": "任务不存在或已过期"}), 404

    report_path = task.get("report_path", "")
    if not report_path or not os.path.exists(report_path):
        return jsonify({"error": "报告文件不存在，可能已被清理"}), 404

    # 安全检查：确保路径在允许的目录下
    allowed_dirs = [str(UPLOAD_DIR), os.path.dirname(str(UPLOAD_DIR))]
    real_path = os.path.realpath(report_path)
    if not any(real_path.startswith(os.path.realpath(d)) for d in allowed_dirs):
        return jsonify({"error": "路径不被允许"}), 403

    report_name = task.get("report_name", os.path.basename(report_path))
    return send_file(
        report_path,
        as_attachment=True,
        download_name=report_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


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
        app.run(host="0.0.0.0", port=args.port, debug=False, threaded=True)
    except OSError as e:
        if "Address already in use" in str(e):
            print(f"❌ 端口 {args.port} 已被占用，换一个:")
            print(f"   python web_app.py --port {args.port + 1}")
        else:
            raise
