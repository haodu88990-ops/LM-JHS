#!/usr/bin/env python3
"""
保费测算工具 — 一键运行。

用法:
    python run.py <费率表.xlsx> [选项]

流程全自动:
    1. 自动识别费率表格式（grid / column）
    2. 自动读取产品信息（算费方向、费率单位...）
    3. 展示测算条件与边界值
    4. 交互式询问 serialNo
    5. API 批量测算
    6. 输出 Excel 对比报告

示例:
    python run.py 费率表.xlsx
    python run.py 费率表.xlsx --serial-no 1491455820292947968
    python run.py 费率表.xlsx --smoke
"""

import sys
import os
import argparse
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.config import ProductProfile
from src.rate_parser import RateTableParser, write_boundary_xlsx
from src.tester import BatchTester
from src.reporter import ReportGenerator


# ================================================================
# 内置默认值（费率表里有产品信息的就用产品信息，没有的用这些）
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
    "product": {
        "product_id": "991452",
        "company_id": "100080",
    },
    "defaults": {
        "insurant_id": 85320,
        "insurant_occ_level": 1,
        "insurant_social_insurance": "1",
        "policy_holder_id": 85321,
        "policy_holder_age": 30,
        "policy_holder_sex": "1",
        "dividend_draw_type": "2",
        "request_type": "prospectus",
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
    "output": {},
}


# ================================================================
def prompt_serial_no() -> str:
    print()
    print("🔑 请输入模拟测算的 serialNo（方案序列号）：")
    print("   此参数从业务系统获取，没有它无法进行保费测算")
    print()
    while True:
        val = input("   serialNo: ").strip()
        if val:
            return val
        print("   ⚠ serialNo 不能为空")


def prompt_account() -> tuple:
    """如果默认账号不对，让用户改"""
    print()
    print("📧 登录账号（回车使用默认）:")
    account = input(f"   账号 [{DEFAULT_PRODUCT['credentials']['account']}]: ").strip()
    if not account:
        account = DEFAULT_PRODUCT['credentials']['account']
    password = input(f"   密码(已加密) [{DEFAULT_PRODUCT['credentials']['password_md5'][:8]}...]: ").strip()
    if not password:
        password = DEFAULT_PRODUCT['credentials']['password_md5']
    return account, password


# ================================================================
# 主流程
# ================================================================

def main():
    parser = argparse.ArgumentParser(
        description="保费测算工具 — 一键对比费率表与 API 测算结果",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("rate_file", help="费率表 Excel 路径")
    parser.add_argument("--serial-no", "-s", help="方案序列号")
    parser.add_argument("--proposal-id", default="", help="投保单号")
    parser.add_argument("--output", "-o", help="报告输出路径")
    parser.add_argument("--account", help="登录账号")
    parser.add_argument("--smoke", action="store_true", help="冒烟测试（前5条）")
    parser.add_argument("--smoke-count", type=int, default=5)
    args = parser.parse_args()

    if not os.path.exists(args.rate_file):
        print(f"❌ 文件不存在: {args.rate_file}")
        sys.exit(1)

    rate_file = os.path.abspath(args.rate_file)

    # ================================================================
    # Phase 1: 解析费率表
    # ================================================================
    print()
    print("=" * 60)
    print("  Phase 1: 解析费率表")
    print("=" * 60)
    print(f"  文件: {rate_file}")


    # 内容驱动解析
    rp = RateTableParser()
    parsed = rp.parse(rate_file)
    rows = parsed["_rows"]

    product_data = dict(DEFAULT_PRODUCT)
    if parsed["product_name"] != "(未标注)":
        product_data["product_name"] = parsed["product_name"]
    product = ProductProfile.from_dict(product_data)

    case_count = parsed["case_count"]

    pname = parsed['product_name']
    dtype = parsed['data_type']
    funit = parsed['fee_unit']
    print(f"  产品: {pname}")
    print(f"  格式: 自动检测")
    print(f"  算费方向: {dtype}")
    print(f"  费率单位: 每{funit}元")
    print(f"  用例数:  {case_count}")
    # Phase 2: 参数确认
    # ================================================================
    print()
    print("=" * 60)
    print("  Phase 2: 参数确认")
    print("=" * 60)

    serial_no = args.serial_no or prompt_serial_no()
    proposal_id = args.proposal_id

    # 登录账号：命令行传了就直接用，没传且是交互模式才询问
    account = args.account or DEFAULT_PRODUCT["credentials"]["account"]
    password = DEFAULT_PRODUCT["credentials"]["password_md5"]
    if not args.serial_no and not args.account:
        account, password = prompt_account()

    product.credentials.account = account
    product.credentials.password_md5 = password

    print(f"\n  ✅ serialNo: {serial_no}")
    print(f"  ✅ API:     {product.api.base_url}")
    print(f"  ✅ 产品:    {product.product_name or '(费率表未标注)'}")

    if not args.serial_no:
        print()
        confirm = input("  开始测算? [Y/n]: ").strip().lower()
        if confirm and confirm != "y":
            print("  已取消")
            return

    # ================================================================
    # Phase 3: API 测算
    # ================================================================
    print()
    print("=" * 60)
    print("  Phase 3: API 保费测算")
    print("=" * 60)

    tmp_boundary = os.path.join(os.path.dirname(rate_file), "_boundary_tmp.xlsx")
    write_boundary_xlsx(rows, tmp_boundary,
                         product.product_name or parsed["product_name"])

    tester = BatchTester(product, serial_no=serial_no, proposal_id=proposal_id)
    results = tester.run(
        boundary_file=tmp_boundary,
        smoke_only=args.smoke,
        smoke_count=args.smoke_count,
    )

    if tmp_boundary and os.path.exists(tmp_boundary):
        os.remove(tmp_boundary)

    if not results:
        return

    # ================================================================
    # Phase 4: 报告
    # ================================================================
    print()
    print("=" * 60)
    print("  Phase 4: 生成报告")
    print("=" * 60)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = args.output or f"保费测算报告_{ts}.xlsx"
    reporter = ReportGenerator(product.product_name or "保险产品")
    reporter.generate(results, output)

    print(f"\n  ✅ 报告: {output}")
    print()


if __name__ == "__main__":
    main()
