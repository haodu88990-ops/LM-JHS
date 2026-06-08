#!/usr/bin/env python3
"""
保费测算验证工具。

流程:
  1. 解析费率表 → 展示算费方向、条件维度、边界值概况
  2. 要求输入 serialNo（模拟测算关键参数，不提供则无法继续）
  3. 逐条调用 API 测算保费，与费率表对比
  4. 输出 Excel 对比报告

用法:
    python run_tests.py -f <格式.yaml> -r <费率表.xlsx> -p <产品.yaml> [选项]

必需:
    -f, --format     费率表格式定义
    -r, --rate-file  费率表 Excel 路径
    -p, --product    产品配置

serialNo 提供方式（优先级从高到低）:
    --serial-no CLI参数   →  直接传入
    交互式输入             →  运行中提示输入

示例:
    python run_tests.py -f formats/grid.yaml -r 费率表.xlsx -p products/my_product.yaml
    python run_tests.py -f formats/grid.yaml -r 费率表.xlsx -p products/my_product.yaml --serial-no 1491455820292947968
"""

import sys
import os
import argparse
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.config import FormatProfile, ProductProfile
from src.rate_reader import RateTableReader, RateTableMetadata
from src.boundary import BoundarySummaryGenerator
from src.tester import BatchTester
from src.reporter import ReportGenerator


def resolve_path(base: str, target: str) -> str:
    if os.path.isabs(target):
        return target
    if os.path.exists(target):
        return os.path.abspath(target)
    p = os.path.join(os.path.dirname(os.path.abspath(base)), target)
    if os.path.exists(p):
        return p
    p = os.path.join(os.path.dirname(os.path.abspath(base)), "..", target)
    if os.path.exists(p):
        return os.path.abspath(p)
    return os.path.join(os.path.dirname(os.path.abspath(base)), target)


def prompt_serial_no() -> str:
    """交互式获取 serialNo —— 没有就不能继续"""
    print()
    print("🔑 请输入模拟测算的 serialNo（方案序列号）：")
    print("   (此参数从业务系统获取，没有它无法进行保费测算)")
    print()
    while True:
        val = input("serialNo: ").strip()
        if val:
            return val
        print("  ⚠ serialNo 不能为空，请重新输入：")


def main():
    parser = argparse.ArgumentParser(
        description="保费测算验证工具 — 费率表 vs API 对比",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--format", "-f", required=True, help="费率表格式定义 (formats/*.yaml)")
    parser.add_argument("--rate-file", "-r", required=True, help="费率表 Excel 路径")
    parser.add_argument("--product", "-p", required=True, help="产品配置 (products/*.yaml)")
    parser.add_argument("--serial-no", "-s", help="方案序列号（不提供则交互询问）")
    parser.add_argument("--proposal-id", default="", help="投保单号（可选）")
    parser.add_argument("--output", "-o", help="报告输出路径")
    parser.add_argument("--smoke", action="store_true", help="冒烟测试（仅前5条）")
    parser.add_argument("--smoke-count", type=int, default=5, help="冒烟测试用例数")
    args = parser.parse_args()

    # ---- 校验文件 ----
    for path, label in [(args.format, "格式文件"), (args.rate_file, "费率表"), (args.product, "产品配置")]:
        p = resolve_path(args.format, path) if label != "格式文件" else path
        if not os.path.exists(os.path.abspath(path) if os.path.isabs(path) else path):
            # try resolve
            resolved = resolve_path(args.format, path)
            if not os.path.exists(resolved):
                print(f"❌ {label}不存在: {path}")
                sys.exit(1)

    # ================================================================
    # Phase 1: 解析费率表
    # ================================================================
    print("=" * 60)
    print("  📊 Phase 1: 解析费率表")
    print("=" * 60)

    fmt = FormatProfile.from_yaml(args.format)
    product = ProductProfile.from_yaml(args.product)

    rate_file = resolve_path(args.format, args.rate_file)
    reader = RateTableReader(fmt, product)

    # 1.1 元数据（算费方向）
    meta = reader.read_metadata(rate_file)
    print(meta.summary())

    # 1.2 条件维度
    rows = reader.read_all_sections(rate_file)
    conditions = _analyze_conditions(rows)
    print(f"\n  📋 算费条件:")
    for dim, values in conditions.items():
        print(f"     {dim}: {values}")

    # 1.3 边界值概况
    print(f"\n  📋 边界值:")
    for row in rows:
        print(f"     {row['保障方案']:12s} | 交费{row['交费期间']:2d}年 | {row['性别']} | "
              f"最小{row['最小年龄']:2d}岁({row['最小年龄费率']}‰) | "
              f"最大{row['最大年龄']:2d}岁({row['最大年龄费率']}‰)")

    # ================================================================
    # Phase 2: 获取 serialNo
    # ================================================================
    print()
    print("=" * 60)
    print("  🔑 Phase 2: 模拟测算参数")
    print("=" * 60)

    serial_no = args.serial_no or prompt_serial_no()
    proposal_id = args.proposal_id

    print(f"\n  serialNo: {serial_no}")
    if proposal_id:
        print(f"  proposalId: {proposal_id}")
    print(f"  产品: {product.product_name}")
    print(f"  API:  {product.api.base_url}")
    print(f"  测试用例数: {len(rows) * 2} 条（每条边界值生成最小/最大年龄两个用例）")

    # 确认
    if not args.serial_no:
        print()
        confirm = input("  确认开始测算? [Y/n]: ").strip().lower()
        if confirm and confirm != "y":
            print("  已取消")
            return

    # ================================================================
    # Phase 3: 执行测算
    # ================================================================
    print()
    print("=" * 60)
    print("  🧪 Phase 3: API 保费测算")
    print("=" * 60)

    # 先生成边界值汇总（作为 tester 的输入）
    tmp_boundary = args.output and args.output.replace(".xlsx", "_boundary.xlsx") or "/tmp/_boundary_tmp.xlsx"
    gen = BoundarySummaryGenerator(fmt, product)
    gen.generate(rate_file=rate_file, output_file=tmp_boundary)

    # 执行测试
    tester = BatchTester(product, serial_no=serial_no, proposal_id=proposal_id)
    results = tester.run(
        boundary_file=tmp_boundary,
        smoke_only=args.smoke,
        smoke_count=args.smoke_count,
    )

    if not results:
        print("⚠ 无测试结果")
        return

    # ================================================================
    # Phase 4: 输出报告
    # ================================================================
    print()
    print("=" * 60)
    print("  📝 Phase 4: 生成报告")
    print("=" * 60)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = args.output or f"保费测算报告_{ts}.xlsx"
    reporter = ReportGenerator(product.product_name)
    reporter.generate(results, output)

    # 清理临时文件
    if tmp_boundary != output and os.path.exists(tmp_boundary):
        os.remove(tmp_boundary)

    print(f"\n✅ 报告: {output}")


def _analyze_conditions(rows: list[dict]) -> dict:
    """分析测算条件的维度与取值"""
    dims = {
        "保障方案": set(),
        "交费期间": set(),
        "性别": set(),
        "责任计划": set(),
        "年龄范围": set(),
    }
    for r in rows:
        dims["保障方案"].add(r.get("保障方案", ""))
        dims["交费期间"].add(r.get("交费期间"))
        dims["性别"].add(r.get("性别"))
        dims["责任计划"].add(r.get("责任计划"))
        dims["年龄范围"].add(f"{r.get('最小年龄')}~{r.get('最大年龄')}岁")

    return {k: sorted(v, key=str) for k, v in dims.items() if v}


if __name__ == "__main__":
    main()
