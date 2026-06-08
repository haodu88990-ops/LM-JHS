#!/usr/bin/env python3
"""
从费率表 Excel 生成边界值汇总。

接收一个费率表格式定义 + 一个费率表文件，产出标准 10 列边界值汇总。
格式定义只管 Excel 布局，不涉及具体产品。

用法:
    python generate_boundary.py --format <格式.yaml> --rate-file <费率表.xlsx> [选项]

选项:
    --format, -f     费率表格式定义文件 (formats/*.yaml)
    --rate-file, -r  费率表 Excel 路径
    --output, -o     输出 Excel 路径
    --product, -p    可选：关联产品配置（用于标题、脚注）
    --grouped, -g    使用分组模式读取

示例:
    python generate_boundary.py -f formats/column.yaml -r 费率表.xlsx
    python generate_boundary.py -f formats/grid.yaml -r 费率表.xlsx -o 边界值汇总.xlsx
    python generate_boundary.py -f formats/column_shifted.yaml -r 费率表.xlsx -p products/my_product.yaml
"""

import sys
import os
import argparse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.config import FormatProfile, ProductProfile
from src.boundary import BoundarySummaryGenerator
from src.rate_reader import RateTableReader


def resolve_path(base: str, target: str) -> str:
    """智能路径解析"""
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


def main():
    parser = argparse.ArgumentParser(
        description="从费率表 Excel 生成边界值汇总",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--format", "-f", required=True, help="费率表格式定义 (formats/*.yaml)")
    parser.add_argument("--rate-file", "-r", required=True, help="费率表 Excel 路径")
    parser.add_argument("--output", "-o", help="输出 Excel 路径")
    parser.add_argument("--product", "-p", help="产品配置（可选，用于标题和脚注）")
    parser.add_argument("--grouped", "-g", action="store_true", help="分组模式读取")
    parser.add_argument("--explore", "-e", action="store_true", help="仅探索费率表结构")
    args = parser.parse_args()

    # 加载格式
    if not os.path.exists(args.format):
        print(f"❌ 格式文件不存在: {args.format}")
        sys.exit(1)

    print(f"📋 格式定义: {args.format}")
    fmt = FormatProfile.from_yaml(args.format)
    print(f"   类型: {fmt.format_type}")
    print(f"   区段: {len(fmt.sections)} 个")
    for s in fmt.sections:
        print(f"     - {s.sheet}: {s.label} (cols {s.column_start or 'auto'}~{s.column_end or 'auto'})")

    # 可选加载产品
    product = None
    if args.product:
        if os.path.exists(args.product):
            product = ProductProfile.from_yaml(args.product)
            print(f"   关联产品: {product.product_name}")

    rate_file = resolve_path(args.format, args.rate_file)

    # 探索模式
    if args.explore:
        reader = RateTableReader(fmt, product)
        info = reader.explore(rate_file)
        for sn, si in info.items():
            print(f"\n{'='*60}")
            print(f"Sheet: {sn}  ({si['dimensions']})")
            for rn, cells in sorted(si["rows"].items()):
                cs = " | ".join(f"C{c}: {v}" for c, v in cells)
                print(f"  Row {rn:2d}: {cs}")
        return

    # 生成
    output_file = args.output or "边界值汇总.xlsx"
    generator = BoundarySummaryGenerator(fmt, product)
    out = generator.generate(rate_file=rate_file, output_file=output_file, use_grouped=args.grouped)
    print(f"\n✅ 输出: {out}")


if __name__ == "__main__":
    main()
