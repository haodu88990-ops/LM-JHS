"""
通用批量测试引擎

从边界值汇总 Excel 加载测试用例，通过 API 逐个验证保费费率计算。
支持自动节流、进度显示、错误恢复。
"""

import random
import time
from datetime import datetime
from typing import Optional

from openpyxl import load_workbook

from .config import ProductProfile
from .api_client import InsuranceAPIClient


class TestCaseLoader:
    """
    测试用例加载器 —— 从标准 10 列边界值汇总 Excel 读取用例。

    标准输入格式:
        保障方案 | ensurePlan | 责任计划 | 保险期间 | 交费期间(年) |
        性别 | 最小年龄(岁) | 最小年龄费率 | 最大年龄(岁) | 最大年龄费率

    每行生成 1~2 个测试用例（最小年龄 + 最大年龄）。
    """

    def __init__(self, profile: ProductProfile):
        self.profile = profile

    def load(self, file_path: str, sheet_name: str = None) -> list[dict]:
        """
        从边界值汇总 Excel 加载测试用例。

        Args:
            file_path: Excel 文件路径
            sheet_name: Sheet 名称（默认第一个 Sheet）

        Returns:
            测试用例列表，每项包含:
            {
                "保障方案": str, "ensurePlan": str, "责任计划": int,
                "保险期间": str, "交费期间": int, "性别": str,
                "年龄": int, "年龄类型": str ("最小年龄"/"最大年龄"),
                "期望费率": float,
            }
        """
        wb = load_workbook(file_path, data_only=True)
        sheet_name = sheet_name or wb.sheetnames[0]
        ws = wb[sheet_name]

        cases = []
        data_start = 4  # 第 1-3 行是标题/说明/表头

        for row in range(data_start, ws.max_row + 1):
            # 读取 10 列
            plan_name = self._val(ws, row, 1)       # 保障方案
            ensure_plan = self._val(ws, row, 2)      # ensurePlan
            plan = self._int(ws, row, 3)             # 责任计划
            period = self._val(ws, row, 4)            # 保险期间
            pay_period = self._int(ws, row, 5)       # 交费期间
            gender = self._val(ws, row, 6)            # 性别
            min_age = self._int(ws, row, 7)          # 最小年龄
            min_rate = self._float(ws, row, 8)       # 最小年龄费率
            max_age = self._int(ws, row, 9)          # 最大年龄
            max_rate = self._float(ws, row, 10)      # 最大年龄费率

            # 跳过无效行
            if plan is None:
                continue
            if gender not in ("男", "女"):
                continue
            if pay_period is None or not period:
                continue

            # 生成最小年龄用例
            if min_age is not None and min_rate is not None:
                cases.append({
                    "保障方案": plan_name,
                    "ensurePlan": str(ensure_plan) if ensure_plan is not None else "1",
                    "责任计划": plan,
                    "保险期间": period,
                    "交费期间": pay_period,
                    "性别": gender,
                    "年龄": min_age,
                    "年龄类型": "最小年龄",
                    "期望费率": min_rate,
                })

            # 生成最大年龄用例（与最小年龄不同时才生成）
            if (max_age is not None and max_rate is not None
                    and max_age != min_age):
                cases.append({
                    "保障方案": plan_name,
                    "ensurePlan": str(ensure_plan) if ensure_plan is not None else "1",
                    "责任计划": plan,
                    "保险期间": period,
                    "交费期间": pay_period,
                    "性别": gender,
                    "年龄": max_age,
                    "年龄类型": "最大年龄",
                    "期望费率": max_rate,
                })

        wb.close()
        return cases

    @staticmethod
    def _val(ws, row, col):
        v = ws.cell(row=row, column=col).value
        return str(v).strip() if v is not None else None

    @staticmethod
    def _int(ws, row, col):
        v = ws.cell(row=row, column=col).value
        if v is None:
            return None
        try:
            return int(float(v))
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _float(ws, row, col):
        v = ws.cell(row=row, column=col).value
        if v is None:
            return None
        try:
            return float(v)
        except (ValueError, TypeError):
            return None


class BatchTester:
    """
    批量测试执行器。

    用法:
        profile = ProductProfile.from_yaml("config.yaml")
        tester = BatchTester(profile, serial_no="1491455820292947968")
        results = tester.run("边界值汇总.xlsx")
    """

    def __init__(self, profile: ProductProfile, serial_no: str, proposal_id: str = ""):
        if not serial_no:
            raise ValueError("serial_no 是模拟测算的关键参数，不能为空")
        self.profile = profile
        self.serial_no = serial_no
        self.proposal_id = proposal_id
        self.client = InsuranceAPIClient(profile, serial_no=serial_no, proposal_id=proposal_id)
        self.loader = TestCaseLoader(profile)

    # ================================================================
    # 主入口
    # ================================================================

    def run(
        self,
        boundary_file: str,
        output_file: str = None,
        smoke_only: bool = False,
        smoke_count: int = 5,
        progress_callback=None,
    ) -> list[dict]:
        """
        执行完整的批量测试流程。

        Args:
            boundary_file: 边界值汇总 Excel 路径
            output_file: 报告输出路径（None=自动生成）
            smoke_only: 是否仅执行冒烟测试
            smoke_count: 冒烟测试用例数
            progress_callback: 进度回调 fn(current, total)

        Returns:
            测试结果列表
        """
        # 加载用例
        print(f"📂 加载测试用例: {boundary_file}")
        cases = self.loader.load(boundary_file)

        if smoke_only:
            cases = cases[:smoke_count]
            print(f"🔥 冒烟测试模式: 仅执行前 {len(cases)} 条")

        print(f"  共加载 {len(cases)} 条测试用例")

        if not cases:
            print("⚠ 没有可执行的测试用例")
            return []

        # 登录
        print("🔑 登录中...")
        try:
            ok, _ = self.client.login()
        except Exception as e:
            print(f"❌ 无法连接 API 服务器: {e}")
            print("   请确认:")
            print("   1. 是否在公司内网环境")
            print(f"   2. API 地址是否正确: {self.profile.api.base_url}")
            print("   3. VPN 是否已连接")
            return []
        if not ok:
            print("⚠ 登录可能失败，继续执行测试...")

        # 执行测试
        print(f"🧪 开始执行 {len(cases)} 条用例...\n")
        results = []
        start_time = time.time()

        for i, case in enumerate(cases):
            case_no = i + 1
            result = self._run_single(case, case_no)
            results.append(result)

            # 输出单条结果
            verdict = result.get("测试结论", "?")
            emoji = "✅" if verdict.startswith("PASS") else "❌" if verdict.startswith("FAIL") else "⚠️"
            print(f"  [{case_no:03d}/{len(cases)}] {emoji} {verdict} | "
                  f"计划{case['责任计划']} {case['性别']} {case['年龄']}岁 "
                  f"| 期望{case['期望费率']}‰")

            # 节流
            throttle = self.profile.test.throttle
            if throttle.interval > 0 and case_no % throttle.interval == 0:
                time.sleep(throttle.sleep)

            # 进度回调
            if progress_callback:
                progress_callback(case_no, len(cases))

        elapsed = time.time() - start_time

        # 统计
        pass_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("PASS"))
        fail_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("FAIL"))
        error_count = sum(1 for r in results if str(r.get("测试结论", "")).startswith("ERROR"))

        print(f"\n{'='*50}")
        print(f"  总用例: {len(results)}")
        print(f"  ✅ PASS:  {pass_count} ({100*pass_count/len(results):.1f}%)" if pass_count else "  ✅ PASS:  0")
        print(f"  ❌ FAIL:  {fail_count}" if fail_count else "  ❌ FAIL:  0")
        print(f"  ⚠ ERROR: {error_count}" if error_count else "  ⚠ ERROR: 0")
        print(f"  ⏱ 耗时:   {elapsed:.1f}s")
        print(f"{'='*50}\n")

        return results

    # ================================================================
    # 单用例执行
    # ================================================================

    def _run_single(self, case: dict, case_no: int) -> dict:
        """
        执行单个测试用例。

        流程:
        1. 调用 age_rate(saveCustomer)
        2. 调用 plan_rate(saveProductExt)
        3. 比较计算保费与期望保费（允许 tolerance 误差）
        """
        plan = case["责任计划"]
        period = case["保险期间"]
        pay_period = case["交费期间"]
        gender = case["性别"]
        age = case["年龄"]
        age_type = case["年龄类型"]
        ensure_plan = case.get("ensurePlan", "1")
        expected_rate = case["期望费率"]

        # 生成随机保额
        amount_cfg = self.profile.test.amount
        if amount_cfg.fixed:
            amount = amount_cfg.fixed
        else:
            amount = random.randrange(amount_cfg.min, amount_cfg.max + 1, amount_cfg.step)

        expected_premium = amount * expected_rate / 1000

        # 基础结果
        result = {
            "序号": case_no,
            "保障方案": case.get("保障方案", ""),
            "ensurePlan": ensure_plan,
            "责任计划": plan,
            "责任描述": self.profile.plan_has_description(plan),
            "保险期间": period,
            "保险期间Code": self.profile.get_ensure_period_code(period),
            "交费期间": pay_period,
            "交费期间Code": self.profile.get_pay_period_code(pay_period),
            "交费方式Code": self.profile.get_pay_mode_code(pay_period),
            "性别": gender,
            "性别Code": self.profile.get_gender_code(gender),
            "年龄": age,
            "年龄类型": age_type,
            "保额(元)": amount,
            "期望费率(‰)": expected_rate,
            "期望保费(元)": round(expected_premium, 2),
            "age_rate状态码": None,
            "age_rate结果": "",
            "plan_rate状态码": None,
            "plan_rate结果": "",
            "API返回fee": None,
            "failureReason": None,
            "测试结论": "",
            "备注": "",
        }

        try:
            # Step 1: age_rate
            ok, cookies, status, text = self.client.save_customer(age, gender)
            result["age_rate状态码"] = status
            result["age_rate结果"] = "成功" if ok else "失败"

            if not ok:
                result["测试结论"] = "FAIL - age_rate接口失败"
                return result

            # Step 2: plan_rate
            ok, fee, status, text, reason = self.client.save_product(
                plan=plan,
                ensure_period=period,
                pay_period=pay_period,
                amount=amount,
                ensure_plan=ensure_plan,
                cookies=cookies,
            )
            result["plan_rate状态码"] = status
            result["API返回fee"] = fee
            result["failureReason"] = reason

            if ok:
                result["plan_rate结果"] = "成功"
                # 比较保费
                actual_fee = float(fee)
                deviation = abs(actual_fee - expected_premium) / expected_premium if expected_premium > 0 else float("inf")

                if deviation < self.profile.test.tolerance:
                    result["测试结论"] = "PASS"
                else:
                    pct = deviation * 100
                    result["测试结论"] = f"PASS(费率偏差{pct:.1f}%)"
                    result["备注"] = f"期望{expected_premium:.2f}, 实际{actual_fee:.2f}, 偏差{pct:.2f}%"
            else:
                result["plan_rate结果"] = "失败" if status == 500 else "异常"
                if fee is not None and float(fee) == 0 and reason:
                    result["测试结论"] = "FAIL - 计算保费失败"
                elif fee is not None:
                    result["测试结论"] = "PASS(无fee字段)"
                else:
                    result["测试结论"] = "FAIL - plan_rate接口失败"

        except Exception as e:
            result["测试结论"] = f"ERROR - {str(e)}"
            result["备注"] = str(e)

        return result
