"""
通用 API 客户端

封装保险核心系统的三个核心接口：登录、年龄费率查询、计划费率计算。
完全由 ProductProfile 配置驱动，不包含任何产品特定硬编码。
"""

import json
import warnings
import requests
from typing import Any, Optional

from .config import ProductProfile

# 禁用 SSL 警告（测试环境证书问题）
warnings.filterwarnings("ignore", message="Unverified HTTPS request")


class InsuranceAPIClient:
    """
    保险核心系统 API 客户端。

    用法:
        profile = ProductProfile.from_yaml("config.yaml")
        client = InsuranceAPIClient(profile)
        client.login()
        ok, cookies, status, text = client.save_customer(age=30, sex="男")
        ok, fee, status, text, reason = client.save_product(
            plan=0, ensure_period="终身", pay_period=5, amount=1000000,
            ensure_plan="1"
        )
    """

    def __init__(self, profile: ProductProfile, serial_no: str, proposal_id: str = ""):
        """
        初始化客户端。

        Args:
            profile: 产品配置档案
            serial_no: 方案序列号（必须由用户提供，模拟测算的关键参数）
            proposal_id: 投保单号（可选，API 调用后会动态更新）
        """
        self.profile = profile
        self.session = requests.Session()
        self.login_cookies = None

        # 用户提供的模拟测算参数
        self.serial_no = serial_no
        self.proposal_id = proposal_id

    # ================================================================
    # 登录
    # ================================================================

    def login(self, account: str = None, password_md5: str = None) -> tuple:
        """
        登录获取会话 Cookie。

        Args:
            account: 账号（默认从配置读取）
            password_md5: MD5 密码（默认从配置读取）

        Returns:
            (success: bool, response: requests.Response)
        """
        url = self.profile.login_url
        account = account or self.profile.credentials.account
        password_md5 = password_md5 or self.profile.credentials.password_md5

        params = {
            "account": account,
            "password": password_md5,
        }

        # 首先尝试 form-urlencoded
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        resp = self.session.post(
            url, data=params, headers=headers,
            verify=self.profile.api.verify_ssl,
            timeout=self.profile.api.timeout,
        )

        # 如果返回"不能为空"，改用 JSON 格式重试
        if "不能为空" in resp.text:
            headers = {"Content-Type": "application/json"}
            resp = self.session.post(
                url, data=json.dumps(params), headers=headers,
                verify=self.profile.api.verify_ssl,
                timeout=self.profile.api.timeout,
            )

        self.login_cookies = resp.cookies

        # 判断登录是否成功
        success = False
        try:
            result = resp.json()
            if result.get("code") == 0 or result.get("success"):
                success = True
        except (ValueError, AttributeError):
            pass

        if "成功" in resp.text:
            success = True

        return success, resp

    # ================================================================
    # 年龄费率查询 (saveCustomer)
    # ================================================================

    def save_customer(
        self,
        insurant_age: int,
        insurant_sex: str,
        cookies=None,
    ) -> tuple:
        """
        POST 年龄费率查询接口。

        Args:
            insurant_age: 被保人年龄（整数）
            insurant_sex: 被保人性别标签（"男"/"女"）
            cookies: 指定 Cookies（默认用 login_cookies）

        Returns:
            (success: bool, cookies, status_code: int, response_text: str)
        """
        url = self.profile.age_rate_url
        d = self.profile.defaults
        sex_code = self.profile.get_gender_code(insurant_sex)

        payload = {
            "insurantName": "",
            "insurantAge": insurant_age,
            "insurantSex": sex_code,
            "insurantOccLevel": d.insurant_occ_level,
            "insurantSocialInsurance": d.insurant_social_insurance,
            "insurantId": d.insurant_id,
            "insurantBirthday": "",
            "policyHolderAge": d.policy_holder_age,
            "policyHolderSex": d.policy_holder_sex,
            "policyHolderId": d.policy_holder_id,
            "policyHolderBirthday": "",
            "serialNo": self.serial_no,
            "type": d.request_type,
        }

        headers = {"Content-Type": "application/json"}
        resp = self.session.post(
            url, data=json.dumps(payload), headers=headers,
            cookies=cookies or self.login_cookies,
            verify=self.profile.api.verify_ssl,
            timeout=self.profile.api.timeout,
        )

        success = self._check_success(resp)

        # 从响应中提取并更新 serialNo / proposalId
        try:
            result = resp.json()
            plan_list = result.get("info", {}).get("proposalPlanVOList", [])
            if plan_list and len(plan_list) > 0:
                plan_vo = plan_list[0]
            else:
                plan_vo = result.get("info", {})

            if plan_vo.get("serialNo"):
                self.serial_no = str(plan_vo["serialNo"])
            if plan_vo.get("proposalId"):
                self.proposal_id = str(plan_vo["proposalId"])
        except (ValueError, AttributeError, KeyError):
            pass

        return success, resp.cookies, resp.status_code, resp.text

    # ================================================================
    # 计划费率计算 (saveProductExt)
    # ================================================================

    def save_product(
        self,
        plan: int,
        ensure_period: str,
        pay_period,
        amount: int,
        ensure_plan: str = "1",
        duty_list: list = None,
        cookies=None,
    ) -> tuple:
        """
        POST 计划费率计算接口。

        Args:
            plan: 责任计划编号 (0-7)
            ensure_period: 保险期间标签（"终身"）
            pay_period: 交费期间（1, 3, 5, 10, 15, 20, 30）
            amount: 保额（元）
            ensure_plan: 承保方案（"1"=标准体, "2"=优选体）
            duty_list: 责任选项列表（None=自动从配置构建）
            cookies: 指定 Cookies

        Returns:
            (success: bool, fee: float|None, status_code: int,
             response_text: str, failure_reason: str|None)
        """
        url = self.profile.plan_rate_url
        d = self.profile.defaults

        # 自动构建责任列表
        if duty_list is None:
            duty_list = self.profile.get_duties_for_plan(plan)

        payload = {
            "serialNo": self.serial_no,
            "productId": self.profile.product.product_id,
            "companyId": self.profile.product.company_id,
            "proposalId": self.proposal_id,
            "dividendDrawType": d.dividend_draw_type,
            "premium": "",
            "ensurePeriodCode": self.profile.get_ensure_period_code(ensure_period),
            "payPeriodCode": self.profile.get_pay_period_code(pay_period),
            "payModeCode": self.profile.get_pay_mode_code(pay_period),
            "ensurePlan": ensure_plan,
            "amountDescr": str(amount),
            "amount": str(amount),
            "fee": "",
            "dutyOptionList": duty_list,
        }

        headers = {"Content-Type": "application/json"}
        resp = self.session.post(
            url, data=json.dumps(payload), headers=headers,
            cookies=cookies or self.login_cookies,
            verify=self.profile.api.verify_ssl,
            timeout=self.profile.api.timeout,
        )

        fee = None
        failure_reason = None
        success = False

        try:
            result = resp.json()
            code = result.get("code")

            # 从响应提取 fee
            fee = self._extract_fee(result)

            # 更新动态参数
            info = result.get("info", {})
            if isinstance(info, dict):
                if info.get("serialNo"):
                    self.serial_no = str(info["serialNo"])
                if info.get("proposalId"):
                    self.proposal_id = str(info["proposalId"])

            # 成功判断
            if code == 200 and fee is not None and float(fee) > 0:
                failure_reason = result.get("info", {}).get("failureReason") if isinstance(result.get("info"), dict) else None
                if not failure_reason:
                    success = True
            elif code == 500:
                failure_reason = result.get("message", "服务器内部错误")
            else:
                # code==200 但 fee 为空/为0 → 检查 failureReason
                if isinstance(result.get("info"), dict):
                    fr = result["info"].get("failureReason")
                    if fr:
                        failure_reason = fr
                if not failure_reason:
                    failure_reason = result.get("message")
                # fee=0 且无 failureReason/message 说明 API 实际计算失败,
                # 不设置 success=True, 由上层判定为 FAIL

        except (ValueError, AttributeError):
            pass

        # DEBUG: 失败时写入日志文件
        if not success:
            import os as _os, json as _json
            _log_dir = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "..", "_uploads")
            _os.makedirs(_log_dir, exist_ok=True)
            _log_path = _os.path.join(_log_dir, "_debug_failure.json")
            _entry = {
                "failure_reason": failure_reason,
                "http_status": resp.status_code,
                "request_payload": {k: v for k, v in payload.items()},
                "response_body": str(resp.text)[:3000],
            }
            with open(_log_path, "w", encoding="utf-8") as _f:
                _json.dump(_entry, _f, ensure_ascii=False, indent=2)
            import sys
            print(f"[DEBUG] Failure logged to: {_log_path}", file=sys.stderr, flush=True)

        return success, fee, resp.status_code, resp.text, failure_reason

    # ================================================================
    # 便捷方法：完整测试流程
    # ================================================================

    def run_case(
        self,
        plan: int,
        ensure_period: str,
        pay_period,
        amount: int,
        insurant_age: int,
        insurant_sex: str,
        ensure_plan: str = "1",
    ) -> dict:
        """
        执行单个测试用例的完整流程（age_rate → plan_rate）。

        Returns:
            包含所有步骤结果的字典
        """
        result = {
            "age_success": False,
            "age_status": None,
            "age_text": "",
            "plan_success": False,
            "plan_status": None,
            "plan_text": "",
            "fee": None,
            "failure_reason": None,
            "error": None,
        }

        try:
            # Step 1: age_rate
            ok, cookies, status, text = self.save_customer(insurant_age, insurant_sex)
            result["age_success"] = ok
            result["age_status"] = status
            result["age_text"] = text

            if not ok:
                result["error"] = "age_rate接口失败"
                return result

            # Step 2: plan_rate
            ok, fee, status, text, reason = self.save_product(
                plan=plan,
                ensure_period=ensure_period,
                pay_period=pay_period,
                amount=amount,
                ensure_plan=ensure_plan,
                cookies=cookies,
            )
            result["plan_success"] = ok
            result["plan_status"] = status
            result["plan_text"] = text
            result["fee"] = fee
            result["failure_reason"] = reason

        except Exception as e:
            result["error"] = str(e)

        return result

    # ================================================================
    # 内部工具方法
    # ================================================================

    def _check_success(self, resp: requests.Response) -> bool:
        """检查 API 响应是否成功"""
        try:
            result = resp.json()
            if result.get("code") in (0, 200) or result.get("success"):
                return True
        except (ValueError, AttributeError):
            pass
        return "成功" in resp.text

    @staticmethod
    def _extract_fee(result: dict) -> Optional[float]:
        """
        从 API 响应中多层次提取 fee 字段。

        尝试路径:
        1. result.info.fee
        2. result.info.proposalPlanVOList[0].fee
        3. result.fee
        """
        info = result.get("info", {})

        # 路径 1
        if isinstance(info, dict) and "fee" in info:
            try:
                return float(info["fee"])
            except (ValueError, TypeError):
                pass

        # 路径 2
        if isinstance(info, dict):
            plan_list = info.get("proposalPlanVOList", [])
            if plan_list and isinstance(plan_list, list):
                plan_vo = plan_list[0]
                if isinstance(plan_vo, dict) and "fee" in plan_vo:
                    try:
                        return float(plan_vo["fee"])
                    except (ValueError, TypeError):
                        pass

        # 路径 3
        if "fee" in result:
            try:
                return float(result["fee"])
            except (ValueError, TypeError):
                pass

        return None

    def reset_session(self, serial_no: str = None, proposal_id: str = ""):
        """重置会话状态"""
        self.session = requests.Session()
        self.login_cookies = None
        if serial_no is not None:
            self.serial_no = serial_no
        self.proposal_id = proposal_id
