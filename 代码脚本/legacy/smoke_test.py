# -*- coding: utf-8 -*-
import sys, time
sys.path.insert(0, r'f:\workbuddy\test_Jhs')
from batch_test import BatchAPITest, load_test_cases, build_duty_option_list, ENSURE_PERIOD_MAP, pay_period_code, pay_mode_code, SEX_MAP, PLAN_DUTY_CODES
import random

cases = load_test_cases()[:5]
api = BatchAPITest()

print(">>> 登录...")
ok, _ = api.login()
print(f"  登录结果: {ok}")

print("\n>>> 冒烟测试前5条...")
for i, c in enumerate(cases, 1):
    plan = c['责任计划']
    gender = c['性别']
    age = c['年龄']
    expected_rate = c['期望费率']
    pay_period = c['交费期间']
    ensure_period = c['保险期间']
    amount = 1000000
    duty_list = build_duty_option_list(plan)

    ar_ok, ar_cookies, ar_code, ar_text = api.age_rate(age, gender)
    pr_ok, fee, pr_code, pr_text, failure = api.plan_rate(
        plan, ensure_period, pay_period, amount, duty_list, cookies=ar_cookies
    )

    expected_premium = amount * expected_rate / 1000
    if fee and fee > 0:
        diff_pct = abs(fee - expected_premium) / expected_premium * 100
        status = "PASS" if diff_pct < 1 else f"偏差{diff_pct:.1f}%"
    else:
        status = f"FAIL: {failure}"

    print(f"  [{i}] 计划{plan} 交{pay_period}年 {gender}{age}岁 保额{amount:,} -> fee={fee}, 期望={expected_premium:.1f}, {status}")
    time.sleep(0.3)
