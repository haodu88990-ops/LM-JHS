# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r'f:\workbuddy\test_Jhs')
from batch_test import load_test_cases

cases = load_test_cases()
print(f'共加载 {len(cases)} 条用例')
std = [c for c in cases if c.get('ensurePlan') == '1']
opt = [c for c in cases if c.get('ensurePlan') == '2']
print(f'  标准体: {len(std)} 条')
print(f'  优选体: {len(opt)} 条')
print()
print('--- 标准体前3条 ---')
for c in std[:3]:
    print(f"  {c['保障方案']}(ensurePlan={c['ensurePlan']}) 计划{c['责任计划']} 交{c['交费期间']}年 {c['性别']} {c['年龄']}岁({c['年龄类型']}) 费率={c['期望费率']}")
print()
print('--- 优选体前3条 ---')
for c in opt[:3]:
    print(f"  {c['保障方案']}(ensurePlan={c['ensurePlan']}) 计划{c['责任计划']} 交{c['交费期间']}年 {c['性别']} {c['年龄']}岁({c['年龄类型']}) 费率={c['期望费率']}")
