# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r'f:\workbuddy\test_Jhs')
from batch_test import load_test_cases

cases = load_test_cases()
print(f'共加载 {len(cases)} 条用例')
for c in cases[:6]:
    print(f"  计划{c['责任计划']} {c['保险期间']} 交{c['交费期间']}年 {c['性别']} {c['年龄']}岁({c['年龄类型']}) 费率={c['期望费率']}")
print('...')
for c in cases[-4:]:
    print(f"  计划{c['责任计划']} {c['保险期间']} 交{c['交费期间']}年 {c['性别']} {c['年龄']}岁({c['年龄类型']}) 费率={c['期望费率']}")
