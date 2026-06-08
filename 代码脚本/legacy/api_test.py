# -*- coding: utf-8 -*-
"""
接口测试类
用于测试保险经纪系统相关接口
"""

import requests
import json


class SimpleAPITest:
    """简洁版接口测试类"""

    def __init__(self):
        self.base_url = "https://kfzxtb.lmbaoxian.com:13080"
        self.session = requests.Session()
        self.login_cookies = None
        self.age_rate_cookies = None

    def login(self, account="15856990088", password="dc483e80a7a0bd9ef71d8cf973673924"):
        """登录接口 - 尝试多种提交方式"""
        url = f"{self.base_url}/broker/api/user/login.html"
        
        # 尝试方式1: 表单格式
        data = {
            "account": account,
            "password": password
        }
        
        headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest"
        }
        
        response = self.session.post(url, data=data, headers=headers, verify=False)
        
        # 如果失败，尝试方式2: JSON格式
        if "不能为空" in response.text:
            headers["Content-Type"] = "application/json"
            response = self.session.post(url, json=data, headers=headers, verify=False)
        
        self.login_cookies = response.cookies
        
        print(f"\n{'='*50}")
        print("【登录接口】请求信息:")
        print(f"URL: {url}")
        print(f"请求参数: {data}")
        print(f"Content-Type: {headers['Content-Type']}")
        print(f"{'='*50}")
        print("响应信息:")
        print(f"状态码: {response.status_code}")
        print(f"Cookie: {dict(response.cookies)}")
        print(f"响应内容: {response.text}")
        print(f"{'='*50}")
        
        return response

    def age_rate(self, cookies=None):
        """年龄_费率接口"""
        url = f"{self.base_url}/broker/api/prospectus/saveCustomer.html"

        payload = \
            {
                "insurantName": "",
                "insurantAge": insurantAge,
                "insurantSex": insurantSex,
                "insurantOccLevel": 1,
                "insurantSocialInsurance": insurantSocialInsurance,
                "insurantId": 85320,
                "insurantBirthday": "",
                "policyHolderAge": 30,
                "policyHolderSex": "1",
                "policyHolderId": 85321,
                "policyHolderBirthday": "",
                "serialNo": "1491455820292947968",
                "type": "prospectus"
            }
        
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest"
        }

        cookies = cookies or self.login_cookies
        response = self.session.post(url, json=payload, headers=headers, cookies=cookies, verify=False)
        self.age_rate_cookies = response.cookies
        
        print(f"\n{'='*50}")
        print("【年龄_费率接口】请求信息:")
        print(f"URL: {url}")
        print(f"请求参数: {json.dumps(payload, ensure_ascii=False)}")
        print(f"Cookie: {dict(cookies) if cookies else 'None'}")
        print(f"{'='*50}")
        print("响应信息:")
        print(f"状态码: {response.status_code}")
        print(f"响应内容: {response.text}")
        print(f"{'='*50}")
        
        return response

    def plan_rate(self, cookies=None):
        """计划_费率接口"""
        url = f"{self.base_url}/broker/api/prospectus/saveProductExt.html"
        
        payload = {
            "serialNo": "1495810487793745920",
            "productId": "991452",
            "companyId": "100080",
            "proposalId": "40472",
            "dividendDrawType": "2",
            "premium": "",
            "ensurePeriodCode": "TO105",
            "payPeriodCode": "1",
            "payModeCode": "1",
            "ensurePlan": "1",
            "amountDescr": "1010000",
            "amount": "1010000",
            "fee": "",
            "dutyOptionList": [

            ]
        }
        
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest"
        }

        cookies = cookies or self.age_rate_cookies
        response = self.session.post(url, json=payload, headers=headers, cookies=cookies, verify=False)
        
        print(f"\n{'='*50}")
        print("【计划_费率接口】请求信息:")
        print(f"URL: {url}")
        print(f"请求参数: {json.dumps(payload, ensure_ascii=False)}")
        print(f"Cookie: {dict(cookies) if cookies else 'None'}")
        print(f"{'='*50}")
        print("响应信息:")
        print(f"状态码: {response.status_code}")
        print(f"响应内容: {response.text}")
        print(f"{'='*50}")
        
        return response


# ======== 使用示例 ========
if __name__ == "__main__":
    import warnings
    warnings.filterwarnings("ignore")

    print("="*60)
    print("开始执行接口测试...")
    print("="*60)
    
    api = SimpleAPITest()
    
    # Step 1: 登录
    print("\n>>> Step 1: 执行登录接口...")
    api.login()
    
    # Step 2: 年龄费率（使用登录cookie）
    print("\n>>> Step 2: 执行年龄_费率接口...")
    api.age_rate()
    
    # Step 3: 计划费率（使用年龄费率cookie）
    print("\n>>> Step 3: 执行计划_费率接口...")
    api.plan_rate()
    
    print("\n" + "="*60)
    print("所有接口测试执行完成!")
    print("="*60)
