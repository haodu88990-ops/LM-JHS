
# 解析plan_rate完整响应，找fee字段位置
import requests, json, warnings
warnings.filterwarnings("ignore")

BASE_URL = "https://kfzxtb.lmbaoxian.com:13080"
session = requests.Session()

# 登录
resp = session.post(f"{BASE_URL}/broker/api/user/login.html",
    data={"account":"15856990088","password":"dc483e80a7a0bd9ef71d8cf973673924"},
    headers={"Content-Type":"application/x-www-form-urlencoded; charset=UTF-8","X-Requested-With":"XMLHttpRequest"},
    verify=False)
print("Login:", resp.status_code)
login_cookies = resp.cookies

# age_rate
resp2 = session.post(f"{BASE_URL}/broker/api/prospectus/saveCustomer.html",
    json={"insurantName":"","insurantAge":0,"insurantSex":"1","insurantOccLevel":1,
          "insurantSocialInsurance":"1","insurantId":85320,"insurantBirthday":"",
          "policyHolderAge":30,"policyHolderSex":"1","policyHolderId":85321,
          "policyHolderBirthday":"","serialNo":"1491455820292947968","type":"prospectus"},
    headers={"Content-Type":"application/json","X-Requested-With":"XMLHttpRequest"},
    cookies=login_cookies, verify=False)
print("AgeRate:", resp2.status_code)
ar_cookies = resp2.cookies if resp2.cookies else login_cookies

# plan_rate 完整响应
resp3 = session.post(f"{BASE_URL}/broker/api/prospectus/saveProductExt.html",
    json={"serialNo":"1495810487793745920","productId":"991452","companyId":"100080",
          "proposalId":"40472","dividendDrawType":"2","premium":"",
          "ensurePeriodCode":"TO105","payPeriodCode":"1","payModeCode":"1",
          "ensurePlan":"0","amountDescr":"1000000","amount":"1000000",
          "fee":"","dutyOptionList":[]},
    headers={"Content-Type":"application/json","X-Requested-With":"XMLHttpRequest"},
    cookies=ar_cookies, verify=False)

print("\n=== plan_rate 完整响应 ===")
print(resp3.status_code)
try:
    data = resp3.json()
    print(json.dumps(data, ensure_ascii=False, indent=2))
except:
    print(resp3.text)
