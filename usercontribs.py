import time
import sys
from datetime import datetime
import json

import base

username, ucend, ucstart = base.usercontribs_get_config()

params = {
    "action": "query",
    "format": "json",
    "list": "usercontribs",
    "formatversion": 2,
    "uclimit": "max",
    "ucuser": username,
    "ucprop": "ids|title|timestamp|comment|size|flags|sizediff|tags",
}

if ucend != "":
    params.update({"ucend": ucend})

if ucstart != "":
    params.update({"ucstart": ucstart})

contribs_result = []
last_uccontinue = ""
loop_count = 0

current_time = datetime.now().strftime("%Y%m%d%H%M%S")

base.login()

print("启动成功", end='\n\n')

while True:
    base.sleep()

    if last_uccontinue != "":
        last_params = params.copy()
        last_params.update({"uccontinue": last_uccontinue})
    else:
        last_params = params

    contribs_data = base.get_data(last_params)

    if 'error' in contribs_data:
        print("API返回错误信息，请检查时间戳格式是否正确。")
        input("按回车键退出")
        sys.exit(1)

    if not contribs_data['query']['usercontribs']:
        print("API未返回有效数据，请检查用户是否存在、用户是否有贡献或时间范围是否正确。")
        input("按回车键退出")
        sys.exit(1)

    contribs_result.extend(contribs_data['query']['usercontribs'])

    loop_count += 1
    print(f"成功获取第{loop_count}组数据")

    if 'continue' not in contribs_data:
        break

    last_uccontinue = contribs_data['continue']['uccontinue']

contribs_result = contribs_result[::-1]

base.output(f'usercontribs-{username}-{current_time}.json', contribs_result, "json")

print(f"结果已保存至usercontribs-{username}-{current_time}.json")
