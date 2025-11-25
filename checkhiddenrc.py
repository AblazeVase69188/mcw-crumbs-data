import time
from bisect import bisect_left

import base

starttime, endtime = base.checkhiddenrc_get_config()

loop = 0

arv_data = []  # (revid)
rc_data = []  # (revid, logparams)

base_params = {
    "action": "query",
    "format": "json",
    "list": "allrevisions|recentchanges",
    "formatversion": 2,
    "arvprop": "ids",
    "arvlimit": 500,
    "rcprop": "ids|loginfo",
    "rclimit": 500,
    "rctype": "edit|new|log"
}

if starttime != "":
    base_params.update({
        "arvstart": starttime,
        "rcstart": starttime
    })

if endtime != "":
    base_params.update({
        "arvend": endtime,
        "rcend": endtime
    })

last_rccontinue = ""
last_arvcontinue = ""

base.login()

print("启动成功", end='\n\n')

while True:
    time.sleep(3)

    if last_rccontinue != "":  # 不是首次循环，使用这个继续
        base_params.update({
            "rccontinue": last_rccontinue,
            "arvcontinue": last_arvcontinue
        })

    data = base.get_data(base_params)

    loop += 1
    print(f"成功获取第{loop}组数据")

    for item in data['query']['allrevisions']:
        for rev in item['revisions']:
            arv_data.append(rev['revid'])

    for item in data['query']['recentchanges']:
        if 'actionhidden' not in item:
            rc_data.append((
                item['revid'],
                item.get('logparams', "")
            ))

    continue_data = data.get('continue', {})
    last_rccontinue = continue_data.get('rccontinue', "")
    last_arvcontinue = continue_data.get('arvcontinue', "")

    if not last_rccontinue or not last_arvcontinue:
        break

print("所有数据已经获取完毕，正在处理中...")

arv_data.sort()
rc_data.sort(key=lambda x: x[0])
min_revid = next(item[0] for item in rc_data if item[0] != 0)

# rc还包含日志，所以必定比相同数量的arv多，因此截断arv的前面部分
idx = bisect_left(arv_data, min_revid)
del arv_data[:idx]

# 所有出现在rc_data中的revid
rc_revids = {item[0] for item in rc_data}

# 特判：移动日志保留重定向（"suppressredirect": false）也只显示一条，revid+1的项为重定向页面的创建
exclude_revids = set()
for revid, logparams in rc_data:
    if logparams and logparams.get('suppressredirect') == False:
        exclude_revids.add(revid + 1)

# 最终筛选
hidden_data = []
for revid in arv_data:
    # 在arv_data但不在rc_data中，不是需要排除的特殊revid
    if revid not in rc_revids and revid not in exclude_revids:
        hidden_data.append(revid)

for revid in hidden_data:
    print(f"{base.WIKI_BASE_URL}/?diff={revid}")
