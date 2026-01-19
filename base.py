import json
import sys
import time
import re
from pathlib import Path

import requests
import datetime


def difftime_get_config():  # difftime.py读取配置
    difftime_config = config.get("difftime", {})
    num_diff = int(difftime_config.get("num_diff"))

    minrev_val = difftime_config.get("minrev")
    minrev = int(minrev_val) if minrev_val and minrev_val.strip() != '' else None
    maxrev_val = difftime_config.get("maxrev")
    maxrev = int(maxrev_val) if maxrev_val and maxrev_val.strip() != '' else None

    return num_diff, minrev, maxrev


def usercontribs_get_config():  # usercontribs.py读取配置
    usercontribs_config = config.get("usercontribs", {})
    queryusername = usercontribs_config.get("username")
    ucend = usercontribs_config.get("starttime")
    ucstart = usercontribs_config.get("endtime")
    return queryusername, ucend, ucstart


def editperiod_get_config():  # editperiod.py读取配置
    editperiod_config = config.get("editperiod", {})
    datafile = editperiod_config.get("datafile")
    file_path = Path("output") / f"{datafile}.json"
    try:
        with open(file_path, "r", encoding="utf-8") as contribs_file:
            contribs_data = json.load(contribs_file)
    except FileNotFoundError:
        print("指定的文件不存在！")
        input("按回车键退出")
        sys.exit(1)
    return datafile, contribs_data


def activeusers_get_config():  # activeusers.py读取配置
    activeusers_config = config.get("activeusers", {})
    usergroup_order = activeusers_config.get("usergroup_order")
    usergroup_mapping = activeusers_config.get("usergroup_mapping")
    mode = activeusers_config.get("mode")
    if mode not in ["standard", "debug"]:
        print("指定的模式不存在，请检查配置文件。")
        input("按回车键退出")
        sys.exit(1)
    return usergroup_order, usergroup_mapping, mode


def checkhiddenrc_get_config():  # checkhiddenrc.py读取配置
    checkhiddenrc_config = config.get("checkhiddenrc", {})
    endtime = checkhiddenrc_config.get("starttime")
    starttime = checkhiddenrc_config.get("endtime")
    return starttime, endtime


def patrolintegration_get_config():  # patrolintegration.py读取配置
    patrolintegration_config = config.get("patrolintegration", {})
    starttime = patrolintegration_config.get("starttime")
    endtime = patrolintegration_config.get("endtime")
    return starttime, endtime


def login():  # 尝试登录
    global session

    if username and password:
        session = requests.Session()
        session.headers.update({"User-Agent": user_agent})

        # 获取登录token
        r1 = session.get(WIKI_API_URL, params={
            'action': 'query',
            'meta': 'tokens',
            'type': 'login',
            'format': 'json'
        })
        login_token = r1.json()['query']['tokens']['logintoken']

        # 提交用户名和密码
        r2 = session.post(WIKI_API_URL, data={
            'action': 'login',
            'lgname': username,
            'lgpassword': password,
            'lgtoken': login_token,
            'format': 'json'
        })

        # 检查登录结果
        result = r2.json()
        if result['login']['result'] == 'Success':
            print("登录成功")
        else:
            print("登录失败：", result['login'])
            session = None

    else:
        print("未提供用户名和密码，将以未登录状态运作")
        session = None
    
    return session


def get_data(params):  # 从Mediawiki API获取数据
    tries = 0
    while 1:
        try:
            if session:
                response = session.get(WIKI_API_URL, params=params)
            else:
                response = requests.get(WIKI_API_URL, params=params, headers={"User-Agent": user_agent})

            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException:
            tries += 1
            if tries > max_retries:
                break

            print(f"未获取到数据，{retry_delay}秒后重试。")
            time.sleep(retry_delay)

    print("重试失败，请检查网络连接。")
    input("按回车键退出")
    sys.exit(1)


def get_last_diff():  # 获取当前最大revid
    params = {
        "action": "query",
        "format": "json",
        "formatversion": 2,
        "list": "recentchanges",
        "rcprop": "ids",
        "rclimit": 1,
        "rctype": "edit|new"
    }
    return get_data(params)


def extract_from_timestamp(timestamp_str):  # 提取日期、时间数据，将UTC时间改为当前时区
    # 时间格式：2025-06-16T12:24:55Z
    dt = datetime.datetime.strptime(timestamp_str, "%Y-%m-%dT%H:%M:%SZ")

    dt = dt + datetime.timedelta(hours=timezone)

    return dt


def is_ip_address(s):
    ipv4_pattern = r'^(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])\.'
    ipv4_pattern += r'(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])\.'
    ipv4_pattern += r'(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])\.'
    ipv4_pattern += r'(25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9]?[0-9])$'

    # IPv6仅匹配大写和完整格式
    ipv6_pattern = r'^([0-9A-F]{1,4}:){7}[0-9A-F]{1,4}$'

    return (
            re.match(ipv4_pattern, s) is not None or
            re.match(ipv6_pattern, s) is not None
    )


def output(filename, data, extension):  # 将文件输出到output文件夹
    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)
    output_filename = out_dir / filename
    if extension == "xlsx":
        data.save(output_filename)
    elif extension == "json":
        with open(output_filename, "w", encoding="utf-8") as output_file:
            json.dump(data, output_file, ensure_ascii=False, indent=4)
    else:
        with open(output_filename, "w", encoding="utf-8") as output_file:
            output_file.write(data)


def sleep():  # 请求间隔
    if request_interval > 0:
        time.sleep(request_interval)


with open("config.json", "r", encoding="utf-8") as config_file:
    config = json.load(config_file)
    wiki = config["wiki"]
    user_agent = config["user_agent"]
    timezone = int(config["timezone"])
    request_interval = float(config["request_interval"])
    max_retries = int(config["max_retries"])
    retry_delay = int(config["retry_delay"])
    username = config["username"] if "username" in config else None
    password = config["password"] if "password" in config else None

wiki_lang = ['de', 'en', 'es', 'fr', 'it', 'ja', 'ko', 'lzh', 'nl', 'pt', 'ru', 'th', 'uk', 'zh', 'meta']

if wiki not in wiki_lang:
    print("不存在此语言的Minecraft Wiki！")
    input("按回车键退出")
    sys.exit(1)

elif wiki == 'en':
    WIKI_BASE_URL = "https://minecraft.wiki"

else:
    WIKI_BASE_URL = f"https://{wiki}.minecraft.wiki"

WIKI_API_URL = WIKI_BASE_URL + "/api.php"

# session = login()
