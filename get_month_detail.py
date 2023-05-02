import configparser
import json

import openpyxl
import requests


def get_data(uid, month, current_page, cur_page_size, cookie, reward_type):
    url = "https://api-takumi.mihoyo.com/event/srledger/month_detail"
    params = {
        "uid": uid,
        "region": "prod_gf_cn",
        "month": month,
        "type": reward_type,
        "current_page": current_page,
        "page_size": cur_page_size
    }
    headers = {
        "Referer": "https://webstatic.mihoyo.com/",
        "Cookie": cookie
    }
    response = requests.get(url, headers=headers, params=params)
    data = json.loads(response.text)
    if data["retcode"] != 0:
        if data["message"] == "请先登录":
            print("Cookie已过期，请重新输入")
            get_cookie(True)
            main()
        else:
            raise Exception(data["message"])
    return data["data"]


def get_cookie(need_new=False):
    config_file = f"{path}/config.ini"
    config = configparser.ConfigParser()
    config.read(config_file)
    if "cookie" in config and not need_new:
        return config["cookie"]["value"]

    cookie = input("请输入Cookie：")
    config["cookie"] = {"value": cookie}
    with open(config_file, "w") as f:
        config.write(f)
    return cookie


def get_all_data_by_reward_type(uid, month, reward_type):
    cookie = get_cookie()
    total = get_data(uid, month, 1, 1, cookie, reward_type)["total"]
    result = []
    total_page_num = (total - 1) // page_size + 2
    for i in range(1, total_page_num):
        print(f"{i}/{total_page_num - 1}\t{reward_type_list[reward_type - 1]}")
        data = get_data(uid, month, i, page_size, cookie, reward_type)["list"]
        result.extend(data)
    return result


path = "month_detail_data/"
reward_type_list = ["星琼", "星轨通票&星轨专票"]
page_size = 100


def main():
    uid = "100205005"
    month = "202304"

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for i in range(len(reward_type_list)):
        result = get_all_data_by_reward_type(uid, month, i + 1)
        ws = wb.create_sheet(reward_type_list[i])
        ws.append(["action", "action_name", "time", "num"])
        for item in result:
            ws.append([item["action"], item["action_name"], item["time"], item["num"]])
        wb.save(f"{path}/{uid}_{month}.xlsx")


if __name__ == "__main__":
    main()
