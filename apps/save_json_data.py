import json
import os

import requests


new_path = "C:/temp/output/西村房屋信息/"
authorization = "Bearer eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiI4NjEzNjdkOC02OWNkLTRhMmEtYTA0Zi0xNzc5ZTgzNWQ1ODgiLCJleHAiOjE2MjU1OTAxODAsImlhdCI6MTYyMzg2MjE4MH0.AN7PAZ6VMVBZjbdJ55x7Bb7EpGXlrvXl9cSfSagNbNeaXaPG5A5p4XBaNCUnR0b7-Rcv5pGsT3T3ICN70KkZHA"
url = "https://ruralhouse.cubigdata.cn/houseInfoManage/queryHouseBaseInfoDetailById?id="


def save():

    # 加载事先取得的，包含所有人员信息的JSON文件。
    users = json.load(open('C:/temp/input/users.txt', 'r', encoding='utf-8'))
    print("共 " + str(len(users['data']['records'])) + " 名人员。")

    for user in users['data']['records']:
        response = requests.get(url + user['id'], headers={'Authorization': authorization})
        if response.status_code != 200:
            print('查询出错！')
            break
        detail = json.loads(response.text)
        # 产权人
        property_owner = detail["data"]["propertyOwner"]
        # 身份证号
        cert_number = detail["data"]["certNumber"]

        file_path = new_path + property_owner + "（" + user['id'] + "）.json"
        if os.path.exists(file_path):
            print('文件已存在：' + property_owner + "（" + user['id'] + "）")
            continue

        target = open(file_path, 'w', encoding='utf-8')
        target.write(response.text)
        target.close()


def main():
    save()


if __name__ == '__main__':
    main()
