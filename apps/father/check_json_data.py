import json
import os

file_path = "C:/temp/output/西村房屋信息/"


def check():
    """
    对住宅信息的JSON文件进行检查。
    因为为了省事，没有对住宅信息全部分歧做判断，所以检查一下没有漏判断的。
    """

    for root, dirs, files in os.walk(file_path):
        # 遍历文件
        for file in files:
            detail = json.load(open(file_path + file, 'r', encoding='utf-8'))
            # 产权人
            AA = detail["data"]["propertyOwner"]
            # 身份证号
            AB = detail["data"]["certNumber"]
            # 建筑层数
            buildStorey = detail["data"]["buildStorey"]
            if buildStorey == "1":
                BA = "■"
            elif buildStorey == "2":
                BB = "■"
            elif buildStorey == "3":
                BC = "■"
            # 建筑面积
            BD = detail["data"]["buildArea"]
            # 建造年代
            buildYear = detail["data"]["buildYear"]
            if buildYear == "1":
                CA = "■"
            elif buildYear == "2":
                CB = "■"
            elif buildYear == "3":
                CC = "■"
            elif buildYear == "4":
                CD = "■"
            elif buildYear == "5":
                CE = "■"
            # 结构类型
            structType = detail["data"]["structType"]
            if structType == "1":  # 砖混结构（预制楼板）
                DA = "■"
            elif structType == "2":  # 砖混结构（现浇楼板） / 砖石结构（非预制板）
                DB = "■"
            elif structType == "3":  # 土木结构
                DC = "■"
            elif structType == "6":  # 钢结构
                DD = "■"
            elif structType == "7":  # 土木结构
                DE = "■"
            else:
                print("出现其他结构类型：" + AA)
                break
            # 建造方式
            buildMode = detail["data"]["buildMode"]
            buildModeOther = detail["data"]["buildModeOther"]
            if buildMode == "1":
                EA = "■"
            elif buildMode == "2":
                EB = "■"
            elif buildMode == "3":
                EC = "■"
            elif buildMode == "4":
                ED = "■"
                EE = buildModeOther
            # 土地性质
            landType = detail["data"]["landType"]
            if landType == "1":
                FA = "■"
            elif landType == "2":
                FB = "■"
            # 宅基地手续
            homesteadPermit = detail["data"]["homesteadPermit"]
            if homesteadPermit == "1":
                GA = "■"
            elif homesteadPermit == "2":
                GB = "■"
            else:
                print("出现其他宅基地手续：" + AA)
                break
            # 规划建设手续
            planPermit = detail["data"]["planPermit"]
            if planPermit == "1":
                GC = "■"
            elif planPermit == "2":
                GD = "■"
            else:
                print("出现其他规划建设手续：" + AA)
                break
            # 竣工验收手续
            completePermit = detail["data"]["completePermit"]
            if completePermit == "1":
                HA = "■"
            elif completePermit == "2":
                HB = "■"
            else:
                print("出现其他竣工验收手续：" + AA)
                break
            # 房屋登记手续
            houseRegisterPermit = detail["data"]["houseRegisterPermit"]
            if houseRegisterPermit == "1":
                HC = "■"
            elif houseRegisterPermit == "2":
                HD = "■"
            else:
                print("出现其他竣工验收手续：" + AA)
                break
            # 是否改造
            houseTransformType = detail["data"]["houseTransformType"]
            if houseTransformType == "1":
                IB = "■"  # 是
                # 改造回数
                houseTransformNum = detail["data"]["houseTransformNum"]
                if houseTransformNum == "1":
                    IC = "■"
                elif houseTransformNum == "2":
                    ID = "■"
                else:
                    print("出现其他改造回数：" + AA)
                    break
                # 改造内容
                houseTransformContent = detail["data"]["houseTransformContent"]
                if len(houseTransformContent) > 1:
                    print("出现多次改造内容：" + AA)
                    break
                if houseTransformContent[0] == "1":
                    JA = "■"
                elif houseTransformContent[0] == "2":
                    JB = "■"
                else:
                    print("出现其他改造内容：" + AA)
                    break
            elif houseTransformType == "2":
                IA = "■"  # 否
            else:
                print("出现其他是否改造：" + AA)
                break
            # 是否用作经营
            manageHave = detail["data"]["manageHave"]
            if manageHave == "1":
                KA = "■"  # 是
                # 经营审批手续
                managePermit = detail["data"]["managePermit"]
                if managePermit == "1":
                    LA = "■"
                elif managePermit == "2":
                    LB = "■"
                else:
                    print("出现其他经营审批手续：" + AA)
                    break
                # 主要用途
                manageType = detail["data"]["manageType"]
                if len(manageType) > 1:
                    print("出现多次主要用途：" + AA)
                    break
                if manageType[0] == "1":
                    MA = "■"
                elif manageType[0] == "3":
                    MB = "■"
                elif manageType[0] == "4":
                    MC = "■"
                else:
                    print("出现其他主要用途：" + AA)
                    break
            elif manageHave == "2":
                KB = "■"  # 否
            else:
                print("出现其他是否用作经营：" + AA)
                break
            # 采暖用能
            heatEnergy = detail["data"]["heatEnergy"]
            if "1" in heatEnergy:
                NA = "■"
            if "2" in heatEnergy:
                NB = "■"
            if "3" in heatEnergy:
                NC = "■"
            if "4" in heatEnergy:
                ND = "■"
            if "5" in heatEnergy:
                NE = "■"
                NF = detail["data"]["heatEnergyOther"]
                break
            # 炊事用能
            cookEnergy = detail["data"]["cookEnergy"]
            if "1" in cookEnergy:
                OA = "■"
            if "2" in cookEnergy:
                OB = "■"
            if "3" in cookEnergy:
                OC = "■"
            if "4" in cookEnergy:
                OD = "■"
            if "5" in cookEnergy:
                OE = "■"
                OF = detail["data"]["cookEnergyOther"]
                break
            # 安全隐患初判
            houseSafeJudge = detail["data"]["houseSafeJudge"]
            if houseSafeJudge == "2":
                print("出现其他安全隐患初判：" + AA)


def main():
    check()


if __name__ == '__main__':
    main()
