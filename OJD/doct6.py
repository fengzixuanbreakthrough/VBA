from os import name
import xlrd
import dictocin as dtc
#import  canoe as can
result=[]
routingmessages=[]


def panduanfasong(fasong, change, i, mouID, result2, srcflag,
                  alldata):  # change可以为源周期，或者是某一网段的周期,srcflag指示是否是对源网段的发送周期进行修改，以区别改源/目标
    global result
    if type(fasong) == type(1.2):  # 判断发送周期类型为float
        print('>>>>>1')
        mouID[change] = int(fasong)
    elif type(fasong) == type(int(1)):  # 判断发送周期类型为int
        print('>>>>>2')
        mouID[change] = int(fasong)
    elif is_number(fasong):  # 若遇到类似于cycle time='30'时
        print('>>>>>3')
        mouID[change] = int(fasong)
    # 处理事件后周期触发这种情况，写"当"时
    elif '当' in fasong:
        print('>>>>>4')
        print("里面有当字")
        print(alldata[i])
        print(alldata[i][3])
        print(type(alldata[i][3]))
        if srcflag == 1:
            guodutime = alldata[i][3].split('当')
        else:
            guodutime = alldata[i][-2].split('当')
        mouID[change] = int(str(guodutime[0]))
        print(22222)
    # 处理事件后周期触发这种情况，写"when"时
    elif 'when' in fasong:
        print('>>>>>5')
        if srcflag == 1:
            guodutime = alldata[i][3].split('when')
        else:
            guodutime = alldata[i][-2].split('when')
        mouID[change] = int(str(guodutime[0]))
    elif '事件' in fasong or 'Event' in fasong or 'event' in fasong:
        print('>>>>>6')
        mouID[change] = int(10)
    elif fasong == '':
        print('>>>>>7')
        # mouID[change]='未填写'

        mouID[change] = '未填'
        result2.append("第" + str(i + 1) + "行中" +'未填写全部的cycle time')
        return mouID[change], result2
    else:
        print('>>>>>8')
        mouID[change] = '故障'
        tempresult = '第' + str(i + 1) + "行中" + "应用报文sheet中的发送周期书写格式有误"
        print('到这4444444444444444444444')
        result2.append(tempresult)

    return mouID[change], result2

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False

def read_excel(path):
    # 打开文件
    alldata = []
    alldiagdata = []
    global result
    global routingmessages
    result=[]
    routingmessages=[]
    df = xlrd.open_workbook(path)
    names = df.sheet_names()
    print(names)
    if 'Summary' in names:
        df_sheet = df.sheet_by_name('Summary')
        # print(1111)
    elif '应用报文' in names:
        df_sheet = df.sheet_by_name('应用报文')
    else:
        allmessage = []
        result.append('无法找到应用报文的sheet，请检查文件内容格式是否按标准填写')
        print(result)
        return result, allmessage
    nRows = df_sheet.nrows
    for i in range(nRows):
        alldata.append(df_sheet.row_values(i))

    if 'Summary_Diag' in names:
        df_sheet2 = df.sheet_by_name('Summary_Diag')
        # print(1111)
    elif '诊断报文' in names:
        df_sheet2 = df.sheet_by_name('诊断报文')
    else:
        allmessage = []
        result.append('无法找到诊断报文的sheet')
        return result, allmessage

    nRows2 = df_sheet2.nrows
    for i in range(nRows2):
        alldiagdata.append(df_sheet2.row_values(i))

    fengeall = []  # 网关所有的网段路由方向
    fengesourdes = []
    allmessage = []  # 所有报文路由的报文
    spaceflag = 0  # 默认不是空行

    diagmessage = []  # 所有诊断CANsheet中的源网段报文ID
    alldiagmessage = []  # 所有诊断CANsheet中报文路由的报文
    # print(len(alldata))
    print("到这1.1")
    for i in range(0, len(alldata)):
        find1 = 'requested signals, source'
        find2 = '源信号'
        print("此时的alldata")
        print(alldata[i])
        print(alldata[i][0])






        if (find1 in alldata[i][0]) or (find2 in alldata[i][0]):  # 寻找是否为新的路由表并获取路由方向
            print('**********1')
            fengesourdes = []
            # sourceflag=1#指示下一行说明了源网段和目的网段
            findflag1 = ':'
            findflag2 = '：'
            # print(3333)
            if findflag1 in alldata[i][0]:
                guodusource = alldata[i][0].split(':')  # 获取源网段

                if guodusource[-1] == '动力' or guodusource[-1] == 'PT':
                    guodusource[-1] = 'PT'
                elif guodusource[-1] == '底盘' or guodusource[-1] == 'Chassis':
                    guodusource[-1] = 'Chassis1'
                elif guodusource[-1] == '底盘2' or guodusource[-1] == 'Chassis2':
                    guodusource[-1] = 'Chassis2'
                elif guodusource[-1] == '舒适' or guodusource[-1] == 'Comf':
                    guodusource[-1] = 'Comf1'
                elif guodusource[-1] == '舒适2':
                    guodusource[-1] = 'Comf2'
                elif guodusource[-1] == '信息' or guodusource[-1] == 'Info':
                    guodusource[-1] = 'Info1'
                elif guodusource[-1] == '驾辅':
                    guodusource[-1] = 'ADas'
                elif guodusource[-1] == '诊断':
                    guodusource[-1] = 'Diag'
                elif guodusource[-1] == 'EV' or guodusource[-1] == '电动':
                    guodusource[-1] = 'EV'
                elif guodusource[-1] == 'Gateway' or guodusource[-1] == '网关' or guodusource[-1] == 'GATEWAY':
                    continue
                else:

                    str1 = "第" + str(i + 1) + "行中" + "应用报文sheet中的源子网1.1" + guodusource[-1] + "书写有误，请检查格式及是否包含空格等"
                    result.append(str1)

                print("到这1.2")

                fengesourdes.append(guodusource[-1])

                guodudes = alldata[i][4].split(':')  # 获取目标网段

                if guodudes[-1] == '动力' or guodudes[-1] == 'PT':
                    print("到这1.3.1")
                    guodudes[-1] = 'PT'
                elif guodudes[-1] == '底盘' or guodudes[-1] == 'Chassis':
                    print("到这1.3.2")
                    guodudes[-1] = 'Chassis1'
                elif guodudes[-1] == '底盘2':
                    print("到这1.3.3")
                    guodudes[-1] = 'Chassis2'
                elif guodudes[-1] == '舒适' or guodudes[-1] == 'Comf':
                    print("到这1.3.4")
                    guodudes[-1] = 'Comf1'
                elif guodudes[-1] == '舒适2':
                    print("到这1.3.5")
                    guodudes[-1] = 'Comf2'
                elif guodudes[-1] == '信息' or guodudes[-1] == 'Info':
                    print("到这1.3.6")
                    guodudes[-1] = 'Info1'
                elif guodudes[-1] == '驾辅':
                    print("到这1.3.7")
                    guodudes[-1] = 'ADas'
                elif guodudes[-1] == '诊断':
                    print("到这1.3.8")
                    guodudes[-1] = 'Diag'
                elif guodudes[-1] == 'EVCAN' or guodudes[-1] == '电动':
                    print("到这1.3.9")
                    guodudes[-1] = 'EV'
                elif guodudes[-1] == 'Gateway' or guodudes[-1] == '网关' or guodudes[-1] == 'GATEWAY':
                    print("到这1.3.10")
                    continue
                else:
                    allmessage = []
                    print("到这1.3")
                    print('以前的result2')
                    print(result)
                    str1 = "第" + str(i + 1) + "行中" + "应用报文sheet中的目标子网1.2" + guodudes[-1] + "书写有误，请检查格式及是否包含空格等"

                    result.append(str1)
                    print("此时的str1：：")
                    print(str1)

                    # return result,allmessage

                fengesourdes.append(guodudes[-1])

                fengeall.append(fengesourdes)

                print("到这1.4")
                continue
            elif findflag2 in alldata[i][0]:
                guodusource = alldata[i][0].split('：')  # 获取源网段
                if guodusource[-1] == '动力' or guodusource[-1] == 'PT':
                    guodusource[-1] = 'PT'
                elif guodusource[-1] == '底盘' or guodusource == 'Chassis':
                    guodusource[-1] = 'Chassis1'
                elif guodusource[-1] == '底盘2':
                    guodusource[-1] = 'Chassis2'
                elif guodusource[-1] == '舒适' or guodusource[-1] == 'Comf':
                    guodusource[-1] = 'Comf1'
                elif guodusource[-1] == '舒适2':
                    guodusource[-1] = 'Comf2'
                elif guodusource[-1] == '信息' or guodusource[-1] == 'Info':
                    guodusource[-1] = 'Info1'
                elif guodusource[-1] == '驾辅':
                    guodusource[-1] = 'ADas'
                elif guodusource[-1] == '诊断':
                    guodusource[-1] = 'Diag'
                elif guodusource[-1] == 'EVCAN' or guodusource[-1] == '电动':
                    guodusource[-1] = 'EV'
                elif guodusource[-1] == 'Gateway' or guodusource[-1] == '网关' or guodusource[-1] == 'GATEWAY':
                    continue
                else:
                    str1 = "第" + str(i + 1) + "行中" + "应用报文sheet中的源子网2.1" + guodusource[-1] + "书写有误，请检查格式及是否包含空格等"
                    result.append(str1)
                fengesourdes.append(guodusource[-1])

                guodudes = alldata[i][4].split('：')  # 获取目标网段
                if guodudes[-1] == '动力':
                    guodudes[-1] = 'PT'
                elif guodudes[-1] == '底盘' or guodudes == 'Chassis':
                    guodudes[-1] = 'Chassis1'
                elif guodudes[-1] == '底盘2':
                    guodudes[-1] = 'Chassis2'
                elif guodudes[-1] == '舒适' or guodudes[-1] == 'Comf':
                    guodudes[-1] = 'Comf1'
                elif guodudes[-1] == '舒适2':
                    guodudes[-1] = 'Comf2'
                elif guodudes[-1] == '信息' or guodudes[-1] == 'Info':
                    guodudes[-1] = 'Info1'
                elif guodudes[-1] == '驾辅':
                    guodudes[-1] = 'ADas'
                elif guodudes[-1] == '诊断':
                    guodudes[-1] = 'Diag'
                elif guodudes[-1] == 'EVCAN' or guodudes[-1] == '电动':
                    guodudes[-1] = 'EV'
                elif guodudes[-1] == 'Gateway' or guodudes[-1] == '网关' or guodudes[-1] == 'GATEWAY':
                    continue
                else:
                    allmessage = []
                    print("是不是又错了")
                    print(guodudes[-1])
                    str1 = "第" + str(i + 1) + "行中" + "应用报文sheet中的目标子网2.2" + guodudes[-1] + "书写有误，请检查格式及是否包含空格等"
                    result.append(str1)
                    return result, allmessage

                fengesourdes.append(guodudes[-1])
                fengeall.append(fengesourdes)

                continue
            else:
                result.append('应用报文sheet中的源子网或目标子网的标点符号书写有误')

                continue
            print("&&&&&&&&&&&&&到这")



        elif alldata[i][1] == '' or alldata[i][-4]=='':  # 检测是否为空行
            print('**********2')
            print(77777)
            spaceflag = 0
            for num in range(0, len(alldata[i])):
                if alldata[i][num] != '':
                    spaceflag = 1  # 为1时表示不是空行

                else:
                    continue
            print("怀疑为空行的alldata【i】")
            print(alldata[i])
            print('是否为空行：：')
            print(spaceflag)
            if spaceflag == 0:  # 是空行
                fengesourdes = []  # 某一具体网段的路由方向设置为空
                print("^^^来此")
                continue

            if alldata[i][1] == '' or alldata[i][-4] == '' and spaceflag == 1:
                print('来此2………………')
                result.append("第" + str(i + 1) + "行中" + '未填写报文名')
                continue



        elif alldata[i][2]=='' or alldata[i][-3]=='':
            print('*************66')
            result.append("第" + str(i + 1) + "行中" + '未填写报文ID')
            continue

        elif (alldata[i][-1] == 'message') or (alldata[i][-1] == '报文'):  # 是报文路由
            print('*************3')

            if alldata[i][2] not in routingmessages:
                print('*************3.1')
                mouID = {"SourceID": alldata[i][2]}
                mouID['SourceChannel'] = fengesourdes[0]

                mouIDchange, result = panduanfasong(alldata[i][3], 'Sourcetime', i, mouID, result, 1, alldata)

                continueflag = 0

                for resultsingle in range(0, len(result)):
                    if result[resultsingle] == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                        print(result[resultsingle])
                        print('*************3.2')
                        allmessage = []
                        continue
                        #return result, allmessage
                    if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                        print('*************3.3')
                        allmessage = []
                        continueflag = 1
                        continue
                if continueflag == 1:
                    print('*************3.4')
                    continue

                mouID['Sourcetime'] = mouIDchange

                mouID['PT'] = 0
                mouID['PTid'] = 0
                mouID['PTtime'] = 0
                #
                mouID['Chassis1'] = 0
                mouID['Chassis1id'] = 0
                mouID['Chassis1time'] = 0
                #
                mouID['Chassis2'] = 0
                mouID['Chassis2id'] = 0
                mouID['Chassis2time'] = 0
                #
                mouID['Comf1'] = 0
                mouID['Comf1id'] = 0
                mouID['Comf1time'] = 0
                #
                mouID['Comf2'] = 0
                mouID['Comf2id'] = 0
                mouID['Comf2time'] = 0
                #
                mouID['Info1'] = 0
                mouID['Info1id'] = 0
                mouID['Info1time'] = 0
                #
                mouID['ADas'] = 0
                mouID['ADasid'] = 0
                mouID['ADastime'] = 0
                #
                mouID['EV'] = 0
                mouID['EVid'] = 0
                mouID['EVtime'] = 0
                #
                mouID['Diag'] = 0
                mouID['Diagid'] = 0
                mouID['Diagtime'] = 0
                routingmessages.append(alldata[i][2])

                if (str(fengesourdes[1]) == 'PT') or (str(fengesourdes[1]) == '动力'):
                    mouID['PT'] = 1
                    mouID['PTid'] = alldata[i][-3]

                    mouIDchange, result = panduanfasong(alldata[i][-2], 'PTtime', i, mouID, result, 0, alldata)

                    continueflag = 0
                    for resultsingle in range(0, len(result)):
                        if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue

                        # return result, allmessage

                    mouID['PTtime'] = mouIDchange

                elif (str(fengesourdes[1]) == 'Chassis') or (str(fengesourdes[1]) == 'Chassis1') or (
                        str(fengesourdes[1]) == '底盘'):
                    mouID['Chassis1'] = 1

                    mouID['Chassis1id'] = alldata[i][-3]
                    mouIDchange, result = panduanfasong(alldata[i][-2], 'Chassis1time', i, mouID, result, 0, alldata)

                    continueflag = 0  # 为了使多个result叠加
                    for resultsingle in range(0, len(result)):
                        if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue

                    mouID['Chassis1time'] = mouIDchange


                elif str(fengesourdes[1]) == 'Chassis2' or (str(fengesourdes[1]) == '底盘2'):
                    mouID['Chassis2'] = 1
                    mouID['Chassis2id'] = alldata[i][-3]
                    mouIDchange, result = panduanfasong(alldata[i][-2], 'Chassis2time', i, mouID, result, 0, alldata)
                    continueflag = 0
                    for resultsingle in range(0, len(result)):
                        if resultsingle =="第" + str(i + 1) + "行中" + '未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue
                    mouID['Chassis2time'] = mouIDchange

                elif (str(fengesourdes[1]) == 'Comf1') or (str(fengesourdes[1]) == 'Comf') or (
                        str(fengesourdes[1]) == '舒适'):
                    mouID['Comf1'] = 1
                    mouID['Comf1id'] = alldata[i][-3]

                    mouIDchange, result = panduanfasong(alldata[i][-2], 'Comf1time', i, mouID, result, 0, alldata)
                    continueflag = 0
                    for resultsingle in range(0, len(result)):
                        if resultsingle =="第" + str(i + 1) + "行中" + '未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue

                    mouID['Comf1time'] = mouIDchange

                elif str(fengesourdes[1]) == 'Comf2' or (str(fengesourdes[1]) == '舒适2'):
                    mouID['Comf2'] = 1
                    mouID['Comf2id'] = alldata[i][-3]
                    mouIDchange, result = panduanfasong(alldata[i][-2], 'Comf2time', i, mouID, result, 0, alldata)
                    continueflag = 0
                    for resultsingle in range(0, len(result)):
                        if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue
                    mouID['Comf2time'] = mouIDchange

                elif str(fengesourdes[1]) == 'EV' or str(fengesourdes[1]) == '电动':
                    mouID['EV'] = 1
                    mouID['EVid'] = alldata[i][-3]
                    mouIDchange, result = panduanfasong(alldata[i][-2], 'EVtime', i, mouID, result, 0, alldata)
                    continueflag = 0
                    for resultsingle in range(0, len(result)):
                        if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue
                    mouID['EVtime'] = mouIDchange

                elif str(fengesourdes[1]) == 'Diag' or str(fengesourdes[1]) == '诊断':
                    mouID['Diag'] = 1
                    mouID['Diagid'] = alldata[i][-3]
                    mouIDchange, result = panduanfasong(alldata[i][-2], 'Diagtime', i, mouID, result, 0, alldata)
                    continueflag = 0
                    for resultsingle in range(0, len(result)):
                        if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue
                    mouID['Diagtime'] = mouIDchange

                elif str(fengesourdes[1]) == 'Info' or (str(fengesourdes[1]) == 'Info1') or str(
                        fengesourdes[1]) == '信息':
                    mouID['Info1'] = 1
                    mouID['Info1id'] = alldata[i][-3]
                    mouIDchange, result = panduanfasong(alldata[i][-2], 'Info1time', i, mouID, result, 0, alldata)
                    continueflag = 0
                    for resultsingle in range(0, len(result)):
                        if resultsingle =="第" + str(i + 1) + "行中" + '未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue

                    mouID['Info1time'] = mouIDchange

                elif str(fengesourdes[1]) == 'ADas' or str(fengesourdes[1]) == '驾辅':
                    mouID['ADas'] = 1
                    mouID['ADasid'] = alldata[i][-3]
                    mouIDchange, result = panduanfasong(alldata[i][-2], 'ADastime', i, mouID, result, 0, alldata)

                    continueflag = 0
                    for resultsingle in range(0, len(result)):
                        if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                            allmessage = []
                            return result, allmessage
                        if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                            allmessage = []
                            continueflag = 1
                            continue
                    if continueflag == 1:
                        continue
                    mouID['ADastime'] = mouIDchange

                allmessage.append(mouID)
            else:  # 该报文ID已经存在于建好的报文路由表中
                for mouID in allmessage:
                    if mouID['SourceID'] == alldata[i][2]:

                        if (str(fengesourdes[1]) == 'PT') or (str(fengesourdes[1]) == '动力'):
                            mouID['PT'] = 1
                            mouID['PTid'] = alldata[i][-3]

                            mouIDchange, result = panduanfasong(alldata[i][-2], 'PTtime', i, mouID, result, 0, alldata)
                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage
                            mouID['PTtime'] = mouIDchange
                            alreadyexist = 1
                        elif (str(fengesourdes[1]) == 'Chassis') or (str(fengesourdes[1]) == 'Chassis1') or (
                                str(fengesourdes[1]) == '底盘'):
                            mouID['Chassis1'] = 1
                            mouID['Chassis1id'] = alldata[i][-3]
                            mouIDchange, result = panduanfasong(alldata[i][-2], 'Chassis1time', i, mouID, result, 0,
                                                                alldata)

                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage

                            mouID['Chassis1time'] = mouIDchange
                            alreadyexist = 1
                        elif str(fengesourdes[1]) == 'Chassis2' or (str(fengesourdes[1]) == '底盘2'):
                            mouID['Chassis2'] = 1
                            mouID['Chassis2id'] = alldata[i][-3]
                            mouIDchange, result = panduanfasong(alldata[i][-2], 'Chassis2time', i, mouID, result, 0,
                                                                alldata)

                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage

                            mouID['Chassis2time'] = mouIDchange
                            alreadyexist = 1
                        elif (str(fengesourdes[1]) == 'Comf1') or (str(fengesourdes[1]) == 'Comf') or (
                                str(fengesourdes[1]) == '舒适'):
                            mouID['Comf1'] = 1
                            mouID['Comf1id'] = alldata[i][-3]
                            mouIDchange, result = panduanfasong(alldata[i][-2], 'Comf1time', i, mouID, result, 0,
                                                                alldata)
                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage
                            mouID['Comf1time'] = mouIDchange
                            alreadyexist = 1
                        elif str(fengesourdes[1]) == 'Comf2' or (str(fengesourdes[1]) == '舒适2'):
                            mouID['Comf2'] = 1
                            mouID['Comf2id'] = alldata[i][-3]
                            mouIDchange, result = panduanfasong(alldata[i][-2], 'Comf2time', i, mouID, result, 0,
                                                                alldata)

                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage
                            mouID['Comf2time'] = mouIDchange
                            alreadyexist = 1
                        elif str(fengesourdes[1]) == 'EV' or str(fengesourdes[1]) == '电动':
                            mouID['EV'] = 1
                            mouID['EVid'] = alldata[i][-3]
                            mouIDchange, result = panduanfasong(alldata[i][-2], 'EVtime', i, mouID, result, 0, alldata)
                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage
                            mouID['EVtime'] = mouIDchange
                            alreadyexist = 1
                        elif str(fengesourdes[1]) == 'Diag' or str(fengesourdes[1]) == '诊断':
                            mouID['Diag'] = 1
                            mouID['Diagid'] = alldata[i][-3]
                            mouIDchange, result = panduanfasong(alldata[i][-2], 'Diagtime', i, mouID, result, 0,
                                                                alldata)
                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage
                            mouID['Diagtime'] = mouIDchange
                            alreadyexist = 1
                        elif (str(fengesourdes[1]) == 'Info') or (str(fengesourdes[1]) == 'Info1') or str(
                                fengesourdes[1]) == '信息':
                            mouID['Info1'] = 1
                            mouID['Info1id'] = alldata[i][-3]
                            mouIDchange, result = panduanfasong(alldata[i][-2], 'Info1time', i, mouID, result, 0,
                                                                alldata)
                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage
                            mouID['Info1time'] = mouIDchange
                            alreadyexist = 1
                        elif str(fengesourdes[1]) == 'ADas' or str(fengesourdes[1]) == '驾辅':
                            mouID['ADas'] = 1
                            mouID['ADasid'] = alldata[i][-3]
                            mouIDchange, result = panduanfasong(alldata[i][-2], 'ADastime', i, mouID, result, 0,
                                                                alldata)
                            for resultsingle in range(0, len(result)):
                                if resultsingle == "第" + str(i + 1) + "行中" +'未填写全部的cycle time':
                                    allmessage = []
                                    return result, allmessage
                                if "第" + str(i + 1) + "行中" + '应用报文sheet中的发送周期书写格式有误' in result[resultsingle]:
                                    allmessage = []
                                    return result, allmessage

                            mouID['ADastime'] = mouIDchange
                            alreadyexist = 1
                        else:
                            allmessage = []
                            str1 = "第" + str(i + 1) + "行中" + "应用报文sheet中的目标子网" + fengesourdes[1] + "书写有误，请检查格式及是否包含空格等"
                            result.append(str1)
                            return result, allmessage
                    else:
                        continue
        else:
            print("pass")

    for i in range(1, len(alldiagdata)):
        if alldiagdata[i][0] not in diagmessage:
            mouID2 = {"SourceID": alldiagdata[i][0]}
            diagmessage.append(alldiagdata[i][0])

            if alldiagdata[i][1] == '诊断' or alldiagdata[i][1] == 'Diag' or alldiagdata[i][1] == 'diag':
                mouID2['SourceChannel'] = 'Diag'
                mouID2['Sourcetime'] = int(10)

            elif alldiagdata[i][1] == '舒适' or alldiagdata[i][1] == 'Comf1' or alldiagdata[i][1] == 'Comf' or \
                    alldiagdata[i][1] == 'comf' or alldiagdata[i][1] == '舒适CAN':
                mouID2['SourceChannel'] = 'Comf1'
                mouID2['Sourcetime'] = int(10)

            elif alldiagdata[i][1] == '舒适2' or alldiagdata[i][1] == 'Comf2' or alldiagdata[i][1] == 'Comf2' or \
                    alldiagdata[i][1] == 'comf2':
                mouID2['SourceChannel'] = 'Comf2'
                mouID2['Sourcetime'] = int(10)

            elif alldiagdata[i][1] == '动力' or alldiagdata[i][1] == 'PT':
                mouID2['SourceChannel'] = 'PT'
                mouID2['Sourcetime'] = int(10)

            elif alldiagdata[i][1] == '信息' or alldiagdata[i][1] == 'Info' or alldiagdata[i][1] == 'Info1':
                mouID2['SourceChannel'] = 'Info1'
                mouID2['Sourcetime'] = int(10)

            elif alldiagdata[i][1] == '驾辅' or alldiagdata[i][1] == 'ADas':
                mouID2['SourceChannel'] = 'ADas'
                mouID2['Sourcetime'] = int(10)

            elif alldiagdata[i][1] == 'EV' or str(alldiagdata[1]) == '电动':
                mouID2['SourceChannel'] = 'EV'
                mouID2['Sourcetime'] = int(10)

            elif alldiagdata[i][1] == '底盘' or alldiagdata[i][1] == 'Chassis' or alldiagdata[i][1] == 'Chassis1':
                mouID2['SourceChannel'] = 'Chassis1'
                mouID2['Sourcetime'] = int(10)

            elif alldiagdata[i][1] == '底盘2' or alldiagdata[i][1] == 'Chassis2':
                mouID2['SourceChannel'] = 'Chassis2'
                mouID2['Sourcetime'] = int(10)
            else:
                str1 = "第" + str(i + 1) + "行中" + "诊断报文sheet中的源子网" + alldiagdata[i][1] + "书写有误，请检查格式及是否包含空格等"
                result.append(str1)
            #
            mouID2['PT'] = 0
            mouID2['PTid'] = 0
            mouID2['PTtime'] = 0
            #
            mouID2['Chassis1'] = 0
            mouID2['Chassis1id'] = 0
            mouID2['Chassis1time'] = 0
            #
            mouID2['Chassis2'] = 0
            mouID2['Chassis2id'] = 0
            mouID2['Chassis2time'] = 0
            #
            mouID2['Comf1'] = 0
            mouID2['Comf1id'] = 0
            mouID2['Comf1time'] = 0
            #
            mouID2['Comf2'] = 0
            mouID2['Comf2id'] = 0
            mouID2['Comf2time'] = 0
            #
            mouID2['Info1'] = 0
            mouID2['Info1id'] = 0
            mouID2['Info1time'] = 0
            #
            mouID2['ADas'] = 0
            mouID2['ADasid'] = 0
            mouID2['ADastime'] = 0
            #
            mouID2['EV'] = 0
            mouID2['EVid'] = 0
            mouID2['EVtime'] = 0
            #
            mouID2['Diag'] = 0
            mouID2['Diagid'] = 0
            mouID2['Diagtime'] = 0

            if alldiagdata[i][-1] == '动力' or alldiagdata[i][-1] == 'PT':
                mouID2['PT'] = 1
                mouID2['PTid'] = mouID2['SourceID']
                mouID2['PTtime'] = mouID2['Sourcetime']
            elif alldiagdata[i][-1] == '诊断' or alldiagdata[i][-1] == 'Diag' or alldiagdata[i][-1] == 'diag':
                mouID2['Diag'] = 1
                mouID2['Diagid'] = mouID2['SourceID']
                mouID2['Diagtime'] = mouID2['Sourcetime']

            elif alldiagdata[i][-1] == '舒适' or alldiagdata[i][-1] == 'Comf1' or alldiagdata[i][-1] == 'Comf' or \
                    alldiagdata[i][-1] == 'comf' or alldiagdata[i][-1] == '舒适CAN':
                mouID2['Comf1'] = 1
                mouID2['Comf1id'] = mouID2['SourceID']
                mouID2['Comf1time'] = mouID2['Sourcetime']

            elif alldiagdata[i][-1] == '舒适2' or alldiagdata[i][-1] == 'Comf2' or alldiagdata[i][-1] == 'Comf2' or \
                    alldiagdata[i][-1] == 'comf2':
                mouID2['Comf2'] = 1
                mouID2['Comf2id'] = mouID2['SourceID']
                mouID2['Comf2time'] = mouID2['Sourcetime']

            elif alldiagdata[i][-1] == '信息' or alldiagdata[i][-1] == 'Info' or alldiagdata[i][-1] == 'Info1':
                mouID2['Info1'] = 1
                mouID2['Info1id'] = mouID2['SourceID']
                mouID2['Info1time'] = mouID2['Sourcetime']

            elif alldiagdata[i][-1] == '驾辅' or alldiagdata[i][-1] == 'ADas':
                mouID2['ADas'] = 1
                mouID2['ADasid'] = mouID2['SourceID']
                mouID2['ADastime'] = mouID2['Sourcetime']

            elif alldiagdata[i][-1] == 'EV' or str(alldiagdata[i][-1]) == '电动':
                mouID2['EV'] = 1
                mouID2['EVid'] = mouID2['SourceID']
                mouID2['EVtime'] = mouID2['Sourcetime']

            elif alldiagdata[i][-1] == '底盘' or alldiagdata[i][-1] == 'Chassis' or alldiagdata[i][-1] == 'Chassis1':
                mouID2['Chassis1'] = 1
                mouID2['Chassis1id'] = mouID2['SourceID']
                mouID2['Chassis1time'] = mouID2['Sourcetime']

            elif alldiagdata[i][-1] == '底盘2' or alldiagdata[i][-1] == 'Chassis2':
                mouID2['Chassis2'] = 1
                mouID2['Chassis2id'] = mouID2['SourceID']
                mouID2['Chassis2time'] = mouID2['Sourcetime']
            else:
                str2 = "第" + str(i + 1) + "行中" + "诊断报文sheet中的目标子网" + alldiagdata[i][-1] + "书写有误，请检查格式及是否包含空格等"
                result.append(str2)

            alldiagmessage.append(mouID2)
        else:  # 该报文已经存在于建好的诊断路由表中
            for mouID2 in alldiagmessage:
                if mouID2['SourceID'] == alldiagdata[i][0]:
                    if alldiagdata[i][-1] == '动力' or alldiagdata[i][-1] == 'PT':
                        mouID2['PT'] = 1
                        mouID2['PTid'] = mouID2['SourceID']
                        mouID2['PTtime'] = mouID2['Sourcetime']
                    elif alldiagdata[i][-1] == '诊断' or alldiagdata[i][-1] == 'Diag' or alldiagdata[i][-1] == 'diag':
                        mouID2['Diag'] = 1
                        mouID2['Diagid'] = mouID2['SourceID']
                        mouID2['Diagtime'] = mouID2['Sourcetime']

                    elif alldiagdata[i][-1] == '舒适' or alldiagdata[i][-1] == 'Comf1' or alldiagdata[i][-1] == 'Comf' or \
                            alldiagdata[i][-1] == 'comf' or alldiagdata[i][-1] == '舒适CAN':
                        mouID2['Comf1'] = 1
                        mouID2['Comf1id'] = mouID2['SourceID']
                        mouID2['Comf1time'] = mouID2['Sourcetime']

                    elif alldiagdata[i][-1] == '舒适2' or alldiagdata[i][-1] == 'Comf2' or alldiagdata[i][
                        -1] == 'Comf2' or alldiagdata[i][-1] == 'comf2':
                        mouID2['Comf2'] = 1
                        mouID2['Comf2id'] = mouID2['SourceID']
                        mouID2['Comf2time'] = mouID2['Sourcetime']

                    elif alldiagdata[i][-1] == '信息' or alldiagdata[i][-1] == 'Info' or alldiagdata[i][-1] == 'Info1':
                        mouID2['Info1'] = 1
                        mouID2['Info1id'] = mouID2['SourceID']
                        mouID2['Info1time'] = mouID2['Sourcetime']

                    elif alldiagdata[i][-1] == '驾辅' or alldiagdata[i][-1] == 'ADas':
                        mouID2['ADas'] = 1
                        mouID2['ADasid'] = mouID2['SourceID']
                        mouID2['ADastime'] = mouID2['Sourcetime']

                    elif alldiagdata[i][-1] == 'EV' or str(alldiagdata[-1]) == '电动':
                        mouID2['EV'] = 1
                        mouID2['EVid'] = mouID2['SourceID']
                        mouID2['EVtime'] = mouID2['Sourcetime']

                    elif alldiagdata[i][-1] == '底盘' or alldiagdata[i][-1] == 'Chassis' or alldiagdata[i][
                        -1] == 'Chassis1':
                        mouID2['Chassis1'] = 1
                        mouID2['Chassis1id'] = mouID2['SourceID']
                        mouID2['Chassis1time'] = mouID2['Sourcetime']

                    elif alldiagdata[i][-1] == '底盘2' or alldiagdata[i][-1] == 'Chassis2':
                        mouID2['Chassis2'] = 1
                        mouID2['Chassis2id'] = mouID2['SourceID']
                        mouID2['Chassis2time'] = mouID2['Sourcetime']
                    else:
                        str2 = "第" + str(i + 1) + "行中" + '诊断报文sheet中的目标子网' + alldiagdata[i][-1] + "书写有误，请检查格式及是否包含空格等"
                        result.append(str2)

                else:
                    continue
    # print('让我看看alldiagmessage：：：：：：：：：：----------------------')
    # print(alldiagmessage)
    for i in alldiagmessage:
        allmessage.append(i)
    # print("把两个sheet的值合并，此时还没把源网段sourcechannel映射成数")
    # print(allmessage)#此时还没把源网段sourcechannel映射成数

    for single in allmessage:

        if single['SourceChannel'] == 'PT':
            single['SourceChannel'] = 1
        elif single['SourceChannel'] == 'Chassis1':
            single['SourceChannel'] = 2
        elif single['SourceChannel'] == 'EV':
            single['SourceChannel'] = 3
        elif single['SourceChannel'] == 'Comf1':
            single['SourceChannel'] = 4
        elif single['SourceChannel'] == 'Info1':
            single['SourceChannel'] = 5
        elif single['SourceChannel'] == 'Diag':
            single['SourceChannel'] = 6
        elif single['SourceChannel'] == 'ADas':
            single['SourceChannel'] = 7
        elif single['SourceChannel'] == 'Chassis2':
            single['SourceChannel'] = 8
        elif single['SourceChannel'] == 'Comf2':
            single['SourceChannel'] = 9
    # print("把名字映射成数字")
    # print(allmessage)
    return result, allmessage


def zhuhanshu(path2,path3):
    path='D:/OJD/'+path2
    print("文件路径：")
    print(path)
    path_cin='D:/OJD/'+path3
    print("path_Cin:")
    print(path_cin)

    global result
    result = []
    errorresult, allmessage = read_excel(path)
    print("结果的errorresult值：")
    print(errorresult)

    if (len(errorresult) != 0):
        return errorresult
    else:
        print('执行后续对allmessage的处理工作')

        print(allmessage)

        print("无错误，现在要返回错误值")
        print(errorresult)
        dtc.dictobin(allmessage,path_cin)
        #
        return '已成功执行完毕'
