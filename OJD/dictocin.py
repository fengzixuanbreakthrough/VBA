import zipfile
import shutil
import os
def dictobin(lastroutingmessage,path):
    fllename_number = path+'number.cin'
    fllename_messge = path+'message.cin'
    print("cin文件路径：")
    print(fllename_messge)
    maxmessge = len(lastroutingmessage)
    with open(fllename_number, 'w') as file_object:
        file_object.write("/*@!Encoding:936*/\n")
        file_object.write("variables\n")
        file_object.write("{\n")
        print(len(lastroutingmessage) - 2)
        file_object.write("  int number = {}; \n".format(maxmessge))
        file_object.write("}\n")

    with open(fllename_messge, 'w') as file_object1:
        file_object1.write("/*@!Encoding:936*/\n")
        file_object1.write("variables\n")
        file_object1.write("{\n")
        file_object1.write("   struct MsgRout\n")
        file_object1.write("{\n")
        # 1
        file_object1.write("    long ID;\n")
        # 2
        file_object1.write("    int cycle;\n")
        # 3
        file_object1.write("    int Source_Chn;\n")

        # 4
        file_object1.write("    int Support_PT;\n")
        # 5
        file_object1.write("    int cycle_PT;\n")
        # 6
        file_object1.write("    long cycle_PTID;\n")

        # 7
        file_object1.write("    int Support_CH;\n")
        # 8
        file_object1.write("    int cycle_CH;\n")
        # 9
        file_object1.write("    long cycle_CHID;\n")

        # 10
        file_object1.write("    int Support_EV;\n")
        # 11
        file_object1.write("    int cycle_EV;\n")
        # 12
        file_object1.write("    long cycle_EVID;\n")

        # 13
        file_object1.write("    int Support_Com;\n")
        # 14
        file_object1.write("    int cycle_Com;\n")
        # 15
        file_object1.write("    long cycle_ComID;\n")

        # 16
        file_object1.write("    int Support_Info;\n")
        # 17
        file_object1.write("    int cycle_Info;\n")
        # 18
        file_object1.write("    long cycle_InfoID;\n")

        # 19
        file_object1.write("    int Support_Diag;\n")
        # 20
        file_object1.write("    int cycle_Diag;\n")
        # 21
        file_object1.write("    long cycle_DiagID;\n")

        # 22
        file_object1.write("    int Support_ADAS;\n")
        # 23
        file_object1.write("    int cycle_ADAS;\n")
        # 24
        file_object1.write("    long cycle_ADASID;\n")

        # 25
        file_object1.write("    int Support_CH2;\n")
        # 26
        file_object1.write("    int cycle_CH2;\n")
        # 27
        file_object1.write("    long cycle_CH2ID;\n")

        # 28
        file_object1.write("    int Support_COM2;\n")
        # 29
        file_object1.write("    int cycle_COM2;\n")
        # 30
        file_object1.write("    long cycle_COM2ID;\n")

        file_object1.write("  }msg[1000]={\n")
        print("你好！")
        for i in range(0, maxmessge - 1):
            file_object1.write("                        {")
            file_object1.write(
                "0x{0}x,{1},{2},{3},{4},0x{5}x,{6},{7},0x{8}x,{9},{10},0x{11}x,{12},{13},0x{14}x,{15},{16},0x{17}x,{18},{19},0x{20}x,{21},{22},0x{23}x,{24},{25},0x{26}x,{27},{28},0x{29}x".format(
                    lastroutingmessage[i]['SourceID'], lastroutingmessage[i]['Sourcetime'],
                    lastroutingmessage[i]['SourceChannel']
                    , lastroutingmessage[i]['PT'], lastroutingmessage[i]['PTtime'], lastroutingmessage[i]['PTid']
                    , lastroutingmessage[i]['Chassis1'], lastroutingmessage[i]['Chassis1time'],
                    lastroutingmessage[i]['Chassis1id']
                    , lastroutingmessage[i]['EV'], lastroutingmessage[i]['EVtime'], lastroutingmessage[i]['EVid']
                    , lastroutingmessage[i]['Comf1'], lastroutingmessage[i]['Comf1time'],
                    lastroutingmessage[i]['Comf1id']
                    , lastroutingmessage[i]['Info1'], lastroutingmessage[i]['Info1time'],
                    lastroutingmessage[i]['Info1id']
                    , lastroutingmessage[i]['Diag'], lastroutingmessage[i]['Diagtime'], lastroutingmessage[i]['Diagid']
                    , lastroutingmessage[i]['ADas'], lastroutingmessage[i]['ADastime'], lastroutingmessage[i]['ADasid']
                    , lastroutingmessage[i]['Chassis2'], lastroutingmessage[i]['Chassis2time'],
                    lastroutingmessage[i]['Chassis2id']
                    , lastroutingmessage[i]['Comf2'], lastroutingmessage[i]['Comf2time'],
                    lastroutingmessage[i]['Comf2id']))
            file_object1.write("}")
            file_object1.write(",")
            file_object1.write("\n")
        file_object1.write("                        {")

        file_object1.write(
            "0x{0}x,{1},{2},{3},{4},0x{5}x,{6},{7},0x{8}x,{9},{10},0x{11}x,{12},{13},0x{14}x,{15},{16},0x{17}x,{18},{19},0x{20}x,{21},{22},0x{23}x,{24},{25},0x{26}x,{27},{28},0x{29}x".format(
                lastroutingmessage[maxmessge - 1]['SourceID'], lastroutingmessage[maxmessge - 1]['Sourcetime'],
                lastroutingmessage[maxmessge - 1]['SourceChannel']
                , lastroutingmessage[maxmessge - 1]['PT'], lastroutingmessage[maxmessge - 1]['PTtime'],
                lastroutingmessage[maxmessge - 1]['PTid']
                , lastroutingmessage[maxmessge - 1]['Chassis1'], lastroutingmessage[maxmessge - 1]['Chassis1time'],
                lastroutingmessage[maxmessge - 1]['Chassis1id']
                , lastroutingmessage[maxmessge - 1]['EV'], lastroutingmessage[maxmessge - 1]['EVtime'],
                lastroutingmessage[maxmessge - 1]['EVid']
                , lastroutingmessage[maxmessge - 1]['Comf1'], lastroutingmessage[maxmessge - 1]['Comf1time'],
                lastroutingmessage[maxmessge - 1]['Comf1id']
                , lastroutingmessage[maxmessge - 1]['Info1'], lastroutingmessage[maxmessge - 1]['Info1time'],
                lastroutingmessage[maxmessge - 1]['Info1id']
                , lastroutingmessage[maxmessge - 1]['Diag'], lastroutingmessage[maxmessge - 1]['Diagtime'],
                lastroutingmessage[maxmessge - 1]['Diagid']
                , lastroutingmessage[maxmessge - 1]['ADas'], lastroutingmessage[maxmessge - 1]['ADastime'],
                lastroutingmessage[maxmessge - 1]['ADasid']
                , lastroutingmessage[maxmessge - 1]['Chassis2'], lastroutingmessage[maxmessge - 1]['Chassis2time'],
                lastroutingmessage[maxmessge - 1]['Chassis2id']
                , lastroutingmessage[maxmessge - 1]['Comf2'], lastroutingmessage[maxmessge - 1]['Comf2time'],
                lastroutingmessage[maxmessge - 1]['Comf2id']))
        file_object1.write("}")
        file_object1.write("\n")
        file_object1.write("                        };")
        file_object1.write("\n")
        file_object1.write("}")
    shutil.copyfile(fllename_messge, r'D:\CANoeConfig\TestModules\message.cin')