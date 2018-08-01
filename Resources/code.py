<<<<<<< HEAD
﻿# -*- coding: utf-8 -*-
"""
Created on Thu Nov 23 00:40:42 2017
@author: Raymond
"""   
def find_file():
    ''' 从文件夹中查找Excel文件 '''
    print("\n    Finding Excel file...")
    import glob
    try:
        filePath = []
        for fileType in ['xlsx','xls']:
            path = glob.glob(r'../*.{}'.format(fileType))
            if path!=[]: filePath.extend(path)
    except Exception as ex:
        print(ex)
        input('Error：调用glob寻找Excel文件失败，请重试！\n按【回车键】退出...')
        exit(-1)
    files = [path.split('\\')[-1] for path in filePath]
    return files

def find_in_fileName(fileName, itemList):
    for i in itemList:
        if fileName.find(i) != -1:
            item = i
            break
    return item


def load_file():
    ''' 第一步：加载Excel文件，从文件名中提取 建筑编号、年份和月份 '''
    files = find_file()
    fileNum = len(files)
    if fileNum < 1:
        print('\nError: 当前文件夹下未检测到Excel文件，请复制要处理的Excel文件到该文件夹下！')
        input('\n按【回车键】退出...')   
        exit(-1)
    else:
        if fileNum == 1:
            fileName = files[0]
        elif fileNum > 1:
            print('    当前文件夹下检测到以下Excel文件:')
            for i in range(fileNum):
                    print('        {}  {}'.format(i+1,files[i]))
            while True:
                try:
                    select = input('    请选择要处理的文件[1-{}]: '.format(fileNum))
                    fileName = files[int(select)-1]
                    break
                except Exception:
                    print('    Error：文件选择错误!\n')  
    
    file = xlrd.open_workbook('../'+fileName)
    print('    已获取Excel文件： {}'.format(fileName))

    buildingList = ['S1','S2','S3','S4']
    monthList = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    # 从文件名中提取建筑编号、年份和月份
    building = find_in_fileName(fileName, buildingList)
    month = find_in_fileName(fileName, monthList)
    year = fileName.split('.')[0].strip()[-4:]
    print('    Building: {}   Month: {} {}'.format(building, month, year))

    month = str(monthList.index(month)+1)
    if len(month)==1:
        month = '0'+ month
    year_month = year+'-'+month
    return file,fileName,building, year_month

def load_sheets(file):
    ''' 加载Excel各表 '''
    for sheetName in file.sheet_names():
        if sheetName.lower() in ['raw']:
            Raw = file.sheet_by_name(sheetName)
        elif sheetName.lower() in ['meter record','meter']:
            Meter = file.sheet_by_name(sheetName)
        elif sheetName.lower() in ['data']:
            Data = file.sheet_by_name(sheetName)
    return Raw, Data, Meter

def quota_out(date):
    if date > 0 and date <= 7:
        quota = 17.85
    elif date > 7 and date <= 15:
        quota = 35.71
    elif date > 15 and date <= 23:
        quota = 53.57
    else:
        quota = 71.42
    return quota

def quota_in(date):
    if date > 0 and date <= 7:
        quota = 53.57
    elif date > 7 and date <= 15:
        quota = 35.71
    elif date > 15 and date <= 23:
        quota = 17.85
    else:
        quota = 0
    return quota

def change_is_in_this_month(change,bed,ym):
    ''' 判断床位变动是否发生在本月 '''
    string = change.strip().split(' ') # 去除字符串前后空格
    monthList = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    month_check = False
    for i in monthList:
        if i.lower() in change.lower():
            month_check = True
            month = str(monthList.index(i)+1)
            if len(month)==1:
                month = '0'+ month
            break
    if month_check == True: 
        if month == ym[-2:]:
            return True
        else:
            return False
    else:
        print('\n\nError： {} 床位的 \"轉房或退宿日期\" 栏内容为： {}'.format(bed,change))
        print('        该拼写可能有误，程序被迫中止，请更正后重新运行该程序！\n')
        input('按【回车键】退出...')
        exit(-1)

def change_list_processor(change,bed):
    ''' 对该床位变动信息进行分类 '''
    string = change.strip().split(' ') # 去除行前空格
    try:
        if string[0] == 'check-in':
            change_type = 'check in'
            day = float(string[2])
            date = string[3]
            quota = 71.42
        elif string[0] == 'stay-in':   # stay-in only on 3 Oct 2017 
            change_type = 'move in'
            day = float(string[3])
            date = string[3:6]
            quota = quota_in(day)
        elif string[0] == 'move':      # move to S308037A on 4 Sep 2017 (charge)
            change_type = 'move out'
            day = float(string[4])
            date = string[4:7]
            quota = quota_out(day)
        elif string[0] == 'check-out': # check-out on 14 Dec 2017 (charge)
            change_type = 'check out'
            day = float(string[2])
            date = string[2:5]
            quota = 71.42
        return change_type, day, date, quota
    except Exception:
        print('\n\nError： {} 床位的 \"轉房或退宿日期\" 栏内容为： {}'.format(bed,change))
        print('        该格式可能有误，程序被迫中止，请更正后重新运行该程序！\n')
        input('按【回车键】退出...')
        exit(-1)

def Index(List,item):
    try:
        return List.index(item)
    except:
        return None

def student_infor_extractor(row,List):
    room = row[Index(List,'room')]
    bed = row[Index(List,'bed')]
    try:
        room_type = row[Index(List,'room_type')]
    except:
        room_type = None
    ID = row[Index(List,'studentID_1')]+row[Index(List,'studentID_2')]
    if ID in [42,84,'',' ']: 
        ID = None
    name_eng = row[Index(List,'name_eng')]
    try:
        name_cn = row[Index(List,'name_cn')]
    except:
        name_cn = None
    try:
        checkInDate = xlrd.xldate.xldate_as_datetime(float(row[Index(List,'checkInDate')]), 0)
        checkInDate = str(checkInDate)[:10]
    except Exception:
        checkInDate = None
    change = row[Index(List,'change')]
    
    return room, bed, room_type,ID,name_eng, name_cn,checkInDate,change

def creat_list():
    change_list = pd.DataFrame(columns=['Room', 'Bed','ID','name_eng',
                                        'name_cn','Check In Date','Change','Type'])
    empty_bed_list = pd.DataFrame(columns=['Room', 'Bed', 'Room Type'])
    Occupied_list = pd.DataFrame(columns=['Room', 'Bed','Room Type',
                                          'ID','Name(eng)',
                                          'Name(cn)','Check-in Date','Change',
                                          'Quota', '个人起始读数','个人结束读数'])
    return [change_list,empty_bed_list,Occupied_list]

def is_check_in(checkInDate,ym):
    return checkInDate[:7] == ym

def locate_item(sheet,strings):
    for line in range(8):
        row_number = line
        row = sheet.row_values(line)
        for i in strings:
            col_number = Index(row,i)
            if col_number != None:
                return [row_number,col_number]

def Meter_read_extractor(Meter):
    print('    加载Meter表的数据')
    row_number, col_number=locate_item(Meter,['房號'])
    Meter_ROW_Begining = row_number+1
    meter = {}
    for line in range(Meter_ROW_Begining, Meter.nrows):
        row = Meter.row_values(line)
        room = row[col_number]
        read = row[col_number+2:col_number+4]
        meter[room] = read
    return meter

def select_itemList(building):
    ''' 因为S1-S4各表的项目栏顺序不一样，故需分别确定，以免对错 '''

    index_S2 = ['bed','room','sex','room_type','studentID_1','studentID_2',
                 'name_eng','checkInDate','change','read_s','read_e']
    index_S3_S4 = ['room','bed','room_type','studentID_1','studentID_2',
                 'name_eng','name_cn','checkInDate','change','read_s','read_e']
    index_S1 = ['room','bed','studentID_1','studentID_2',
                 'name_eng','name_cn','checkInDate','change','read_s','read_e']  
    if building == 'S2':
        itemList = index_S2
    elif building in ['S3','S4']:
        itemList = index_S3_S4
    else:
        itemList = index_S1
    return itemList

def row_infor_extractor(row,meter,lists, itemList,year_month):
    ''' 单行信息提取器：对一行信息进行提取，并将该行信息并入对应表中'''

    room, bed, room_type,ID,name_eng, name_cn,checkInDate,change = student_infor_extractor(row,itemList)
    change_list,empty_bed_list,Occupied_list = lists
    print('\r    加载RAW表数据: {}'.format(bed),end='', flush=True)
    
    # 该床位无住客（表格内容为#N/A），录入empty_bed_list
    if ID is None:  
        empty_bed_list.loc[empty_bed_list.shape[0] + 1] = {
            'Room': room,
            'Bed': bed,
            'Room Type': room_type
        }
    # 有住客，录入 Occupied_list
    else: 
        # 床位无变动,但不是check in，普通处理
        if change in ['',None] and is_check_in(checkInDate,year_month)==False:  
            quota = 71.42
            ele_s,ele_e = meter[room]
        # 床位有变化，但不是这个月的，故普通处理
        elif change not in ['',None] and change_is_in_this_month(change,bed,year_month)==False:
            quota = 71.42
            ele_s,ele_e = meter[room]
        # 床位该月有变化，录入change_list中
        else:
            if change in ['',None] and is_check_in(checkInDate,year_month)==True:
                change_type, quota = 'check in',71.42
            else:
                change_type,_,_,quota = change_list_processor(change,bed)
            if change_type in ['check in','move in']:
                try:
                    read_when_change = float(row[itemList.index('read_s')])
                except Exception:
                    print('\n    Error: {} 的宿生{t}，请在填入[{t}时电费度数]！'.format(bed,t=change_type))
                    input('\n\n按【回车键】退出...')
                    exit(-1)
                ele_s = read_when_change
                ele_e = meter[room][1]
            elif change_type in ['check out','move out']:
                try:
                    read_when_change = float(row[itemList.index('read_e')])
                except Exception:
                    print('\n    Error: {} 的宿生{t}，请填入[{t}时的电费度数]！'.format(bed,t=change_type))
                    input('\n\n按【回车键】退出...')
                    exit(-1)  
                ele_s = meter[room][0]
                ele_e = read_when_change
            change_list.loc[change_list.shape[0] + 1] = {
                    'Room': room,
                    'Bed': bed,
                    'ID': ID,
                    'name_eng': name_eng,
                    'name_cn': name_cn,
                    'Check In Date': checkInDate,
                    'Change': change,
                    'Type' : change_type
                    }

        Occupied_list.loc[Occupied_list.shape[0] + 1] = {
                'Room': room,
                'Bed': bed,
                'Room Type':room_type,
                'ID':ID,
                'Name(eng)':name_eng,
                'Name(cn)': name_cn, 
                'Check-in Date':checkInDate,
                'Change': change,
                'Quota': quota,
                '个人起始读数': ele_s,
                '个人结束读数': ele_e
                }
    
    return Occupied_list, empty_bed_list, change_list



def infor_extractor(Raw,meter,itemList,year_month):
    ''' 信息提取器：对Raw表遍历地进行信息读取，将所有信息并入总表中，为下一步计算做好准备 '''

	# 先生成三个空表：change_list,empty_bed_list,Occupied_list 
    lists = creat_list() 
    # Raw表：跳过表头，直达数据开始行
    Raw_ROW_Begining = locate_item(Raw,['轉房或退宿日期','Remarks'])[0]+1

    # 遍历各行，并将各行信息分别并入三表中
    for line in range(Raw_ROW_Begining, Raw.nrows):
        row = Raw.row_values(line)  # 读取某一行
        # Occupied_list,empty_bed_list,change_list 
        Lists = row_infor_extractor(row,meter,lists,itemList,year_month)
    print('\r    加载RAW表数据: Done!        ')
    print('    所有数据已加载并预处理完毕！\n') 

    # 将 Occupied_list,empty_bed_list合并得到全表
    Full_list = pd.concat([Lists[0], Lists[1]], axis=0, 
                              join='outer', ignore_index=True)
    columns = ['Room', 'Bed','Room Type','ID','Name(eng)',
              'Name(cn)','Check-in Date','Change',
              'Quota', '个人起始读数','个人结束读数']
    Full_list = Full_list.reindex_axis(columns, axis=1)
    # 将DataFrame按床位号重新排序，将None的位置填入''
    Full_list = Full_list.sort_values(by = 'Bed').fillna('').reset_index(drop=True)
    
    return Full_list,Lists[0],Lists[1],Lists[2]



def calculator(meter, Full_list, Occupied_list):
    '''信息提取好后，开始计算，这也是最核心的部分'''

    warning = [] #异常信息将放在这里

    # FeeList：最终结算的表，将直接转成结果Excel
    FeeList = pd.DataFrame(columns=['Room', 'Bed','Room type', 'ID',
                                    'Name(eng)','Name(cn)','Check-in Date','Change',
                                    'Room Start','Room End','Room Used',
                                    'Personal Start','Personal End','Difference',
                                    'Personal Used','Quota', 
                                    'Paying Degree', 'Fee'])

    for index,row in Full_list.iterrows():
        room = row['Room']
        bed = row['Bed']
        ID = row['ID']
        quota = row['Quota']
        print('\r    Calculating: {}'.format(bed),end='', flush=True)
        room_start,room_end = meter[room]
        try:
            room_used = room_end - room_start # Meter表中没有记录该房间
        except Exception:
            room_used = None 
            
        if ID in ['',None,'nan']:
            # 该床位无住客
            room_start,room_end, room_used = None, None, None
            #ele_s, ele_e = room_start,room_end
            ele_s, ele_e = None, None
            #dif = room_used
            dif = None
            used_ind, quota, degree, fee = None, None, None, None
            
        else:
            # 该床有住客
            this_room = Occupied_list[Occupied_list['Room'] == room]
            items = this_room.shape[0] # 这个房间的条目数
            this_person = this_room[this_room['Bed'] == bed].reset_index(drop=True)
            ele_s, ele_e = this_person.iloc[0,-2:]
            dif = ele_e - ele_s
            if items == 1: # 该房间只有一个人住（单人间或独享双人间）
                used_ind = dif
                degree = used_ind - quota
            else:
                roomate = this_room[this_room['Bed'] != bed].reset_index(drop=True)
                roomateNumber = roomate.shape[0]
                
                if roomateNumber == 1:
                    # 只有一个室友，情况比较简单
                    ele_s_rm, ele_e_rm = roomate.iloc[0, -2:]
                    overlap = min(ele_e, ele_e_rm) - max(ele_s, ele_s_rm)
                elif roomateNumber == 0:
                    overlap = 0
                    warning.append([bed, '异常：该床位信息有重复了两次！'])
                elif roomateNumber == 2:
                    if roomate.ix[0,'ID'] == roomate.ix[1,'ID']:
                        ele_s_rm, ele_e_rm = roomate.iloc[0, -2:]
                        overlap = min(ele_e, ele_e_rm) - max(ele_s, ele_s_rm)
                        warning.append([bed, '异常：该床位的舍友信息重复了两次'])
                    else:
                        ele_s_rm1, ele_e_rm1 = roomate.iloc[0, -2:]
                        ele_s_rm2, ele_e_rm2 = roomate.iloc[1, -2:]
                        overlap1 = min(ele_e, ele_e_rm1) - max(ele_s, ele_s_rm1)
                        overlap2 = min(ele_e, ele_e_rm2) - max(ele_s, ele_s_rm2)
                        overlap = overlap1 + overlap2
                        warning.append([bed, '异常：该同学在该月有两个不同的舍友'])
                elif roomateNumber > 2:
                    ID_list = list(set(roomate['ID'].values.tolist()))
                    if len(ID_list) ==1:
                        # 有重复，实际只有一个室友
                        ele_s_rm, ele_e_rm = roomate.iloc[0, -2:]
                        overlap = min(ele_e, ele_e_rm) - max(ele_s, ele_s_rm)
                        warning.append([bed, '异常：该床位的舍友信息重复了三次'])
                    elif len(ID_list) == 2:
                        # 无重复，确有两个室友
                        ele_s_rm1, ele_e_rm1 = roomate.iloc[0, -2:]
                        ele_s_rm2, ele_e_rm2 = roomate.iloc[1, -2:]
                        overlap1 = min(ele_e, ele_e_rm1) - max(ele_s, ele_s_rm1)
                        overlap2 = min(ele_e, ele_e_rm2) - max(ele_s, ele_s_rm2)
                        overlap = overlap1 + overlap2
                        warning.append([bed, '异常：该同学在该月有两个室友，并且有个舍友信息重复了两次，最好手动检查一下！'])
                    else:
                        overlap = 0
                        warning.append([bed, '异常：该同学在该月有两个以上室友，情况太复杂，系统无法计算，请手动计算！'])
                own = dif - overlap
                used_ind = overlap / 2 + own
                degree = used_ind - quota
            if degree > 0:
                fee = degree * 1.4
            else:
                fee = 0

            dif = round(dif,2)
            used_ind = round(used_ind,2)
            degree =round(degree,2)
            room_used = round(room_used,2)
            fee = round(fee,2)
                
        FeeList.loc[FeeList.shape[0] + 1] = {
            'Room': room,
            'Bed': bed,
            'Room type': row['Room Type'],
            'ID':ID,
            'Name(eng)': row['Name(eng)'],
            'Name(cn)': row['Name(cn)'],
            'Check-in Date':row['Check-in Date'],
            'Change': row['Change'],
            'Room Start': room_start,
            'Room End': room_end,
            'Room Used':room_used,
            'Personal Start': ele_s,
            'Personal End': ele_e,
            'Difference': dif,
            'Personal Used':used_ind,
            'Quota': quota,
            'Paying Degree': degree,
            'Fee': fee
        }
        
    warning = pd.DataFrame(warning)
    print('\r    Calculating: Done!           ', flush=True)
    return FeeList,warning


        

def main():
    ''' 主函数 '''  
    print('\n运行时请关闭被操作的Excel文件！\n\n程序开始运行...')
    
    # 1.加载数据并预处理
    print('\n1.加载数据并预处理')
    file,fileName,building, year_month = load_file()
    Raw, Data, Meter = load_sheets(file)
    meter = Meter_read_extractor(Meter)
    itemList = select_itemList(building)
    Full_list,OccupiedList,emptyList,changeList = infor_extractor(Raw,meter,itemList,year_month)

    # 2.计算电费
    print('\n2.开始计算电费\n')
    FeeList,warning = calculator(meter,Full_list, OccupiedList)

    # 3.计算结果写入Excel表中
    print('\n\n3.生成电费结果文件（Excel）')
    nameStr ='【Fee】{}'.format(fileName)
    writer = pd.ExcelWriter('../Result/'+nameStr)
    FeeList.to_excel(writer, sheet_name = '结算')
    warning.to_excel(writer, sheet_name ='异常提醒')
    emptyList.to_excel(writer, sheet_name ='空床位汇总')
    changeList.to_excel(writer, sheet_name ='床位变动汇总')
    OccupiedList.to_excel(writer, sheet_name ='现有住客汇总')
    Full_list.to_excel(writer, sheet_name ='全部汇总')

    while True:
        try:
            writer.save()
            print('\n    输出结果请见Result文件夹下:\n    {}'.format(nameStr))
            break
        except Exception as ex:
            print(ex)
            print('\nError：写入文件失败！请检查当前是否打开了该同名文件:')
            print('      {}'.format(nameStr))
            input('\n如有打开请关闭后按【回车键】继续......')
            
    print('\n\nAwesome! 大功告成！\n')
    print('-----------------------------')
    print('     Powered by Raymond')
    print('-----------------------------')
    input('\n\n按【回车键】退出...')


if __name__ == '__main__':
    try:
        import numpy as np
        import pandas as pd
        import xlrd
    except Exception as ex:
        print(ex)
        input('\nError：加载python库失败，请检查重试！\n按【回车键】退出...')
        exit(-1)

    try:
        main()
    except Exception as ex:
        print('\nError: {}'.format(ex))
        input('\n\n按任意键退出...')
             
        
=======
﻿# -*- coding: utf-8 -*-
"""
Created on Thu Nov 23 00:40:42 2017
@author: Raymond
"""   
def find_file():
    ''' 从文件夹中查找Excel文件 '''
    print("\n    Finding Excel file...")
    import glob
    try:
        filePath = []
        for fileType in ['xlsx','xls']:
            path = glob.glob(r'../*.{}'.format(fileType))
            if path!=[]: filePath.extend(path)
    except Exception as ex:
        print(ex)
        input('Error：调用glob寻找Excel文件失败，请重试！\n按【回车键】退出...')
        exit(-1)
    files = [path.split('\\')[-1] for path in filePath]
    return files

def find_in_fileName(fileName, itemList):
    for i in itemList:
        if fileName.find(i) != -1:
            item = i
            break
    return item


def load_file():
    ''' 第一步：加载Excel文件，从文件名中提取 建筑编号、年份和月份 '''
    files = find_file()
    fileNum = len(files)
    if fileNum < 1:
        print('\nError: 当前文件夹下未检测到Excel文件，请复制要处理的Excel文件到该文件夹下！')
        input('\n按【回车键】退出...')   
        exit(-1)
    else:
        if fileNum == 1:
            fileName = files[0]
        elif fileNum > 1:
            print('    当前文件夹下检测到以下Excel文件:')
            for i in range(fileNum):
                    print('        {}  {}'.format(i+1,files[i]))
            while True:
                try:
                    select = input('    请选择要处理的文件[1-{}]: '.format(fileNum))
                    fileName = files[int(select)-1]
                    break
                except Exception:
                    print('    Error：文件选择错误!\n')  
    
    file = xlrd.open_workbook('../'+fileName)
    print('    已获取Excel文件： {}'.format(fileName))

    buildingList = ['S1','S2','S3','S4']
    monthList = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    # 从文件名中提取建筑编号、年份和月份
    building = find_in_fileName(fileName, buildingList)
    month = find_in_fileName(fileName, monthList)
    year = fileName.split('.')[0].strip()[-4:]
    print('    Building: {}   Month: {} {}'.format(building, month, year))

    month = str(monthList.index(month)+1)
    if len(month)==1:
        month = '0'+ month
    year_month = year+'-'+month
    return file,fileName,building, year_month

def load_sheets(file):
    ''' 加载Excel各表 '''
    for sheetName in file.sheet_names():
        if sheetName.lower() in ['raw']:
            Raw = file.sheet_by_name(sheetName)
        elif sheetName.lower() in ['meter record','meter']:
            Meter = file.sheet_by_name(sheetName)
        elif sheetName.lower() in ['data']:
            Data = file.sheet_by_name(sheetName)
    return Raw, Data, Meter

def quota_out(date):
    if date > 0 and date <= 7:
        quota = 17.85
    elif date > 7 and date <= 15:
        quota = 35.71
    elif date > 15 and date <= 23:
        quota = 53.57
    else:
        quota = 71.42
    return quota

def quota_in(date):
    if date > 0 and date <= 7:
        quota = 53.57
    elif date > 7 and date <= 15:
        quota = 35.71
    elif date > 15 and date <= 23:
        quota = 17.85
    else:
        quota = 0
    return quota

def change_is_in_this_month(change,bed,ym):
    ''' 判断床位变动是否发生在本月 '''
    string = change.strip().split(' ') # 去除字符串前后空格
    monthList = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    month_check = False
    for i in monthList:
        if i.lower() in change.lower():
            month_check = True
            month = str(monthList.index(i)+1)
            if len(month)==1:
                month = '0'+ month
            break
    if month_check == True: 
        if month == ym[-2:]:
            return True
        else:
            return False
    else:
        print('\n\nError： {} 床位的 \"轉房或退宿日期\" 栏内容为： {}'.format(bed,change))
        print('        该拼写可能有误，程序被迫中止，请更正后重新运行该程序！\n')
        input('按【回车键】退出...')
        exit(-1)

def change_list_processor(change,bed):
    ''' 对该床位变动信息进行分类 '''
    string = change.strip().split(' ') # 去除行前空格
    try:
        if string[0] == 'check-in':
            change_type = 'check in'
            day = float(string[2])
            date = string[3]
            quota = 71.42
        elif string[0] == 'stay-in':   # stay-in only on 3 Oct 2017 
            change_type = 'move in'
            day = float(string[3])
            date = string[3:6]
            quota = quota_in(day)
        elif string[0] == 'move':      # move to S308037A on 4 Sep 2017 (charge)
            change_type = 'move out'
            day = float(string[4])
            date = string[4:7]
            quota = quota_out(day)
        elif string[0] == 'check-out': # check-out on 14 Dec 2017 (charge)
            change_type = 'check out'
            day = float(string[2])
            date = string[2:5]
            quota = 71.42
        return change_type, day, date, quota
    except Exception:
        print('\n\nError： {} 床位的 \"轉房或退宿日期\" 栏内容为： {}'.format(bed,change))
        print('        该格式可能有误，程序被迫中止，请更正后重新运行该程序！\n')
        input('按【回车键】退出...')
        exit(-1)

def Index(List,item):
    try:
        return List.index(item)
    except:
        return None

def student_infor_extractor(row,List):
    room = row[Index(List,'room')]
    bed = row[Index(List,'bed')]
    try:
        room_type = row[Index(List,'room_type')]
    except:
        room_type = None
    ID = row[Index(List,'studentID_1')]+row[Index(List,'studentID_2')]
    if ID in [42,84,'',' ']: 
        ID = None
    name_eng = row[Index(List,'name_eng')]
    try:
        name_cn = row[Index(List,'name_cn')]
    except:
        name_cn = None
    try:
        checkInDate = xlrd.xldate.xldate_as_datetime(float(row[Index(List,'checkInDate')]), 0)
        checkInDate = str(checkInDate)[:10]
    except Exception:
        checkInDate = None
    change = row[Index(List,'change')]
    
    return room, bed, room_type,ID,name_eng, name_cn,checkInDate,change

def creat_list():
    change_list = pd.DataFrame(columns=['Room', 'Bed','ID','name_eng',
                                        'name_cn','Check In Date','Change','Type'])
    empty_bed_list = pd.DataFrame(columns=['Room', 'Bed', 'Room Type'])
    Occupied_list = pd.DataFrame(columns=['Room', 'Bed','Room Type',
                                          'ID','Name(eng)',
                                          'Name(cn)','Check-in Date','Change',
                                          'Quota', '个人起始读数','个人结束读数'])
    return [change_list,empty_bed_list,Occupied_list]

def is_check_in(checkInDate,ym):
    return checkInDate[:7] == ym

def locate_item(sheet,strings):
    for line in range(8):
        row_number = line
        row = sheet.row_values(line)
        for i in strings:
            col_number = Index(row,i)
            if col_number != None:
                return [row_number,col_number]

def Meter_read_extractor(Meter):
    print('    加载Meter表的数据')
    row_number, col_number=locate_item(Meter,['房號'])
    Meter_ROW_Begining = row_number+1
    meter = {}
    for line in range(Meter_ROW_Begining, Meter.nrows):
        row = Meter.row_values(line)
        room = row[col_number]
        read = row[col_number+2:col_number+4]
        meter[room] = read
    return meter

def select_itemList(building):
    ''' 因为S1-S4各表的项目栏顺序不一样，故需分别确定，以免对错 '''

    index_S2 = ['bed','room','sex','room_type','studentID_1','studentID_2',
                 'name_eng','checkInDate','change','read_s','read_e']
    index_S3_S4 = ['room','bed','room_type','studentID_1','studentID_2',
                 'name_eng','name_cn','checkInDate','change','read_s','read_e']
    index_S1 = ['room','bed','studentID_1','studentID_2',
                 'name_eng','name_cn','checkInDate','change','read_s','read_e']  
    if building == 'S2':
        itemList = index_S2
    elif building in ['S3','S4']:
        itemList = index_S3_S4
    else:
        itemList = index_S1
    return itemList

def row_infor_extractor(row,meter,lists, itemList,year_month):
    ''' 单行信息提取器：对一行信息进行提取，并将该行信息并入对应表中'''

    room, bed, room_type,ID,name_eng, name_cn,checkInDate,change = student_infor_extractor(row,itemList)
    change_list,empty_bed_list,Occupied_list = lists
    print('\r    加载RAW表数据: {}'.format(bed),end='', flush=True)
    
    # 该床位无住客（表格内容为#N/A），录入empty_bed_list
    if ID is None:  
        empty_bed_list.loc[empty_bed_list.shape[0] + 1] = {
            'Room': room,
            'Bed': bed,
            'Room Type': room_type
        }
    # 有住客，录入 Occupied_list
    else: 
        # 床位无变动,但不是check in，普通处理
        if change in ['',None] and is_check_in(checkInDate,year_month)==False:  
            quota = 71.42
            ele_s,ele_e = meter[room]
        # 床位有变化，但不是这个月的，故普通处理
        elif change not in ['',None] and change_is_in_this_month(change,bed,year_month)==False:
            quota = 71.42
            ele_s,ele_e = meter[room]
        # 床位该月有变化，录入change_list中
        else:
            if change in ['',None] and is_check_in(checkInDate,year_month)==True:
                change_type, quota = 'check in',71.42
            else:
                change_type,_,_,quota = change_list_processor(change,bed)
            if change_type in ['check in','move in']:
                try:
                    read_when_change = float(row[itemList.index('read_s')])
                except Exception:
                    print('\n    Error: {} 的宿生{t}，请在填入[{t}时电费度数]！'.format(bed,t=change_type))
                    input('\n\n按【回车键】退出...')
                    exit(-1)
                ele_s = read_when_change
                ele_e = meter[room][1]
            elif change_type in ['check out','move out']:
                try:
                    read_when_change = float(row[itemList.index('read_e')])
                except Exception:
                    print('\n    Error: {} 的宿生{t}，请填入[{t}时的电费度数]！'.format(bed,t=change_type))
                    input('\n\n按【回车键】退出...')
                    exit(-1)  
                ele_s = meter[room][0]
                ele_e = read_when_change
            change_list.loc[change_list.shape[0] + 1] = {
                    'Room': room,
                    'Bed': bed,
                    'ID': ID,
                    'name_eng': name_eng,
                    'name_cn': name_cn,
                    'Check In Date': checkInDate,
                    'Change': change,
                    'Type' : change_type
                    }

        Occupied_list.loc[Occupied_list.shape[0] + 1] = {
                'Room': room,
                'Bed': bed,
                'Room Type':room_type,
                'ID':ID,
                'Name(eng)':name_eng,
                'Name(cn)': name_cn, 
                'Check-in Date':checkInDate,
                'Change': change,
                'Quota': quota,
                '个人起始读数': ele_s,
                '个人结束读数': ele_e
                }
    
    return Occupied_list, empty_bed_list, change_list



def infor_extractor(Raw,meter,itemList,year_month):
    ''' 信息提取器：对Raw表遍历地进行信息读取，将所有信息并入总表中，为下一步计算做好准备 '''

	# 先生成三个空表：change_list,empty_bed_list,Occupied_list 
    lists = creat_list() 
    # Raw表：跳过表头，直达数据开始行
    Raw_ROW_Begining = locate_item(Raw,['轉房或退宿日期','Remarks'])[0]+1

    # 遍历各行，并将各行信息分别并入三表中
    for line in range(Raw_ROW_Begining, Raw.nrows):
        row = Raw.row_values(line)  # 读取某一行
        # Occupied_list,empty_bed_list,change_list 
        Lists = row_infor_extractor(row,meter,lists,itemList,year_month)
    print('\r    加载RAW表数据: Done!        ')
    print('    所有数据已加载并预处理完毕！\n') 

    # 将 Occupied_list,empty_bed_list合并得到全表
    Full_list = pd.concat([Lists[0], Lists[1]], axis=0, 
                              join='outer', ignore_index=True)
    columns = ['Room', 'Bed','Room Type','ID','Name(eng)',
              'Name(cn)','Check-in Date','Change',
              'Quota', '个人起始读数','个人结束读数']
    Full_list = Full_list.reindex_axis(columns, axis=1)
    # 将DataFrame按床位号重新排序，将None的位置填入''
    Full_list = Full_list.sort_values(by = 'Bed').fillna('').reset_index(drop=True)
    
    return Full_list,Lists[0],Lists[1],Lists[2]



def calculator(meter, Full_list, Occupied_list):
    '''信息提取好后，开始计算，这也是最核心的部分'''

    warning = [] #异常信息将放在这里

    # FeeList：最终结算的表，将直接转成结果Excel
    FeeList = pd.DataFrame(columns=['Room', 'Bed','Room type', 'ID',
                                    'Name(eng)','Name(cn)','Check-in Date','Change',
                                    'Room Start','Room End','Room Used',
                                    'Personal Start','Personal End','Difference',
                                    'Personal Used','Quota', 
                                    'Paying Degree', 'Fee'])

    for index,row in Full_list.iterrows():
        room = row['Room']
        bed = row['Bed']
        ID = row['ID']
        quota = row['Quota']
        print('\r    Calculating: {}'.format(bed),end='', flush=True)
        room_start,room_end = meter[room]
        try:
            room_used = room_end - room_start # Meter表中没有记录该房间
        except Exception:
            room_used = None 
            
        if ID in ['',None,'nan']:
            # 该床位无住客
            room_start,room_end, room_used = None, None, None
            #ele_s, ele_e = room_start,room_end
            ele_s, ele_e = None, None
            #dif = room_used
            dif = None
            used_ind, quota, degree, fee = None, None, None, None
            
        else:
            # 该床有住客
            this_room = Occupied_list[Occupied_list['Room'] == room]
            items = this_room.shape[0] # 这个房间的条目数
            this_person = this_room[this_room['Bed'] == bed].reset_index(drop=True)
            ele_s, ele_e = this_person.iloc[0,-2:]
            dif = ele_e - ele_s
            if items == 1: # 该房间只有一个人住（单人间或独享双人间）
                used_ind = dif
                degree = used_ind - quota
            else:
                roomate = this_room[this_room['Bed'] != bed].reset_index(drop=True)
                roomateNumber = roomate.shape[0]
                
                if roomateNumber == 1:
                    # 只有一个室友，情况比较简单
                    ele_s_rm, ele_e_rm = roomate.iloc[0, -2:]
                    overlap = min(ele_e, ele_e_rm) - max(ele_s, ele_s_rm)
                elif roomateNumber == 0:
                    overlap = 0
                    warning.append([bed, '异常：该床位信息有重复了两次！'])
                elif roomateNumber == 2:
                    if roomate.ix[0,'ID'] == roomate.ix[1,'ID']:
                        ele_s_rm, ele_e_rm = roomate.iloc[0, -2:]
                        overlap = min(ele_e, ele_e_rm) - max(ele_s, ele_s_rm)
                        warning.append([bed, '异常：该床位的舍友信息重复了两次'])
                    else:
                        ele_s_rm1, ele_e_rm1 = roomate.iloc[0, -2:]
                        ele_s_rm2, ele_e_rm2 = roomate.iloc[1, -2:]
                        overlap1 = min(ele_e, ele_e_rm1) - max(ele_s, ele_s_rm1)
                        overlap2 = min(ele_e, ele_e_rm2) - max(ele_s, ele_s_rm2)
                        overlap = overlap1 + overlap2
                        warning.append([bed, '异常：该同学在该月有两个不同的舍友'])
                elif roomateNumber > 2:
                    ID_list = list(set(roomate['ID'].values.tolist()))
                    if len(ID_list) ==1:
                        # 有重复，实际只有一个室友
                        ele_s_rm, ele_e_rm = roomate.iloc[0, -2:]
                        overlap = min(ele_e, ele_e_rm) - max(ele_s, ele_s_rm)
                        warning.append([bed, '异常：该床位的舍友信息重复了三次'])
                    elif len(ID_list) == 2:
                        # 无重复，确有两个室友
                        ele_s_rm1, ele_e_rm1 = roomate.iloc[0, -2:]
                        ele_s_rm2, ele_e_rm2 = roomate.iloc[1, -2:]
                        overlap1 = min(ele_e, ele_e_rm1) - max(ele_s, ele_s_rm1)
                        overlap2 = min(ele_e, ele_e_rm2) - max(ele_s, ele_s_rm2)
                        overlap = overlap1 + overlap2
                        warning.append([bed, '异常：该同学在该月有两个室友，并且有个舍友信息重复了两次，最好手动检查一下！'])
                    else:
                        overlap = 0
                        warning.append([bed, '异常：该同学在该月有两个以上室友，情况太复杂，系统无法计算，请手动计算！'])
                own = dif - overlap
                used_ind = overlap / 2 + own
                degree = used_ind - quota
            if degree > 0:
                fee = degree * 1.4
            else:
                fee = 0

            dif = round(dif,2)
            used_ind = round(used_ind,2)
            degree =round(degree,2)
            room_used = round(room_used,2)
            fee = round(fee,2)
                
        FeeList.loc[FeeList.shape[0] + 1] = {
            'Room': room,
            'Bed': bed,
            'Room type': row['Room Type'],
            'ID':ID,
            'Name(eng)': row['Name(eng)'],
            'Name(cn)': row['Name(cn)'],
            'Check-in Date':row['Check-in Date'],
            'Change': row['Change'],
            'Room Start': room_start,
            'Room End': room_end,
            'Room Used':room_used,
            'Personal Start': ele_s,
            'Personal End': ele_e,
            'Difference': dif,
            'Personal Used':used_ind,
            'Quota': quota,
            'Paying Degree': degree,
            'Fee': fee
        }
        
    warning = pd.DataFrame(warning)
    print('\r    Calculating: Done!           ', flush=True)
    return FeeList,warning


        

def main():
    ''' 主函数 '''  
    print('\n运行时请关闭被操作的Excel文件！\n\n程序开始运行...')
    
    # 1.加载数据并预处理
    print('\n1.加载数据并预处理')
    file,fileName,building, year_month = load_file()
    Raw, Data, Meter = load_sheets(file)
    meter = Meter_read_extractor(Meter)
    itemList = select_itemList(building)
    Full_list,OccupiedList,emptyList,changeList = infor_extractor(Raw,meter,itemList,year_month)

    # 2.计算电费
    print('\n2.开始计算电费\n')
    FeeList,warning = calculator(meter,Full_list, OccupiedList)

    # 3.计算结果写入Excel表中
    print('\n\n3.生成电费结果文件（Excel）')
    nameStr ='【Fee】{}'.format(fileName)
    writer = pd.ExcelWriter('../Result/'+nameStr)
    FeeList.to_excel(writer, sheet_name = '结算')
    warning.to_excel(writer, sheet_name ='异常提醒')
    emptyList.to_excel(writer, sheet_name ='空床位汇总')
    changeList.to_excel(writer, sheet_name ='床位变动汇总')
    OccupiedList.to_excel(writer, sheet_name ='现有住客汇总')
    Full_list.to_excel(writer, sheet_name ='全部汇总')

    while True:
        try:
            writer.save()
            print('\n    输出结果请见Result文件夹下:\n    {}'.format(nameStr))
            break
        except Exception as ex:
            print(ex)
            print('\nError：写入文件失败！请检查当前是否打开了该同名文件:')
            print('      {}'.format(nameStr))
            input('\n如有打开请关闭后按【回车键】继续......')
            
    print('\n\nAwesome! 大功告成！\n')
    print('-----------------------------')
    print('     Powered by Raymond')
    print('-----------------------------')
    input('\n\n按【回车键】退出...')


if __name__ == '__main__':
    try:
        import numpy as np
        import pandas as pd
        import xlrd
    except Exception as ex:
        print(ex)
        input('\nError：加载python库失败，请检查重试！\n按【回车键】退出...')
        exit(-1)

    try:
        main()
    except Exception as ex:
        print('\nError: {}'.format(ex))
        input('\n\n按任意键退出...')
             
        
>>>>>>> 6d8f4927b391e9258ec9e9b2209f61b4c060efd9
        