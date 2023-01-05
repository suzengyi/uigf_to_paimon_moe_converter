import json5
import urllib.request
import openpyxl
from datetime import datetime

_BANNER_URL="https://github.com/MadeBaruna/paimon-moe/raw/fe3298c16f840fc6c3e47c96ff83718c8a801279/src/data/banners.js"
_LOCALE_URL="https://github.com/MadeBaruna/paimon-moe/raw/main/src/locales/items/zh.json"
_ENCODE="utf-8"
_PAIMON_DATE_FORMAT="%Y-%m-%d %H:%M:%S"
_UIGF_DATE_FORMAT=_PAIMON_DATE_FORMAT
_PAIMON_TITLE=["Type","Name","Time","⭐","Pity","#Roll","Group","Banner","Part"]
_GACHA_TYPE_DICT={"100":["Beginners' Wish","新手祈愿"],
                  "200":["Standard", "常驻祈愿"],
                  "301":["Character Event","角色活动祈愿"],
                  "400":["Character Event","角色活动祈愿"],
                  "302":["Weapon Event","武器活动祈愿"]}
_PULL_TYPE_DICT={
    "Weapon":"武器",
    "Character":"角色"
}

def _get_key (dict, value):
    for k, v in dict.items():
         if v == value:
            return k
    return None
  
def _get_url(url:str)->str:
    url=urllib.request.urlopen(url)
    return url.read().decode(_ENCODE)

def get_banners()->dict:
    banner_js_str=_get_url(_BANNER_URL)
    banners_dict=json5.loads(banner_js_str.split('=')[1].split(';')[0])
    for banners in banners_dict.values():
        for wish in banners:
            for wish_key in wish:
                if wish_key != 'start' and wish_key != 'end':
                    continue
                wish[wish_key]=datetime.strptime(wish[wish_key], _PAIMON_DATE_FORMAT)
    return banners_dict

def get_locale()->dict:
    locale_raw = _get_url(_LOCALE_URL)
    return json5.loads(locale_raw)

def get_uigf_from_file(path:str)->dict:
    with open(path,"r",encoding=_ENCODE) as f:
        uigf_dict=json5.load(f,encoding=_ENCODE)
    return uigf_dict

def _init_workbook()->openpyxl.workbook.Workbook:
    ret=openpyxl.workbook.Workbook()
    ws = ret.active
    titles = [x[0] for x in _GACHA_TYPE_DICT.values()]
    ws.title = titles[0]
    ws.append(_PAIMON_TITLE)
    for i in range(1,len(titles)):
        if titles[i] in ret.sheetnames:
            continue
        ret.create_sheet(titles[i])
        ws=ret[titles[i]]
        ws.append(_PAIMON_TITLE)
    return ret

def get_pity(rank_type:str,
             ws:openpyxl.worksheet.worksheet.Worksheet)->int:
    pity_col = _PAIMON_TITLE.index("Pity")
    row_idx=ws.max_row
    if rank_type=="3":
        return 1
    elif row_idx < 2:
        return 1
    else:
        ret = 1
        while row_idx >1 and int(ws[row_idx][pity_col-1].value) <4:
            ret += 1
            row_idx -=1
        if ret >10:
            print(f"{ws.title}数据错误，Pity值大于10")
        return ret

def get_banner_by_date(
        banners_dict:dict,  #banner字典
        time:datetime,        #UIGF时间
        gacha_type:str      #UIGF gacha_type
        )->tuple:           #返回（标题:str，是否切换池子:bool）
    flag = False
    if gacha_type == '302':
        #武器活动祈愿
        idx = get_banner_by_date.weapon_idx
        banners = banners_dict['weapons']
    else:
        #角色活动祈愿
        idx = get_banner_by_date.character_idx
        banners = banners_dict['characters']
    while not banners[idx]['start'] < time < banners[idx]['end']:
        flag=True
        if time < banners[idx]['start']:
            print("向前移动")
            idx-=1
        else :
            idx+=1
    if gacha_type == '302':
        #武器活动祈愿
        get_banner_by_date.weapon_idx=idx
    else:
        #角色活动祈愿
        get_banner_by_date.character_idx = idx
    return banners[idx]['name'],flag
get_banner_by_date.character_idx=0
get_banner_by_date.weapon_idx=0

def _get_valid_col(row):
    ret = 0
    for i in row:
        if i.value is None:
            break
        else:
            ret +=1
    return ret

def fix_pity_row_group_by_sheet(ws:openpyxl.worksheet.worksheet.Worksheet):
    time_col=_PAIMON_TITLE.index("Time")
    star_col=_PAIMON_TITLE.index("⭐")
    pity_col=_PAIMON_TITLE.index("Pity")
    roll_col=_PAIMON_TITLE.index("#Roll")
    group_col=_PAIMON_TITLE.index("Group")
    banner_col=_PAIMON_TITLE.index("Banner")
    max_col=len(_PAIMON_TITLE)
    if ws.max_row > 1:
        ws[2][pity_col].value=1
        ws[2][roll_col].value=1
        ws[2][group_col].value=1
    if ws.max_row < 2:
        return
    for i in range(3,ws.max_row):
        #先填group
        if ws[i][time_col].value == ws[i-1][time_col].value:
            #时间相同为一组
            ws[i][group_col].value = ws[i-1][group_col].value
            time_same=True
        else :
            #时间不同为下一组，但是不确定是从1开始还是+1
            time_same = False
        #再比较banner和列数
        if ws[i][banner_col].value == ws[i-1][banner_col].value and\
            _get_valid_col(ws[i])<=max_col :
            #与上一卡池一致
            ws[i][roll_col].value = ws[i-1][roll_col].value+1
            if not time_same:
                ws[i][group_col].value = ws[i-1][group_col].value+1
            if ws[i][star_col].value == "3":
                ws[i][pity_col].value = 1
            
            j=i-1 #j指向i的上一行行号
            count=1
            while j>1 and \
                int(ws[j][star_col].value) < int(ws[i][star_col].value) and \
                ws[j][banner_col].value == ws[i][banner_col].value and \
                _get_valid_col(ws[j])<=max_col:
                j-=1
                count+=1
            ws[i][pity_col].value=count
        else:
            #新卡池
            ws[i][pity_col].value = 1
            ws[i][roll_col].value = 1
            ws[i][group_col].value = 1



def fix_pity_row_group(wb:openpyxl.workbook.Workbook):
    for sheet_name in wb.sheetnames:
        fix_pity_row_group_by_sheet(wb[sheet_name])
    wb.save("test_output2.xlsx")

def main():
    banners_dict=get_banners()
    locale_dict = get_locale()
    uigf_dict=get_uigf_from_file("uigf.json")
    workbook=_init_workbook()
    for gacha in uigf_dict['list']:
        worksheet=workbook[_GACHA_TYPE_DICT[gacha['gacha_type']][0]]
        row=[]
        #Type
        row.append(_get_key(_PULL_TYPE_DICT, gacha['item_type']))
        #Name
        name=_get_key(locale_dict, gacha['name'])
        if name == None:
            print(f"{gacha['name']}未找到对应翻译")
            row.append("")
        else:
            row.append(name)
        #Time
        time=datetime.strptime(gacha['time'], _UIGF_DATE_FORMAT)
        row.append(time.strftime(_PAIMON_DATE_FORMAT))
        #⭐
        row.append(gacha['rank_type'])
        #Pity 暂时不在这里添加
        row.append("")
        ##Roll 暂时不在这里添加
        row.append("")
        #Group 暂时不在这里添加
        row.append("")
        #Banner
        diff_flag = False
        if gacha['gacha_type'] == "100":
            row.append(banners_dict['beginners'][0]['name'])
        elif gacha['gacha_type'] == "200":
            row.append(banners_dict['standard'][0]['name'])
        else:
            banner, diff_flag = get_banner_by_date(banners_dict, time, gacha['gacha_type'])
            row.append(banner)
        #Part
        if gacha['gacha_type'] == "400":
            row.append("Wish 2")
        else :
            row.append("")
        if diff_flag == True:
            row.append("1")
        worksheet.append(row)
    workbook.save("test_output.xlsx")



if __name__=="__main__":
    #main()
    fix_pity_row_group(openpyxl.load_workbook("test_output.xlsx"))
    #fix_pity_row_group_by_sheet(openpyxl.load_workbook("test_output.xlsx")["Character Event"])