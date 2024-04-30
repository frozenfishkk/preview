
import xlrd
from read_excel import formatValue

map =  [
            ["攻击威力", "自身攻击%"],
            ["附加伤害", ""],
            ["目标当前生命百分比", "其当前生命%"],
            ["目标最大生命百分比", "其最大生命%"],
            ["自身当前生命百分比", "自身当前生命%"],
            ["自身最大生命百分比", "自身最大生命%"],
            ["附加护甲", "自身护甲%"],
            ["目标已损失生命值", "其已损失生命%"],
            ["百分比伤害上限", "自身攻击%"],
            ["附加已损失生命值", "自身已损失生命值%"],
            ["上一次造成的伤害", "上一次造成伤害的%"],
            ["自身护盾百分比", "自身护盾%"],
            ["上一次造成的治疗", "%"],
            ["上一次受到的伤害", "%"],
            ["目标负面状态数量关联伤害", "自身攻击%"],
            ["状态层数关联伤害", "自身攻击%"],
            ["目标状态ID", "自身攻击%"],
            ["自身状态ID", "自身攻击%"],
            ["暴击伤害额外提高", "爆头伤害额外提高%"],
            ["单位距离伤害修正", ""],
        ]

query_map = {'能量回复速度':'mana_gen', '物理防御':'p_def', '魔法防御':'m_def', '百分比加物理防御':'p_def_percent',
                 '百分比加魔法防御':'m_def_percent', '物理穿甲率':'arp_percent', '魔法穿甲率%':'spp_percent', '攻击百分比':'atk_percent',
                 '命中率':'hit_rate', '爆头率':'cri_head_rate', '爆头伤害':'cri_head_dmg', '闪避率':'miss',
                 '伤害提升率':'damage_percent', '攻速加成百分比':'attack_speed_up', '伤害减免率':'damage_reduce_percent'
                 }

workbook_stateconf = xlrd.open_workbook(filename=rf'D:\workspace\aquaman\design\00_导表\X_Data\战斗\属性配置表.xlsx')
sheet = workbook_stateconf.sheet_by_name(sheet_name='Sheet1')
crows = sheet.nrows
ccols = sheet.ncols
sheetHeader = sheet.row_values(0)
state_conf_data = {}
for row in range(0,crows):
    row_value = sheet.row_values(row)
    state_enid = row_value[1]
    state_conf_data[state_enid] = {k: v for k, v in zip(sheetHeader, row_value)}
#————————————————————————————————————
workbook = xlrd.open_workbook(filename=rf'D:\workspace\aquaman\design\00_导表\X_Data\战斗\状态表.xlsx')
sheet = workbook.sheet_by_name(sheet_name='状态表')
rows = sheet.nrows
cols = sheet.ncols
sheetHeader = sheet.row_values(0)
IDIndex = sheetHeader.index("状态ID")
ID = '1310503'
state_data = {}
for row in range(5,rows):
    row_value = sheet.row_values(row)
    state_id = row_value[IDIndex]
    state_data[state_id] = {k: v for k, v in zip(sheetHeader, row_value)}
data = state_data.get(int(ID))
text = ''
for key, replaceValue in query_map.items():
    value = data.get(key)
    cname = state_conf_data[replaceValue]['显示名称']
    print(value)
    if not value:
        continue
    value = formatValue(value)
    print(value)
    if '百分比' in key or '率' in key:
        value = int(value/100)
    if value > 0:
        text += f'{cname}提高{value}%'
    elif value < 0:
        text += f'{cname}降低{int(value)}%'.replace('-','')
print(text)














