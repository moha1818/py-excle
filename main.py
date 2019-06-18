# coding=utf-8
from openpyxl import load_workbook
import requests
import json



def requestApi(data):
    url="http://115.231.208.38:8099/mnbyd/parkpot/parkpotAction.do?"+data
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36',
        'Referer': 'http://115.231.208.38:8099/mnbyd/acl/userAction.do?doaction=lgok',
        'Connection': 'keep-alive',
        'Content-Type':'application/x-www-form-urlencoded; charset=utf-8',
        'Cookie':'JSESSIONID=27D666118E73570F4975DF09AB843753',
        'Host':'115.231.208.38:8099',
        'X-Requested-With' : 'XMLHttpRequest'
    }
    req = requests.get(url, headers=headers)
    print(req.status_code)

wb = load_workbook('停车场列表.xlsx')
sheet = wb.active
street = {'蓝天路':'055','外潜龙街':'057','雷公巷':'059','东胜路':'060','马园路':'074','三市路':'079','会展路':'083','麦德龙路':'080','前河南路':'087','星海南路':'088','光华路':'089','惠风东路':'111','鄞县大道':'113','诚信路':'115','泰安中路':'116','机场路':'118','古林镇':'120','启文路':'119','宁东路':'130','青少年宫南路':'132','金山八路与金山路交叉口':'134','启明路':'137','首南中路':'139','华山路':'140','彩虹北路':'141','三兴桥东路':'143','朝晖路':'144','兴海中路':'148','新马路':'149','长兴路':'150','西河街':'151','科技路':'153','长春路':'154','洪塘中路':'157','徐戎路':'159','启运路':'162','建设村西南方向':'164','泰康中路':'165','桑田路':'166','新市路':'168','下应段':'170','学士路与日丽中路交叉':'175','世纪大道':'177','岳林东路':'179','中山路':'180','江东北路':'181','康庄南路':'183','新义路':'186','钟包路':'187','盛海路':'188','肖家巷':'049','章耆街':'050','孝闻街':'051','文昌街':'052','国医街':'053','厂堂街':'054','潜龙路':'056','彩虹桥巷':'058','淇漕街':'061','青林湾街':'064','滨江大道':'066','槐树路':'067','南塘老街':'068','悠云路':'070','昌兴路':'071','四明中路':'112','学士路':'114','前河北路':'117','丽园南路':'063','广安路':'048','沁园街 ':'062','福明路':'065','鄞丰路':'069','甬江大道':'072','马衙街':'073','永丰北路':'076','栎社航空路':'078','文教路':'081','航辉路':'082','北柳街':'084','天童南路':'085','百丈东路':'086','环城南路':'091','梅墟路':'095','江南路':'096','永达路':'100','北海路':'101','宁南北路':'106','环城北路':'108','钱湖北路':'110','咸塘街':'026','开明街':'023','大来街':'024','大梁街':'025','柳汀街':'027','广济街':'028','大沙泥街':'029','西北街':'032','呼童街':'033','宁穿路':'034','大庆南路':'035','中马路':'001','人民路':'002','桃渡路':'003','车站路':'004','扬善路':'005','和义路':'006','苍水街':'007','县前街':'008','药局巷':'009','老实巷':'010','华楼巷':'011','碶闸街':'012','日新街':'013','车轿街':'014','东渡路':'015','石板巷':'016','灵桥路':'017','中山东路':'018','解放南路':'019','药行街':'020','解放北路':'039','狮子街':'022','白沙路':'036','镇明路':'030','南站东路':'031','苍松路':'037','三支街':'038','大闸南路':'040','县学街':'041','仓桥街':'042','望京路':'043','清河北路':'044','江安路':'045','鄞慈镇路':'046','槐新路':'047','聚贤路':'077','鄞州中山东路':'075','竺家巷':'090','姚隘路':'092','丽江东路':'093','大闸路':'094','江东北路':'097','句章东路':'098','元吉巷':'099','嵩江中路':'102','学院路':'103','兴宁路':'104','宁慈西路':'105','西草马路':'107','联丰中路':'109','海晏北路':'121','中河路':'123','民安东路与海晏北路交汇':'127','民安东路':'128','长兴东路':'193','百丈路':'194','梅竹路':'195','宁大步行街':'198','江宁路':'199','宁徐路':'200','天广路':'201','孙塘路':'202','风华路':'203','首南西路':'205','永丰路':'206','柳丁街':'208','北明程路':'209','舜水南路':'211','四明西路与宁南北路交叉口':'212','贸城西路':'213','高塘一村':'214','气象北路':'215','堇山中路':'216','荣骆路':'218','凤洋一路':'219','蝶缘路':'220'}

f = open('data.json',encoding='utf-8')
data = json.load(f)
parkpotids = set()
for n in range(0,150):
    parkpotids.add(data['data'][n]['parkpotid'])

for row_cell in sheet['A3':'O107']:
    list = ["doaction=saveParkpotInfo","&issave=0","&parkpotclassify=01","&parkpotstate=1","&minfreeparklotcount=-10","&paymode=1","&chargemode=4","&chargemode=4"]
    areaid = ""
    parkpotid = ""
    for cell in row_cell:
        if(cell.column == 'B'):
            parkpotid = str(cell.value)
            list.append("&parkpotid=")
            list.append(parkpotid)
            list.append("&mapno=P")
            list.append(str(cell.value))
        if(cell.column == 'C'):
            list.append("&parkpotname=")
            list.append(str(cell.value))
        if(cell.column == 'E'):
            list.append("&areaid=")
            area = str(cell.value)
            areaid = street.get(area[3:],"")
            list.append(areaid)
        if(cell.column == 'F'):
            list.append("&address=")
            list.append(str(cell.value))
        if(cell.column == 'G'):
            list.append("&telephone=")
            list.append(str(cell.value))
        if(cell.column == 'H'):
            list.append("&parkinnum=")
            parkinnum = str(cell.value)
            list.append(parkinnum[0])
        if(cell.column == 'I'):
            list.append("&parkoutnum=")
            parkoutnum = str(cell.value)
            list.append(parkoutnum[0])
        if(cell.column == 'J'):
            list.append("&totalparklotcount=")
            list.append(str(cell.value))
        if(cell.column == 'L'):
            list.append("&chargestandard=")
            list.append(str(cell.value))
        if(cell.column == 'O'):
            list.append("&department=")
            list.append(str(cell.value))
    if(areaid == ""):
        print("没有匹配到areaid的ID:"+parkpotid)
    requestApi(''.join(list))





