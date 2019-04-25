import xlrd
from datetime import datetime
import mysql.connector
conn = mysql.connector.connect(host="localhost", user='root', password='root', use_unicode=True, database ='data')
cursor = conn.cursor()
data = xlrd.open_workbook('../data/PATIENT_RECORDS_MX.xlsx')
import re
re_name = r'.*姓名：(.+)出.*生.*地:(.+)'
re_sex = r'.*性别：(男|女)'
table=data.sheets()[0]
num = 0
illness=''
for i in range( table.nrows):
    illness = ''
    row_value =table.row_values(i)
    s1 = table.row(i)[0].value
    search_name = re.search(r'.*姓.*名.*：(.+).*', s1)
    search_sex = re.search(r'.*性.*别.*：.*([男|女])', s1)
    search_age = re.search(r'年.*龄.*：(.+?)[岁|月]', s1)
    search_job = re.search(r'职.*业.*：(.+)', s1)
    search_nation = re.search(r'民.*族.*：(.+)族', s1)
    treate_date = re.search(r'入.*院.*[日|时].*[期|间].*：(.+)', s1)
    birth_place = re.search(r'出.*[生|身].*地.*?[：|:]?(.+)', s1)
    if(search_name and search_sex and search_age and search_job and search_nation and treate_date):
        print(i)
        name = search_name.group(1)
        sex = search_sex.group(1)
        age = search_age.group(1)
        job = search_job.group(1)
        nation = search_nation.group(1)
        date_time = treate_date.group(1)
        bir_place = birth_place.group(1)
        print(bir_place)
        for sj in row_value[1:]:
            if sj != '无明确诊断意见' and sj != '':
                illness = illness + " " + sj
        re_treate_date = re.match(r'(\d\d\d\d)[年|-](\d+)[月|-](\d+).*?(\d+)[时|:](\d+)', date_time)
        if re_treate_date:
            treate_date_time = datetime.strptime(re_treate_date.group(1)+'-'+re_treate_date.group(2) + '-'+re_treate_date.group(3)+' '+
                                             re_treate_date.group(4)+':'+re_treate_date.group(5), '%Y-%m-%d %H:%M'

                                             )
        elif (re.match(r'(\d\d\d\d)[年|-](\d+)[月|-](\d+)', date_time)):
            re_treate_date=re.match(r'(\d\d\d\d)[年|-](\d+)[月|-](\d+)', date_time)
            treate_date_time = datetime.strptime(
                re_treate_date.group(1) + '-' + re_treate_date.group(2) + '-' + re_treate_date.group(3) ,'%Y-%m-%d')

        insert_sql = 'insert into patient_info (patientNo,patientName, illness, age, sex, job, nation, treate_date,birth_place) values(%s, %s, %s, %s, %s, %s, %s, %s, %s)'
        cursor.execute(insert_sql,[i+1, name.strip(), illness.strip(), age.strip(), sex.strip(), job.strip(), nation.strip(), treate_date_time,bir_place.strip()])
        conn.commit()




    # if search_name:
    #     print(search_name.group(1))
    #     num += 1
    # if search_sex:
    #     print(search_sex.group(1))
    # if search_age:
    #     print(search_age.group(1))
    # if search_nation:
    #     print(search_nation.group(1))
    # if treate_date:
    #     print(treate_date.group(1))
    #     re_treate_date =re.match(r'(\d\d\d\d)[年|-](\d\d)[月|-](\d\d).*?(\d\d)[时|:](\d\d)',treate_date.group(1) )
    #     for i in range(1, 6):
    #         print(re_treate_date.group(i),end='')

    # print(illness)

# print(num)
#
# def time_format(str):
