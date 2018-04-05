import pandas as pd
import pypyodbc

wellDict={'Appaloosa #7H': 'N9F182LUKU',
 'Cobra #23H': 'N8VK2P3RM4',
 'Cobra #26H': 'OCCKJI8UNR',
 'Cobra #27H': 'OCCKJIIUOR',
 'District Lonestar #2H': 'O2G0INWS2E',
 'District Lonestar #6H': 'OCA2HB8I6S',
 'Doris #7H': 'OCCLCHCMJM',
 'Horton Tree #7H': 'N8VKI4BUF9',
 'Horton Tree #12H': 'NA8KK9X81A',
 'Horton Tree #13H': 'NA8KLSOB28',
 'Horton Tree #14H': 'NA8KNBPA3H',
 'Horton Tree #15H': 'NA8KO4C24E',
 'Osprey #22H': 'N8VKD78HPB',
 'Overcoming Faith # 3H': 'OBJ2GQTDRB',
 'Overcoming Faith # 4H': 'OBJ2HF06SB',
 'Overcoming Faith # 5H': 'OBJ2H339TB',
 'Overcoming Faith # 6H': 'N8VKCRJJOI',
 'Overcoming Faith # 7H': 'N8VLJ04TMG',
 'Overcoming Faith #13H': 'OBJ2HF9DUB',
 'Palomino #5H': 'OCCI4LJQ9U',
 'Palomino #9H': 'N9F1Q38OLR',
 'Quarterhorse #16H': 'N9EL038KSG',
 'Quarterhorse-DLS #11H': 'OCCKCPSGVR',
 'Quarterhorse-DLS #12H': 'OCCKCPNGWR',
 'Quarterhorse-DLS #13H': 'OCCKCPXHXR',
 'Quarterhorse-DLS #14H': 'OCCKCPRHYR',
 'Ranger #11H': 'N8VK7GNSN1',
 'SE Mansfield # 7H': 'OAEM4TTJ5P',
 'SE Mansfield # 8H': 'OAEM4DEM4P',
 'SE Mansfield # 9H': 'OAEM3PEQ3T',
 'SE Mansfield #10H': 'OAEM3SVJ2X',
 'SE Mansfield #11H': 'OAEM32CJ1X',
 'SE Mansfield #12H': 'N8VJ2J7R1Q',
 'Sowell #12H': 'N8VKEO1IC7',
 'Sowell #13H': 'N9SK1IWR1P',
 'Thoroughbred #2H': 'N8VKI2KGEA',
 'Thoroughbred #4H': 'OCCLK01L9W',
 'Thoroughbred #5H': 'OCCLK0HLAW',
 'Viper #15H': 'OCCKJKVQ3R',
 'Viper #9H': 'N8VK2MHQL4'}

#  load all the sheets in the single spreadsheet
data=pd.read_excel("C:\\Users\\qiufangda\\Desktop\\GHAB\\ProductionReport\\Master\\GHA Barnett Wells_Excel_M.xlsx", sheetname=None,usecols=["CHP","Choke","Comment","DP","D_DATE", "FTP", "Flow_Hours", "Gas", "Injection_Gas", "LP","Net_Gas","Oil","Temperature","Water","Wells"])
# data=pd.read_excel("ProductionReport/Master/GHA Barnett Wells_Excel_M.xlsx", sheetname=None)
# daily update spreadsheet
Dailydata=pd.read_excel('C:\\Users\\qiufangda\\Desktop\\GHAB\\ProductionReport\\Daily\\GHA Barnett Wells_Excel.xlsm', sheetname=None,header=0,usecols=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15])
#Dailydata=pd.read_csv('C:\\Users\\qiufangda\\Desktop\\GHAB\\ProductionReport\\Daily\\GHA Barnett Wells_Excel.csv', sheetname=None,usecols=["CHP","Choke", "Comment", "DP", "D_DATE", "FTP", "Flow_Hours", "Gas","Injection_Gas","LP","Net_Gas","Oil","Temperature","Water","Wells"])
# Dailydata=pd.read_excel('ProductionReport/Daily/GHA Barnett Wells_Excel.xlsx', sheetname=None)

ProdDB = pd.DataFrame()

MasterReportWriter = pd.ExcelWriter('C:\\Users\\qiufangda\\Desktop\\GHAB\\ProductionReport\\Master\\GHA Barnett Wells_Excel_M.xlsx')
# MasterReportWriter = pd.ExcelWriter('ProductionReport/Master/GHA Barnett Wells_Excel_M.xlsx')
# read only the sheet in the wellDict dictionary
for key, item in wellDict.items():
    #print(key,item)
    try:
        if key in sorted(data.keys()):  # data.keys() refers to the sheet name in the spreadsheet#
        #  Update the Master data sheet for each well with the daily update spreadsheet
            # Get individual well data from the master data sheet
            welldata = data[key]
            # Get individual well data from the daily update data sheet
            Dailywelldata = Dailydata[key]
            # combine the two dataframe together and drop the duplicated value on D_DATE column
            Newwelldata = pd.concat([welldata, Dailywelldata]).drop_duplicates('D_DATE')
            # Output the updated master data
            Newwelldata.to_excel(MasterReportWriter, sheet_name=key, index=False)
        # Create the AC_DAILY table
            # create a data frame containing PROPNUM column to be used in AC DAILY sheet as the UID.
            uid = pd.DataFrame({"PROPNUM": item, "NOD": range(len(data[key]))})
            # For the master data , create/merge two dataframe based on the number of rows to create the db for AC_DAILY
            db = pd.merge(uid, welldata, left_index=True, right_index=True)
            # For the daily update data, create/merge two dataframe based on the number of rows
            Dailydb = pd.merge(uid, Dailywelldata, left_index=True, right_index=True)
            # combine the master dataframe and daily-update dataframe and drop the duplicated value
            # based on D_DATE column
            ndb = pd.concat([db, Dailydb]).drop_duplicates('D_DATE')
            # update the NOD row based on the new length of the ndb
            ndb['NOD'] = range(len(ndb))
            # store each well's ndb to ProdDB to prepare to output
            ProdDB = ProdDB.append(ndb)
            # DailyProdDB=DailyProdDB.append(Dailydb)
    except KeyError:
        print("Missing :" + key)
# save the excel sheet and closed the Excel Writer
MasterReportWriter.save()

# Output the Excel sheet
ProdDB.to_excel('C:\\Users\\qiufangda\\Desktop\\GHAB\\ProductionReport\\AC_DAILY.xlsx', sheet_name='AC_DAILY', index=False)
# ProdDB.to_excel('ProductionReport/AC_DAILY.xlsx', sheet_name='AC_DAILY', index=False)

#ProdDB.to_csv('C:\\Users\\qiufangda\\Desktop\\GHAB\\ProductionReport\\AC_DAILY.csv', sep=',')
# ProdDB.to_csv('AC_DAILY.csv', sep=',')
col_list = ProdDB.columns.values.tolist()

DATA_TYPE = {
    None: 'Text',
    'nod': 'Long',
    'd_date': 'DateTime',
    'oil': 'Long',
    'gas':  'Long',
    'water': 'Long',
    'wells':  'Long',
    'ftp': 'Long',
    'chp': 'Long',
    'lp': 'Long',
    'dp': 'Long',
    'flow_hours': 'Long',
    'temperature': 'Long',
    'injection_gas': 'Long',
    'net_gas': 'Long'
}

def get_type(name):
    t = DATA_TYPE.get(name.lower())

    if not t:
        return DATA_TYPE.get(None)

    return t;


sql = u"""CREATE TABLE AC_Daily(""" + ','.join(['{}' for _ in range(len(col_list))]).format(*[str(item) + " {}".format(get_type(str(item))) for item in col_list]) + """);""";
#print (sql)
conn = pypyodbc.win_connect_mdb("C:\\Users\\qiufangda\\Desktop\\GHAB\\GHAOFM\\BEACON0114.mdb")
# conn = pypyodbc.win_connect_mdb(
# r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\\Users\\qiufangda\\Desktop\\GHAB\\GHAOFM\\test.mdb;'
# )

cur = conn.cursor()

try:
    cur.execute(sql)
    conn.commit()
except pypyodbc.ProgrammingError as e:
    if 'already exists' not in str(e):
        raise
    else:
        cur.execute('DROP TABLE AC_Daily;')
        cur.execute(sql)
        conn.commit()

cols = ','.join(['{}' for _ in range(len(col_list))]).format(*[str(item) for item in col_list])

def clean(val):
    val = val.strip()

    if not val:
        return 0

    if val.lower() == 'nan':
        return 0

    return val


for index, row in ProdDB.iterrows():
    values = [str(row[item]) for item in col_list]
    sql = u"""INSERT INTO AC_Daily ({}) VALUES ({});""".format(cols, ','.join(['?' for x in values]))
    #print(sql)
    values = [clean(x) for x in values]
    #print(values)
    cur.execute(sql, values)
conn.commit()
cur.close()
conn.close()
