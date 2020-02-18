##############################
# For EnotManager version 7.x
#
# Leonid Dudnik
# Feb. 18 2020
##############################

import zipfile
from pathlib import Path
import os
import pyodbc

# setup Appinfo
app_version = '7.6.7.4'
app_date    = '18.02.2020'
app_release = """- Оптимизирована перепроводка зарплаты, 45+ минут сокращено до 3х секунд
- Добавлена сверка зарплаты менеджерам позаказно
- Окно клиентов с индивидуальными работами (Контрагенты->Индивидуальные работы клиентов)
- В автоматическую спецификацию добавлены индивидуальные работы клиентов (Упаковка в картон)"""

#setup zip
source_path = r'D:\PROJECTS\GCLOUD\EnotSQL\EnotManager7\bin'
trim_size   = source_path.__len__()
dest_path   = r'\\SQLSERVER\Share\Enot2\Dropbox\update'
exclude = ['app.publish',
'composer',
'de',
'es',
'fr',
'it',
'ja',
'ko',
'pt',
'Release',
'res',
'Dapper.xml', 
'DevExpress.Data.v19.1.xml', 
'DevExpress.DataAccess.v19.1.xml', 
'DevExpress.Map.v19.1.Core.xml', 
'DevExpress.Office.v19.1.Core.xml', 
'DevExpress.Pdf.v19.1.Core.xml', 
'DevExpress.Printing.v19.1.Core.xml', 
'DevExpress.RichEdit.v19.1.Core.xml', 
'DevExpress.Sparkline.v19.1.Core.xml', 
'DevExpress.Utils.v19.1.xml', 
'DevExpress.Xpo.v19.1.xml', 
'DevExpress.XtraBars.v19.1.xml', 
'DevExpress.XtraEditors.v19.1.xml', 
'DevExpress.XtraGrid.v19.1.xml', 
'DevExpress.XtraLayout.v19.1.xml', 
'DevExpress.XtraMap.v19.1.xml', 
'DevExpress.XtraPrinting.v19.1.xml', 
'DevExpress.XtraRichEdit.v19.1.xml', 
'DevExpress.XtraTreeList.v19.1.xml', 
'EnotManager.application', 
'EnotManager.exe.config', 
'EnotManager.exe.manifest', 
'EntityFramework.SqlServer.xml', 
'EntityFramework.xml', 
'EPPlus.xml', 
'itextsharp.xml', 
'Newtonsoft.Json.xml', 
'SqlServiceBrokerListener.xml', 
'System.ValueTuple.xml', 
'conn.bin',
'ConnStringManager.dll',
'ConnStringManager.pdb',
'DbViewStatus.xml',
'EnotUpdater.exe',
'log.xml',
'SlipTemplate.xlsm',
'ttn_izum.xlsx',
'wb2_izum.xlsx',
'wb2_kiev.xlsx',
'wb2_makeevka.xlsx',
'wb2_odessa.xlsx',
'UsageExample.txt']

def createZip():
    print('** CREATE ZIP **')
    with zipfile.ZipFile(dest_path+'\\'+app_version+'.zip', 'w', zipfile.ZIP_DEFLATED) as zipObj:
        for dirpath, dirnames, filenames in os.walk(source_path):
            for filename in filenames:
                path = Path((dirpath[(dirpath.__len__() - trim_size) * -1:])).parts
                if path[1] not in exclude:
                    if filename not in exclude:
                        fullname = os.path.join(dirpath, filename)
                        zipObj.write(fullname, fullname[(fullname.__len__() - trim_size) * -1:])
                        print(fullname)

def updateAppInfo():
    print('** UPDATE APPINFO **')
    conn = pyodbc.connect('DRIVER={SQL Server}; '
                         'SERVER=0.0.0.0; '
                         'DATABASE=OldEnot; '
                         'UID=user;'
                         'PWD=*****')
    cursor = conn.cursor()
    cursor.execute("UPDATE AppInfo SET AppVersion='{0}', LastDate='{1}', Release='{2}'".format(app_version, app_date, app_release))
    cursor.commit()
    conn.close()


if __name__ == '__main__':
    createZip()
    updateAppInfo()
    print('** DONE **')