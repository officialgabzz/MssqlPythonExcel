import xlrd


import pyodbc
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=ABDULMALIK\MSSQL14;'
                      'Database=EMS_ZONE_Patch;'
                      'Trusted_Connection=yes;'
                      'UID=username;'
                      'PWD=dbpassword;'
                      )
cursor = conn.cursor()

loc = ("MEMMCOL AMI Integrated OBIS_list Template-V1.0(2020.06.19) (1).xlsx")
try:
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(2)

    for i in range(4, sheet.nrows):
        print(sheet.cell_value(i, 2), sheet.cell_value(i, 3))
        query = "insert into EventList(MeterModel,EventID,Description) values(\'MMX-313-CT\', {}, \'{}\');".format(
            int(sheet.cell_value(i, 2)), sheet.cell_value(i, 3))
        cursor.execute(query)
except:
    print("Some went wrong")

finally:
    conn.commit()
    cursor.close()
