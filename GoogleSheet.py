import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import os

class Googlesheet:
    def __init__(self):
        os.chdir("D:\\Programs\\GoogleSheet_Integration")    #your drive path
        self.jsonPath =os.getcwd()+"\\blackhat1901.json"     # path of json File that contains credentials
        self.scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']  
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(self.jsonPath, self.scope)
        client = gspread.authorize(self.creds)
        self.sheet = client.open('AttendanceData')
        self.worksheet = self.sheet.get_worksheet(0)
    
    def getSheetdata(self):
        self.records_data = self.worksheet.get_all_records()
        return  pd.DataFrame.from_dict(self.records_data)
    
    def getDataCount(self):
        self.worksheet.col_count
        self.Row_list = len(self.worksheet.col_values(1))
        self.Col_list = len(self.worksheet.row_values(1))
        self.dic={'Total Rows':self.Row_list,'Toatl Columns': self.Col_list}
        return self.dic
    
    def insertIntoSheet(self, data):
        self.col_Count = (self.worksheet.col_values(1))
        if type(data)==list:
            if self.col_Count==0:
                self.colLen= len(data)
                self.colName=chr(64+self.colLen)+str(1)
                self.worksheet.format('A1:'+self.colName, {'textFormat': {'bold': True}})
        else:
            return "Please provide data of list t#pe"
        self.worksheet.update('A'+str(len(self.col_Count)+1) , data)
        return "Data updated."

    
x=Googlesheet()
y=x.insertIntoSheet([['12/16/2023', 'Sunil Chand', '11:44:35'],['12/16/2023', 'Sushil Sharma', '11:44:35']])
print(y)
#m=x.getSheetdata()
#print(m)
#z=x.getDataCount()
#print(z)


