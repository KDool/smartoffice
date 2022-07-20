from robot.api import logger
from robot.libraries.BuiltIn import BuiltIn
import pandas as pd
import os
from datetime import date,datetime,timedelta

class MyLibrary:
    """Give this library a proper name and document it."""

    def example_python_keyword(self):
        logger.info("This is Python!")

    def readFile(self,path=''):
        # path = self.FileLocation()
        if path == '':
            path = os.getcwd()
        file = pd.read_excel(path,sheet_name='Sheet1',skiprows=[0],usecols="A,B,C")
        file.rename(columns = {'Unnamed: 0' : 'STT','SÁNG':'AM','CHIỀU':'PM'}, inplace = True)
        range_len = len(file)

        for i in range(range_len):
            # print(file.loc[i,'STT'])
            if pd.isnull(file['STT'].iloc[i]):
                file.at[i,'STT'] = file.loc[i-1,'STT']
        file = file.dropna(subset=['AM','PM'],how='all')
        df1 = file[['STT','AM']].dropna(subset=['AM'],how='all')
        df2 = file[['STT','PM']].dropna(subset=['PM'],how='all')
        df1.rename(columns = {'AM':'infor'}, inplace = True)
        df2.rename(columns = {'PM':'infor'}, inplace = True)
        df = pd.concat([df1,df2],axis=0)
        df = df.reset_index(drop=True)
        print(df)
        list_infor = self.extractData(df)
        # print(list_infor)
        return list_infor

    def extractData(self,df: pd.DataFrame):
        num_events = int(len(df)/2)
        print(num_events)
        
        total_list = []

        for i in range(num_events):
            print(i)
            date = ''
            time = ''
            title = ''
            location = ''
            des = ''

            date = self.extract_date(df,i*2)
            time,title =  self.extract_f1(df,i*2)
            location,des = self.extract_f2(df,i*2+1)
            list_info = [date,time,title,location,des]
            if 'BigData'.lower() in des.lower() or 'Nguyễn Tất Hậu'.lower() in des.lower() or 'Nguyễn Thúc Cương'.lower() in des.lower():
                total_list.append(list_info)
            else:
                continue
        return total_list


    def extract_f1(self,df,i):
        text = df['infor'].iloc[i]
        a,time,title = text.split(':',2)
        time = a +':'+ time
        title = title.strip()

        return time,title

    def extract_f2(self,df,i):
        text = df['infor'].iloc[i]
        location,des = text.split('\n',1)
        s1,location = location.split(':',1)
        return location.strip(),des

    def extract_date(self,df,i):
        s = df['STT'].iloc[i]
        res = s[s.find('(')+1:s.find(')')]
        current_year = str(date.today().year)
        day,month = res.split('/',1)
 
        datetime_object = datetime.strptime(month, "%m")
        month_name = datetime_object.strftime("%b")

        res = month_name + ' ' + day + ', ' + current_year
        return res
    
    def FileLocation(self,filename=''):
        # with open(path, 'r') as f:
        #     first_line = f.readline()
        
        # txt= first_line.split("=",1)[1]
        path = "C:\\Users\\admin\\Desktop\\lichtuan"
        txt = path + "\\" + filename
        return txt
    
    def get_schedule_week(self):
        today = date.today()
        date_monday = today - timedelta(days=today.weekday())
        date_friday = date_monday + timedelta(days=4)
        a = str(date_monday)
        b = str(date_friday)
        monday = datetime.strptime(a, "%Y-%m-%d").strftime("%d/%m/%Y")
        friday = datetime.strptime(b, "%Y-%m-%d").strftime("%d/%m/%Y")
        return [monday,friday]

    def date_timestamp(self,string_date):
        b = str(int(datetime.strptime(string_date, "%d/%m/%Y").timestamp()))
        return b

    def create_link_download(self,a='',b=''):
        a = self.date_timestamp(a)
        b = self.date_timestamp(b)
        link = 'https://smartoffice.mobifone.vn/BusinessService/report/lichTuanChuaDuyetReport?fromTime='+a+'000'+'&endTime='+b+'000'+'&userDeptRoleId=9627b6fa-be51-4c38-9580-8a9e355773e9&exportType=processed&deptId=TTCNTT-TCT'
        return link
    
    def get_file_name(self,in_date,current_time=''):
        date_format = in_date
        file = 'lichTuanBanHanh_' + date_format +'_'+ current_time + '.xlsx'
        return file
    
    def get_current_time(self):
        now = datetime.now()
        hour = str(now.hour)
        if int(now.minute) < 10:
            minute = '0' + str(now.minute)
        else:
            minute = str(now.minute)
        if int(now.hour) < 10:
            hour = '0' + str(now.hour)
        else:
            hour = str(now.hour)

        current_time = hour+minute
       
        today = date.today()
        a = str(today)
        current_date = datetime.strptime(a, "%Y-%m-%d").strftime("%d%m%Y")
        print("Current Date: ",current_date)
        print("Current Time: ",current_time)
        return [current_time,current_date]
