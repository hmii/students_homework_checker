from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
from IPython.display import display
import time
import requests
import re
import glob
import os
import re
import numpy as np

class Homework:
     
    def __init__(self):
        self.member_dict = {}
        
    
    def setting(self, class_dic, driver, month, week):
        self.class_dic = class_dic
        self.driver = driver
        self.month = month
        self.week = week 
        
    def get_member_info(self, class_id):
        
        url = self.class_dic[class_id]+'/member'
        self.driver.get(url)
        time.sleep(1.5)
        
        name_list = []
        names = self.driver.find_elements_by_css_selector("span.ellipsis")

        for name in names[2:]:
            name_list.append(name.text)

        self.member_dict[class_id] = pd.Series(sorted(name_list))

    def export_member_info(self):
        df = pd.DataFrame(self.member_dict) # 학생명단 
        df.to_excel(f"{self.week}주차_member.xlsx", index=False)
        
# ------------------------------------            
    def url(self, tag):
        self.tag = tag
        url = self.class_dic[self.class_id] + f'/hashtag/{tag}'
        self.driver.get(url)
        time.sleep(1)  
        
    def click_comment(self):
        try: 
            self.driver.find_element_by_class_name('commentToggle._commentToggleBtn').click()
            time.sleep(0.5)

            for _ in range(3):
                try:
                    self.driver.find_element_by_class_name('prevComment').click()
                    time.sleep(0.2)
                except :
                    pass
        except :
            pass
        
# -------------------------------
        
    def page_source(self):
        html = self.driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        students = soup.select('.itemWrap')
        return students
    
    def get_comment(self):
        try:
            print(self.class_id, self.tag, "댓글 수: ", 
                  self.driver.find_element_by_css_selector("span.count").text)
            student_data = {}
            students = self.page_source() # page_source 함수 실행 
            print("--------------------", len(students))
            for student in students :
                name = student.select('button > strong')[0].text
                try :
                    score = int(student.select('p')[0].text)
                except :
                    score = "O"
                student_data[name] = score 
            self.df[self.tag] = student_data
        except:
            pass
        
        
# ---------------------------------------       
    def get_test(self): 
        band_name = self.driver.find_element_by_class_name('uriText').text
        try : 
            # 퀴즈박스 클릭
            self.driver.find_element_by_class_name('item.-quiz').click() 

            scroll_area = self.driver.find_element_by_xpath('//*[@id="wrap"]/div[2]/div/div/section')
            action = webdriver.ActionChains(self.driver)
            action.move_to_element(scroll_area).perform()
            self.driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scroll_area)
            time.sleep(1)

            # 퀴즈명 가져오기 
            quiz_name = self.driver.find_element_by_class_name('addTitle')
            quiz_name = quiz_name.text.replace("~", "_")

            self.driver.find_element_by_class_name('button._btnOpenTakerList').click() # 응답자 보기 클릭
            time.sleep(0.5)
            self.driver.find_element_by_class_name('btnMoreMenu._btnHeaderMoreMenu').click() # 점 3개 클릭
            time.sleep(0.5)
            self.driver.find_element_by_class_name('btnMenuText._btnDownloadGradingResult').click() # 채점결과 다운로드 

            time.sleep(2)

            # 저장된 파일 불러오기 
            add = f'C:/Users/hmkang/Downloads/{band_name}_채점결과_{quiz_name}_' + r'{}_{}.xlsx'
            filename = glob.glob(add.format("*", "*"))[0]
            test_df = pd.read_excel(filename, header=3) # 3번째 줄부터 불러와라 

            # 멤버 및 점수만 뽑아내기 
            # test_df['멤버'] = test_df['멤버'].str.extract(r'([가-힣].*)')
            test_df = test_df.dropna()
            test_df = test_df[['멤버', '점수']]
            test_df = test_df.rename(columns = {'멤버':'name', '점수':self.tag})
 
            return test_df 
        except :
            pass
        
        
# --------------------------------------
    def pds_tag(self, num):
        return [f'p-{self.month}-{self.week}-{i}' for i in range(1, num+1)]

    def voca_tag(self, num):
        return [f'v-{self.month}-{self.week}-{i}' for i in range(1, num+1)]

    def special_tag(self, num):
        return [f's-{self.month}-{self.week}-{i}' for i in range(1, num+1)]
    
    def test_tag(self): # 
        return [f't-{self.month}-{self.week}']
    
    def pareto_tag(self):
        return [f'1500-{self.month}-{self.week}']
    
# --------------------------------------------------    
    def comment_scrap(self, tags): 
        self.df = {}
        for tag in tags:
            self.url(tag)
            self.click_comment()
            self.get_comment()
        return pd.DataFrame(self.df).reset_index().rename(columns={'index':"name"})  


    def excel_scrap(self, test_tag): 
        for tag in test_tag:
            self.url(tag)
            return self.get_test()
        
# --------------------------------------------------    
    def export_excel(self, class_id, tags, test_tag=0):
        self.class_id = class_id
        
        comment_df = self.comment_scrap(tags)
        
        self.member_dict[self.class_id].name = 'name' # 프레임으로 변경시 컬럼 이름이 됨 
        name_df = self.member_dict[self.class_id].to_frame() # 시리즈 형태로 저장된 밸류를 프레임으로 변환 
        name_df = name_df.dropna()
        
        df = pd.merge(name_df, comment_df, on='name', how='outer')
    
        try : 
            test_df = self.excel_scrap(test_tag) # 테스트가 있다면 ... 
            df = pd.merge(df, test_df, on='name', how='outer')
        except:
            pass 
        
        df = df.fillna("X")
        df.to_excel(f'{self.class_id}_{self.week}주차.xlsx', index=False)
        
        
        
        
class Export:
    
    def __init__(self, class_id, month, week, pre=0):
        self.class_id = class_id
        self.month = month
        self.week = week 
        self.pre = pre

    def week_df(self):
        df = pd.read_excel(f'./{self.week}주차/{self.class_id}_{self.week}주차.xlsx')
        try :
            week_pre = pd.read_excel(f'./{self.pre}주차/{self.pre}주차_comment.xlsx', 
                                 sheet_name=f'{self.class_id}').drop(columns=['life'])
            df = pd.merge(week_pre, df, on='name', how='outer')
        except :
            pass
        return df
        
    def heart(self):
        df = self.week_df()
        df['life'] = df.isin(["X"]).sum(axis=1)
        for j in range(self.month-1, self.month+1):
            for i in range(1, 6):
                try :
                    df['life'] = np.where(df[f't-{j}-{i}'] == 15, df['life'] -1, df['life'])
                    df['life'] = np.where(df[f't-{j}-{i}'] == 16, df['life'] -1, df['life'])
                    df['life'] = np.where(df[f't-{j}-{i}'] == 17, df['life'] -1, df['life'])
                except :
                    pass
        df['life'] = df['life'].apply(lambda x : (7-x) * "🧡")
        df = df.sort_values(by=['name'])
        return df    

