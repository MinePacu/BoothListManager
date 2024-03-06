import gspread
import google.auth.transport.requests
from gspread.cell import a1_to_rowcol
from gspread.cell import rowcol_to_a1
import requests
import urllib.parse
import pandas as pd
from datetime import date
from datetime import datetime
import time


spreadsheetId = "1CJ-K_6nBLhgyPbVSKSuq5T9tHAF-RT80WKLDmAI5NeQ"
sheetName = "선입금, 통판, 인포 목록의 사본"
sheetNumber = 1

class LinkCollector():
    """
    부스 목록 시트에서 인포 라벨에 있는 하이퍼링크 및 선입금 라벨에 있는 하이퍼링크들만 수집하기 위한 클래스 
    """
    
    @staticmethod
    def GetBasicData(spreadsheetId: str, sheetName: str):
        """
        해당 시트에서 하이퍼링크가 설정된 셀들만의 정보들을 가져와 딕셔너리를 만들어 반환합니다.
        이 함수에서 반환되는 리스트(딕셔너리)은 이 클래스의 다른 함수들을 사용하는데 있어, 가장 Raw 데이터입니다.

        - 매개 변수
            spreadsheetId : 셀들을 정보를 가져올 스프레드시트의 Id입니다.
            sheetName : 해당 스프레드시트에서 로드할 시트의 이름입니다.
        """
        client_ = gspread.service_account()

        print("LinkCollector.GetBasicData : 셀 전체 데이터를 가져오는 중...")
        #sh1_cell_list = sh.get_worksheet(sheetNumber).get_all_cells()

        # print(sh1_cell_list)

        print("LinkCollector.GetBasicData : 선입금 및 통판 링크만 가져오는 중...")

        request = google.auth.transport.requests.Request()
        client_.http_client.auth.refresh(request)
        access_token = client_.http_client.auth.token

        fields = "sheets(data(rowData(values(formattedValue,hyperlink,textFormatRuns))))"
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheetId}?ranges={urllib.parse.quote(sheetName)}&fields={urllib.parse.quote(fields)}"
        res = requests.get(url, headers={"Authorization": "Bearer " + access_token})


        obj = res.json()
        rows = obj["sheets"][0]['data'][0]['rowData']
        res = []

        # print(rows)

        for i, r in enumerate(rows):
            temp1 = []
            cols = r.get("values", [])
            for j, c in enumerate(cols):
                a1Notation = gspread.utils.rowcol_to_a1(i + 1, j + 1)
                if "hyperlink" in c:
                    temp1.append({"cell": a1Notation, "hyperlinks": [{"text": c.get("formattedValue", ""), "hyperlink": c.get("hyperlink", "")}]})
                elif "textFormatRuns" in c:
                    formattedValue = c.get("formattedValue", "")
                    textFormatRuns = c.get("textFormatRuns", [])
                    temp2 = {"cell": a1Notation, "hyperlinks": []}
                    for k, e in enumerate(textFormatRuns):
                        startIdx = e.get("startIndex", 0)
                        f = e.get("format", {})
                        if "link" in f:
                            temp2["hyperlinks"].append({"text": formattedValue[startIdx:textFormatRuns[k + 1]["startIndex"] if k + 1 < len(textFormatRuns) else len(formattedValue)], "hyperlink": f.get("link", {}).get("uri", "")})
                    if temp2["hyperlinks"] != []:
                        temp1.append(temp2)
            if temp1 != []:
                res.append(temp1)
                
        return res
    
    @staticmethod
    def GetInfoLinks(res: list, convertToPandaDataFrame = False):
        """
        매개 변수 res에 있는 데이트에서 인포가 있는 셀들만의 정보들을 모아 리스트로 만들어 반환합니다.
        
        - 매개 변수
            res : 하이퍼링크가 있는 셀들만 모은 Raw 리스트 데이터
            convertToPandaDataFrame : Panda 라이브러리에서 DataFrame를 만들어 반환할지 여부입니다. 기본값은 False입니다.
        """
    # 인포 링크 수집기
        F_Data_pd = []
        for data in res:
            # 두 번째 또는 그 이상의 인포인 경우
            if len(data) < 2 and "F" in str(data[0]["cell"]):
                F_Data_pd.append({'Cell': data[0]["cell"], 'Info_Text': data[0]["hyperlinks"][0]["text"], 'Info_URL': data[0]["hyperlinks"][0]["hyperlink"]})

            # 일반적인 한 개의 인포를 가진 경우
            elif len(data) >= 2 and "F" in str(data[0]["cell"]):
                F_Data_pd.append({'Cell': data[0]["cell"], 'Info_Text': data[0]["hyperlinks"][0]["text"], 'Info_URL': data[0]["hyperlinks"][0]["hyperlink"]})

            # 인포를 가지지 않으면서 선입금 또는 통판 링크를 가진 경우
            elif len(data) >= 2 and "F" not in str(data[1]["cell"]):
                continue

            # 두 번째 또는 그 이상의 선입금 또는 통판 링크를 가진 경우
            elif len(data) < 2 and "F" not in str(data[0]["cell"]):
                continue
            else:
                F_Data_pd.append({'Cell': data[1]["cell"], 'Info_Text': data[1]["hyperlinks"][0]["text"], 'Info_URL': data[1]["hyperlinks"][0]["hyperlink"]})
            
        if (convertToPandaDataFrame == True):
            F_Data_pdf = pd.DataFrame(F_Data_pd, index=list(range(1, len(F_Data_pd) + 1)))
            return F_Data_pdf
    
        else:
            return F_Data_pd

    def GetPreOrderAndMailOrderLinks(res: list, sh1_cell_list: list, convertToPandaDataFrame = False):
        """
        매개 변수 res에 있는 데이터에서 선입금 또는 통판 링크가 있는 셀들을 모아 리스트를 만들어 반환합니다.
        
        - 매개 변수
            res: 하이퍼링크가 있는 셀들만 모은 Raw 리스트 데이터
            sh1_cell_list : 해당 시트의 모든 셀들의 값들을 담은 데이터, 일반적으로 worksheet.get_all_values 함수에 의해 반환된 리스트입니다.
            convertToPandaDataFrame : Panda 라이브러리에서 DataFrame를 만들어 반환할지 여부입니다. 기본값은 False입니다.
            
        """
        # 선입금 또는 통판 링크 가져오기
        H_Data_pd = []

        for data in res:
            data_temp = {}

            # 일반적인 부스 번호 + 인포 + 선입금 링크
            if len(data) >= 3 and "H" in str(data[2]["cell"]):
                data_temp = {'Cell': data[2]["cell"], 'text': data[2]["hyperlinks"][0]["text"], 'link': data[2]["hyperlinks"][0]["hyperlink"]}

            # 인포 + 선입금 링크
            elif len(data) == 2 and "H" in str(data[1]["cell"]):
                data_temp = {'Cell': data[1]["cell"], 'text': data[1]["hyperlinks"][0]["text"], 'link': data[1]["hyperlinks"][0]["hyperlink"]}

            # 선입금 링크
            elif len(data) == 1 and "H" in str(data[0]["cell"]):
                data_temp = {'Cell': data[0]["cell"], 'text': data[0]["hyperlinks"][0]["text"], 'link': data[0]["hyperlinks"][0]["hyperlink"]}

            else:
                continue

            # H_Data_pd.append(data_temp)

            search_row = a1_to_rowcol(data_temp["Cell"])[0]
            search_column = a1_to_rowcol(data_temp["Cell"])[1]
            matches = [x for x in sh1_cell_list if x.row == search_row and x.col == search_column - 1][0]

            H_Data_pd.append({'Reservation_finish_data_cell': matches.address, 'date': matches.value, 'link_cell': data_temp["Cell"], 'link_text':data_temp["text"], 'link': data_temp["link"]})

        if (convertToPandaDataFrame == True):
            H_Data_pdf = pd.DataFrame(H_Data_pd, index=list(range(1, len(H_Data_pd) + 1)))
            return H_Data_pdf
    
        else:
            return H_Data_pd

    @staticmethod
    def GetPreOrderAndMailOrderDate(res: list, sh1_cell_list: list, convertToPandaDataFrame: bool):
        """
        매개 변수 res에 있는 데이터에서 선입금 또는 통판 링크 및 마감 일자가 있는 셀들을 모아 리스트를 만들어 반환합니다.
        
        - 매개 변수
            res: 하이퍼링크가 있는 셀들만 모은 Raw 리스트 데이터
            sh1_cell_list : 해당 시트의 모든 셀들의 값들을 담은 데이터, 일반적으로 worksheet.get_all_values 함수에 의해 반환된 리스트입니다.
            convertToPandaDataFrame : Panda 라이브러리에서 DataFrame를 만들어 반환할지 여부입니다. 기본값은 False입니다.
            
        """
        H_Data_pd = []
        for data in res:
            data_temp = {}

            # 일반적인 부스 번호 + 인포 + 선입금 링크
            if len(data) >= 3 and "H" in str(data[2]["cell"]):
                data_temp = {'Cell': data[2]["cell"], 'text': data[2]["hyperlinks"][0]["text"], 'link': data[2]["hyperlinks"][0]["hyperlink"]}

            # 인포 + 선입금 링크
            elif len(data) == 2 and "H" in str(data[1]["cell"]):
                data_temp = {'Cell': data[1]["cell"], 'text': data[1]["hyperlinks"][0]["text"], 'link': data[1]["hyperlinks"][0]["hyperlink"]}

            # 선입금 링크
            elif len(data) == 1 and "H" in str(data[0]["cell"]):
                data_temp = {'Cell': data[0]["cell"], 'text': data[0]["hyperlinks"][0]["text"], 'link': data[0]["hyperlinks"][0]["hyperlink"]}

            else:
                continue

    
        # 선입금 또는 통판 마감일 가져오기
        reservation_finish_date_list = []
        for data in H_Data_pd:
            search_row = a1_to_rowcol(data["Cell"])[0]
            search_column = a1_to_rowcol(data["Cell"])[1]
            matches = [x for x in sh1_cell_list if x.row == search_row and x.col == search_column - 1][0]

            reservation_finish_date_list.append({'cell': rowcol_to_a1(matches.row, matches.col) ,'date': matches.value, 'LinkTextCell': rowcol_to_a1(matches.row, matches.col + 2), 'LinkText': data["text"], 'Link': data["link"]})

        if (convertToPandaDataFrame == True):
            reservation_finish_date_list_pd = pd.DataFrame(reservation_finish_date_list, index=list(range(1, len(reservation_finish_date_list) + 1)))
            return reservation_finish_date_list_pd
    
        else:
            return reservation_finish_date_list
    
    @staticmethod
    def CompareTimeTable(pdf_list: list, convertToPandaDataFrame: bool):
        Comparedtimetable = []
        for data in pdf_list:
            #print(data) #for Debugging
            Comparedtimetable.append({'date' : data["date"], 'ComparedResult': str(ConvertDateTimeFromStr(data["date"]) < datetime.today()), 'Link_cell': data['link_cell'], 'LinkText' : data['link_text'], 'link': data['link']})
        
        if (convertToPandaDataFrame == True):
            Comparedtimetable_pdf = pd.DataFrame(Comparedtimetable, index=list(range(1, len(Comparedtimetable) + 1)))
            return Comparedtimetable_pdf

        else:
            return Comparedtimetable
        
def ConvertDateTimeFromStr(string: str):
    if ("?" in string or string == ""):
        return datetime(2024, 12, 31, 23, 59, 59)
    
    else:
        textlist = string.split('/')
        month = textlist[0]

        textlist2 = textlist[1].split('(')
        day = textlist2[0]

        time = "23"
        minute = "59"
        if "시" in str(textlist2[1]):
            textlist3 = textlist[1].split('\n')
            time = textlist3[1].split('시')[0]

            minute = "0"
            if "분" in str(textlist3[1]):
                minute = textlist3[1].split(" ")[1].replace("분", "")

        #print(month, day, time, minute)
        return datetime(2024, int(month), int(day), int(time), int(minute), 0)
