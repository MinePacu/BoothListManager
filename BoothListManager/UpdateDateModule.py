import gspread
import google.auth.transport.requests
from gspread.cell import a1_to_rowcol
import requests
import urllib.parse
import pandas as pd
from datetime import date
from datetime import datetime
import time
# from IPython.display import display

# ================================================================= Sheet Selecting Section ===============================================================
spreadsheetId = "1gq4mzPObtg7zxCk7FrlZ3IYJYRsBoBZHkDtodp8rMX4"
sheetName = "선입금, 통판, 인포 목록"
sheetNumber = 0

# ================================================================= Reading Sheets and Cells Data =========================================================

client_ = gspread.service_account()

print("셀 전체 데이터를 가져오는 중...")
sh = client_.open_by_key(spreadsheetId)
sh1_cell_list = sh.get_worksheet(sheetNumber).get_all_cells()

# print(sh1_cell_list)

print("선입금 및 통판 링크만 가져오는 중...")

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
        
# print(res)
# 1 = F
# print(res[0])

# ====================================================================== Functions ========================================================================

def GetInfoLinks(res, convertToPandaDataFrame):
# 인포 링크 수집기
    F_Data_pd = []
    for data in res:
        # 두 번째 또는 그 이상의 인포인 경우
        if len(data) < 2 and "F" in str(data[0]["cell"]):
            F_Data_pd.append({'Cell': data[0]["cell"], 'Info_URL': data[0]["hyperlinks"][0]["hyperlink"]})

        # 일반적인 한 개의 인포를 가진 경우
        elif len(data) >= 2 and "F" in str(data[0]["cell"]):
            F_Data_pd.append({'Cell': data[0]["cell"], 'Info_URL': data[0]["hyperlinks"][0]["hyperlink"]})

        # 인포를 가지지 않으면서 선입금 또는 통판 링크를 가진 경우
        elif len(data) >= 2 and "F" not in str(data[1]["cell"]):
            continue

        # 두 번째 또는 그 이상의 선입금 또는 통판 링크를 가진 경우
        elif len(data) < 2 and "F" not in str(data[0]["cell"]):
            continue
        else:
            F_Data_pd.append({'Cell': data[1]["cell"], 'Info_URL': data[1]["hyperlinks"][0]["hyperlink"]})
            
    if (convertToPandaDataFrame == True):
        F_Data_pdf = pd.DataFrame(F_Data_pd, index=list(range(1, len(F_Data_pd) + 1)))
        return F_Data_pdf
    
    else:
        return F_Data_pd

def GetPreOrderAndMailOrderLinks(res, convertToPandaDataFrame):
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

def GetPreOrderAndMailOrderDate(res, convertToPandaDataFrame):
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

        reservation_finish_date_list.append({'cell': rowcol_to_a1(matches.row, matches.col) ,'date': matches.value, 'LinkTextCell': rowcol_to_al(matches.row, matches.col + 2), 'LinkText': data["text"], 'Link': data["link"]})

    if (convertToPandaDataFrame == True):
        reservation_finish_date_list_pd = pd.DataFrame(reservation_finish_date_list, index=list(range(1, len(reservation_finish_date_list) + 1)))
        return reservation_finish_date_list_pd
    
    else:
        return reservation_finish_data_list
    
def ConvertDateTimeFromStr(string):
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
    

def CompareTimeTable(pdf_list, convertToPandaDataFrame):
    Comparedtimetable = []
    for data in pdf_list:
        #print(data)
        Comparedtimetable.append({'date' : data["date"], 'ComparedResult': str(ConvertDateTimeFromStr(data["date"]) < datetime.today()), 'Link_cell': data['link_cell'], 'LinkText' : data['link_text'], 'link': data['link']})
        
    if (convertToPandaDataFrame == True):
        Comparedtimetable_pdf = pd.DataFrame(Comparedtimetable, index=list(range(1, len(Comparedtimetable) + 1)))
        return Comparedtimetable_pdf

    else:
        return Comparedtimetable
    
def updatedate(timetable):
    print("만료된 링크 정리 중...")
    print("", end="")
    i = 0
    for data in timetable:
        i = i + 1
        if (data['ComparedResult'] == str(True) and "(만료)" not in str(data['LinkText'])):
            link = data['link']
            celltext = data['LinkText'] + " (만료)"
            print("\r셀 " + data['Link_cell'] + " 업데이트 중 / 진행도 : " + str(round((i / len(timetable)) * 100, 2)) + "%", end="")
            sh.get_worksheet(sheetNumber).update_acell(data['Link_cell'], f'=HYPERLINK("{link}", "{celltext}")')
            time.sleep(1.5)
        else:
            continue
    print("\n업데이트 완료.")
    
# ======================================================== After Reading Data, Excuting Code =========================================================
    
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)

#display(GetPreOrderAndMailOrderLinks(res,True))
#display(CompareTimeTable(GetPreOrderAndMailOrderLinks(res, False), True))
updatedate(CompareTimeTable(GetPreOrderAndMailOrderLinks(res, False), False))


