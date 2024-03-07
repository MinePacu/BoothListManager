from datetime import datetime
from locale import LC_ALL 
import tkinter as tk
from tkinter import Entry, ttk
import re
import clipboard
from tkinter import messagebox
from typing import Collection
import gspread
from gspread.cell import rowcol_to_a1
from gspread.cell import a1_to_rowcol
from gspread.utils import MergeType, ValueInputOption
from string import ascii_uppercase

from gspread_formatting.models import Borders
import gspread_formatting

from LinkCollector import LinkCollector
from UpdateLogger import UpdateLogger
from UpdateLogger import LogType


def find_duplicating_Indexes(_List, searchWord: str):
    iterated_index_position_list = [
        i for i in range(len(_List)) if _List[i] == searchWord
        ]
    
    return iterated_index_position_list
       
def GetRecommandLocation(booth_list_tmp: list[str], searchBoothNum: str):
    """
    매개 변수인 부스 번호가 어느 셀에 들어가야할지를 계산하여 반환합니다.
    이미 있는 부스 번호인 경우, 해당 부스 번호가 있는 셀의 위치의 아래 행 위치를 반환합니다.
    단, 해당 알파벳 열의 부스가 하나도 등록되지 않은 경우, 0을 반환합니다. (이건 생각하기 힘들어서... 어차피 나 혼자 쓰는데 뭐..)
    
    부스 번호는 [알파벳 + 숫자] 형식을 가집니다. (예 : W10)
    
    - 매개 변수
        :param booth_list_tmp : 부스 번호들의 리스트, 이 리스트에 빈 요소는 없어야 합니다.
        :param searchBoothNum : 추가할 부스 번호
    """
    print("GetRecommandLocation : 부스를 추가할 셀 위치 계산 중...")
    # Recommanded Cell's Location
    conclusionLocation = ""
    IsFind = False
    IsAlreadyAdded = False
    global IsAlredyExisted
    IsAlredyExisted = False
    AlreadyExistedLocation = ""

    if ',' in searchBoothNum:
        searchBoothNum = searchBoothNum.split(", ")[0]
    
    for k in range(0, len(booth_list_tmp)):
        if (booth_list_tmp[k].count(searchBoothNum) >= 1):
            # 중복인 경우, 이미 등록된 부스가 있는 열 + 1의 값을 반환하여 바로 추가할 수 있도록 함
            IsAlreadyAdded = True
            AlreadyExistedLocation = str(k + 1 + 1) # 열 번호는 1부터 시작 + 다음 위치로 계산해서 +1 한번 더

    if IsAlreadyAdded == False:
        userSector = re.findall('[a-zA-Z]', searchBoothNum)
        userSectorNum = re.findall('\d', searchBoothNum)
    
        userSectorNum_tmp = ""
        for k in userSectorNum:
            userSectorNum_tmp += k  

        #print(userSector)
        #print(userSectorNum_tmp)

        # 해당 부스의 이전 값부터 탐색
        for i in range(1, int(userSectorNum_tmp) - 1):
            try:
                # 추가하려는 부스의 이전 번호의 부스를 검색 (이는 2칸 전, 3칸 전일 수 있음)
                autoSearchBoothNum = userSector[0] + str(int(userSectorNum_tmp) - i)
                print("GetRecommandLocation : 검색 중인 부스 번호 : " + autoSearchBoothNum)
                for j in range(0, len(booth_list_tmp)):
                    try:
                        Index_ = booth_list_tmp[j].index(autoSearchBoothNum)
                        Index = j
                        break;
                    except:
                        continue;
                    
                # 검색 중인 부스가 여러 개의 행을 병합한 경우
                if (booth_list_tmp.count(booth_list_tmp[Index]) > 1):
                    iterated_Indexes = find_duplicating_Indexes(booth_list_tmp, booth_list_tmp[Index])
                    conclusionLocation = str(booth_list_tmp[iterated_Indexes[len(iterated_Indexes) - 1]] + 1)
                    IsFind = True
            
                # 검색 중인 부스가 한 개의 행을 보유한 경우
                elif (booth_list_tmp.count(booth_list_tmp[Index]) == 1):
                    conclusionLocation = str(Index + 1) # 0으로 시작해서 + 1
                    IsFind = True
                break;
            except:
                continue;
        
        # 찾지 못한 경우, 해당 부스의 다음 값으로 탐색
        if IsFind == False:
            # 예) W10 이라 하면 1 ~ (25(W열의 최대값) - 10 - 1 = 14) 으로 W11부터 W25까지 탐색 
            for i in range(1, alphabet_max_count[alphabet_list.index(userSector[0])] - int(userSectorNum_tmp) - 1):
                try:
                    autoSearchBoothNum = userSector[0] + str(int(userSectorNum_tmp) + i)
                    print("GetRecommandLocation : 검색 중인 부스 번호 : " + autoSearchBoothNum)
                    for j in range(0, len(booth_list_tmp)):
                        try:
                            Index_ = booth_list_tmp[j].index(autoSearchBoothNum)
                            Index = j
                            break;
                        except:
                            continue;
                
                    # 검색 중인 부스가 여러 개의 행을 병합한 경우
                    if (booth_list_tmp.count(booth_list_tmp[Index]) > 1):
                        iterated_Indexes = find_duplicating_Indexes(booth_list_tmp, booth_list_tmp[Index])
                        conclusionLocation = str(booth_list_tmp[iterated_Indexes[len(iterated_Indexes) - 1]] + 1)
                        IsFind = True
            
                    # 검색 중인 부스가 한 개의 행을 보유한 경우
                    elif (booth_list_tmp.count(booth_list_tmp[Index]) == 1):
                        conclusionLocation = str(Index + 1) # 0으로 시작해서 + 1에 새로 한 행을 만들어야하므로 + 1 한 번 더
                        IsFind = True
                    break;
                except:
                    continue;
                
        # 그냥 없으면 수동으로 하자.
        if IsFind == False:
            return 0

        print("GetRecommandLocation : 계산된 열 위치 : " + conclusionLocation)
        return conclusionLocation
    
    else:
        print("GetRecommandLocation : 이미 있는 부스입니다. 이 부스가 있는 다음 열로 배정합니다. 계산된 열 위치 : " + AlreadyExistedLocation)
        IsAlredyExisted = True
        return AlreadyExistedLocation
    
def Add_new_BoothData(BoothNumber: str, BoothName: str, Genre: str, Yoil: str, InfoLabel: str, InfoLink: str, Pre_Order_Date: str, Pre_Order_label: str, Pre_Order_Link: str):
    """
    새 부스 데이터를 시트에 추가합니다.
    부스 번호가 있으면, 시트 내에 존재하는 부스 번호와 비교하여 적절한 위치에 새 데이터를 추가합니다.
    부스 번호가 없는 경우나 해당 알파벳 열의 부스가 하나도 없는 경우, 위치를 계산하지 않은 채로, 가장 아래에 추가됩니다.

    - 매개 변수
        :param BoothNumber: 추가할 부스 번호
        :param BoothName: 추가할 부스 이름
        :param Genre: 추가할 부스가 취급하는 장르
        :param Yoil: 추가할 부스의 참가 요일
        :param InfoLabel: 추가할 부스의 인포 링크의 라벨
        :param InfoLink: 추가할 부스의 인포 링크
        :param Pre_Order_Date: 추가할 부스의 선입금 또는 통판의 마감 일자
        :param Pre_Order_label: 추가할 부스의 선입금 또는 통판 링크의 라벨
        :param Pre_Order_Link: 추가할 부스의 선입금 또는 통판 링크
    """
    
    # 부스 장르 함수 생성 (다중 줄 포함)
    NewBoothGenre = f''
    if '//' in Genre:
        NewBoothGenre = f'=TEXTJOIN(CHAR(10), 0, '
        SplitedGenre = re.split('//', Genre)
        i = 0;
        for OnelineGenre in SplitedGenre:
            NewBoothGenre += f'"{OnelineGenre}'
            if i != len(SplitedGenre) - 1:
                NewBoothGenre += f'", '
            else:
                NewBoothGenre += f'")'
            i = i + 1
    else:
        NewBoothGenre = Genre

    # 부스 선입금 마감 일자 함수 생성 (다중 줄 포함)
    NewPreOrderDate = f''
    if '//' in Pre_Order_Date:
        NewPreOrderDate = f'=TEXTJOIN(CHAR(10), 0, '
        SplitedPreOrderDate = re.split('//', Pre_Order_Date)
        i = 0
        for OnelinePreOrderDate in SplitedPreOrderDate:
            NewPreOrderDate += f'"{OnelinePreOrderDate}'
            if i != len(SplitedPreOrderDate) - 1:
                NewPreOrderDate += f'", '
            else:
                NewPreOrderDate += f'")'
            i = i + 1
        #print(NewPreOrderDate)
    else:
        NewPreOrderDate = Pre_Order_Date

    # 부스 인포 링크 함수 생성
    NewInfoLink = f''
    if (InfoLink != ''):
        NewInfoLink = f'=HYPERLINK("{InfoLink}", "{InfoLabel}")'

    # 부스 선입금 링크 함수 생성
    NewPreOrderLink = f''
    if (Pre_Order_Link != ''):
        NewPreOrderLink = f'=HYPERLINK("{Pre_Order_Link}", "{Pre_Order_label}")'

    #print(BoothNumber, BoothName, Genre, Yoil, InfoLabel, InfoLink, Pre_Order_label, Pre_Order_Link)

    print("Add_new_BoothData : 셀 전체 데이터를 가져오는 중...")
    sh = client_.open_by_key(spreadsheetId)
    sheet = sh.get_worksheet(sheetNumber)
    updatesheet = sh.get_worksheet(UpdateLogSheetNumber)
    booth_list = sh.get_worksheet(sheetNumber).get(f"{BoothNumber_Col_Alphabet}1:{BoothNumber_Col_Alphabet}")

    fmt = gspread_formatting.CellFormat(
        borders=Borders(
            top=gspread_formatting.Border("SOLID"), 
            bottom=gspread_formatting.Border("SOLID"), 
            left=gspread_formatting.Border("SOLID"), 
            right=gspread_formatting.Border("SOLID")
            ),
        horizontalAlignment='CENTER',
        verticalAlignment='MIDDLE'
        )
    
    if (BoothNumber != ''):
        booth_list_tmp = booth_list.copy()
        #print(booth_list_tmp[13])
        j = 0
        for boothnum in booth_list_tmp:
            if len(boothnum) == 0:
                booth_list_tmp[j] = booth_list_tmp[j - 1]
            if ',' in str(booth_list_tmp[j][0]):
                booth_list_tmp[j] = booth_list_tmp[j][0].split(", ")
            j = j + 1

        RecommandLocation = int(GetRecommandLocation(booth_list_tmp, BoothNumber))
        
        if int(RecommandLocation) == 0:
            NewRowData = [BoothNumber, BoothName, NewBoothGenre, Yoil, NewInfoLink, NewPreOrderDate, NewPreOrderLink, '']            
            sheet.append_row(NewRowData, value_input_option=ValueInputOption.user_entered)
            gspread_formatting.format_cell_range(sheet, f"{BoothNumber_Col_Alphabet}{len(booth_list) + 1}:{Etc_Point_Col_Alphabet}{len(booth_list) + 1}", fmt)
            
            updatetime = UpdateLastestTime()
            UpdateLogger.AddUpdateLog(updatesheet, LogType.Pre_Order, updatetime, sheet.id, f'{Pre_Order_link_Col_Alphabet}{len(booth_list) + 1}', BoothNumber)
            
            if MapSheetNumber != None:
                SetLinkToMap(BoothNumber)

        else: 
            NewRowData = ['', BoothNumber, BoothName, NewBoothGenre, Yoil, NewInfoLink, NewPreOrderDate, NewPreOrderLink, '']
            sheet.insert_row(NewRowData, RecommandLocation, value_input_option=ValueInputOption.user_entered)
            gspread_formatting.format_cell_range(sheet, f"{BoothNumber_Col_Alphabet}{RecommandLocation}:{Etc_Point_Col_Alphabet}{RecommandLocation}", fmt)
 
            updatetime = UpdateLastestTime()
            UpdateLogger.AddUpdateLog(updatesheet, LogType.Pre_Order, updatetime, sheet.id, f'{Pre_Order_link_Col_Alphabet}{RecommandLocation}', BoothNumber)
            
            global IsAlredyExisted
            if IsAlredyExisted == True:
                k = 0
                for k in range(RecommandLocation - 1 - 1, -1, -1):
                    if len(booth_list[k]) != 0:
                        break;

                sheet.merge_cells(f"{BoothNumber_Col_Alphabet}{k + 1}:{Yoil_Col_Alphabet}{RecommandLocation}", MergeType.merge_columns)
            
            if MapSheetNumber != None:
                SetLinkToMap(BoothNumber)
    else:
       NewRowData = [BoothNumber, BoothName, NewBoothGenre, Yoil, NewInfoLink, NewPreOrderDate, NewPreOrderLink, '']    
       sheet.append_row(NewRowData, value_input_option=ValueInputOption.user_entered)
       gspread_formatting.format_cell_range(sheet, f"{BoothNumber_Col_Alphabet}{len(booth_list)}:{Etc_Point_Col_Alphabet}{len(booth_list)}", fmt)
    
       updatetime = UpdateLastestTime()
       UpdateLogger.AddUpdateLog(updatesheet, LogType.Pre_Order, updatetime, sheet.id, f'{Pre_Order_link_Col_Alphabet}{len(booth_list) + 1}', None, BoothName)
            
    print("Add_new_BoothData : 부스 추가 완료")
       
def Remove_Row(BoothNumber: str = None, BoothName: str = None, Genre: str = None, Yoil: str = None, InfoLabel: str = None, Pre_Order_Date: str = None, Pre_Order_label: str = None, Cell_row: int = None):
    """
    매개 변수에 있는 요소들 중에 하나로 검색된 열을 삭제합니다.
    셀의 열 번호를 넣으면 검색하지 않고 바로 삭제합니다. 이 경우, 매개 변수를 'Cell_row = ' 식으로 지정하세요.
    검색되지 않으면 -1을 반환합니다.
    
    매개 변수에서 가장 앞에 있는 변수 순으로 검색하며, 하나의 매개 변수라도 사용하여 검색이 되면, 그 이후 매개 변수는 검색에 사용되지 않습니다.
    
    - 매개 변수
        (이 함수의 모든 매개 변수의 기본값은 None 입니다.)
        :param BoothNumber : 검색할 부스 번호
        :param BoothName : 검색할 부스 이름
        :param Genre : 검색할 부스가 취급하는 장르
        :param Yoil : 검색할 부스의 참가 요일 
        :param InfoLabel : 검색할 부스의 인포 라벨
        :param Pre_Order_Date : 검색할 부스의 선입금 또는 통판 마감 일자
        :param Pre_Order_label : 검색할 부스의 선입금 또는 통판 링크의 라벨
        :param Cell_row = 삭제할 부스의 열 번호
    """
    sheet = client_.open_by_key(spreadsheetId)
    BoothListSheet = sheet.get_worksheet(sheetNumber)
    
    if BoothNumber != None:
        searchedCell = BoothListSheet.find(BoothNumber)
        BoothListSheet.delete_rows(searchedCell.row)
        
    elif BoothName != None:
        searchedCell = BoothListSheet.find(BoothName)
        BoothListSheet.delete_rows(searchedCell.row)
        
    elif Genre != None:
        searchedCell = BoothListSheet.find(Genre)
        BoothListSheet.delete_rows(searchedCell.row)
        
    elif Yoil != None:
        searchedCell = BoothListSheet.find(Yoil)
        BoothListSheet.delete_rows(searchedCell.row)
        
    elif InfoLabel != None:
        searchedCell = BoothListSheet.find(InfoLabel)
        BoothListSheet.delete_rows(searchedCell.row)
        
    elif Pre_Order_Date != None:
        searchedCell = BoothListSheet.find(Pre_Order_Date)
        BoothListSheet.delete_rows(searchedCell.row)
        
    elif Pre_Order_label != None:
        searchedCell = BoothListSheet.find(Pre_Order_label)
        BoothListSheet.delete_rows(searchedCell.row)
        
    elif Cell_row != None:
        BoothListSheet.delete_rows(Cell_row)

    else:
        return -1
    
    return 0

def Modify_Existed_Row(BoothNum_Cell_Row: int, BoothNumber: str, BoothName: str, Genre: str, Yoil: str, InfoLabel: str, InfoLink: str, Pre_Order_Date: str, Pre_Order_label: str, Pre_Order_Link: str):
    """
    정수 BoothNum_Cell_Row 열의 데이터를 매개 변수에 있는 데이터로 수정합니다.
    부스 번호가 수정되면 원래 있던 열을 삭제하고, 열 위치를 다시 계산하여 등록하며 부스 번호가 그대로인 경우, 같은 열에서 내용만 수정됩니다.
    
    - 매개 변수
        :param BoothNum_Cell_Row : 수정 전의 부스 번호가 있는 열 번호
        :param BoothNumber : 수정 후의 부스 번호
        :param Genre : 수정 후의 장르
        :param Yoil : 수정 후의 참가 요일
        :param InfoLabel : 수정 후의 인포 링크의 라벨
        :param InfoLink : 수정 후의 인포 링크
        :param Pre_Order_Data : 수정 후의 선입금 또는 통판 마감 일자
        :param Pre_Order_Label : 수정 후의 선입금 또는 통판 링크의 라벨
        :param Pre_Order_Link : 수정 후의 선입금 도는 통판 링크
    """
    sheet = client_.open_by_key(spreadsheetId)
    BoothListSheet = sheet.get_worksheet(sheetNumber)
    
    if BoothListSheet.get(rowcol_to_a1(BoothNum_Cell_Row, 2)) != BoothNumber:
        Remove_Row(Cell_row = BoothNum_Cell_Row)

    update_cell = []
    update_cell.append(gspread.Cell(BoothNum_Cell_Row, BoothNumber_Col_Number, ''))
    
    NewBoothGenre = f''
    if '//' in Genre:
        NewBoothGenre = f'=TEXTJOIN(CHAR(10), 0, '
        SplitedGenre = re.split('//', Genre)
        i = 0;
        for OnelineGenre in SplitedGenre:
            NewBoothGenre += f'"{OnelineGenre}'
            if i != len(SplitedGenre) - 1:
                NewBoothGenre += f'", '
            else:
                NewBoothGenre += f'")'
            i = i + 1
    else:
        NewBoothGenre = Genre
    
    # 부스 선입금 마감 일자 함수 생성 (다중 줄 포함)
    NewPreOrderDate = f''
    if '//' in Pre_Order_Date:
        NewPreOrderDate = f'=TEXTJOIN(CHAR(10), 0, '
        SplitedPreOrderDate = re.split('//', Pre_Order_Date)
        i = 0
        for OnelinePreOrderDate in SplitedPreOrderDate:
            NewPreOrderDate += f'"{OnelinePreOrderDate}'
            if i != len(SplitedPreOrderDate) - 1:
                NewPreOrderDate += f'", '
            else:
                NewPreOrderDate += f'")'
            i = i + 1
        #print(NewPreOrderDate)
    else:
        NewPreOrderDate = Pre_Order_Date

    # 부스 인포 링크 함수 생성
    NewInfoLink = f''
    if (InfoLink != ''):
        NewInfoLink = f'=HYPERLINK("{InfoLink}", "{InfoLabel}")'

    # 부스 선입금 링크 함수 생성
    NewPreOrderLink = f''
    if (Pre_Order_Link != ''):
        NewPreOrderLink = f'=HYPERLINK("{Pre_Order_Link}", "{Pre_Order_label}")'
        
    # 새로운 셀 생성
    BoothNumberCell = gspread.Cell(BoothNum_Cell_Row, BoothNumber_Col_Number, BoothNumber)
    BoothNameCell = gspread.Cell(BoothNum_Cell_Row, BoothName_Col_Number, BoothName)
    GenreCell = gspread.Cell(BoothNum_Cell_Row, Genre_Col_Number, NewBoothGenre)
    YoilCell = gspread.Cell(BoothNum_Cell_Row, Yoil_Col_Number, Yoil)
    InfoCell = gspread.Cell(BoothNum_Cell_Row, InfoLink_Col_Number, NewInfoLink)
    Pre_Order_Date_Cell = gspread.Cell(BoothNum_Cell_Row, Pre_Order_Date_Col_Number, NewPreOrderDate)
    Pre_order_Link_Cell = gspread.Cell(BoothNum_Cell_Row, Pre_Order_link_Col_Number, NewPreOrderLink)
    
    # 새로운 열에 추가
    update_cell.append(BoothNumberCell)
    update_cell.append(BoothNameCell)
    update_cell.append(GenreCell)
    update_cell.append(YoilCell)
    update_cell.append(InfoCell)
    update_cell.append(Pre_Order_Date_Cell)
    update_cell.append(Pre_order_Link_Cell)
    
    BoothListSheet.update_cells(update_cell, value_input_option=ValueInputOption.user_entered)
    return
    
def EditInfoCell(infoCell: str, InfoLabel: str, InfoLink: str, PreOrderLinkCell_Count: int = 1, mode: int = 0):
    """
    특성 셀의 인포를 수정합니다.
    
    - 매개 변수
        :param infoCell: 수정하려는 인포가 있는 셀 (a1)
        :param InfoLabel: 수정하려는 인포 라벨
        :param InfoLink: 수정하려는 인포 링크
        :param PreOrderLinkCell_Count (선택): 수정하려는 인포가 있는 열에 있는 선입금 링크의 수입니다. 기본값은 1입니다. 
    """
    sheet = client_.open_by_key(spreadsheetId)
    BoothListSheet = sheet.get_worksheet(sheetNumber)
    
    rowcol = a1_to_rowcol(infoCell)
    BoothListSheet.update_cell(rowcol[0], rowcol[1], f'=HYPERLINK("{InfoLink}", "{InfoLabel}")')
    
    if PreOrderLinkCell_Count != 1:
        BoothListSheet.merge_cells(f'{rowcol_to_a1(rowcol[0], rowcol[1])}:{rowcol_to_a1(rowcol[0] + (PreOrderLinkCell_Count - 1), rowcol[1])}', MergeType.merge_columns)
    
def EditPreOrderCell(PreOrderCell: str, PreOrder_Date: str, PreOrder_Label: str, PreOrder_Link: str, mode: int = 1):
    """
    특정 셀의 선입금 마감 일자 및 선입금 링크를 수정합니다. 이 함수는 부스가 이미 있는 상태에서 새로 선입금 부스를 추가하거나 수정할 때 사용됩니다.
    
    - 매개 변수
        :param PreOrderCell: 선입금 링크가 있는 셀의 a1Notation 값
        :param PreOrder_Date: 수정된 선입금 마감 일자
        :param PreOrder_Label: 수정된 선입금 링크의 라벨
        :param PreOrder_Link: 수정된 선입금 링크
        :param mode (선택): 이 값이 1이면 PreOrderCell 자리에 선입금 링크를 업데이트하며, 0이면 PreOrderCell 셀의 열의 다음 열에 새 선입금 링크를 추가합니다. 기본값은 1입니다.
    """
    sheet = client_.open_by_key(spreadsheetId)
    BoothListSheet = sheet.get_worksheet(sheetNumber)
    
    PreOrderCell_rowcol = a1_to_rowcol(PreOrderCell)
    
    # 부스 선입금 마감 일자 함수 생성 (다중 줄 포함)
    NewPreOrderDate = f''
    if '//' in PreOrder_Date:
        NewPreOrderDate = f'=TEXTJOIN(CHAR(10), 0, '
        SplitedPreOrderDate = re.split('//', PreOrder_Date)
        i = 0
        for OnelinePreOrderDate in SplitedPreOrderDate:
            NewPreOrderDate += f'"{OnelinePreOrderDate}'
            if i != len(SplitedPreOrderDate) - 1:
                NewPreOrderDate += f'", '
            else:
                NewPreOrderDate += f'")'
            i = i + 1
        #print(NewPreOrderDate)
    else:
        NewPreOrderDate = PreOrder_Date
        
    if mode == 0:
        BoothListSheet.insert_row(['', '', '', '', '', '', NewPreOrderDate, f'=HYPERLINK("{PreOrder_Link}", "{PreOrder_Label}")', ''], PreOrderCell_rowcol[0] + 1, ValueInputOption.user_entered)
    else:
        BoothListSheet.update_cell(PreOrderCell_rowcol[0], PreOrderCell_rowcol[1] - 1, NewPreOrderDate)
        BoothListSheet.update_cell(PreOrderCell_rowcol[0], PreOrderCell_rowcol[1], f'=HYPERLINK("{PreOrder_Link}", "{PreOrder_Label}")')

def UpdateLastestTime():
    """
    시트 내에 있는 마지막 업데이트 시간 셀을 업데이트한 후, 업데이트한 시간을 반환합니다.
    해당 셀의 위치는 a1Notation을 기준으로 셀의 행을 나타내는 전역 변수 BoothNumber_Col_Alphabet, 열을 나타내는 전역 변수 UpdateTime_Row_Number 값에 의해 결정됩니다.
    """
    print("UpdateLastestTime : 셀 전체 데이터를 가져오는 중...")
    sh = client_.open_by_key(spreadsheetId)
    sheet = sh.get_worksheet(sheetNumber)
    
    print("UpdateLastestTime : 업데이트 시간 반영 중...")
    updatetime = datetime.now()
    sheet.update_acell(f'{BoothNumber_Col_Alphabet}{UpdateTime_Row_Number}', f"마지막 업데이트 시간 : {updatetime.year}. {updatetime.month}. {updatetime.day} {updatetime.hour}:{str(updatetime.minute).zfill(2)}:{str(updatetime.second).zfill(2)}")
    
    return updatetime

def SetLinkToMap(BoothNumber: str):
    """
    부스 번호 셀과 부스 지도에서의 해당 부스 위치 셀을 서로 링크합니다.
    
    - 매개 변수
        :param BoothNumber: 서로 링크할 부스 번호
    """
    sheet = client_.open_by_key(spreadsheetId)
    BoothListSheet = sheet.get_worksheet(sheetNumber)
    BoothMapSheet = sheet.get_worksheet(MapSheetNumber)
    
    BoothNumberCell_Data = BoothMapSheet.find(BoothNumber)
    
    if ',' in BoothNumber:
        BoothNumber_splited = BoothNumber.split(',')
        for i in len(BoothNumber_splited):
           BoothNumber_splited[i].replace(" ", "")
           
        # key => 지도에서의 해당 부스의 a1 위치 값, value => 부스 위치에서의 a1 위치 값
        BoothLocations = []
        for Number in BoothNumber_splited:
            MapLocationData = BoothMapSheet.find(Number)
            BoothLocations.append(rowcol_to_a1(MapLocationData.row, MapLocationData.col))
            
            # 각 지도 셀에 부스 리스트의 셀 위치를 링크
            BoothMapSheet.update_acell(rowcol_to_a1(MapLocationData.row, MapLocationData.col), f'=HYPERLINK("#gid{BoothListSheet.id}&range={rowcol_to_a1(BoothNumberCell_Data.row, BoothNumberCell_Data.col)}", "{MapLocationData.value}")')
            
        BoothListSheet.update_acell(rowcol_to_a1(BoothNumberCell_Data.row, BoothNumberCell_Data.col), f'=HYPERLINK("#gid={BoothMapSheet.id}&range={BoothLocations[0]}:{BoothLocations[len(BoothLocations) - 1]}", "{BoothNumber}")')

def CopyInfoHyperLinkToClipBoard(InfoLabel: str, InfoLink: str):
    """
    설정한 인포 라벨, 인포 링크로 하이퍼링크를 만들어 클립보드에 복사합니다.
    
    - 메게 변수
        :param InfoLabel: 하이퍼링크로 만들 인포 라벨
        :param InfoLink: 하이퍼링크로 만들 인포 링크
    """
    infoText = f'=HYPERLINK("{InfoLink}", "{InfoLabel}")'
    clipboard.copy(infoText)
    print("CopyInfoHyperLinkToClipBoard : 인포 링크가 클립보드에 복사되었습니다.")

# 부스들의 열 리스트 설정 (해당 변수들은 예시로 제3회 일러스타 페스의 값들임)
thrid_illustarfes_alphabet_list = list(ascii_uppercase)
thrid_illustarfes_alphabet_list.append('가')
thrid_illustarfes_alphabet_list.append('나')

seoul_comic_alphabet_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']

alphabet_list = thrid_illustarfes_alphabet_list


#  부스 열별 최댓값 리스트
third_illustar_fes_booth_max_count = [25, 25, 25, 35, 35, 35, 35, 35, 35, 35, 35, 25, 25, 23, 23, 25, 25, 25, 25, 35, 35, 35, 35, 25, 16, 25, 25]
seoul_comic_world_booth_max_count = [41, 24, 24, 26, 26, 26, 24, 24, 24, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28]

alphabet_max_count = third_illustar_fes_booth_max_count

# 예시 딕셔너리
seoul_comic_alphabet_max_count_dict = dict(zip(seoul_comic_alphabet_list, seoul_comic_world_booth_max_count))
# example_al = 'A'
# print("Debug Dict : " + str(seoul_comic_alphabet_max_count_dict[example_al]))


# 부스 목록 시트의 Id
test_illustar_fes_sheet = "1CJ-K_6nBLhgyPbVSKSuq5T9tHAF-RT80WKLDmAI5NeQ"
test_seoul_comic_sheet_id = "1-lbwfQONKVZ9wD5HXpCjiQ6gElYYU-RZ3ziD0LdrnyE"

spreadsheetId = test_illustar_fes_sheet


# 부스 목록 시트 안에서 선입금 시트
sheetName_illustar_fes = "선입금, 통판, 인포 목록의 사본"
sheetNumber_illustar_fes = 1

seoul_comic_sheetName = "선입금 목록"
seoul_comic_sheetNumber = 0

sheetName = sheetName_illustar_fes
sheetNumber = sheetNumber_illustar_fes

# 선입금 시트 내의 행 인덱싱
BoothNumber_Col_Number = 2
BoothName_Col_Number = 3
Genre_Col_Number = 4
Yoil_Col_Number = 5
InfoLabel_Col_Number = 6
InfoLink_Col_Number = 6
Pre_Order_Date_Col_Number = 7
Pre_Order_Label_Col_Number = 8
Pre_Order_link_Col_Number = 8

BoothNumber_Col_Alphabet = 'B'
Yoil_Col_Alphabet = 'E'
Pre_Order_link_Col_Alphabet = 'H'
Etc_Point_Col_Alphabet = 'I'

# 선입금 시트 내의 열 인덱싱
UpdateTime_Row_Number = 1

# 업데이트 로그 시트
UpdateLogSheetName = "업데이트 내용"
UpdateLogSheetNumber = 1

# 행사 지도 시트
MapSheetNumber = None

# 기타 비교용 변수
IsAlredyExisted = False


# 근본 gspread 클라이언트
client_ = gspread.service_account()

# 메인 윈도우
class MainWindow(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("부스 정리 도구")
        
        Add_Row_Button = tk.Button(self, text='새 부스 추가하기', command= lambda: self.Open_AddorModify_Pre_Order_boothData_window(0))
        Modify_Existed_Row_Button = tk.Button(self, text='부스 번호로 검색하여 부스 수정하기', command=self.Open_searchWindow_With_BoothNum)
        Modify_Existed_Row_Button2 = tk.Button(self, text='부스 이름으로 검색하여 부스 수정하기', command=self.Open_searchWindow_With_BoothName)
        GetInfoHyperLink_Button = tk.Button(self, text='인포 하이퍼링크 바로 만들기', command=self.Open_GetInfoHyperLink_Window)

        Add_Row_Button.pack(side=tk.LEFT, padx=10, pady=10)
        Modify_Existed_Row_Button.pack(side=tk.LEFT, padx=10, pady=10)
        Modify_Existed_Row_Button2.pack(side=tk.LEFT, padx=10, pady=10)
        GetInfoHyperLink_Button.pack(side=tk.LEFT, padx=10, pady=10)
        
    def Close_newwindow(self):
        self.wm_attributes("-disabled", False)
        
        self.new_window.destroy()
        self.deiconify()
        
    def Close_ThridWindow(self):
        if self.new_window == None:
            self.wm_attributes("-disabled", False)
        
        self.third_window.destroy()
        self.deiconify()
        
    def Open_AddorModify_Pre_Order_boothData_window(self, mode: int, Cell_Row: int = None, BoothNumber: str = None, BoothName: str = None, Genre: str = None, Yoil: str = None, InfoLabel: str = None, InfoLink: str = None, Pre_Order_Date: str = None, Pre_Order_label: str = None, Pre_Order_Link: str = None):
        """
        부스 정보를 추가하거나 수정하는 tkinter 기반의 창을 엽니다.

        - 매개 변수
            :param mode : 부스 정보를 추가하는지, 수정하는지를 구분합니다. 0은 추가, 1은 수정 모드입니다.
                   
            - 매개 변수 mode 값이 1인 경우, 아래와 같은 매개 변수가 사용됩니다. 이 매개 변수들의 기본값은 None입니다.
                :param Cell_Row : 수정하려는 부스의 정보가 있는 셀의 열 번호
                :param BoothNumber : 수정하려는 부스 번호
                :param BoothName : 수정하려는 부스 이름
                :param Genre : 수정하려는 부스가 취급하는 장르
                :param Yoil : 수정하려는 부스의 참여 요일
                :param InfoLabel : 수정하려는 부스의 인포 링크의 라벨
                :param InfoLink : 수정하려는 부스의 인포 링크
                :param Pre_Order_Date : 수정하려는 부스의 선입금 또는 통판의 마감 일자
                :param Pre_Order_Label : 수정하려는 부스의 선입금 또는 통판 링크의 라벨
                :param Pre_Order_link : 수정하려는 부스의 선입금 또는 통판 링크
        """
        self.wm_attributes("-disabled", True)
        
        self.new_window = tk.Toplevel(self)
        self.new_window.minsize(500, 550)

        self.new_window.transient(self)
        
        self.new_window.protocol("WM_DELETE_WINDOW", self.Close_newwindow)
    
        self.new_window.title("부스 추가")
    
        # ========================== Settings Yoil Option ==============================
        OptionList = ["토/일", "토", "일"]
        Yoil_variable = tk.StringVar(self.new_window)
        Yoil_variable.set(OptionList[0])

        # ========================== Define UI elements ================================
        label_boothNum = tk.Label(self.new_window, text='부스 번호')
        Text_BoothNum = tk.Entry(self.new_window)

        label_boothName = tk.Label(self.new_window, text='부스 이름')
        Text_BoothName = tk.Entry(self.new_window)
    
        label_Genre = tk.Label(self.new_window, text='장르')
        Text_Genre = tk.Entry(self.new_window)
    
        label_Yoil = tk.Label(self.new_window, text='참가 요일')
        Option_Yoil = tk.OptionMenu(self.new_window, Yoil_variable, *OptionList)
        Option_Yoil.config(width=20)
    
        label_Info_Label = tk.Label(self.new_window, text='인포 라벨')
        Text_Info_Label = tk.Entry(self.new_window)
    
        label_Info = tk.Label(self.new_window, text='인포 링크')
        Text_Info = tk.Entry(self.new_window)
    
        label_Pre_Order_Date = tk.Label(self.new_window, text='선입금 마감 요일')
        Text_Pre_Order_Date = tk.Entry(self.new_window)
    
        label_Pre_Order_Label = tk.Label(self.new_window, text='선입금 링크 라벨')
        Text_Pre_Order_Label = tk.Entry(self.new_window)
    
        label_Pre_Order_Link = tk.Label(self.new_window, text='선입금 링크')
        Text_Pre_Order_Link = tk.Entry(self.new_window)
    
        Add_Row_To_Sheet_button = None
        if mode == 0:
            Add_Row_To_Sheet_button = tk.Button(self.new_window, text='부스 추가하기',
                                           command= lambda : Add_new_BoothData(Text_BoothNum.get(), Text_BoothName.get(), Text_Genre.get(), Yoil_variable.get(),
                                                                         Text_Info_Label.get(), Text_Info.get(), Text_Pre_Order_Date.get(), Text_Pre_Order_Label.get(), Text_Pre_Order_Link.get()))

        else:
            self.new_window.title("부스 수정")
            Add_Row_To_Sheet_button = tk.Button(self.new_window, text='부스 수정하기',
                                                command= lambda: Modify_Existed_Row(Cell_Row, Text_BoothNum.get(), Text_BoothName.get(), Text_Genre.get(), Yoil_variable.get(),
                                                                         Text_Info_Label.get(), Text_Info.get(), Text_Pre_Order_Date.get(), Text_Pre_Order_Label.get(), Text_Pre_Order_Link.get()))
        
        entryList = [Text_BoothNum, Text_BoothName, Text_Genre, Text_Info_Label, Text_Info, Text_Pre_Order_Date, Text_Pre_Order_Label, Text_Pre_Order_Link]
        Empty_all_Entry = tk.Button(self.new_window, text='모든 칸 비우기', command= lambda: self.Empty_all_Entries(entryList))

        # ============================ placing element =================================
        label_boothNum.grid(column=0, row=0, padx=10, pady=2.5, sticky="w")
        Text_BoothNum.grid(column=0, row=1, padx=10, pady=2.5, ipadx = 170)
    
        label_boothName.grid(column=0, row=2, padx=10, pady=2.5, sticky="w")
        Text_BoothName.grid(column=0, row=3, padx=10, pady=2.5, ipadx = 170)
    
        label_Genre.grid(column=0, row=4, padx=10, pady=2.5, sticky="w")
        Text_Genre.grid(column=0, row=5, padx=10, pady=2.5, ipadx = 170)
    
        label_Yoil.grid(column=0, row=6, padx=10, pady=2.5, sticky="w")
        Option_Yoil.grid(column=0, row=7, padx=10, pady=2.5, sticky="w")
    
        label_Info_Label.grid(column=0, row=8, padx=10, pady=2.5, sticky="w")
        Text_Info_Label.grid(column=0, row=9, padx=10, pady=2.5, ipadx = 170)
    
        label_Info.grid(column=0, row=10, padx=10, pady=2.5, sticky="w")
        Text_Info.grid(column=0, row=11, padx=10, pady=2.5, ipadx = 170)
    
        label_Pre_Order_Date.grid(column=0, row=12, padx=10, pady=2.5, sticky="w")
        Text_Pre_Order_Date.grid(column=0, row=13, padx=10, pady=2.5, ipadx = 170)
    
        label_Pre_Order_Label.grid(column=0, row=14, padx=10, pady=2.5, sticky="w")
        Text_Pre_Order_Label.grid(column=0, row=15, padx=10, pady=2.5, ipadx = 170)
    
        label_Pre_Order_Link.grid(column=0, row=16, padx=10, pady=2.5, sticky="w")
        Text_Pre_Order_Link.grid(column=0, row=17, padx=10, pady=2.5, ipadx = 170)
    
        Add_Row_To_Sheet_button.grid(column=0, row=18, padx=10, pady=5, sticky="e")
        
        Empty_all_Entry.grid(column=0, row=19, padx=10, pady=5, sticky="e")
    
        if mode == 1:
            Text_BoothNum.insert(0, BoothNumber)
            Text_BoothName.insert(0, BoothName)
            Text_Genre.insert(0, Genre)
            Yoil_variable.set(Yoil)
            if InfoLabel != None:
                Text_Info_Label.insert(0, InfoLabel)
                Text_Info.insert(0, InfoLink)
            if Pre_Order_label != None:
                Text_Pre_Order_Date.insert(0, Pre_Order_Date)
                Text_Pre_Order_Label.insert(0, Pre_Order_label)
                Text_Pre_Order_Link.insert(0, Pre_Order_Link)
                
        self.new_window.focus_set()
        
    def Empty_all_Entries(self, EntryList: list[Entry]):
        """
        매개 변수 EntryList에 있는 모든 Entry의 문자열을 빈 문자열로 바꿉니다.
        
        - 매개 변수
            :parma EntryList: 빈 문자열로 바꿀 Entry 요소의 리스트
        """
        for entry in EntryList:
            entry.delete(0, len(entry.get()) - 1)
        
    def Open_searchWindow_With_BoothNum(self):
        """
        부스 번호를 검색하기 위한 tkinter 기반의 창을 엽니다.
        """
        self.wm_attributes("-disabled", True)
        
        self.new_window = tk.Toplevel(self)
        self.new_window.minsize(300, 100)

        self.new_window.transient(self)
        
        self.new_window.protocol("WM_DELETE_WINDOW", self.Close_newwindow)
    
        self.new_window.title("부스 번호로 검색")
        
        OptionList = ["인포", "선입금 링크"]
        Option_variable = tk.StringVar(self.new_window)
        Option_variable.set(OptionList[0])
    
        label_SearchToBoothNum = tk.Label(self.new_window, text='검색할 부스 번호')
        Text_SearchToBoothNum = tk.Entry(self.new_window)
    
        label_Modifything = tk.Label(self.new_window, text='수정할 요소')
        Option_Thing = tk.OptionMenu(self.new_window, Option_variable, *OptionList)
        Option_Thing.config(width=20)
        
        Button_Search = tk.Button(self.new_window, text='검색하기', command= lambda: self.Search_Booth_WithBoothNumber(Text_SearchToBoothNum.get(), Option_variable.get()))
    
        label_SearchToBoothNum.grid(column = 0, row=0, padx=10, pady=2.5, sticky="w")
        Text_SearchToBoothNum.grid(column=0, row=1, padx=10, pady=2.5, ipadx=60)
        
        label_Modifything.grid(column=0, row=2, padx=10, pady=2.5, sticky="w")
        Option_Thing.grid(column=0, row=3, padx=10, pady=2.5, ipadx=20)
    
        Button_Search.grid(column=0, row=4, padx=10, pady=5, sticky="e")
        
        self.new_window.focus_set()
     
    def Open_searchWindow_With_BoothName(self):
        """
        부스 이름으로 검색하기 위한 tkinter 기반의 창을 엽니다.
        """
        self.wm_attributes("-disabled", True)
        
        self.new_window = tk.Toplevel(self)
        self.new_window.minsize(300, 100)

        self.new_window.transient(self)
        
        self.new_window.protocol("WM_DELETE_WINDOW", self.Close_newwindow)
    
        self.new_window.title("부스 이름으로 검색")
        
        OptionList = ["인포", "선입금 링크"]
        Option_variable = tk.StringVar(self.new_window)
        Option_variable.set(OptionList[0])
    
        label_SearchToBoothName = tk.Label(self.new_window, text='검색할 부스 이름')
        Text_SearchToBoothName = tk.Entry(self.new_window)
    
        label_Modifything = tk.Label(self.new_window, text='수정할 요소')
        Option_Thing = tk.OptionMenu(self.new_window, Option_variable, *OptionList)
        Option_Thing.config(width=20)
    
        Button_Search = tk.Button(self.new_window, text='검색하기', command= lambda: self.Search_Booth_WithBoothName(Text_SearchToBoothName.get(), Option_variable.get()))
    
        label_SearchToBoothName.grid(column = 0, row=0, padx=10, pady=2.5, sticky="w")
        Text_SearchToBoothName.grid(column=0, row=1, padx=10, pady=2.5, ipadx=60)
        
        label_Modifything.grid(column=0, row=2, padx=10, pady=2.5, sticky="w")
        Option_Thing.grid(column=0, row=3, padx=10, pady=2.5, ipadx=20)
    
        Button_Search.grid(column=0, row=4, padx=10, pady=5, sticky="e")
        
        self.new_window.focus_set()
  
    def Search_Booth_WithBoothNumber(self, BoothNumber: str, Option: str):
        """
        해당 부스 번호를 검색하여 수정하기 위한 창을 엽니다. 수정할 요소는 인포, 선입금 링크 둘 중 하나입니다.
        검색 결과가 없는 경우, -1을 반환합니다.
    
        - 매개 변수
            :param BoothNumber : 검색하려는 부스의 번호
            :param Option : 검색하여 가져오려는 정보의 종류로 "인포", "선입금 링크" 두 문자열 중 하나입니다.
        """

        print("Search_Booth_WithBoothNumber : 셀 전체 데이터를 가져오는 중...")
        sh = client_.open_by_key(spreadsheetId)
        sheet = sh.get_worksheet(sheetNumber)
    
        booth_list = sh.get_worksheet(sheetNumber).get(f"{BoothNumber_Col_Alphabet}1:{BoothNumber_Col_Alphabet}")

        booth_list_tmp = booth_list.copy()
        #print(booth_list_tmp[13])
        j = 0
        for boothnum in booth_list_tmp:
            if len(boothnum) == 0:
                booth_list_tmp[j] = booth_list_tmp[j - 1]
            if ',' in str(booth_list_tmp[j][0]):
                booth_list_tmp[j] = booth_list_tmp[j][0].split(", ")
            j = j + 1
            
        """
        Search_booth_number = ""
        # print(str(booth_list_tmp))
        if ',' in BoothNumber:
            booth_split = BoothNumber.split(",")
            for i in range(0, len(booth_list_tmp)):
                #print("debug i : " + str(i))
                for k in range(0, len(booth_list_tmp[i])):
                    #print("debug k : " + str(k))
                    if booth_list_tmp[i][k] == booth_split[0]:
                        for l in range(0, len(booth_list_tmp[i])):
                            if l == len(booth_list_tmp[i]) - 1:
                                Search_booth_number += booth_list_tmp[i][l]
                            else:
                                Search_booth_number += booth_list_tmp[i][l] + ', '
                        k = len(booth_list_tmp[i])
                        i = len(booth_list_tmp)
                    else:
                        continue
                        
        else:
            Search_booth_number = BoothNumber
        """
        Search_booth_number = BoothNumber
        
        print(f"Search_Booth_WithBoothNumber : 부스 번호 {Search_booth_number} 검색 중...")
        result_cell = sheet.find(Search_booth_number)
    
        if result_cell == None:
            print("Search_Booth_WithBoothNumber : 검색 결과가 없습니다.")
            self.printDebugLine()
            messagebox.showinfo('검색', message='검색 결과가 없습니다.')
            return -1
        
        print("Search_Booth_WithBoothNumber : 결과 재구성 중...")
        result_list = []
        result_list.append(sheet.cell(result_cell.row, 1).value)
        for i in range(0, 7):
            result_list.append(sheet.cell(result_cell.row, result_cell.col + i).value)
        
        res = LinkCollector.GetBasicData(spreadsheetId, sheetName)
        InfoLink_list = LinkCollector.GetInfoLinks(res)
        OrderLink_list = LinkCollector.GetPreOrderAndMailOrderLinks(res, sheet.get_all_cells())
        
        InfoLink_tuple = tuple()
        OrderLink_tuple = tuple()
        match_Info = [data1 for data1 in InfoLink_list if data1["Cell"] == rowcol_to_a1(result_cell.row, InfoLink_Col_Number)]
        match_OrderLink = [data2 for data2 in OrderLink_list if data2["link_cell"] == rowcol_to_a1(result_cell.row, Pre_Order_link_Col_Number)]
        if Option == "인포":
            k = 0
            for k in range(result_cell.row, len(booth_list)):
                if len(booth_list[k]) == 0:
                    for data2 in InfoLink_list:
                        if data2["Cell"] == rowcol_to_a1(k + 1, InfoLink_Col_Number):
                            match_Info.append(data2)
                else:
                    break;
            
            for k in range(result_cell.row, len(booth_list)):
                if len(booth_list[k]) == 0:
                    for data2 in OrderLink_list:
                        if data2["link_cell"] == rowcol_to_a1(k + 1, Pre_Order_link_Col_Number):
                            match_OrderLink.append(data2)
                else:
                    break;
                            
            if len(match_Info) != 0:
                for n in range(len(match_Info)):
                    InfoLink_tuple += ([result_list[1], result_list[2], match_Info[n]["Cell"], match_Info[n]["Info_Text"], match_Info[n]["Info_URL"]], )
                    print("Search_Booth_WithBoothNumber : 디버그 : match_Info : " + str(match_Info[n]))
            
                self.printDebugLine()
                self.Close_newwindow()
                self.Open_Modify_Info_Window(InfoLink_tuple, Title=f'{result_list[1]} : {result_list[2]} 부스의 인포 검색 결과')
                
            else:
                print("Search_Booth_WithBoothNumber : 부스는 시트에 있지만 인포가 없습니다. 해당 부스의 인포를 추가하는 창을 출력합니다.")
                self.printDebugLine()
                self.AddInfoWindow(f'{rowcol_to_a1(result_cell.row, InfoLabel_Col_Number)}', PreOrderLinkCell_Count= len(match_OrderLink), Title=f'{result_list[1]} : {result_list[2]} 부스의 인포 추가')
        elif Option == "선입금 링크":
            #print("Search_Booth_WithBoothNumber : 디버그 : match_Info : " + str(match_OrderLink))
            #print("Search_Booth_WithBoothNumber : 디버그 : result_cell.row : " + str(result_cell.row))
            k = 0
            for k in range(result_cell.row, len(booth_list)):
                #print("Search_Booth_WithBoothNumber : 디버그 : boothlist[k] : " + str(booth_list[k]))
                if len(booth_list[k]) == 0:
                    for data2 in OrderLink_list:
                        if data2["link_cell"] == rowcol_to_a1(k + 1, Pre_Order_link_Col_Number):
                            match_OrderLink.append(data2)
                else:
                    break;
            if len(match_OrderLink) != 0:
                for n in range(len(match_OrderLink)):
                    OrderLink_tuple += ([result_list[1], result_list[2], match_OrderLink[n]["link_cell"], match_OrderLink[n]["date"], match_OrderLink[n]["link_text"], match_OrderLink[n]["link"]], )
                    print("Search_Booth_WithBoothNumber : 디버그 : match_Info[n] : " + str(match_OrderLink[n]))
                    
                self.printDebugLine()
                self.Close_newwindow()
                self.Open_Modify_PreOrder_Window(OrderLink_tuple, Title=f'{result_list[1]} : {result_list[2]} 부스의 선입금 링크 검색 결과')
            
            else:
                print("Search_Booth_WithBoothNumber : 부스는 시트에 있지만 선입금 링크가 없습니다. 해당 부스의 선입금 링크를 추가하는 창을 출력합니다.")
                self.printDebugLine()
                self.AddPreOrderWindow(f'{rowcol_to_a1(result_cell.row, Pre_Order_link_Col_Number)}', Title=f'{result_list[1]} : {result_list[2]} 부스의 선입금 링크 추가')

        #print("Search_Booth_WithBoothNumber : 디버그 : match_Info : " + str(OrderLink_tuple))

    def Search_Booth_WithBoothName(self, BoothName: str, Option: str):
        """
        부스 이름으로 부스 정보들을 검색하여 수정하기 위한 창을 엽니다. 수정할 요소는 인포, 선입금 링크 둘 중 하나입니다.
        검색 결과가 없는 경우, -1을 반환합니다.
    
        - 매개 변수
            :param BoothName : 검색하려는 부스의 이름
            :param Option : 검색하여 가져오려는 정보의 종류로 "인포", "선입금" 두 문자열 중 하나입니다.
        """
        print("Search_Booth_WithBoothName : 셀 전체 데이터를 가져오는 중...")
        sh = client_.open_by_key(spreadsheetId)
        sheet = sh.get_worksheet(sheetNumber)
    
        booth_list = sh.get_worksheet(sheetNumber).get(f"{BoothNumber_Col_Alphabet}1:{BoothNumber_Col_Alphabet}")

        booth_list_tmp = booth_list.copy()
        #print(booth_list_tmp[13])
        j = 0
        for boothnum in booth_list_tmp:
            if len(boothnum) == 0:
                booth_list_tmp[j] = booth_list_tmp[j - 1]
            if ',' in str(booth_list_tmp[j][0]):
                booth_list_tmp[j] = booth_list_tmp[j][0].split(", ")
            j = j + 1
    
            
        print(f"Search_Booth_WithBoothName : 부스 이름 {BoothName} 검색 중...")
        result_cell = sheet.find(BoothName)
    
        if result_cell == None:
            print("Search_Booth_WithBoothName : 검색 겱과가 없습니다.")
            self.printDebugLine()
            messagebox.showinfo('검색', message='검색 결과가 없습니다.')
            return -1
        
        result_list = []
        result_list.append(sheet.cell(result_cell.row, 1).value)
        # col = 3 -> BoothName, range(-1, 5) -> result_coll.col + i -> 2 ~ 8
        for i in range(-1, 5):
            result_list.append(sheet.cell(result_cell.row, result_cell.col + i).value)
        
        res = LinkCollector.GetBasicData(spreadsheetId, sheetName)
        InfoLink_list = LinkCollector.GetInfoLinks(res)
        OrderLink_list = LinkCollector.GetPreOrderAndMailOrderLinks(res, sheet.get_all_cells())
        
        print("Search_Booth_WithBoothName : 결과 재구성 중...")
        InfoLink_tuple = tuple()
        OrderLink_tuple = tuple()
        match_Info = [data1 for data1 in InfoLink_list if data1["Cell"] == rowcol_to_a1(result_cell.row, InfoLink_Col_Number)]
        match_OrderLink = [data2 for data2 in OrderLink_list if data2["link_cell"] == rowcol_to_a1(result_cell.row, Pre_Order_link_Col_Number)]
        if Option == "인포":
            k = 0
            for k in range(result_cell.row, len(booth_list)):
                if len(booth_list[k]) == 0:
                    for data2 in InfoLink_list:
                        if data2["Cell"] == rowcol_to_a1(k + 1, InfoLink_Col_Number):
                            match_Info.append(data2)
                else:
                    break;
            if len(match_Info) != 0:
                for n in range(len(match_Info)):
                    InfoLink_tuple += ([result_list[1], result_list[2], match_Info[n]["Cell"], match_Info[n]["Info_Text"], match_Info[n]["Info_URL"]], )
                    print("Search_Booth_WithBoothName : 디버그 : match_Info : " + str(match_Info[n]))
            
                self.printDebugLine()
                self.Close_newwindow()
                self.Open_Modify_Info_Window(InfoLink_tuple, Title=f'{result_list[1]} : {result_list[2]} 부스의 인포 검색 결과')
                
            else:
                print("Search_Booth_WithBoothName : 부스는 시트에 있지만 인포가 없습니다. 해당 부스의 인포를 추가하는 창을 출력합니다.")
                self.printDebugLine()
                self.AddInfoWindow(f'{rowcol_to_a1(result_cell.row, InfoLabel_Col_Number)}', PreOrderLinkCell_Count= len(match_OrderLink), Title=f'{result_list[1]} : {result_list[2]} 부스의 인포 추가')
            # link_list.append(match_Info[0]["Info_URL"])
        elif Option == "선입금 링크":
            #print("Search_Booth_WithBoothNumber : 디버그 : match_Info : " + str(match_OrderLink))
            #print("Search_Booth_WithBoothNumber : 디버그 : result_cell.row : " + str(result_cell.row))
            k = 0
            for k in range(result_cell.row, len(booth_list)):
                #print("Search_Booth_WithBoothNumber :k 디버그 : boothlist[k] : " + str(booth_list[k]))
                if len(booth_list[k]) == 0:
                    for data2 in OrderLink_list:
                        if data2["link_cell"] == rowcol_to_a1(k + 1, Pre_Order_link_Col_Number):
                            match_OrderLink.append(data2)
                else:
                    break;
            if len(match_OrderLink) != 0:
                for n in range(len(match_OrderLink)):
                    OrderLink_tuple += ([result_list[1], result_list[2], match_OrderLink[n]["link_cell"], match_OrderLink[n]["date"], match_OrderLink[n]["link_text"], match_OrderLink[n]["link"]], )
                    print("Search_Booth_WithBoothName : 디버그 : match_Info[n] : " + str(match_OrderLink[n]))
                    
                self.printDebugLine()
                self.Close_newwindow()
                self.Open_Modify_PreOrder_Window(OrderLink_tuple, Title=f'{result_list[1]} : {result_list[2]} 부스의 선입금 링크 검색 결과')
                
            else:
                print("Search_Booth_WithBoothName : 부스는 시트에 있지만 선입금 링크가 없습니다. 해당 부스의 선입금 링크를 추가하는 창을 출력합니다.")
                self.printDebugLine()
                self.AddPreOrderWindow(f'{rowcol_to_a1(result_cell.row, Pre_Order_link_Col_Number)}', Title=f'{result_list[1]} : {result_list[2]} 부스의 선입금 링크 추가')
        
    def Open_Modify_Info_Window(self, InfoList: tuple, Title: str = None):
        """
        인포를 수정하는 tkinter 기반의 창을 엽니다.
        
        - 매개 변수
            :param InfoList: 인포들을 가진 튜플입니다. 이 튜플은 [인포 셀 위치, 인포 라벨, 인포 링크]로 구성되어야 합니다.
            :param Title (선택): 여는 창의 제목입니다. 기본값은 None입니다.
        """
        self.wm_attributes("-disabled", True)
        
        self.new_window = tk.Toplevel(self)
        self.new_window.minsize(1000, 335)

        self.new_window.transient(self)
        
        self.new_window.protocol("WM_DELETE_WINDOW", self.Close_newwindow)
    
        if Title != None:
            self.new_window.title(Title)
        else:
            self.new_window.title("부스 인포 수정")
        
        treeview = ttk.Treeview(self.new_window, columns=["BoothNumber", "BoothName", "CellLocation", "InfoLabel", "InfoLink"], displaycolumns=["BoothNumber", "BoothName", "CellLocation", "InfoLabel", "InfoLink"])
        treeview.grid(column = 0, row = 0, padx = 10, ipadx=500)
        
        treeview.column("BoothNumber", width=50, anchor="center")
        treeview.heading("BoothNumber", text="부스 번호", anchor="center")
        
        treeview.column("BoothName", width=100, anchor="center")
        treeview.heading("BoothName", text="부스 이름", anchor="center")
        
        treeview.column("CellLocation", width=50, anchor="center")
        treeview.heading("CellLocation", text="인포 셀 위치", anchor="center")
        
        treeview.column("InfoLabel", width=50, anchor="center")
        treeview.heading("InfoLabel", text="인포 라벨", anchor="center")
        
        treeview.column("InfoLink", width=750 ,anchor="center")
        treeview.heading("InfoLink", text="인포 링크", anchor="center")

        treeview["show"] = "headings"
        

        Button_Add_info = tk.Button(self.new_window, text='인포 추가하기', command=self.AddInfoWindow)
        Button_Edit_info = tk.Button(self.new_window, text="인포 수정하기")
        
        Button_Add_info.grid(column=0, row=1, padx=10, pady=5, sticky="e")
        Button_Edit_info.grid(column=0, row=2, padx=10, pady=5, sticky="e")
        
        for i in range(len(InfoList)):
            treeview.insert('', 'end', text="", value=InfoList[i], iid=i)
            
        self.new_window.focus_set()

    def Open_Modify_PreOrder_Window(self, PreOrderList: tuple, Title: str = None):
        """
        선입금 링크들을 수정하는 tkinter 기반의 창을 엽니다.
        
        - 매개 변수
            :param PreOrderList: 선입금 링크들을 가진 튜플입니다. 이 튜플은 [선입금 링크 셀 위치, 선입금 마감 일자, 선입금 링크 라벨, 선입금 링크] 로 구성되어야 합니다.
            :param Title (선택): 여는 창의 제목입니다. 기본값은 None입니다.
        """
        self.wm_attributes("-disabled", True)
        
        self.new_window = tk.Toplevel(self)
        self.new_window.minsize(1000, 335)

        self.new_window.transient(self)
        
        self.new_window.protocol("WM_DELETE_WINDOW", self.Close_newwindow)
    
        if Title != None:
            self.new_window.title(Title)
        else:
            self.new_window.title("부스 선입금 링크 수정")
        
        treeview = ttk.Treeview(self.new_window, columns=["BoothNumber", "BoothName", "CellLocation", "PreOrderDate", "PreOrderLabel", "PreOrderLink"], displaycolumns=["BoothNumber", "BoothName", "CellLocation", "PreOrderDate", "PreOrderLabel", "PreOrderLink"])
        treeview.grid(column = 0, row = 0, padx = 10, ipadx=500)
        
        treeview.column("BoothNumber", width=100, anchor="center")
        treeview.heading("BoothNumber", text="부스 번호", anchor="center")
        
        treeview.column("BoothName", width=100, anchor="center")
        treeview.heading("BoothName", text="부스 이름", anchor="center")
        
        treeview.column("CellLocation", width=50, anchor="center")
        treeview.heading("CellLocation", text="선입금 링크 셀 위치", anchor="center")
        
        treeview.column("PreOrderDate", width=50, anchor="center")
        treeview.heading("PreOrderDate", text="선입금 마감 일", anchor="center")
        
        treeview.column("PreOrderLabel", width=100, anchor="center")
        treeview.heading("PreOrderLabel", text="선입금 링크 라벨", anchor="center")
        
        treeview.column("PreOrderLink", anchor="center")
        treeview.heading("PreOrderLink", text="선입금 링크", anchor="center")

        treeview["show"] = "headings"
        
        
        Button_Add_info = tk.Button(self.new_window, text="선입금 링크 추가하기", command= lambda: self.AddPreOrderWindow(treeview.item(treeview.focus()).get('values')[2], 
                                                                                                                 None, 
                                                                                                                 f'{treeview.item(treeview.focus()).get("values")[0]} : {treeview.item(treeview.focus()).get("values")[1]} 부스의 선입금 링크 추가'))
        Button_Edit_info = tk.Button(self.new_window, text="선입금 링크 수정하기", command= lambda: self.AddPreOrderWindow(treeview.item(treeview.focus()).get('values')[2], 
                                                                                                                  treeview.item(treeview.focus()).get('values'), 
                                                                                                                  f'{treeview.item(treeview.focus()).get("values")[0]} : {treeview.item(treeview.focus()).get("values")[1]} 부스의 {treeview.item(treeview.focus()).get("values")[2]} 셀의 선입금 링크 수정'))
        
        Button_Add_info.grid(column=0, row=1, padx=10, pady=5, sticky="e")
        Button_Edit_info.grid(column=0, row=2, padx=10, pady=5, sticky="e")
        
        for i in range(len(PreOrderList)):
            treeview.insert('', 'end', text="", value=PreOrderList[i], iid=i)
            
        self.new_window.focus_set()
        
    def EditInfoLink(self, Treeview: ttk.Treeview):
        """
        해당 트리뷰에서 선택한 행의 인포 링크를 수정하여 시트에 반영합니다.
        
        - 매개 변수
            :param Treeview: 수정 창에 있는 트리뷰
        """
        selectedItem = Treeview.focus()
        getValue = Treeview.item(selectedItem).get('values')
        
        EditInfoCell(getValue[0], getValue[1], getValue[2])
        
    def EditPreOrderLink(self, Treeview: ttk.Treeview):
        """
        해당 트리뷰에서 선택한 행의 선입금 링크를 수정하여 시트에 반영합니다.
        
        - 매개 변수
            :param Treeview: 수정 창에 있는 트리뷰
        """
        selectedItem = Treeview.focus()
        getValue = Treeview.item(selectedItem).get('values')
        
        EditPreOrderCell(getValue[0], getValue[1], getValue[2], getValue[3])
        
    def AddInfoWindow(self, infoCell_a1: str, PreOrderLinkCell_Count: int = 1, Title: str = None):
        """
        이미 있는 부스의 인포를 추가하는 tkinter 기반의 창을 엽니다.
        
        - 매개 변수
            :param infoCell_a1: 추가하려는 인포가 있는 셀의 a1Notation 값
            :param PreOrderLinkCell_Count (선택) : 추가혀려는 인포가 있는 열의 선입금 링크의 수입니다. 기본값은 1입니다.
            :param Title (선택): 여는 창의 제목입니다. 기본값은 None입니다.
        """
        self.wm_attributes("-disabled", True)
        
        self.third_window = tk.Toplevel(self)
        self.third_window.minsize(300, 200)

        self.third_window.transient(self)
        
        self.third_window.protocol("WM_DELETE_WINDOW", self.Close_ThridWindow)
    
        if Title != None:
            self.third_window.title(Title)
        else:
            self.third_window.title("인포 추가")
        
        Add_Info_Label = tk.Label(self.third_window, text='인포 라벨')
        Add_Info_Label_Entry = tk.Entry(self.third_window)
        
        Add_Info_Link_Label = tk.Label(self.third_window, text='인포 링크')
        Add_Info_Link_Entry = tk.Entry(self.third_window)
        
        Add_Info_Button = tk.Button(self.third_window, text='인포 추가', command= lambda: EditInfoCell(infoCell_a1, Add_Info_Label_Entry.get(), Add_Info_Link_Entry.get(), PreOrderLinkCell_Count))
        
        Add_Info_Label.grid(column=0, row=0, padx=10, pady=5, sticky='w')
        Add_Info_Label_Entry.grid(column=0, row=1, padx=10, pady=5, ipadx=170)
        
        Add_Info_Link_Label.grid(column=0, row=2, padx=10, pady=5, sticky='w')
        Add_Info_Link_Entry.grid(column=0, row=3, padx=10, pady=5, ipadx=170)

        Add_Info_Button.grid(column=0, row=4, padx=10, pady=5, sticky='e')
        
    def AddPreOrderWindow(self, PreOrderCell_a1: str, selectedItem: list = None, Title: str = None):
        """
        이미 있는 부스의 선입금 링크를 추가하는 tkinter 기반의 창을 엽니다.
        
        - 매개 변수
            :param PreOrderCell_a1: 선입금 링크가 추가될 셀의 a1Notation 값
            :param selectedItem (선택): 호출한 함수에서 treeview를 사용하는 경우, 추가 또는 수정한 선입금 링크 데이터의 리스트입니다.
            :param Title (선택): 여는 창의 제목입니다. 기본값은 None입니다.
        """
        self.wm_attributes("-disabled", True)
        
        self.third_window = tk.Toplevel(self)
        self.third_window.minsize(300, 200)

        self.third_window.transient(self)
        
        self.third_window.protocol("WM_DELETE_WINDOW", self.Close_ThridWindow)
    
        if Title != None:
            self.third_window.title(Title)
        else:
            self.third_window.title("선입금 링크 추가")
        
        Add_PreOrderDate_Label = tk.Label(self.third_window, text='선입금 마감 일자')
        Add_PreOrderDate = tk.Entry(self.third_window)

        Add_PreOrder_Label = tk.Label(self.third_window, text='선입금 라벨')
        Add_PreOrder_Label_Entry = tk.Entry(self.third_window)

        Add_PreOrder_Link_Label = tk.Label(self.third_window, text='선입금 링크')
        Add_PreOrder_Link_Entry = tk.Entry(self.third_window)
        
        button_Text = ''
        if "수정" in Title:
            button_Text = '선입금 링크 수정'
        else:
            button_Text = '선입금 링크 추가'
            
        mode = 0
        if "수정" in button_Text:
            mode = 1
        else:
            mode = 0
            
        Add_PreOrder_Button = tk.Button(self.third_window, text=button_Text, command= lambda: EditPreOrderCell(PreOrderCell_a1, Add_PreOrderDate.get(), Add_PreOrder_Label_Entry.get(), Add_PreOrder_Link_Entry.get(), mode))
        
        Add_PreOrderDate_Label.grid(column=0, row=0, padx=10, pady=5, sticky='w')
        Add_PreOrderDate.grid(column=0, row=1, padx=10, pady=5, ipadx=170)
        
        Add_PreOrder_Label.grid(column=0, row=2, padx=10, pady=5, sticky='w')
        Add_PreOrder_Label_Entry.grid(column=0, row=3, padx=10, pady=5, ipadx=170)
        
        Add_PreOrder_Link_Label.grid(column=0, row=4, padx=10, pady=5, sticky='w')
        Add_PreOrder_Link_Entry.grid(column=0, row=5, padx=10, pady=5, ipadx=170)

        Add_PreOrder_Button.grid(column=0, row=6, padx=10, pady=5, sticky='e')
        
        if selectedItem != None:
            Add_PreOrderDate.insert(0, selectedItem[3])
            Add_PreOrder_Label_Entry.insert(0, selectedItem[4])
            Add_PreOrder_Link_Entry.insert(0, selectedItem[5])


    def Open_GetInfoHyperLink_Window(self):
        """
        인포 하이퍼링크를 클립보드에 복사하는 tkinter 기반의 창을 엽니다.
        """
        self.wm_attributes("-disabled", True)
        
        self.new_window = tk.Toplevel(self)
        self.new_window.minsize(300, 200)

        self.new_window.transient(self)
        
        self.new_window.protocol("WM_DELETE_WINDOW", self.Close_newwindow)
    
        self.new_window.title('인포 하이퍼링크 제작 도구')
        
        Add_Info_Label = tk.Label(self.new_window, text='인포 라벨')
        Add_Info_Label_Entry = tk.Entry(self.new_window)
        
        Add_Info_Link_Label = tk.Label(self.new_window, text='인포 링크')
        Add_Info_Link_Entry = tk.Entry(self.new_window)
        
        Add_Info_Button = tk.Button(self.new_window, text='클립보드에 복사', command= lambda: CopyInfoHyperLinkToClipBoard(Add_Info_Label_Entry.get(), Add_Info_Link_Entry.get()))
        
        Add_Info_Label.grid(column=0, row=0, padx=10, pady=5, sticky='w')
        Add_Info_Label_Entry.grid(column=0, row=1, padx=10, pady=5, ipadx=170)
        
        Add_Info_Link_Label.grid(column=0, row=2, padx=10, pady=5, sticky='w')
        Add_Info_Link_Entry.grid(column=0, row=3, padx=10, pady=5, ipadx=170)

        Add_Info_Button.grid(column=0, row=4, padx=10, pady=5, sticky='e')

    def printDebugLine(self):
        print("=========================================================================")
        
app = MainWindow()
app.mainloop()
