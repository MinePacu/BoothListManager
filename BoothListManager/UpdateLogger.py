from datetime import datetime
import gspread
from enum import Enum
from gspread.utils import ValueInputOption
from gspread_formatting.models import Borders
import gspread_formatting

class LogType(Enum):
    Pre_Order = 1
    Mail_Order = 2
    Info = 3

class UpdateLogger():
    """
    업데이트 로그에 대한 클래스
    """
    @staticmethod
    def AddUpdateLog(sheet: gspread.Worksheet, logtype: LogType, updatetime: datetime, sheetid_hyperlink: str, HyperLinkCell: str, BoothNumber: str = None, BoothName: str = None, IsOwnAuthor: bool = False, AuthorNickName: str = None):
        """
        업데이트 로그를 추가합니다.
        
        - 매개 변수
            :param sheet: 업데이트 로그를 추가할 시트
            :param logtype: 업데이트 로그의 추가 타입으로 선입금인 Pre_Order, 통판인 Mail_Order 둘 중 하나의 값입니다.
            :param updatetime: 업데이트 시간
            :param sheetid_hyperlink: 업데이트된 부스 정보가 위치한 시트의 Id
            :param HyperLinkCell: 업데이트된 부스 정보의 선입금 또는 통판 또는 인포 링크가 담긴 셀의 a1Notation 값
            :param BoothNumber: (선택) 업데이트된 부스의 부스 번호, 이 값이 None인 경우, BoothName은 None이 아니여야 합니다.
            :param BoothName: (선택) 업데이트된 부스의 부스 이름, 이 값이 None인 경우, BoothNumber은 None이 아니여야 합니다.
            :param IsOwnAuthor: (선택) 업데이트된 정보가 특정 작가님별로 업데이트된 정보인지 여부입니다. 기본값은 False입니다.
            :param AuthorNickName: (선택) IsOwnAuthor의 값이 True인 경우에 사용되며, 작가님의 닉네임입니다. 기본값은 None입니다.
        """
        
        # 업데이트 로그의 가장 최신 로그가 자리할 열의 위치
        updatelog_Lastest_Low = 5

        # 업데이트 로그가 B행부터 시작이면 이렇게, 아닌 경우 이걸 수정
        updatelog_data = ['']
        
        # 업데이트 로그의 시각과 내용이 들어갈 행의 알파벳
        UpdateLogtime_ColAlphabet = 'B'
        UpdateLog_ColAlphabet = 'C'
        
        updatelog_time = datetime.now()
        updatelog_data.append(f'{updatelog_time.month}.{updatelog_time.day} {updatelog_time.hour}:{str(updatelog_time.minute).zfill(2)}:{str(updatelog_time.second).zfill(2)}')
        
        updatelog_string = f''
        if logtype == LogType.Pre_Order:
            if IsOwnAuthor == True:
                if BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothNumber} 부스의 {AuthorNickName} 작가님의 선입금 링크 추가")'
                elif BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothName} 부스의 {AuthorNickName} 작가님의 선입금 링크 추가")'
            else:
                if BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothNumber} 부스의 선입금 링크 추가")'
                elif BoothName != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothName} 부스의 선입금 링크 추가")'

        elif logtype == LogType.Mail_Order:
            if IsOwnAuthor == True:
                if BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothNumber} 부스의 {AuthorNickName} 작가님의 통판 링크 추가")'
                elif BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothName} 부스의 {AuthorNickName} 작가님의 통판 링크 추가")'
            else:
                if BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothNumber} 부스의 통판 링크 추가")'
                elif BoothName != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothName} 부스의 통판 링크 추가")'
                    
        elif logtype == LogType.Info:
            if IsOwnAuthor == True:
                if BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothNumber} 부스의 {AuthorNickName} 작가님의 인포 추가")'
                elif BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothName} 부스의 {AuthorNickName} 작가님의 인포 추가")'
            else:
                if BoothNumber != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothNumber} 부스의 인포 추가")'
                elif BoothName != None:
                    updatelog_string = f'=HYPERLINK("#gid={sheetid_hyperlink}&range={HyperLinkCell}", "{BoothName} 부스의 인포 추가")'
                    
        updatelog_data.append(updatelog_string)
        
        sheet.insert_row(updatelog_data, updatelog_Lastest_Low, value_input_option=ValueInputOption.user_entered)
        fmt = gspread_formatting.CellFormat(
            borders=Borders(
                top=gspread_formatting.Border("SOLID"), 
                bottom=gspread_formatting.Border("SOLID"), 
                left=gspread_formatting.Border("SOLID"), 
                right=gspread_formatting.Border("SOLID")
                ),
            horizontalAlignment='CENTER',
            verticalAlignment='MIDDLE',
            )
        
        gspread_formatting.format_cell_range(sheet, f"{UpdateLogtime_ColAlphabet}{updatelog_Lastest_Low}:{UpdateLog_ColAlphabet}{updatelog_Lastest_Low}", fmt)
        gspread_formatting.set_row_height(sheet, str(updatelog_Lastest_Low), 30)