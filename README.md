# BoothListManager

![Static Badge](https://img.shields.io/badge/%ED%98%84%EC%9E%AC_%EC%A0%81%EC%9A%A9%EB%90%9C_%EB%8F%99%EC%9D%B8_%ED%96%89%EC%82%AC-%EC%97%86%EC%9D%8C-yellow)


그저 개발자가 필요해서 만든 동인 행사에 사용되는 부스 목록 관리 프로그램입니다.


## 기능

### 부스 추가
- 부스 추가 시에 추가 열을 자동 분석하여 추가
- 부스를 추가하는 과정에서, 업데이트 로그 자동 수정 및 지도 자동 링크 (단, 지도가 있는 시트를 따로 지정해야 합니다.)
### 부스 검색
- 부스 번호 또는 부스 이름으로 시트에서 부스를 검색
### 부스 정보 수정
- 이미 있는 부스에 대한 인포 및 선입금 링크 추가 및 수정 기능
- 수동 관리를 위한 인포 및 선입금 링크 관련 함수 자동 생성 및 클립보드 복사

<br/>

## 라이브러리

이 프로그램을 사용하기 위해 필요한 라이브러리입니다.
긱 라이브러리는 다음과 같은 명령어로 설치할 수 있습니다.

```sh
pip install [라이브러리 이름]
```

| 라이브러리 | 설명 |
| ------ | ------ |
| gspread | 구글 스프레드시트와 파이썬을 연동하기 위한 라이브러리 |
| gspread.formatting | 스프레드시트에서 셀의 서식을 쉽게 지정하기 위한 라이브러리 |
| clipboard | 클립보드 관련 라이브러리 |

<br/>

## 코드 파일

| 파일 | 파일에 있는 기능 |
| ------ | ------ |
| [Window.py][Windowfile] | 창 관련 기능 및 대부분의 기능들이 포함된 파일 |
| [LinkCollector.py][LinkCollectorfile] | 시트에서 하이퍼링크가 지정된 셀들만을 선택하여 정리하도록 하는 기능을 포함하는 파일 |
| [UpdateLogger.py][UpdateLoggerfile] | 업데이트 기록 시트에 업데이트 내용을 자동으로 추가하는 기능을 포함하는 파일 |

<br/>

## 테스트
기능들은 행사를 기준으로 각 시기 별로 나눕니다. 각 시기들의 설명을 다음과 같습니다.

| 시기 | 설명 |
| ------ | ------- |
| 1 | 부스 번호 및 지도가 추가되지 않은 상태에서 선입금 및 인포가 추가되는 시기 |
| 1.5 | 공식에서 부스 번호 및 지도 공개, 이미 등록된 부스들의 부스 번호를 갱신 및 정렬하는 시기 |
| 2 | 부스 번호 및 지도가 공개된 이후에 새로 발견한 부스들을 추가하는 시기 |
| 3 | 공식 행사 종료 후 통판이 추가되는 시기 |
| 상시 | 검색, 정보 수정 등 시기와 관계없이 기능들이 상시에 속함 |

### 시기 1
<table>
    <thead>
        <tr>
            <th>테스트 여부</th>
            <th>테스트 기능</th>
            <th>세부 확인 기능</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>&#9744;</td>
            <td>부스 번호 없이 인포 또는 선입금 링크 혼자 추가하기</td>
            <td rowspan=2>
                <ul>
                    <li>'Add_new_BoothData' 함수에서 'GetRecommandLocation' 함수가 적절한 번호의 반환 </li>
                    <li>'Add_new_BoothData' 함수에서 Hyperlink 함수와 Textjoin 함수의 적용</li>
                    <li>'UpdateLogger.AddUpdateLog()' 함수에서 부스 번호가 아닌 부스 이름으로의 로그 추가</li>
                </ul>
            </td>
        </tr>
        <tr>
            <td>&#9744;</td>
            <td>부스 번호 없이 한 부스가 인포와 선입금 링크를 동시에 가진 채로 추가하기</td>
        </tr>
    </tbody>
</table>

### 시기 1.5
<table>
    <thead>
        <tr>
            <th>테스트 여부</th>
            <th>테스트 기능</th>
            <th>확인 요소</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>&#9744;</td>
            <td>부스 번호 전체 링크하기</td>
            <td>
                <ul>
                    <li>부스 번호가 지도에 정확하게 링크되었는가</li>
                </ul>
            </td>
        </tr>
        <tr>
            <td>&#9744;</td>
            <td>부스 번호 입력 후 전체 정렬하기 (추가 검토)</td>
            <td/>
        </tr>
    </tbody>
</table>

### 시기 2, 3
<table>
    <thead>
        <tr>
            <th>테스트 여부</th>
            <th>테스트 기능</th>
            <th>확인 요소</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>&#9744;</td>
            <td>부스 번호와 함께 인포 또는 선입금 링크 혼자 추가하기</td>
            <td rowspan=2>
                <ul>
                    <li>정확한 칸에 추가되었는가</li>
                    <li>지도에 정확하게 링크되었는가</li>
                    <li>업데이트 로그가 잘 생성되고, 업데이트 시간이 잘 갱신되었는가</li>
                </ul>
            </td>
        </tr>
        <tr>
            <td>&#9744;</td>
            <td>부스 번호와 함께 한 부스가 인포와 선입금 링크를 동시에 가진 채로 추가하기</td>
        </tr>
    </tbody>
</table>

### 상시
<table>
    <thead>
        <tr>
            <th>테스트 여부</th>
            <th>테스트 기능</th>
            <th>확인 요소</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td>&#9744;</td>
            <td>부스 번호 또는 부스 이름으로 부스의 인포 및 선입금 링크 목록 검색</td>
            <td>
                <ul>
                    <li>부스를 정확히 찾아낼 수 있는가</li>
                    <li>목록을 생략없이 출력할 수 있는가</li>
                </ul>
            </td>
        </tr>
        <tr>
            <td>&#9744;</td>
            <td>인포 또는 선입금 링크 목록에서 하나를 선택해 수정하기</td>
            <td>
                <ul>
                    <li>수정된 인포 또는 선입금 링크를 시트에 정확히 반영하는가</li>
                    <li>반영된 목록을 다시 업데이트하여 표기할 수 있는가 (검토 중)</li>
                </ul>
            </td>
        </tr>
    </tbody>
</table>

<br/>

## 라이선스

[MIT 라이선스][mit]

[//]: # (These are reference links used in the body of this note and get stripped out when the markdown processor does its job. There is no need to format nicely because it shouldn't be seen. Thanks SO - http://stackoverflow.com/questions/4823468/store-comments-in-markdown-syntax)

   [dill]: <https://github.com/joemccann/dillinger>
   [git-repo-url]: <https://github.com/joemccann/dillinger.git>
   [john gruber]: <http://daringfireball.net>
   [df1]: <http://daringfireball.net/projects/markdown/>
   [markdown-it]: <https://github.com/markdown-it/markdown-it>
   [Ace Editor]: <http://ace.ajax.org>
   [node.js]: <http://nodejs.org>
   [Twitter Bootstrap]: <http://twitter.github.com/bootstrap/>
   [jQuery]: <http://jquery.com>
   [@tjholowaychuk]: <http://twitter.com/tjholowaychuk>
   [express]: <http://expressjs.com>
   [AngularJS]: <http://angularjs.org>
   [Gulp]: <http://gulpjs.com>
   [mit]: <https://markdown-ui.mit-license.org/>

   [Windowfile]: <https://github.com/MinePacu/BoothListManager/blob/master/BoothListManager/Window.py>
   [LinkCollectorfile]: <https://github.com/MinePacu/BoothListManager/blob/master/BoothListManager/LinkCollector.py>
   [UpdateLoggerfile]: <https://github.com/MinePacu/BoothListManager/blob/master/BoothListManager/UpdateLogger.py>
   [PlOd]: <https://github.com/joemccann/dillinger/tree/master/plugins/onedrive/README.md>
   [PlMe]: <https://github.com/joemccann/dillinger/tree/master/plugins/medium/README.md>
   [PlGa]: <https://github.com/RahulHP/dillinger/blob/master/plugins/googleanalytics/README.md>
