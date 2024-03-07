# BoothListManager

![Static Badge](https://img.shields.io/badge/%ED%98%84%EC%9E%AC_%EC%A0%81%EC%9A%A9%EB%90%9C_%EB%8F%99%EC%9D%B8_%ED%96%89%EC%82%AC-%EC%97%86%EC%9D%8C-yellow)


그저 개발자가 필요해서 만든 동인 행사에 사용되는 부스 목록 관리 프로그램입니다.


## 기능

- 부스 추가, 들어가야 하는 열의 위치는 주변 셀들을 분석하여 자동으로 배치
- 부스를 추가하는 과정에서, 업데이트 로그 자동 수정 및 지도 자동 링크 (단, 지도가 있는 시트를 따로 지정해야 합니다.)
- 부스 번호 또는 부스 이름으로 시트에서 부스를 검색
- 이미 있는 부스에 대한 인포 및 선입금 링크 추가 및 수정 기능
- 수동 관리를 위한 인포 및 선입금 링크 관련 함수 자동 생성 및 클립보드 복사

## 라이브러리

이 프로그램을 사용하기 위해 필요한 라이브러리입니다.

| 라이브러리 | 설명 |
| ------ | ------ |
| gspread | 구글 스프레드시트와 파이썬을 연동하기 위한 라이브러리 |
| gspread.formatting | 스프레드시트에서 셀의 서식을 쉽게 지정하기 위한 라이브러리 |
| clipboard | 클립보드 관련 라이브러리 |

## 코드 파일

| 파일 | 파일에 있는 기능 |
| ------ | ------ |
| [Window.py][Windowfile] | 창 관련 기능 및 대부분의 기능들이 포함된 파일 |
| [LinkCollector.py][LinkCollectorfile] | 시트에서 하이퍼링크가 지정된 셀들만을 선택하여 정리하도록 하는 기능을 포함하는 파일 |
| [UpdateLogger.py][UpdateLoggerfile] | 업데이트 기록 시트에 업데이트 내용을 자동으로 추가하는 기능을 포함하는 파일 |


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
