# 국세청 API를 이용한 사업자 상태 조회
[공공데이터포털 국세청 API 공식 메뉴얼](https://www.data.go.kr/data/15081808/openapi.do)

## Python 설치
최신버전의 Python은 [정식 사이트](https://www.python.org/downloads/)에서 다운로드 후 설치해주세요

## 국세청 API 서비스키 발급
프로그램에 입력할 서비스키는 [공공데이터포털](https://www.data.go.kr/data/15081808/openapi.do)에서 "활용신청"을 클릭해 받아주세요
![활용신청 버튼](./assets/obtain_key.png)

## 사업자등록번호 엑셀 준비
프로그램에서 읽을 엑셀 파일은 이렇게 준비해 주세요:
1. 사업자등록번호를 입력할 열 첫번째 행에 "사업자등록번호" 입력
2. 그 다음 열부터 숫자로 된 사업자등록번호 입력 ("-" 제외)

예시:
![사업자등록번호 엑셀](./assets/input_excel.png)