import requests
import pandas as pd
from styleframe import StyleFrame


def read_service_key():
    try:
        serviceKeyFile = open("./serviceKey.txt", "r")
    except:
        return ""

    serviceKey = serviceKeyFile.read()
    serviceKeyFile.close()

    return serviceKey


def save_service_key(key):
    serviceKeyFile = open("./serviceKey.txt", "w")
    serviceKeyFile.write(key)
    serviceKeyFile.close()


def check_CRN_status(service_key, input_dir, output_dir):
    inputData = pd.read_excel(r"" + input_dir)
    inputDF = pd.DataFrame(inputData)

    try:
        CRNList = inputDF["사업자등록번호"].values.tolist()
    except:
        return '사업자등록번호가 적힌 엑셀 열 첫번째 행에 "사업자등록번호" 제목 입력 후 다시 시도해주세요'

    CRNStringList = list(map(str, CRNList))

    url = "http://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=" + service_key

    maxChunkSize = 100
    CRNChunks = [CRNStringList[i:i + maxChunkSize] for i in range(0, len(CRNStringList), maxChunkSize)]

    resultCRNList = []
    resultCRNStatus = []
    resultCRNTaxType = []
    resultCRNEndDate = []

    for CRNChunk in CRNChunks:
        body = {"b_no": CRNChunk}
        try:
            response = requests.post(url, json=body)
        except requests.exceptions.RequestException:
            return "인터넷 연결을 확인하세요"

        if response.status_code != 200:
            if response.status_code == 400:
                return "유효한 서비스키를 입력 후 저장하세요"
            else:
                return "국세청 서버 오류: 잠시 후 다시 시도해주세요"

        responseJSON = response.json().get("data", [])

        for entry in responseJSON:
            taxType = entry.get("tax_type", "")
            if taxType == "국세청에 등록되지 않은 사업자등록번호입니다.":
                taxType = "국세청에 등록되지 않음"

            status = entry.get("b_stt", "-")

            # 폐업일 안전 처리
            raw_end_date = entry.get("end_dt", "")
            if raw_end_date:
                try:
                    endDate = datetime.strptime(raw_end_date, "%Y%m%d").strftime("%Y-%m-%d")
                except ValueError:
                    endDate = ""
            else:
                endDate = ""

            resultCRNList.append(entry.get("b_no", ""))
            resultCRNStatus.append(status)
            resultCRNTaxType.append(taxType)
            resultCRNEndDate.append(endDate)

    output = {
        "사업자등록번호": resultCRNList,
        "납세자상태": resultCRNStatus,
        "과세유형": resultCRNTaxType,
        "폐업일": resultCRNEndDate
    }

    outputDF = pd.DataFrame(output)

    try:
        excel_writer = StyleFrame.ExcelWriter(output_dir + '/사업자상태.xlsx')
        sf = StyleFrame(outputDF)
        sf.to_excel(excel_writer=excel_writer, best_fit=[
            "사업자등록번호", "납세자상태", "과세유형", "폐업일"], row_to_add_filters=0)
        excel_writer.save()
    except Exception as e:
        return f"엑셀 저장 중 오류 발생: {e}"

    return ""
