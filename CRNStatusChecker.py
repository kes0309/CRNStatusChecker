import requests
import pandas as pd
from styleframe import StyleFrame


def checkCRNStatus(input_dir, output_dir):
    inputData = pd.read_excel(r"" + input_dir)
    inputDF = pd.DataFrame(inputData)

    CRNList = inputDF["사업자등록번호"].values.tolist()
    CRNStringList = list(map(str, CRNList))

    # 공공데이터포털 OpenAPI
    # 국세청 관리
    # Documentation: https://www.data.go.kr/data/15081808/openapi.do
    serviceKeyFile = open("./serviceKey.txt", "r")
    serviceKey = serviceKeyFile.read()
    serviceKeyFile.close()
    url = "http://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=" + serviceKey

    maxChunkSize = 100
    CRNChunks = [CRNStringList[i:i + maxChunkSize]
                 for i in range(0, len(CRNStringList), maxChunkSize)]

    resultCRNList = []
    resultCRNStatus = []
    resultCRNTaxType = []

    for CRNChunk in CRNChunks:
        print(len(CRNChunk))

        body = {"b_no": CRNChunk}
        response = requests.post(url, json=body)

        responseJSON = response.json()["data"]

        for entry in responseJSON:
            taxType = entry["tax_type"]
            if taxType == "국세청에 등록되지 않은 사업자등록번호입니다.":
                taxType = "국세청에 등록되지 않음"

            status = entry["b_stt"]

            if (status == ""):
                status = "-"

            resultCRNList.append(entry["b_no"])
            resultCRNStatus.append(status)
            resultCRNTaxType.append(taxType)

    output = {
        "사업자등록번호": resultCRNList,
        "납세자상태": resultCRNStatus,
        "과세유형": resultCRNTaxType
    }

    outputDF = pd.DataFrame(output)

    excel_writer = StyleFrame.ExcelWriter(output_dir + '/사업자상태.xlsx')
    sf = StyleFrame(outputDF)
    sf.to_excel(excel_writer=excel_writer, best_fit=[
                "사업자등록번호", "납세자상태", "과세유형"], row_to_add_filters=0)
    excel_writer.save()
