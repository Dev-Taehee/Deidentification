# Deidentification
</br>
대학원 수업 중 과제로 문서 비식별화 서비스를 만들던 코드입니다.</br>
해당 문서들에 있는 정보들은 수업시간에 제공받은 것으로 임의의 정보를 생성한 것으로 알고있습니다.</br>
해당 과제에서 다른 팀원분들께서 AI 모델링을 진행해주셨고 저는 문서를 AI 모델에 맞는 텍스트 형태로 변환하고 결과값을 문서에 재적용하는 역할을 맡았습니다.</br>
</br?
# deidentification.py
전반적인 작업 과정을 맡습니다.</br>
1. 사용자가 제공한 파일의 형태(한글, 엑셀, 워드 등)에 따라 pdf 파일로 변환합니다. </br>
<img src="/img/original_pdf.png">
2. 변환된 pdf 파일을 AI 모델에 적합한 텍스트 파일로 변환합니다. </br>
<img src="/img/pdf2txt.png">
3. 변환된 텍스트 파일을 AI 모델에 넣은 후 결과값을 받습니다. </br>
<img src="/img/result.png">
4. 결과값을 토대로 비식별화할 정보를 파악하여 PDF 문서에 검은줄을 치도록합니다. </br>
(이 때 pdf_highlighter.py가 쓰입니다.) </br>
<img src="/img/deidentification.png">
