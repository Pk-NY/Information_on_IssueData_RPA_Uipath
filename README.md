### [의약품 안전나라] ###

의약품 안전나라 사이트 행정처분정보에 접속하여 취합한 자료 중 이슈 내역(거짓,허위,다르게 제조,이물질,부적합,미달 등)을 정리하여 담당자에게 메일 발송

nedrug.mfds.go.kr


**[사전작업]**

Input 폴더 : 기본엑셀양식(3개의 시트 : 현재연월, 이슈사항, 당월이슈), VBA 코드 파일

Config.xlsx : output 폴더 경로, 사이트URL, 메일계정, 메일제목, 메일내용, 결과파일 이름, 엑셀템플릿경로, 필터키워드, vba 코드파일 경로


**[개인작업파일 Business 폴더]**

   CheckOutputFolder - output폴더확인 후 생성
  
   CopyTemplate -  템플릿파일복사후 결과파일생성 

   WriteExcel - 전체 데이터에서 이슈키워드로 필터링 한 후에 다시 당월데이터만 필터링 해서 결과파일에 차례대로 작성
   
   SendEmail - 당월이슈데이터가 있을 경우(당월이슈 DT를 표로 변화하여 메일 내용에 담음)와 없을 경우(메일 내용에 표 없음)  메일 내용을 다르게 해서 
              메일 발송을 처리
   
   
   DataTableToHtmlText - 데이터테이블을 HTML로 변환  
  

### [작업과정] ###

1. **INITIALIZE PROCESS**
   
    output폴더확인, 템플릿파일복사결과파일 생성, 사이트 접속 전월부터 당월까지의 데이터 수집

    업체별 상세페이지 URL -> dt_TransactionData에 담기
   
2. **GET TRANSACTION DATA**

    Transaction Item 타입 :  DataRow 업체별 URL

3. **PROCESS TRANSACTION**

   행정처분 받은 업체별 상세페이지 URL에 접속하여 업체명,처분일자,위반내용,처분사항 정보를 추출, DT에 담기

    
4. **END PROCESS**
   
    전체 데이터에서 이슈키워드로 필터링 한 후에 다시 당월데이터만 필터링 해서 결과파일 차례대로 작성

    메일에 결과파일 첨부해서 발송
   

