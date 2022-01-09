pacman::p_load(tidyverse, lubridate, XLConnect, reshape)
setwd("C:/Users/jason/OneDrive - Kakao Corp/1. KEP 업무")




############ 1. Colo_Info 작업 ###############

# 1.1 엑셀 파일을 읽어옴
wb <- loadWorkbook("cost.xlsx", create=FALSE)

# 1.2 Colo_Info 시트에서 기초데이터를 저장
Colo_Info <- readWorksheet(wb, sheet = "Colo_Info", region = "A5:V")

# 1.3 상면, 존 코드를 합쳐서 식별자를 만든다
Colo_Info <- mutate(Colo_Info,Ref_ID=paste(Colo_Info$IDC.상면.Code,Colo_Info$Zone.구분,sep=""))

# 1.4 Zone_Info를 AE3 셀에 붙여넣기
writeWorksheet(wb, Colo_Info, sheet = "Colo_Info", startRow = 5, startCol = 1) 

# 1.5 계약관련 컬럼명 벡터를 만든다		
colname <- c("계약시작일","계약종료일","기간","상면료","전기료","랙비용","합계")			
colname <- as.data.frame(colname)			
index <- c(1:nrow(colname))			
colname <- cbind(index,colname)			

# 1.6 1.5의 작업을 pivot형태로 만들고 붙여넣음
colname <- pivot_wider(colname, names_from = 'index', values_from = 'colname') # Zone_Info를 pivot형태로 전환				
writeWorksheet(wb, colname, sheet = "Colo_Info", startRow = 3, startCol = 24)			

# 1.7 sum함수를 만들어 붙여넣기				
str <- '=sum(AA6,AB6,AC6)'	# 함수 문자열을 str에 저장		
writeWorksheet(wb, str, sheet = "Colo_Info", startRow = 5, startCol = 30)	# str 함수 문자열을 데이터프레임 형태로 AD5셀에 붙여넣기		

# 1.8 1.6의 컬럼을 복사하여 아래로 이동하여 붙여넣기
rng <- readWorksheet(wb, sheet = "Colo_Info", region = "X4:AD4", autofitRow = TRUE) # 컬럼명을 복사하여 AE5셀에 붙여넣기
writeWorksheet(wb, rng, sheet = "Colo_Info", startRow = 5, startCol = 24)

# 1.9 불필요한 컬럼 삭제
clearRangeFromReference(wb, reference = c("Colo_Info!X3:AD4")) # 불필요한 내용 삭제

# 1.10 상면 정보를 가져와서 중복제거
Zone_Info <- readWorksheet(wb, sheet = "Colo_Info", region = "F5:F") %>% unique()

# 1.11 1.10의 중복제건된 데이터에 대한 인덱스 생성하여 병합
index <- c(1:nrow(Zone_Info))
Zone_Info <- cbind(index,Zone_Info) # index번호와 Zone_Info 병합

# 1.12 1.11 작업을 pivot 형태로 변환하여 붙여넣기
Zone_Info <- pivot_wider(Zone_Info, names_from = 'index', values_from = 'Zone.구분') # Zone_Info를 pivot형태로 전환
writeWorksheet(wb, Zone_Info, sheet = "Colo_Info", startRow = 3, startCol = 31) # Zone_Info를 AE3 셀에 붙여넣기
writeWorksheet(wb, str, sheet = "Colo_Info", startRow = 5, startCol = 31)

# 1.12에서 작업된 컬럼을 아래로 이동하여 붙여넣기
rng <- readWorksheet(wb, sheet = "Colo_Info", region = "AE4:BF4", autofitRow = TRUE) # 컬럼명을 복사하여 AE5셀에 붙여넣기
writeWorksheet(wb, rng, sheet = "Colo_Info", startRow = 5, startCol = 31)

# 1.13 불필요한 컬럼삭제
clearRangeFromReference(wb, reference = c("Colo_Info!AE3:BF4")) # 불필요한 내용 삭제

# 1.14 저장하고 메모리 해제
saveWorkbook(wb)
rm(list=ls())
xlcFreeMemory()




############ 2. Colo_Code 작업 ###############

# 2.1 엑셀 파일을 읽어옴
wb <- loadWorkbook("cost.xlsx", create=FALSE)

# 2.2 상면관련 데이터만 가져와서 중복제거하고 붙여넣기
df <- readWorksheet(wb, sheet = "Colo_Info", region = "B5:H") %>% unique()
writeWorksheet(wb, df, sheet = "Colo_Code", startRow = 5, startCol = 1)

# 2.3 저장하고 메모리 해제
saveWorkbook(wb)
rm(list=ls())
xlcFreeMemory()
 



############ 3_1. Colo_Qty 작업 ###############
wb <- loadWorkbook("cost.xlsx", create=FALSE)
df <- readWorksheet(wb, sheet = "Colo_Info", region = "A5:E")
df <- df[,c(1,5)]
writeWorksheet(wb, df, sheet = "Colo_Qty", startRow = 5, startCol = 1)

saveWorkbook(wb)
rm(list=ls())
xlcFreeMemory()



############ 0. clear ###############
wb <- loadWorkbook("cost.xlsx", create=FALSE)
clearRangeFromReference(wb, reference = c("Colo_Info!W3:BF1000"))
clearRangeFromReference(wb, reference = c("Colo_Code!A3:G100"))
saveWorkbook(wb)
rm(list=ls())
xlcFreeMemory()
