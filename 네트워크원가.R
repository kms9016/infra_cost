pacman::p_load(tidyverse, lubridate, reshape, XLConnect, googlesheets4)


############ 0. 초기데이터 삭제 ###############

# NW 원가데이터
gs4_deauth()
wa <- "https://docs.google.com/spreadsheets/d/1rsNdjz-Erbz3juL0ZK47yabmu8fknTgRz460t7oNuTk/edit?usp=sharing"
gs4_auth(email = "jason.gang@kakaoenterprise.com")

# 공통비용 계산 시트
gs4_deauth()
wb = "https://docs.google.com/spreadsheets/d/1XpMsDCxUfFPMgb-iUqtrqyBYyj_wqSwpXIt2frhh7F8/edit?usp=sharing"
gs4_auth(email = "jason.gang@kakaoenterprise.com")

# 상면비용 계산 시트
gs4_deauth()
wc = "https://docs.google.com/spreadsheets/d/1tlDK_0gAxF_mAIL6ai0Zr9GLqh4fXy_cWvOaR9EAIww/edit?usp=sharing"
gs4_auth(email = "jason.gang@kakaoenterprise.com")

range_delete(wb, range="Nw_Info!A5:BG", shift = "up")
range_delete(wb, range="Nw_Price!A5:AD", shift = "up")
range_delete(wb, range="Colo_Code!A5:H", shift = "up")
range_delete(wb, range="Colo_Qty!A5:AC", shift = "up")
range_delete(wb, range="Colo_Ratio!A5:AE", shift = "up")
range_delete(wb, range="Nw_MPrice!A4:ED", shift = "up")

############ 1. NW_Info 작업 ###############


Ref_Nw <- range_read(wa, range="Nw_Info!A1:V")
Ref_Nw <- mutate(Ref_Nw,Ref_ID=paste(Ref_Nw$`IDC 상면 Code`,Ref_Nw$`Zone 구분`,sep=""))

# Nw_Info 시트로 Ref_Nw 저장
range_write(wb, data = Ref_Nw, range="Nw_Info!A5")


############ 2. 계약정보 작업 ###############

# NW데이터에서 계약기간, 구매가격 정보 가져옴
Ref_Cost <- range_read(wa, range="Nw_Info!W1:AB")
range_write(wb, data = Ref_Cost, range="Nw_Info!X5")

# Colo 시트와 컬럼을 맞추기위해 컬럼추가
colname <- c("합계","BL12")
colname <- as.data.frame(colname)
index <- c(1:nrow(colname))
colname <- cbind(index,colname)

# Zone_Info를 pivot형태로 전환
colname <- pivot_wider(colname, names_from = 'index', values_from = 'colname') 
range_write(wb, data = colname, range="Nw_Info!AD3")

# 함수 문자열을 str에 저장 하고 AE5셀에 붙여넣기
str <- '=(AA6+(AA6*AB6*2))/Z6'
range_write(wb, data = as.data.frame(str), range="Nw_Info!AD5")

# Nw_Info에서 X4~AD4범위를 가져와서 X5셀에 붙여넣기
rng <- range_read(wb, range="Nw_Info!AD4:AE4")
range_write(wb, data = rng, range="Nw_Info!AD5")

#불필요한 3, 4행삭제
range_delete(wb, range="Nw_Info!AD3:3", shift = "left")
range_delete(wb, range="Nw_Info!AD4:4", shift = "left")


############ 3. Zone 구분 작업 ###############

# Zone 구분 컬럼 데이터를 Colo파일에서 읽어오고 중복제거
Zone_Info <- range_read(wc, range="Colo_Info!F5:F") %>% unique()

# Zone 수량만큼 인덱스번호 생성
index <- c(1:nrow(Zone_Info))
Zone_Info <- cbind(index,Zone_Info)
Zone_Info <- pivot_wider(Zone_Info, names_from = 'index', values_from = 'Zone 구분') 

# Zone_Info를 AE3 셀에 붙여넣기
range_write(wb, data = Zone_Info, range="Nw_Info!AF3")

# 함수 문자열을 str에 저장
str <- '=arrayformula(if($A6:$A<>"",if($F6:$F="공통","사용","x"),""))'
range_write(wb, data = as.data.frame(str), range="Nw_Info!AF5")

# AE4셀 부분 행전체를 복사
rng <- range_read(wb, range="Nw_Info!AF4:4")
range_write(wb, data = rng, range="Nw_Info!AF5")

#불필요한 3행, 4행삭제
range_delete(wb, range="Nw_Info!AF3:3", shift = "left")
range_delete(wb, range="Nw_Info!AF4:4", shift = "left")


############ 4. Colo_Code 작업 ###############

#Colo파일의 Colo_Code시트를 그대로 복사해옴
rng <- range_read(wc, range="Colo_Code!A5:H")
range_write(wb, data = rng, range="Colo_Code!A5")



############ 5. Colo_Qty 작업 ###############

# 공통비용은 상면과 달리 IDC코드를 기준으로 작업할 것
Ref_Colo <- range_read(wb, range="Colo_Code!A5:H")
Ref_Colo <- Ref_Colo[,c(1,5,8)]

# IDC코드 존코드 기준을 랙 수량정보 정리
Ref_Colo <- aggregate(Ref_Colo$Qty, by=list(IDC_Code=Ref_Colo$`IDC Code`, IDC_Zone=Ref_Colo$`Zone 구분`), sum, na.rm=T)

# 랙 수량정보를 unpivot한 후 Colo_Qty시트에 추가
Ref_Colo <- pivot_wider(Ref_Colo, names_from = 'IDC_Zone', values_from = 'x') 
Ref_Colo[is.na(Ref_Colo)]<-0
range_write(wb, data = Ref_Colo, range="Colo_Qty!A5")


############ 5. Colo_Ratio 작업 ###############

# Nw_Info시트에서 컬럼 정보를 가져와서 Colo_Ratio에 복사
rng <- range_read(wb, range="Nw_Info!A5:E")
rng <- rng[,c(1,2)]
range_write(wb, data = rng, range="Colo_Ratio!A5")

# Nw_Info에서 AF5~BG5범위를 가져와서 Colo_Qty D4셀에 붙여넣기
rng <- range_read(wb, range="Nw_Info!AF5:BG5")
range_write(wb, data = rng, range="Colo_Ratio!D4")

합계 <- '=sum(D6:AE6)'
range_write(wb, data = as.data.frame(합계), range="Colo_Ratio!C5")

str <- '=if(Nw_Info!AF6="사용",index(Colo_Qty!$B$6:$AC$10,match($B6,Colo_Qty!$A$6:$A$10,0),match(D$5,Colo_Qty!$B$5:$AC$5,0)),0)'
range_write(wb, data = as.data.frame(str), range="Colo_Ratio!D5")

# Colo_Info에서 C4행부터 행전체를 가져와서 Colo_Qty C4셀에 붙여넣기
rng <- range_read(wb, range="Colo_Ratio!D4:4")
range_write(wb, data = rng, range="Colo_Ratio!D5")

#불필요한 3, 4행삭제
range_delete(wb, range="Colo_Ratio!D4:4", shift = "left")



############ 6. Nw_Price 작업 ###############
rng <- range_read(wb, range="Nw_Info!A5:E")
rng <- rng[,c(1,2)]
range_write(wb, data = rng, range="Nw_Price!A5")

# Colo_Info에서 AE5~BF5범위를 가져와서 Colo_Qty C4셀에 붙여넣기
rng <- range_read(wb, range="Nw_Info!AF5:BG5")
range_write(wb, data = rng, range="Nw_Price!C4")

# 함수 문자열을 str에 저장 C5셀에 저장
str <- '=round(Nw_Info!$AD6*(Colo_Ratio!D6/Colo_Ratio!$C6),0)'
range_write(wb, data = as.data.frame(str), range="Nw_Price!C5")

# Colo_Info에서 C4행부터 행전체를 가져와서 Colo_Qty C4셀에 붙여넣기
rng <- range_read(wb, range="Nw_Price!C4:4")
range_write(wb, data = rng, range="Nw_Price!C5")

#불필요한 3, 4행삭제
range_delete(wb, range="Nw_Price!C4:4", shift = "left")


############ 7. Nw_Monthly_Price 작업 ###############
rng <- range_read(wb, range="Nw_Info!A5:E")
rng <- rng[,c(1,2)]
range_write(wb, data = rng, range="Nw_MPrice!A5")

# 함수 문자열을 str에 저장
str <- '=if(AND(C$5>=Nw_Info!$X6,C$5<=Nw_Info!$Y6),Nw_Info!$AD6,0)'

# str 함수 문자열을 데이터프레임 형태로 C5셀에 붙여넣기
range_write(wb, data = as.data.frame(str), range="Nw_MPrice!C5")

# Nw_Info에서 X5~AY5범위를 가져와서 Colo_Qty C4셀에 붙여넣기 --> 시트에서 속성을 날짜로 바꾼다
rng <- range_read(wb, range="Nw_MPrice!C1:ED2")
range_write(wb, data = rng, range="Nw_MPrice!C4")

