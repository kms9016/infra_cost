pacman::p_load(tidyverse, lubridate, reshape, XLConnect, googlesheets4)


############ 0. 초기데이터 삭제 ###############

# 기존 시트 데이터 삭제
gs4_deauth()
wb = "https://docs.google.com/spreadsheets/d/1tlDK_0gAxF_mAIL6ai0Zr9GLqh4fXy_cWvOaR9EAIww/edit?usp=sharing"
gs4_auth(email = "jason.gang@kakaoenterprise.com")

range_delete(wb, range="Colo_Sum!A5:K", shift = "up")
range_delete(wb, range="Colo_Info!A4:BG", shift = "up")
range_delete(wb, range="Colo_Code!A4:H", shift = "up")
range_delete(wb, range="Colo_Qty!A4:AC", shift = "up")
range_delete(wb, range="Colo_Ratio!A4:AE", shift = "up")
range_delete(wb, range="Colo_Price!A4:AD", shift = "up")
range_delete(wb, range="Colo_MPrice!A4:ED", shift = "up")
range_delete(wb, range="Ref_Colo!A4:H", shift = "up")
range_delete(wb, range="단가정보!K5:M", shift = "up")



############ 1_1. Colo_Info 작업 ###############

# Daniel작성한 상면관리 파일에서 데이터 불러오기
gs4_deauth()
wa <- "https://docs.google.com/spreadsheets/d/1cUj_emuhBuW4mwr-ojEo3Bdqi5wqDbBEAEvtFajMtfo/edit#gid=1241473732"
gs4_auth(email = "jason.gang@kakaoenterprise.com")
Ref_Colo <- range_read(wa, range="백데이터!A1:V")
Ref_Colo <- mutate(Ref_Colo,Ref_ID=paste(Ref_Colo$`IDC 상면 Code`,Ref_Colo$`Zone 구분`,sep=""))

# Colo_Info로 Ref_Colo 저장
range_write(wb, data = Ref_Colo, range="Colo_Info!A5")




############ 1_2. 계약정보 작업 ###############

colname <- c("계약시작일","계약종료일","기간","상면료","전기료","랙비용","기타","합계")
colname <- as.data.frame(colname)
index <- c(1:nrow(colname))
colname <- cbind(index,colname)

# Zone_Info를 pivot형태로 전환
colname <- pivot_wider(colname, names_from = 'index', values_from = 'colname') 
range_write(wb, data = colname, range="Colo_Info!X3")

# 함수 문자열을 str에 저장하고 C5셀에 붙여넣기
str <- '=if($K6="O",datevalue("2022. 1. 1"),"")'
range_write(wb, data = as.data.frame(str), range="Colo_Info!X5")

str <- '=if($K6="O",datevalue(today()),"")'
range_write(wb, data = as.data.frame(str), range="Colo_Info!Y5")

str <- '=if($K6="O",DATEDIF(X6,Y6,"M")+1,"")'
range_write(wb, data = as.data.frame(str), range="Colo_Info!Z5")

# 함수 문자열을 str에 저장하고 C5셀에 붙여넣기
str <- '=if($K6="O",vlookup(textjoin("",,$B6,$C6,$H6),Ref_Colo!$D$5:$H$8,match(AA$5,Ref_Colo!$E$5:$H$5,0)+1,false),0)'
range_write(wb, data = as.data.frame(str), range="Colo_Info!AA5")

# 함수 문자열을 str에 저장 하고 AE5셀에 붙여넣기
str <- '=sum(AA6,AB6,AC6,AD6)'
range_write(wb, data = as.data.frame(str), range="Colo_Info!AE5")

# Colo_Info에서 X4~AD4범위를 가져와서 X5셀에 붙여넣기
rng <- range_read(wb, range="Colo_Info!X4:AE4")
range_write(wb, data = rng, range="Colo_Info!X5")

#불필요한 3, 4행삭제
range_delete(wb, range="Colo_Info!X3:3", shift = "left")
range_delete(wb, range="Colo_Info!X4:4", shift = "left")





############ 1_3. Zone 구분 작업 ###############

# Zone 구분 컬럼 데이터를 읽어와서 중복제거
Zone_Info <- range_read(wb, range="Colo_Info!F5:F") %>% unique()

# Zone 수량만큼 인덱스번호 생성
index <- c(1:nrow(Zone_Info))

# index번호와 Zone_Info 병합
Zone_Info <- cbind(index,Zone_Info)

# Zone_Info를 pivot형태로 전환
Zone_Info <- pivot_wider(Zone_Info, names_from = 'index', values_from = 'Zone 구분') 

# Zone_Info를 AE3 셀에 붙여넣기
range_write(wb, data = Zone_Info, range="Colo_Info!AF3")

# 함수 문자열을 str에 저장
str <- '=arrayformula(if($A6:$A<>"",if($F6:$F=AF$5,"사용","x"),""))'

# str 함수 문자열을 데이터프레임 형태로 X5셀에 붙여넣기
range_write(wb, data = as.data.frame(str), range="Colo_Info!AF5")

# AE4셀 부분 행전체를 복사
rng <- range_read(wb, range="Colo_Info!AF4:4")

# AE5셀에 붙여넣기
range_write(wb, data = rng, range="Colo_Info!AF5")

#불필요한 3행, 4행삭제
range_delete(wb, range="Colo_Info!AF3:3", shift = "left")
range_delete(wb, range="Colo_Info!AF4:4", shift = "left")





############ 2. Colo_Code 작업 ###############
rng <- range_read(wb, range="Colo_Info!B5:H") %>% unique()
range_write(wb, data = rng, range="Colo_Code!A5")

# 함수 문자열을 str에 저장 H5셀에 붙여넣기
Qty <- '=countif(Colo_Info!$W6:$W,textjoin("",,D6,E6))'
range_write(wb, data = as.data.frame(Qty), range="Colo_Code!H5")






############ 3. Colo_Qty 작업 ###############
Ref_Colo <- range_read(wb, range="Colo_Code!A5:H")
Ref_colo <- Ref_Colo[,c(4,5,8)]
Ref_colo <- pivot_wider(Ref_colo, names_from = 'Zone 구분', values_from = 'Qty') 
Ref_colo[is.na(Ref_colo)]<-0
range_write(wb, data = Ref_colo, range="Colo_Qty!A5")





############ 4. Colo_Ratio 작업 ###############

# Colo_Info에서 1열, 5열만 가져와서 Colo_Qty에 붙여넣기
rng <- range_read(wb, range="Colo_Info!A5:E")
rng <- rng[,c(1,5)]
range_write(wb, data = rng, range="Colo_Ratio!A5")

# 함수 문자열을 str에 저장하고 C5셀에 붙여넣기
합계 <- '=sum(D6:AE6)'
range_write(wb, data = as.data.frame(합계), range="Colo_Ratio!C5")

# Colo_Info에서 AE5~BF5범위를 가져와서 Colo_Qty D4셀에 붙여넣기
rng <- range_read(wb, range="Colo_Info!AF5:BG5")
range_write(wb, data = rng, range="Colo_Ratio!D4")

# 함수 문자열을 str에 저장 D5셀에 붙여넣기
str <- '=if(Colo_Info!AF6="사용",index(Colo_Qty!$B$6:$AC$10,match($B6,Colo_Qty!$A$6:$A$10,0),match(D$5,Colo_Qty!$B$5:$AC$5,0)),0)'
range_write(wb, data = as.data.frame(str), range="Colo_Ratio!D5")

# Colo_Qty의 4행을 가져와서 C5셀에 붙여넣고 기존의 4행은 삭제 
rng <- range_read(wb, range="Colo_Ratio!D4:4")
range_write(wb, data = rng, range="Colo_Ratio!D5")

#불필요한 4행삭제
range_delete(wb, range="Colo_Ratio!D4:4", shift = "left")




############ 5. Colo_Price 작업 ###############
rng <- range_read(wb, range="Colo_Info!A5:E")
rng <- rng[,c(1,5)]
range_write(wb, data = rng, range="Colo_Price!A5")

# Colo_Info에서 AE5~BF5범위를 가져와서 Colo_Qty C4셀에 붙여넣기
rng <- range_read(wb, range="Colo_Info!AF5:BG5")
range_write(wb, data = rng, range="Colo_Price!C4")

# 함수 문자열을 str에 저장 C5셀에 저장
str <- '=round(Colo_Info!$AE6*(Colo_Ratio!D6/Colo_Ratio!$C6),0)'
range_write(wb, data = as.data.frame(str), range="Colo_Price!C5")

# Colo_Info에서 C4행부터 행전체를 가져와서 Colo_Qty C4셀에 붙여넣기
rng <- range_read(wb, range="Colo_Price!C4:4")
range_write(wb, data = rng, range="Colo_Price!C5")

#불필요한 3, 4행삭제
range_delete(wb, range="Colo_Price!C4:4", shift = "left")





############ 6. Colo_Monthly_Price 작업 ###############
rng <- range_read(wb, range="Colo_Info!A5:E")
rng <- rng[,c(1,5)]
range_write(wb, data = rng, range="Colo_MPrice!A5")

# 함수 문자열을 str에 저장
str <- '=if(AND(C$5>=Colo_Info!$X6,C$5<=Colo_Info!$Y6),Colo_Info!$AE6,0)'

# str 함수 문자열을 데이터프레임 형태로 C5셀에 붙여넣기
range_write(wb, data = as.data.frame(str), range="Colo_MPrice!C5")

# Colo_Info에서 X5~AY5범위를 가져와서 Colo_Qty C4셀에 붙여넣기 --> 시트에서 속성을 날짜로 바꾼다
rng <- range_read(wb, range="Colo_MPrice!C1:ED2")
range_write(wb, data = rng, range="Colo_MPrice!C4")


############ 7. Ref_Colo 작업 ###############
rng <- range_read(wb, range="Colo_Info!A5:H") 
rng <- rng[,c(2,3,8)] %>% unique()
rng <- mutate(rng,Ref_ID=paste(rng$`IDC Code`,rng$`IDC 구분`,rng$전력구분,sep=""))
range_write(wb, data = rng, range="Ref_Colo!A5")

colname <- c("상면료","전기료","랙비용","기타")
colname <- as.data.frame(colname)
index <- c(1:nrow(colname))
colname <- cbind(index,colname)

# Zone_Info를 pivot형태로 전환
colname <- pivot_wider(colname, names_from = 'index', values_from = 'colname') 
range_write(wb, data = colname, range="Ref_Colo!E3")

# 함수 문자열을 str에 저장 E5셀에 저장
str <- '=vlookup($D6,단가정보!$D$6:$H$8,match(E$5,단가정보!$E$5:$H$5,0)+1,false)'
range_write(wb, data = as.data.frame(str), range="Ref_Colo!E5")

# Colo_Info에서 X5~AY5범위를 가져와서 Colo_Qty C4셀에 붙여넣기
rng <- range_read(wb, range="Ref_Colo!E4:H4")
range_write(wb, data = rng, range="Ref_Colo!E5")

#불필요한 3, 4행삭제
range_delete(wb, range="Ref_Colo!E3:H4", shift = "left")






############ 8. 단가정보 전기료 산정 작업 ###############
rng <- range_read(wb, range="Colo_Info!B5:L")
rng <- rng[,c(2,11)]

# 평규사용비용 구하기
평균사용 <- tapply(rng$`실 전력량`, INDEX=list(rng$`IDC 구분`), FUN=mean, na.rm=TRUE)
평균사용 <- as.data.frame(평균사용)

IDC_Code <- c("SEL11","SEL21")
IDC_Code <- as.data.frame(IDC_Code)
index <- c(1:nrow(IDC_Code))
평균사용 <- cbind(IDC_Code,평균사용)

range_write(wb, data = as.data.frame(평균사용), range="단가정보!K5")


# 전기료 합계 구하기
합계 <- '=round(L6*N6*O6*P6,0)'
range_write(wb, data = as.data.frame(합계), range="단가정보!M5")







############ 9. 상면비용 최종정리 ###############

# Colo_Info에서 필요한 컬럼 가져와서 고유ID 생성
Ref <- range_read(wb, range="Colo_Info!B5:K")
Ref <- mutate(Ref, ID=paste(Ref$`IDC Code`, Ref$`IDC 상면 Code`, Ref$`Zone 구분`,Ref$`사용중인 상면인가?`))

# ID별 수량 취합
x <- table(Ref$ID)
x <- as.data.frame(x)

# 수량정보 추가
Ref <- left_join(Ref, x, by=c('ID'='Var1'))
Ref <- select (Ref, 'IDC Code', 'IDC 상면 Code', 'Zone 구분', '전력구분', '사용중인 상면인가?', 랙수량='Freq', 'ID') %>% 
  unique()


# Colo_Info에서 필요한 컬럼 가져와서 각 비용 컬럼만 추출
Ref2 <- range_read(wb, range="Colo_Info!B5:AE")
Ref2 <- mutate(Ref2, ID=paste(Ref2$`IDC Code`, Ref2$`IDC 상면 Code`, Ref2$`Zone 구분`,Ref2$`사용중인 상면인가?`))
Ref2 <- select (Ref2, 'ID', '상면료', '전기료', '랙비용', '기타', '합계') %>% 
  unique()

# Ref + Ref2 join
Ref <- left_join(Ref, Ref2, by=c('ID'))
Ref <- select(Ref, -ID)

# Colo_Sum에 상면금액만 정리
range_write(wb, data = Ref, range="Colo_Sum!A5")


############## The End ################
## 모든 스크립트가 종료 후
## "단가정보"시트를 제일먼저 수식 실행하면됨
## "Colo_MPrice" 시트C5행 전체는 반드시 날짜 형식으로 변경