library(readxl)


read_excel_allsheets <- function(filename, tibble = FALSE) {
  # I prefer straight data.frames
  # but if you like tidyverse tibbles (the default with read_excel)
  # then just pass tibble = TRUE
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}

mysheets_L <- read_excel_allsheets("D:/ST426 RESEARCH PROJECT/FINAL RESEARCH/ANALYSIS MATERIALS/ANALYZE 2ND PART-correlation/R CODE OCT 06/leptodata_for_cor.xlsx")
mysheets_R <- read_excel_allsheets("D:/ST426 RESEARCH PROJECT/FINAL RESEARCH/ANALYSIS MATERIALS/ANALYZE 2ND PART-correlation/R CODE OCT 06/rainfall_for_corr.xlsx")
mysheets_T<-read_excel_allsheets("D:/ST426 RESEARCH PROJECT/FINAL RESEARCH/ANALYSIS MATERIALS/ANALYZE 2ND PART-correlation/R CODE OCT 06/temp_corr.xlsx")
mysheets_H<-read_excel_allsheets("D:/ST426 RESEARCH PROJECT/FINAL RESEARCH/ANALYSIS MATERIALS/ANALYZE 2ND PART-correlation/R CODE OCT 06/Hum_corr.xlsx")


co_2010<-function(sheetL,sheetR)
  {
i<-1
j<-1
c2010<-c()
for(i in 1:ncol(mysheets_L$`2010`[1,]))
  {
  
  for(j in 1:ncol(mysheets_L$`2010`[1,])){
    
    if(i==j)
    {
      c2010[i]<-cor(sheetL[,i], sheetR[,j],  method = "spearman")
    }
    
  }
  
  
}
return(data.frame(c2010))
}



#co_2010(mysheets_L$`2010`,mysheets_R$`2010`),co_2010(mysheets_L$`2011`,mysheets_R$`2011`),co_2010(mysheets_L$`2012`,mysheets_R$`2012`),co_2010(mysheets_L$`2013`,mysheets_R$`2013`)
           #,co_2010(mysheets_L$`2014`,mysheets_R$`2014`),co_2010(mysheets_L$`2015`,mysheets_R$`2015`),co_2010(mysheets_L$`2016`,mysheets_R$`2016`),
          # co_2010(mysheets_L$`2017`,mysheets_R$`2017`),co_2010(mysheets_L$`2018`,mysheets_R$`2018`)

R1<-co_2010(mysheets_L$`2010`,mysheets_R$`2010`)
R2<-co_2010(mysheets_L$`2011`,mysheets_R$`2011`)
R3<-co_2010(mysheets_L$`2012`,mysheets_R$`2012`)
R4<-co_2010(mysheets_L$`2013`,mysheets_R$`2013`)  
R5<-co_2010(mysheets_L$`2014`,mysheets_R$`2014`)
R6<-co_2010(mysheets_L$`2015`,mysheets_R$`2015`) 
R7<-co_2010(mysheets_L$`2016`,mysheets_R$`2016`)
R8<-co_2010(mysheets_L$`2017`,mysheets_R$`2017`)
R9<-co_2010(mysheets_L$`2018`,mysheets_R$`2018`)
rainfall_corr<-data.frame(R1,R2,R3,R4,R5,R6,R7,R8,R9)



T1<-co_2010(mysheets_L$`2010`,mysheets_T$`2010`)
T2<-co_2010(mysheets_L$`2011`,mysheets_T$`2011`)
T3<-co_2010(mysheets_L$`2012`,mysheets_T$`2012`)
T4<-co_2010(mysheets_L$`2013`,mysheets_T$`2013`)  
T5<-co_2010(mysheets_L$`2014`,mysheets_T$`2014`)
T6<-co_2010(mysheets_L$`2015`,mysheets_T$`2015`) 
T7<-co_2010(mysheets_L$`2016`,mysheets_T$`2016`)
T8<-co_2010(mysheets_L$`2017`,mysheets_T$`2017`)
T9<-co_2010(mysheets_L$`2018`,mysheets_T$`2018`)
Temp_lepto<-data.frame(T1,T2,T3,T4,T5,T6,T7,T8,T9)


H1<-co_2010(mysheets_L$`2010`,mysheets_H$`2010`)
H2<-co_2010(mysheets_L$`2011`,mysheets_H$`2011`)
H3<-co_2010(mysheets_L$`2012`,mysheets_H$`2012`)
H4<-co_2010(mysheets_L$`2013`,mysheets_H$`2013`)  
H5<-co_2010(mysheets_L$`2014`,mysheets_H$`2014`)
H6<-co_2010(mysheets_L$`2015`,mysheets_H$`2015`) 
H7<-co_2010(mysheets_L$`2016`,mysheets_H$`2016`)
H8<-co_2010(mysheets_L$`2017`,mysheets_H$`2017`)
H9<-co_2010(mysheets_L$`2018`,mysheets_H$`2018`)
Hum_lepto<-data.frame(H1,H2,H3,H4,H5,H6,H7,H8,H9)

install.packages("writexl")

library(writexl)

write_xlsx(x = Hum_lepto, path = "D:/ST426 RESEARCH PROJECT/FINAL RESEARCH/ANALYSIS MATERIALS/ANALYZE 2ND PART-correlation/R CODE OCT 06/HUM_LEPTO.xlsx", col_names = FALSE)
