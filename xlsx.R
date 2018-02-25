library(openxlsx)

#write data.frame to xlsx file
#parameters including data.frame, xlsx file name, and sheet name
df2xlsx<-function(df,xlsx,sheet=NULL){
  # if xlsx file exists, then load the xlsx
  # if not exists, then create the xlxs file
  if (file.exists(xlsx)){
    wb<-loadWorkbook(xlsx)
  }
  else{
    wb<-createWorkbook()
  }
  
  #if not define the sheet name,use the name of df as sheet name
  if(is.null(sheet)){
    sheet=deparse(substitute(df))
  }
  
  if(sheet %in% sheets(wb)){
    sheetorder<-sheets(wb)
    response<-readline(prompt="the worksheet already exists! you want to overwrite? y|n: ")
    if (response=="y"){
      removeWorksheet(wb,sheet)
    }
    else{
      stop("worksheet is not overwrite.")
    }
  }
  else{
    sheetorder=c(sheets(wb),sheet)
  }
  
  addWorksheet(wb,sheet)
  writeData(wb,sheet,df,rowNames=T)
  worksheetOrder(wb)<-unname(sapply(sheetorder,function(x) which(x==sheets(wb),arr.ind=T)))
  
  if (file.exists(xlsx)){
    saveWorkbook(wb,xlsx,overwrite = T)
  }
  else{
    saveWorkbook(wb,xlsx)
  }
}


#read data from xlsx file to data.frame
#parametes including xlsx file and sheet name
xlsx2df<-function(xlsx,sheet=1L){
  stopifnot(file.exists(xlsx))
  wb<-loadWorkbook(xlsx)
  sheets<-sheets(wb)
  stopifnot(is.integer(sheet) || sheet %in% sheets)
  df<-read.xlsx(wb,sheet,colNames = T,rowNames = T)
  return(df)
}


