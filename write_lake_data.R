# import packages
library(RDCOMClient)

# define variables
file <- "C:/Users/b4ecmmkp/Desktop/2019Gren_R.xlsx"
month <- "July"
day <- "2"
rain <- 0.5
pool <- 232.29
evap <- 0.5
discharge <- 0
tail <- 163.69
gate1 <- 0
gate2 <- 4
gate3 <- 4
gate4 <- 0
gate_op1 <- "0-14-0 @ 0910"
gate_op2 <- "0-14-0 @ 0910"
gate_op3 <- "0-14-0 @ 0910"


# define the edit_lake_excels function
edit_lake_excels <- function(file, month, day, rain, pool, evap, discharge,
                             tail, gate1, gate2, gate3, gate4, gate_op1, 
                             gate_op2, gate_op3){
  # based on the lake name, open the corresponding excel file 
  xlApp <- COMCreate("Excel.Application")
  wb <- xlApp[["Workbooks"]]$Open(file)

  # based on the month, open the corresponding excel sheet 
  sheet <- wb$Worksheets(month)

  # based on the day, obtain the row numbers to edit
  days = rep(1:31, each =3)
  rows_to_read = c(5:52, 5:49)

  rows_df <- data.frame(days, rows_to_read)

  rows_to_edit_per_day <- rows_df$rows_to_read[which(rows_df$days == day)]

  # based on the day, obtain the columns to edit 
  days2 = c(rep(1:31, each = 13))
  cols_to_read = c(rep(3:15, 16), rep(19:31, 15))

  cols_df <- data.frame(days2, cols_to_read)

  cols_to_edit_per_day <- cols_df$cols_to_read[which(cols_df$days2 == day)]

  # change the values of the cells to edit 
  # 24 hour rain 
  cell1 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[1])
  cell1[["Value"]] <- rain

  # Pool elev
  cell2 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[2])
  cell2[["Value"]] <- pool

  # Evap
  cell3 <- sheet$Cells(rows_to_edit_per_day[2],cols_to_edit_per_day[5])
  cell3[["Value"]] <- evap

  # Discharge
  cell4 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[6])
  cell4[["Value"]] <- discharge
  
  # Tailwater
  cell5 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[8])
  cell5[["Value"]] <- tail
  
  # Gates
  cell6 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[9])
  cell6[["Value"]] <- gate1
  cell7 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[10])
  cell7[["Value"]] <- gate2
  cell8 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[11])
  cell8[["Value"]] <- gate3
  cell9 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[12])
  cell9[["Value"]] <- gate4
  
  # Gate Operation
  cell10 <- sheet$Cells(rows_to_edit_per_day[1],cols_to_edit_per_day[13])
  cell10[["Value"]] <- gate_op1
  cell10 <- sheet$Cells(rows_to_edit_per_day[2],cols_to_edit_per_day[13])
  cell10[["Value"]] <- gate_op2
  cell10 <- sheet$Cells(rows_to_edit_per_day[3],cols_to_edit_per_day[13])
  cell10[["Value"]] <- gate_op3
  
  # save the workbook
  wb$Save()
  
  #save as a new workbook
  #wb$SaveAS("new_file.xls")
  
  # close the workbook
  xlApp$Quit()

}

# run the edit_lake_excels function
edit_lake_excels(file, month, day, rain, pool, evap, discharge, tail, 
                 gate1, gate2, gate3, gate4, gate_op1, gate_op2, gate_op3)
