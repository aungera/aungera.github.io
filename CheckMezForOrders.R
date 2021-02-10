### Check Mez For Orders ###

# this program is intended to utilize the spreadsheet
# of doors that need to be produced and shipped out by 
# product service and cross-reference a user generated
# excel sheet of inventory from the mezzanine to inform
# the user that there is an order that they can fulfill
# from the inventory on hand
#
# Note: user is responsible for updating the inventory
# sheet of doors located on the mezzanine
# 
# Remember, the computer isn't intelligent, you are, so
# use it as a tool. 
#
# Inputs:
# - 'Inter Org Past Due Any Org to RVR.xls' FOR DEMAND
# - 'Door Inventory Tracking.xlsx' FOR INVENTORY ON MEZ
# 
# Outputs:
# The program is designed to output 1 pdfs to your working directory
# which informs you of any doors you have orders for according to your
# spreadsheet called "InventoryOrders_DATE" where date is the current date


# program created by Annah Aunger, PS Co-op Spring 2021



#-- make sure you have these packages installed on your machine--#

#install.packages("readxl")
#install.packages("stringi")
#install.packages("tidyselect")
#install.packages("stringr")
#install.packages("gt")
#install.packages("gridExtra")
#install.packages("grid")

# open the library to access the packages
library("readxl")
library("stringi")
library("tidyselect")
library("stringr")
library("gt")
library("gridExtra")
library("grid")


# the program assumes you have a file named 'Inter Org Past Due Any Org to JEF.xls'
# from which you use as your demand spreadsheet.

filename <- "Inter Org Past Due Any Org to RVR.xls"

#-- import the data from the file --#

"make sure the file you would like to read in is in the working directory!"
paste("the working directory is ", getwd(), "")

rawData <- as.data.frame(read_excel(filename))

#-- format the data --#

titles <- as.data.frame(rawData[c(1),])
names(rawData) <- titles
rawData<- rawData[-c(1),]

#-- filter the data for AP5 and old/current Doors --#

AP5Data <- subset(rawData, rawData$`Ship ORG` =="AP5")

# import the file that has a list of the inventory on the mez
doorInventory <- as.data.frame(read_excel("Door Inventory Tracking.xlsx"))

# filter out empty rows & columns
doorInventory <- na.omit(doorInventory[,-c(2,6:8)])

#-- cross reference the inventory with the doors having demand --#

for(i in 1:length(AP5Data$Item)){
  
  for( j in 1:length(doorInventory$`catalog number`)){
    #does the current row from the AP5 data contain the current row's model #?
    
    if(stri_detect_fixed(AP5Data$Item[i], doorInventory$`catalog number`[j], negate=FALSE)){
      AP5Data$'demand'[i] <- doorInventory$inventory
      
      break ## once you find that the entry is indeed in the catalog, you can stop searching
      #break will exit the inner for loop
    } else {
      AP5Data$'demand'[j] <- 0
      
    }
  }
}

if(any(AP5Data$`demand` != 0)){ # that means we have demand, so generate a pdf
  
  #grab all the rows which you want (either all old or all current)
  Doors <- AP5Data[which(AP5Data$demand != 0),]
  
  #-- simplify the columns --#
  
  #limit the columns in the new table
  Doors <- Doors[c("Order Number", "Item", "Quantity Ordered", "Item Description", "demand")]
  #rename the columns
  names(Doors) <- c("Order Num", "Item", "Qty Demanded", "Descrip", "On Hand Inventory")
  
  #-- exporting the tables as pdfs --#
  
  #export the manual check table as a pdf
  pdf(paste("InventoryOrders_", Sys.Date(), ".pdf", sep = ""),height = 8.5,
      width = 11, title = "Inventory Orders", paper = "A4r")
  grid.table(Doors)
  dev.off()
  
} else {
  print("There are no doors on the mez that have demand according to the inventory sheet.")
}
