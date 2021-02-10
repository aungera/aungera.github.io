### Suggest Door Labels ###

# this program is intended to utilize the spreadsheet
# of doors that need to be produced and shipped out by 
# product service and suggest to the user which door
# labels that they should print on the current day, 
# based on the data in the sheet, filtered by current
# versus old production
#
# Note: user is responsible for updating the catalog of
# doors to indicate to the program which are current
# and old production
# 
# Remember, the computer isn't intelligent, you are, so
# use it as a tool. 
#
# Running the program:
# You can run the program line by line, or in three chunks due to user inputs
# (1) lines 1-59, then enter the Y/N asked by the system
# (2) lines 60-97, then enter 1 or 2, asked by the system
# (3) the remainder of the program 
#
# Outputs:
# The program is designed to output 4 pdfs to your working directory.
# 3 of the pdfs are categorized by RH, LH, FZ, another by anything that did
# not belong to the RH/LH/FZ category which are labeled as "manual", and the
# last pdf is if you chose to base printing labels strictly on critical ratio
# instead of 50LH, 40RH and 20FZ, which is labeled as "suggestedSorted"


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


#-- grab the filename that you would like to extract data from --#

# validFile <- readline("Is the filename 'Inter Org Past Due Any Org to JEF.xls' ? 
#                       Enter Y if yes, N if no.")
# 
# if(validFile == "Y"){
#   
#   filename <- "Inter Org Past Due Any Org to RVR.xls"
#   
# } else {
#   
# filename <- readline("Enter the filename with extension: ")
# 
# }

########################################### hard coded value
filename <- "Inter Org Past Due Any Org to RVR.xls"




#-- import the data from the file --#

"make sure the file you would like to read in is in the working directory!"
paste("the working directory is ", getwd(), "")

rawData <- as.data.frame(read_excel(filename))


#-- format the data --#

titles <- as.data.frame(rawData[c(1),])
names(rawData) <- titles
rawData<- rawData[-c(1),]


#-- filter old production from current --#

print("Do you want to print labels for (1) Old production or (2) current?")

option <- "NA" #initialize the variable
option <- 2 ############################################ hard coded
# 
# ## need for to user to enter valid input so the screen will prompt until valid
# while((option!= "1" && option != "2")){
# 
#   option <- readline("Enter 1 or 2: ")
# 
# }

#import the file that has a list of the doors labeled old and current
doorInfo <- as.data.frame(read_excel("Door Catalog.xlsx"))

# get boolean vector of which doors the user wants
if(option == "1"){ # the user wants to see old production
  ## you'll need to filter to see old production
  chosenDoors <- stri_detect_fixed(doorInfo$`production`, "OLD", negate=FALSE)
  
} else { # the user wants to see current production
  ## you'll need to filter to see new production
  chosenDoors <- stri_detect_fixed(doorInfo$`production`, "CURRENT", negate=FALSE)
  
}


#-- filter the data for AP5 and old/current Doors --#

AP5Data <- subset(rawData, rawData$`Ship ORG` =="AP5")

#the door Info will now be only old or only current
doorInfo <- doorInfo[c(chosenDoors),]


#-- cross reference the old new w/ the doors in the spreadsheet --#

for(i in 1:length(AP5Data$Item)){
  
  for( j in 1:length(doorInfo$`catalog number`)){
    
    #does the current row from the AP5 data contain the current row's catalog #?
    if(stri_detect_fixed(AP5Data$Item[i], doorInfo$`catalog number`[j], negate=FALSE)){
      
      AP5Data$'want?'[i] <- TRUE
      AP5Data$'doorID'[i] <- doorInfo$`door ID`[j]
      break ## once you find that the entry is indeed in the catalog, you can stop searching
      #break will exit the inner for loop
      
    } else {
      
      AP5Data$'want?'[i] <- FALSE
      AP5Data$'doorID'[i] <- NA
      
    }
  }
}

#grab all the rows which you want (either all old or all current)
Doors <- AP5Data[c(AP5Data$`want?`),]

#-- simplify the columns --#

#limit the columns in the new table
DoorsSimple <- Doors[c("Order Number", "Item", "Quantity Ordered", "Past Due Status", "Need By Date", "Item Description", "doorID")]
#rename the columns
names(DoorsSimple) <- c("Order Num", "Item", "Qty", "Status", "Need By", "Descrip", "ID")

#-- Use Critical Ratio to Prioritize door scheduling --#

#create a column for the critical ratio
DoorsSimple$`CR` <- "NA"


#change the date format of the need by column so you can manipulate the date &
# calculate CR
for (x in 1:(length(DoorsSimple$`Need By`))) {
  
  #format the date
  DoorsSimple$`Need By`[x] <-as.character(
    as.Date(DoorsSimple$`Need By`[x], format = '%d-%b-%Y'))
  
  # CR is order day - current day / the quantity of the order
  DoorsSimple$`CR`[x] <- round(( as.numeric(as.Date(
    DoorsSimple$`Need By`[x], format = '%Y-%m-%d') - 
      Sys.Date())) / as.numeric(DoorsSimple$`Qty`[x]), digits= 3)
  
}

#-- Need to sort out RH, LH and Freezer --#

#grabbing LH doors
LH <- str_detect(DoorsSimple$`Descrip`, "LH")
LEFT <- str_detect(DoorsSimple$`Descrip`, "LEFT")
left <- str_detect(DoorsSimple$`Descript`, "left")

#appending LH doors into one df
DoorsLH <- rbind(DoorsSimple[c(LH),], DoorsSimple[c(LEFT),], DoorsSimple[c(left),] )

#grabbing RH doors
RH <- str_detect(DoorsSimple$`Descrip`, "RH")
RIGHT <- str_detect(DoorsSimple$`Descrip`, "RIGHT")

#appending RH doors into one df
#DoorsRH <- DoorsSimple[c(RH),]
DoorsRH <- rbind(DoorsSimple[c(RH),], DoorsSimple[c(RIGHT),] )

#grabbing FZ doors
FZ <- str_detect(DoorsSimple$`Descrip`, "FZ")
FREEZER <- str_detect(DoorsSimple$`Descrip`, "FREEZER")
FREEZR <- str_detect(DoorsSimple$`Descrip`, "FREEZR")

#appending FZ doors into one df
#DoorsFZ <- DoorsSimple[c(FZ),]
DoorsFZ <- rbind(DoorsSimple[c(FZ),], DoorsSimple[c(FREEZER),], DoorsSimple[c(FREEZR),] )
#DoorsFZ <- rbind(DoorsFZ, (DoorsSimple[c(FREEZR),]))


##there may still be doors in the list that are not labeled by FZ, RH, LH, etc..
#eliminate the LH doors
Doors_Other <- DoorsSimple[c(!LH),]
Other_LEFT <- str_detect(Doors_Other$`Descrip`, "LEFT")
Doors_Other <- Doors_Other[c(!Other_LEFT),]
Other_left <- str_detect(Doors_Other$`Descrip`, "left")
Doors_Other <- Doors_Other[c(!Other_left),]

#eliminate the RH doors
Other_RH <- str_detect(Doors_Other$`Descrip`, "RH")
Doors_Other <- Doors_Other[c(!Other_RH),]
Other_RIGHT <- str_detect(Doors_Other$`Descrip`, "RIGHT")
Doors_Other <- Doors_Other[c(!Other_RIGHT),]

#eliminate the FZ doors
Other_FZ <- str_detect(Doors_Other$`Descrip`, "FZ")
Doors_Other <- Doors_Other[c(!Other_FZ),]
Other_FREEZER <- str_detect(Doors_Other$`Descrip`, "FREEZER")
Doors_Other <- Doors_Other[c(!Other_FREEZER),]
Other_FREEZR <- str_detect(Doors_Other$`Descrip`, "FREEZR")
Doors_Other <- Doors_Other[c(!Other_FREEZR),]


#-- Order the doors of each category to see priorities --#
# sort the doors from smallest to largest so you can see priorities (most negative)
# (column 8 is the one with the critical ratio)
DoorsSorted <- DoorsSimple[order(as.numeric(DoorsSimple[,8])),]
DoorsLH <- DoorsLH[order(as.numeric(DoorsLH[,8])),]
DoorsRH <- DoorsRH[order(as.numeric(DoorsRH[,8])),]
DoorsFZ <- DoorsFZ[order(as.numeric(DoorsFZ[,8])),]

#cumulative sums LH
sumLH <- cumsum(as.numeric((DoorsLH$`Qty`)))

x=1

while(sumLH[x] < 50){
  x = x+1 # will index to 1 past 50 or 50
}

suggestedLH <- cbind(DoorsLH[c(1:x),], "run total" = sumLH[1:x])

#cumulative sums RH
sumRH <- cumsum(as.numeric((DoorsRH$`Qty`)))

x=1

while(sumRH[x] < 40){
  x = x+1 # will index to 1 past 40 or 40
}

suggestedRH <- cbind(DoorsRH[c(1:x),], "run total" = sumRH[1:x])

#cumulative sums FZ
sumFZ <- cumsum(as.numeric((DoorsFZ$`Qty`)))

x=1

while(sumFZ[x] < 20){
  x = x+1 # will index to 1 past 20 or 20
}

suggestedFZ <- cbind(DoorsFZ[c(1:x),], "run total" = sumFZ[1:x])

#cumulative sums whole sorted list
sumSorted <- cumsum(as.numeric((DoorsSorted$`Qty`)))

x=1

while(sumSorted[x] < 110){
  x = x+1 # will index to 1 past 110 or 110
}

suggestedSorted <- cbind(DoorsSorted[c(1:x),], "run total" = sumSorted[1:x])


#check if the "other" doors fall between the CR ranges for the 3 categories
#return value is the indexes

Other_LH <- which(as.numeric(Doors_Other$`CR`) <= max(as.numeric(suggestedLH$`CR`)))
Other_RH <- which(as.numeric(Doors_Other$`CR`) <= max(as.numeric(suggestedRH$`CR`)))
Other_FZ <- which(as.numeric(Doors_Other$`CR`) <= max(as.numeric(suggestedFZ$`CR`)))


#remove duplicate entries & combine the Manual checks
Manual <- unique(rbind(Doors_Other[c(Other_LH),], Doors_Other[c(Other_RH),],
                       Doors_Other[c(Other_FZ),]))

#remove the order number to help w printing, was only there to 
# help eliminate duplicate columns
suggestedLH <- suggestedLH[,-c(1)]
suggestedRH <- suggestedRH[,-c(1)]
suggestedFZ <- suggestedFZ[,-c(1)]
suggestedSorted <- suggestedSorted[,-c(1)]
 

#-- exporting the tables as pdfs --#

#export the manual check table as a pdf
pdf(paste("manualCheck_", Sys.Date(), ".pdf", sep = ""),height = 8.5,
    width = 11, title = "Manually Check Doors", paper = "A4r")
grid.table(Manual)
dev.off()

#export the suggested LH doors
pdf(paste("suggestedLH_", Sys.Date(), ".pdf", sep = ""),height = 8.5,
    width = 11, title = "Suggested LH Doors")
grid.table(suggestedLH)
dev.off()

#export the suggested RH doors
pdf(paste("suggestedRH_", Sys.Date(), ".pdf", sep = ""),height =8.5,
    width = 11, title = "Suggested RH Doors")
grid.table(suggestedRH)
dev.off()

#export the suggested FZ doors
pdf(paste("suggestedFZ_", Sys.Date(), ".pdf", sep = ""),height = 8.5,
    width = 11, title = "Suggested FZ Doors")
grid.table(suggestedFZ)
dev.off()

#export the whole sorted list, prioritized by critical ratio
## this pdf is formatted different bc how many doors will be there

pdf(paste("Suggested_CR_", Sys.Date(), ".pdf", sep = ""),height = 11,
    width = 8.5, title = "Suggested Doors by Critical Ratios")

mytheme <- gridExtra::ttheme_default( core = list(fg_params=list(cex = 0.55)),
                                    colhead = list(fg_params=list(cex = 0.5)),
                                    rowhead = list(fg_params=list(cex = 0.5)))

myt <- gridExtra::tableGrob(suggestedSorted, theme = mytheme)

grid.draw(myt)
dev.off()

## make pdf of all CR's calculated
pdf(paste("All_CR_", Sys.Date(), ".pdf", sep = ""),height = 11,
    width = 8.5, title = "All Doors by Critical Ratios")

mytheme <- gridExtra::ttheme_default( core = list(fg_params=list(cex = 0.5)),
                                      colhead = list(fg_params=list(cex = 0.4)),
                                      rowhead = list(fg_params=list(cex = 0.4)))

myt <- gridExtra::tableGrob(DoorsSorted, theme = mytheme)

grid.draw(myt)
dev.off()
