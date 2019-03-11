load_packages <- function() {
        library(RSelenium)
        library(wdman)
        library(data.table)
        library(plyr)
        library(dplyr)
        library(RDCOMClient)
        library(stringr)
        library(readxl)
        library(tidyverse)
}

setup_server <- function() {
        selCommand <- wdman::selenium(jvmargs = c("-Dwebdriver.chrome.verboseLogging=true"), 
                                      retcommand = TRUE)
        cat(selCommand) ## run out in CMD
        
}

get_id_data <- function(masterfileDir, downloadDir) {
        eCaps <- list(
                chromeOptions = 
                        list(prefs = list(
                                "profile.default_content_settings.popups" = 0L,
                                "download.prompt_for_download" = FALSE,
                                "download.default_directory" = "r:/heart/auto",
                                # "download.default_directory" = "r:/heart/auto",
                                "download.directory_upgrade" = TRUE,
                                "plugins.always_open_pdf_externally" = TRUE
                        )
                        )
        )
        remDr <- remoteDriver(extraCapabilities = eCaps, port = 4567L)
        remDr$open()
        
        # remDr$navigate("https://infodirect.ca/Search#/Messages")
        
        webElem <- remDr$findElement(using = 'name', value = "DirectorySearchFilter.SearchType")
        webElem$sendKeysToElement(list("Residential"))
        
        webElem2 <- remDr$findElement(using = 'name', value = "DirectorySearchFilter.City")
        webElem2$sendKeysToElement(list("Kamloops"))
        
        
        
        # master <- read_csv(masterfileDir)
        master <- read_csv("r:/allheartoct12/MasterFile2.csv")
        kam <- master[master$City == "KAMLOOPS",]
        kam$file <- str_match(kam$filenamed, "[^/]*$")
        
        streets <- c("Summit dr", "Springhill dr", "Sovereign Crt", "Bestwick", "Gleneagles")
        streetsKam <- unique(kam$`Street name`)[50:100]
        
        for (st in streetsKam) {
                webElem3 <- remDr$findElement(using = 'name', value = "DirectorySearchFilter.Street")
                webElem3$sendKeysToElement(list(st, key = "enter"))
                webElem4 <- remDr$findElement(using = 'id', "button-export")
                Sys.sleep(4)
                webElem4$clickElement()
                Sys.sleep(4)
                webElem3$clearElement()
                
        }
        # remDr$close()
} ## gets all streets from infodirect.ca

open_excel <- function(downloadDir) {
        list_files2 <- list.files(downloadDir, pattern="*.xlsx", full.names=TRUE)
        for (filez in list_files2) {
                shell.exec(filez) ## open excel sheets
        }
}

save_close_excel <- function(downloadDir) {
        list_files2 <- list.files(downloadDir, pattern="*.xlsx", full.names=TRUE)
        wsh <- COMCreate("Wscript.Shell")
        wsh$SendKeys("% n")
        Sys.sleep(2)
        for (x in 1:length(list_files2)) {
                wsh <- COMCreate("Wscript.Shell")
                Sys.sleep(1)
                wsh$SendKeys("^{s}") 
                wsh$SendKeys("%{F4}") 
                Sys.sleep(1)
                wsh$SendKeys("%{TAB}")
                wsh$SendKeys("{ENTER}")
                Sys.sleep(2)
                # if(x == length(list_files2) - 1)
                
        }
}

data_out <- function(downloadDir, dncFile, masterfileDir) {
        # list_files <- list.files(downloadDir, pattern="*.xlsx", full.names=TRUE)
        list_files <- list.files("r:/heart/auto", pattern="*.xlsx", full.names=TRUE)
        # list_filesdnc <- list.files(dncFile, pattern="*.xlsx", full.names=TRUE)
        # list_filesdnc <- list.files("r:/heart/dnc/dnc.xlsx", pattern="*.xlsx", full.names=TRUE)
        dataset <- lapply(list_files, read_excel)
        dnc <- read_excel("r:/heart/dnc/dnc.xlsx", sheet=2)
        # master <- read_csv(masterfileDir)
        master <- read_csv("r:/allheartoct12/MasterFile2.csv")
        kam <- master[master$City == "KAMLOOPS",]
        kam$file <- str_match(kam$filenamed, "[^/]*$")
        dataset <- ldply(dataset, data.frame)
        dataset <- dataset %>% select(Name, House, Street, Apt, City, Postal, Phone)
        # names(dataset) <- c("Name","House","Street","Apt", "City","Prov","Postal","Phone")
        dataset <- dataset[order(dataset$House),]
        dataset$Apt <- sapply(dataset$Apt, as.character)
        dnc$`Last Name` <- tolower(dnc$`Last Name`)
        dnc$`Street Name` <- tolower(dnc$`Street Name`)
        dnc$`Home Phone` <- gsub("-", "", dnc$`Home Phone`)
        dataset$Phone <- gsub("-", "", dataset$Phone)
        data <- as.data.table(dataset)
        data <- data[, 
                     if (!(is.na(House) & is.na(Apt))) 
                             .(
                                     Name = Name %>% unique %>% paste(collapse = ", AND "), 
                                     Phone = Phone %>% unique %>% paste(collapse = " OR ")
                             )
                     else
                             .(Name, Phone)
                     , by=.(House, Street, Apt, City, Postal)]
        
        data <- data %>%
                select(colnames(data)) %>%
                filter(!Phone %in% dnc$`Home Phone`)
        # data <- data[c("Name", "House", "Street", "Apt", "City", "Postal", "Phone")]
        data$Apt[is.na(data$Apt)] <- " "
        data$House[is.na(data$House)] <- " "
        data <- data[with(data, order(House, Street, City)), ]
        
        kam3 <- kam %>% select(file, `Street name`)
        names(kam3)[2] <- "Street"
        tester <- left_join(data, kam3)
        tester <- tester[!duplicated(tester),]
        # write.xlsx(as.data.frame(tester), "tester.xlsx", sheetName="Info Direct Data", append=F, col.names=TRUE, row.names=FALSE)
        # final <- merge(data, kam3)
        final <- tester %>% select(Name, House, Street, Apt, City, Postal, Phone, file)
        final$file <- as.character(final$file)
        # xz <- as.data.frame(unique(final$file))
        
        for (district in (unique(final$file))) {
                final2 <- final %>% filter(file == district)
                districtName <- paste(unique(final2$file), ".csv", sep="")
                fwrite(as.data.frame(final2), districtName)
        }
} ## gets area by district, combines all files, concatenates based on critera, filters out DNC list, outputs data by district name

load_packages()
setup_server()

get_id_data("r:/allheartoct12/MasterFile2.csv", "r:/heart/auto")

open_excel("r:/heart/auto")
Sys.sleep(6)
save_close_excel("r:/heart/auto")

data_out("r:/heart/auto", "r:/heart/dnc/dnc.xlsx", "r:/allheartoct12/MasterFile2.csv")




remDr$close()

















# library(RDCOMClient)

# test2 <- function () {
#         wsh2 <- COMCreate("Wscript.Shell")
#         wsh2$SendKeys("% n")
#         Sys.sleep(2)
#         count = 0
#         while(length(list_files2) + 1 > count) {
#                 wsh <- COMCreate("Wscript.Shell")
#                 wsh$SendKeys("^{s}") 
#                 wsh$SendKeys("%{F4}") 
#                 wsh$SendKeys("%{TAB 2}") 
#                 wsh$SendKeys("{ENTER}")
#                 wsh$SendKeys("^{s}")  ## 
#                 wsh$SendKeys("%{F4}") ##
#                 count = count + 1 
#                 }
# }
# 
# test33 <- function () {
#         #wsh2 <- COMCreate("Wscript.Shell")
#         wsh <- COMCreate("Wscript.Shell")
#                 
#         # wsh$SendKeys("%(TAB)") 
#         wsh$SendKeys("%{TAB 4}") 
#         # wsh$SendKeys("%{TAB}") 
#         wsh$SendKeys("{ENTER}") 
#                 # wsh$SendKeys("{ENTER}")
#                 # wsh$SendKeys("^{s}")  ## 
#                 # wsh$SendKeys("%{F4}") ##
#                 # count = count + 1 
#         }

# open_close_excel <- function() {



# wsh <- COMCreate("Wscript.Shell")
# wsh$SendKeys("^({ESC}D)")
# wsh$SendKeys("{n}")


        
        



