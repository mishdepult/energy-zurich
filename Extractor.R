# Date: 06.12.2021
# Author: Michelle Wäger
# Title: Extractor

# library
library(dplyr)
library(tidyverse)
library(xlsx)

#liest Masterfile ein
d.masterfile <- read_excel("Kopie_Masterfile.xlsx")

#Spalten im Masterfile umbennenen
d.masterfile <- d.masterfile %>% 
  rename(Kontext = Kontext(),
          = )

# leeres Blatt
remove(d.deckblatt)
d.deckblatt <- data.frame(matrix(NA, nrow = 500, ncol = 2))
d.deckblatt <- d.deckblatt %>% rename(DA = X1,
                                      Top5 = X2)

j <- 1

# füllt neue Tabelle ab
for (i in 1:nrow(d.masterfile)){
  if (is.na(d.masterfile$Top5[i])){ #überspringt NAs
    next
  } else {
      if (d.masterfile$Top5[i] == "ja") {
        d.deckblatt$DA[j] <- d.masterfile$DA[i]
        d.deckblatt$Top5[j] <- d.masterfile$Massnahmenbeschrieb[i]
        j <- j+1
    }
  }
}

# zähle alle, neue und abgeschlossene Massnahmen nach DA
d.anzahl <- d.masterfile %>% count(DA)
d.abgeschlossen <- subset(d.masterfile, Status == "a") %>% count(DA)
d.abgeschlossen <- d.abgeschlossen %>% 
  rename(abgeschlossen = n)
d.neu <- subset(d.masterfile, Status == "n") %>% count(DA)
d.neu <- d.neu %>% 
  rename(neu = n)

# fasse Anzahlmassnahmen in einem File zusammen
d.anzahl <- merge(d.anzahl, d.abgeschlossen, all.x = TRUE)
d.anzahl <- merge(d.anzahl, d.neu, all.x = TRUE)
d.anzahl

#schreibe Daten in Excel
write.xlsx(d.anzahl, file = "/cloud/project/Deckblatt.xlsx", 
           col.names = TRUE, row.names = FALSE,
           sheetName = "Anzahl")
