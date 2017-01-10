## ---- cache=FALSE, include=FALSE-----------------------------------------
library(knitr)
library(ggplot2)
library(tidyr)
library(dplyr)
library(xlsx)
library(r2excel)

opts_knit$set(root.dir = '/Volumes/Stockage/Google Drive/skaze/Clients/Son-video/')

## ---- results='hide', warning=FALSE--------------------------------------
Oxom::makeReport(save.dir = "Rapport.xlsx",
                 pic.dir = "sv-logo.png")

## ---- results='hide', warning=FALSE--------------------------------------
Oxom::makeReport(save.dir = "Rapport-normalise.xlsx",
                 pic.dir = "sv-logo.png",
                 budget = 5000)

## ------------------------------------------------------------------------
list.files(pattern = "^(Rapport)")

