#' loadWeborama.Ad
#'
#' A function for loading data exported from the Weborama interface
#' 
loadWeborama.ad <- function(client,
                            #' @param client 
                            datadir,
                            #' @param datadir the directory of the data
                            dfs.names){
  #' @param dfs.names 
  
  if(missing(datadir)) {
    # directory of the weborama data
    datadir <- paste0("Weborama/")
  }
  
  # get the list of the excel fils
  filenames <- list.files(path = datadir, pattern = "^(Report)(.*)(.xlsx)$", full.names = TRUE)
  
  # loop onto this list
  perf.ad <- do.call(dplyr::bind_rows, lapply(filenames, function(filename){
    
    cat("   - loading file", tail(unlist(strsplit(filename, split = "/")), 1), "\n")
    
    # get info of the opened file
    info <- strsplit(strsplit(unlist(c(readxl::read_excel(filename)[3,])), split = "Campaigns: ")[[1]][2], split = " - ")[[1]]
    campaign <- info[1]
    type <- gsub(info[2], pattern = "[[:digit:][:punct:][:blank:]]", replacement = "")
    
    # create the dplyr::tbl_df
    tmp <- readxl::read_excel(filename, sheet = "DataView") %>%
      # some weborama variables do not have the same name
      rename(Size = Insertion,
             Impressions = Imp.,
             Ad = Creative,
             CTR = `CTR/ Imp.`,
             Goals = `Conv.`) %>%
      # make some transformations
      mutate(Date = lubridate::ymd(Date),
             # the ad name may change its pattern later on. This should be adapted
             # accordingly
             # Ad = factor(sapply(strsplit(Ad, split = "[_-]"), function(v) {
             #   if(length(v)==1) {res <- "1";}
             #   else {res <- tail(v, 1);}
             #   if(nchar(res)==1) res <- paste0("V", res);
             #   res <- gsub(x = res, pattern = 'V', replacement = "")
             #   return(res)
             # }
             Ad = factor(sapply(strsplit(Ad, split = "_"), function(v) {
               tail(v, 1)
             })),
             Size = sapply(strsplit(Size, split = " - "), function(v) v[1]),
             # clicks is a numeric and weborama places '-' instead of 0. There is no
             # NA but only no clicks, is 0 clicks
             Clicks = as.numeric(gsub(pattern = "-", replacement = "0", x = Clicks)),
             # for this function, DSP is always weborama
             DSP = "Weborama",
             # same as clicks
             Impressions = as.numeric(gsub(pattern = "-", replacement = "0", x = Impressions)),
             CTR = as.numeric(CTR),
             # Type = prospect or retargeting
             Type = type,
             Campaign = campaign,
             Advertiser = client,
             # missing info in weborama export by size
             IO = '',
             `Ad Group` = getAdGroup(as.character(Ad)),
             Redirect = getRedirect(as.character(Ad)),
             Goals = 0*NA,
             CPC = 0*NA,
             CPA = 0*NA,
             `Goals PV` = 0*NA,
             `Goals PC` = 0*NA,
             Spent = 0*NA,
             CPM = 0*NA) %>%
      arrange(Date)
    tmp
  }))[dfs.names] %>%
    arrange(Date)
  
  #' @return An object from dplyr with colnames given by the input \code{dfs.names}
  return(perf.ad)
}

#' loadWeborama.goals
#'
#' A function for loading data got from the Weborama report
#' 
loadWeborama.goals <- function(client,
                               #' @param client 
                               datadir,
                               #' @param datadir the directory of the data
                               dfs.names){
  # import data
  dt.prosp <- readxl::read_excel(datadir, sheet = "Prospecting")
  dt.prosp$Type <- 'Prospecting'
  
  dt.retar <- readxl::read_excel(datadir, sheet = "Retargeting")
  dt.retar$Type <- 'Retargeting'
  
  campaign <- strsplit(x = unlist(readxl::read_excel(datadir, sheet = 2)[4, 1]), split = ' - ')[[1]][1]
  
  dt <- dt.prosp %>%
    bind_rows(dt.retar) %>%
    rename(Spent = Budget,
           Goals = `TOTAL Vente`,
           `Goals PV` = `Vente PV`,
           `Goals PC` = `Vente PC`) %>%
    mutate(Date = lubridate::ymd(Date),
           DSP = "Weborama",
           Advertiser = client,
           IO = "",
           Campaign = campaign,
           `Ad Group` = "",
           Ad = "",
           Redirect ='',
           Size ='',
           CPC = Spent/Clicks,
           CPA = as.numeric(gsub(x = CPA, replacement = "", pattern = "-"))) %>%
    arrange(Date)
  dt <- dt[dfs.names]
  dt
}