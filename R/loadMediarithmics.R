#' loadMediarithmics.Ad
#'
#' A function for loading data exported from the Mediarithmics interface
#' 
#' @import tidyr
#' @import dplyr
#' 
loadMediarithmics.ad <- function(client,
                              #' @param client 
                              datadir,
                              #' @param datadir the directory of the data
                              dfs.names){
  #' @param dfs.names 
  
  if(missing(datadir)) {
    # directory of the weborama data
    datadir <- paste0("Mediarithmics/")
  }
  
  # get the list of the excel fils
  filenames <- list.files(path = datadir, pattern = "^([^\\~])(.*)(.xlsx)$", full.names = TRUE)
  
  tmp <- lapply(filenames, function(filename){
    cat("   - loading file", tail(unlist(strsplit(filename, split = "/")), 1), "\n")
    dfs.tmp <- list(
      overview.tmp = xlsx::read.xlsx2(file = filename, sheetIndex = 1, startRow = 6),
      ads.tmp = xlsx::read.xlsx2(file = filename, sheetIndex = 2, startRow = 6),
      ad_group.tmp = xlsx::read.xlsx2(file = filename, sheetIndex = 3, startRow = 6),
      sites.tmp = xlsx::read.xlsx2(file = filename, sheetIndex = 4, startRow = 6)
    )
    tmp <- getInfo(filename)
    day <- tmp$date
    campaign <- tmp$campaign
    
    dfs.tmp <- lapply(dfs.tmp, function(df){
      tryCatch({
        df$Date <- day
        df$Campaign <- campaign
        df
      }, error = function(cond){})
    })
    dfs.tmp
  })
  
  dfs <- lapply(1:4, function(name){
    do.call(rbind, lapply(tmp, function(l) l[[name]]))
  })
  names(dfs) <- c('overview','ads', 'ad_group', 'sites')
  
  # dfs$overview <- dfs$overview %>%
  #   mutate_at(.cols = vars(CPA, CPC, CTR, CPM, Spent), .funs = function(v) as.numeric(as.character(v))) %>%
  #   mutate(Ventes = round(Spent/CPA),
  #          Clicks = Spent/CPC,
  #          Impressions = Spent/CPM*1000) %>%
  #   mutate_at(.cols = vars(Ventes, Clicks, Impressions), .funs = function(v) ifelse(is.nan(v) | v==Inf, 0, v))
  # 
  # dfs$ad_group <- dfs$ad_group %>%
  #   mutate_at(.cols = vars(Imp., CPM, Spent, Clicks, CTR, CPC, CPA), .funs = function(v) as.numeric(as.character(v))) %>%
  #   rename(Impressions = Imp.,
  #          Ad = Name) %>%
  #   mutate(Goals = round(Spent/CPA))
  # 
  # dfs$sites <- dfs$sites %>%
  #   mutate_at(.cols = vars(Imp., CPM, Spent, Clicks, CTR, CPC, CPA), .funs = function(v) as.numeric(as.character(v))) %>%
  #   rename(Impressions = Imp.) %>%
  #   mutate(Goals = round(Spent/CPA))
  
  dfs$ads <- dfs$ads %>% rename(Impressions = Imp.,
                            Ad = Name,
                            Size = Format)
  
  missing.names <- dfs.names[!(dfs.names %in% names(dfs$ads))]
  
  for(name in missing.names){
    dfs$ads[name] <- NA
  }
  
  dfs$ads <- dfs$ads %>%
    mutate_at(.cols = vars(Impressions, Spent, Clicks, CPM, CTR, CPC, CPA), .funs = function(v) {as.numeric(as.character(v))}) %>%
    mutate_at(.cols = vars(Impressions, Spent, Clicks), .funs = function(v) ifelse(is.na(v), 0, v)) %>%
    mutate(Goals = round(Spent/CPA),
           `Goals PV` = 0*NA,
           `Goals PC` = 0*NA,
           Type = sapply(Campaign, function(v){
             v <- as.character(v)
             kwd <- grep(x = v, pattern = "Kwd");
             rex <- grep(x = v, pattern = "REx");
             v[kwd] <- "Prospecting"
             v[rex] <- "Retargeting"
             v
           }),
           Ad = factor(sapply(strsplit(as.character(Ad), split = "_"), function(v) tail(v,1))),
           `Ad Group` = getAdGroup(as.character(Ad)),
           Redirect = getRedirect(as.character(Ad)),
           Advertiser = client,
           DSP = "Mediarithmics",
           IO = "") %>%
    arrange(Date)
  
  missing.names <- dfs.names[!(dfs.names %in% names(dfs$ads))]
  
  for(name in missing.names){
    dfs$ads[name] <- NA
  }
  
  dfs$ads <- dfs$ads[dfs.names]
  
  #' @return A data.frame with header as given in input
  return(dfs$ads)
}

getInfo <- function(name){
  tmp <- xlsx::read.xlsx2(name, sheetIndex = 1)
  from <- lubridate::dmy(strsplit(as.character(tmp[2,1]), split = " ")[[1]][2])
  to <- lubridate::dmy(strsplit(as.character(tmp[2,2]), split = " ")[[1]][2])
  if(!identical(from, to)) stop('file contains data for more than one day')
  campaign <- as.character(tmp[1,2])
  return(list(date = from, campaign = campaign))
}

#' loadMediarithmics.goals
#'
#' A function for loading goals written in an external file
#' 
loadMediarithmics.goals <- function(datadir
                                    #' @param datadir the directory of the data
){
  df <- readxl::read_excel(datadir) %>% mutate(Date = lubridate::as_date(Date))
  df
}