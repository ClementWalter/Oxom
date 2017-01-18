#' writeExcel
#'
writeExcel <- function(webo.ad,
                       #' @param webo.ad the data.frame for ad data from Weborama
                       media.ad,
                       #' @param media.ad the data.frame for ad data from Mediarithmics
                       webo.goals,
                       #' @param webo.goals the data.frame for goals data from Weborama
                       media.goals,
                       #' @param media.goals the data.frame for goals data from Mediarithmics
                       marge.dir,
                       #' @param marge.dir an excel file for specifying marge or budget by period of time
                       dfs.names = names(webo.ad),
                       #' @param dfs.names the header of the excel file
                       pic.dir,
                       #' @param pic.dir the directory of a picture to be added at the top left corner
                       save.dir
                       #' @param save.dir the saving directory with full name, ie <directory>/<filename>.xlsx
) {
  
  # Definition des fonctions pour résumer les tableaux
  summarise.df <- function(...){
    summarise_all(..., .funs = function(v) ifelse(class(v)[1]=="numeric", sum(v), "")) %>%
      mutate(CPM = Spent/Impressions*1000,
             CTR = Clicks/Impressions*100,
             CPC = Spent/Clicks,
             CPA = Spent/Goals) %>%
      mutate_at(.cols = vars(Spent, CPM, CTR, CPC, CPA), .funs = function(v) round(v, digits = 2))
  }
  
  summarise_by_ <- function(df, .cat){
    df.tmp <- as.data.frame(df %>% group_by_(.dots = .cat) %>%
                              summarise.df())
    df.tmp <- df.tmp[dfs.names]
    df.tmp
  }
  
  write.header <- function(wb, sheet, head, dfs.names){
    # Ecriture du titre
    r2excel::xlsx.addHeader(wb, sheet, value = head, startRow = 6, level = 3)
    
    # Ecriture du header général
    r2excel::xlsx.addTable(wb, sheet, data = matrix(dfs.names, nrow = 1),
                           startRow = 8, startCol = 1, row.names = FALSE, col.names = FALSE,
                           fontColor = rgb(0, 0, 0), fontSize = 12, rowFill = skaze.darkblue)
    
    start.row <- 9
    start.row
  }
  
  write.df <- function(df, start.row, wb, sheet, start.col = 1, date = TRUE){
    df <- as.data.frame(df)
    if(date){
      r2excel::xlsx.addTable(wb, sheet, data = df['Date'],
                             startRow = start.row, startCol = 1, row.names = FALSE, col.names = FALSE,
                             fontColor='grey30', fontSize=12,
                             rowFill=c(skaze.lightblue, skaze.grey)
      )
    }
    if(is.character(start.col)) {
      start.col <- which(names(df)==start.col)
      df <- df[, -c(1:(start.col-1))]
    }
    r2excel::xlsx.addTable(wb, sheet, data = df,
                           startRow = start.row, startCol = start.col, row.names = FALSE, col.names = FALSE,
                           fontColor='grey30', fontSize=12,
                           rowFill=c(skaze.lightblue, skaze.grey)
    )
    
    start.row <- start.row + nrow(df) + 1
    start.row
  }
  
  # Fusion des données mediarithmics et weborama
  if(!is.null(media.goals) & !is.null(media.ad)) {
    media.goals <- media.ad %>%
      group_by(Date) %>%
      summarise.df() %>%
      mutate(Campaign = NULL, Goals = NULL) %>%
      full_join(media.goals %>% group_by(Date) %>% summarise(Goals = sum(Goals), Campaign = 'campaign'), by = "Date") %>%
      mutate(`Goals PC` = 0*NA,
             `Goals PC` = 0*NA,
             DSP = "Mediarithmics")
  }
  if(!is.null(webo.goals)){
    if(!is.null(media.goals)) {
      df.spent <- webo.goals %>% bind_rows(media.goals) %>% arrange(Date)
    } else {
      df.spent <- webo.goals %>% arrange(Date)
    }
  } else {
    df.spent <- media.goals %>% arrange(Date)
  }
  df.spent <- df.spent[dfs.names]
  
  if(!is.null(webo.ad)){
    if(!is.null(media.ad)){
      df.ad <- webo.ad %>% bind_rows(media.ad) %>% arrange(Date)
    } else {
      df.ad <- webo.ad %>% arrange(Date)
    }
  } else {
    df.ad <- media.ad %>% arrange(Date)
  }
  
  df.ad <- df.ad[dfs.names]
  head <- paste("Résultats globaux du", min(df.ad$Date, df.spent$Date), "au", max(df.ad$Date, df.spent$Date))
  
  # Define some colors
  skaze.darkblue <- rgb(36/255, 116/255, 152/255)
  skaze.grey <- rgb(217/255, 217/255, 217/255)
  skaze.lightblue <- rgb(114/255, 179/255, 199/255)
  skaze.palette <- colorRampPalette(c(skaze.lightblue, skaze.darkblue))
  
  # Load the template
  wb <- xlsx::loadWorkbook(system.file('extdata', 'Template.xlsx', package = "Oxom"))
  sheets <- xlsx::getSheets(wb)
  suivi.global <- sheets$`Suivi Global`
  suivi.quot <- sheets$`Suivi Quotidien`
  suivi.hebdo <- sheets$`Suivi Hebdo`
  suivi.mens <- sheets$`Suivi Mensuel`
  
  ###########################################
  ### Normalisation des dépenses
  ###########################################
  
  # Obtention du tableau le plus général
  df.tmp <- as.data.frame(df.spent %>% summarise.df())
  
  # Obtenir le coef pour arriver au total du budget
  if(!missing(marge.dir)) {
    coef <- xlsx::read.xlsx(file = marge.dir, sheetIndex = 1) %>%
      mutate(Du = lubridate::as_date(Du),
             Au = lubridate::as_date(Au)) %>%
      filter(!is.na(Coef))
    
    # Modification de toutes les données avec des euros
    Coef <- sapply(df.spent$Date, function(d){
      ind <- tail(which(d>=coef$Du), 1)
      coef[ind,3]
    })
    
    df.spent <- df.spent %>%
      mutate_at(.cols = vars(Spent, CPM, CTR, CPC, CPA), .funs = function(v) Coef*v)
    
    Coef <- sapply(df.ad$Date, function(d){
      ind <- tail(which(d>=coef$Du), 1)
      coef[ind,3]
    })
    df.ad <- df.ad %>%
      mutate_at(.cols = vars(Spent, CPM, CTR, CPC, CPA), .funs = function(v) Coef*v)
  }
  
  ############################################
  #### Ecriture du logo
  ############################################
  
  # Placer le logo si donné
  if(!missing(pic.dir)) {
    lapply(sheets, function(sheet) {
      xlsx::addPicture(file = pic.dir, sheet = sheet, startRow = 2, startColumn = 2)
    })
  }
  
  ############################################
  #### Ecriture des tableaux dans suivi.global
  ############################################
  cat('   - ecriture de l\'onglet \'Suivi Global\' \n')
  # Initialisation
  start.row <- write.header(head = head, wb = wb, sheet = suivi.global, dfs.names = dfs.names)
  
  # Ecriture de la ligne résumée la plus global
  start.row <- df.spent %>% summarise.df() %>% write.df(start.row = start.row, wb = wb, sheet = suivi.global, date = FALSE)
  
  # Ecriture du résumé par DSP
  start.row <- df.spent %>%
    summarise_by_(.cat = "DSP") %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.global, start.col = "DSP", date = FALSE)
  
  # Ecriture des tableaux comparatif par autres catégories
  categories <- c('Type', 'Size', 'Ad', 'Redirect')
  for(categorie in categories){
    start.row <- df.ad %>%
      summarise_by_(.cat = categorie) %>%
      write.df(start.row = start.row, wb = wb, sheet = suivi.global, start.col = categorie, date = FALSE)
  }
  
  ############################################
  #### Ecriture des tableaux dans graphe.quot
  ############################################
  cat('   - ecriture de l\'onglet \'Graphe Quotidien\' \n')
  
  # Initialisation
  df.tmp <- df.spent %>% summarise_by_(.cat = 'Date') %>%
    mutate(Date = as.character(Date)) %>%
    select(-Advertiser, -IO, -DSP, -Type, -Campaign, -`Ad Group`, -Size, -Ad, -Redirect)
  start.row <- write.header(head = head, dfs.names = names(df.tmp), wb = wb, sheet = sheets$`Graphe Quotidien`)
  
  # Ecriture du tableau le plus général
  start.row <- df.tmp %>%
    write.df(start.row = start.row, wb = wb, sheet = sheets$`Graphe Quotidien`)
  
  ############################################
  #### Ecriture des tableaux dans suivi.quot
  ############################################
  cat('   - ecriture de l\'onglet \'Suivi Quotidien\' \n')
  # Initialisation
  start.row <- write.header(head = head, wb = wb, sheet = suivi.quot, dfs.names = dfs.names)
  
  # Ecriture du tableau le plus général
  start.row <- df.spent %>% summarise_by_(.cat = 'Date') %>%
    mutate(Date = as.character(Date)) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.quot)
  
  # Ecriture du tableau par DSP
  start.row <- df.spent %>% summarise_by_(.cat = c('Date', 'DSP')) %>%
    mutate(Date = as.character(Date)) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.quot, start.col = "DSP")
  
  # Ecriture du tableau par Size
  start.row <- df.ad %>% summarise_by_(.cat = c('Date', 'Size')) %>%
    mutate(Date = as.character(Date)) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.quot, start.col = "Size")
  
  # Ecriture du tableau par Ad
  start.row <- df.ad %>% summarise_by_(.cat = c('Date', 'Ad')) %>%
    mutate(Date = as.character(Date)) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.quot, start.col = "Ad")
  
  ############################################
  #### Ecriture des tableaux dans graphe.hebdo
  ############################################
  cat('   - ecriture de l\'onglet \'Graphe Hebdo\' \n')
  # Initialisation
  df.tmp <- df.spent %>% mutate(Date = format(Date, "%Y-S%V")) %>% summarise_by_(.cat = 'Date') %>%
    select(-Advertiser, -IO, -DSP, -Type, -Campaign, -`Ad Group`, -Size, -Ad, -Redirect)
  start.row <- write.header(head = head, dfs.names = names(df.tmp), wb = wb, sheet = sheets$`Graphe Hebdo`)
  
  # Ecriture du tableau le plus général
  start.row <- df.tmp %>%
    write.df(start.row = start.row, wb = wb, sheet = sheets$`Graphe Hebdo`)
  
  ############################################
  #### Ecriture des tableaux dans suivi.hebdo
  ############################################
  cat('   - ecriture de l\'onglet \'Suivi Hebdo\' \n')
  # Initialisation
  start.row <- write.header(head = head, wb = wb, sheet = suivi.hebdo, dfs.names = dfs.names)
  
  # Ecriture du tableau le plus général
  start.row <- df.spent %>% mutate(Date = format(Date, "%Y-S%V")) %>% summarise_by_(.cat = 'Date') %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.hebdo)
  
  # Ecriture du tableau par DSP
  start.row <- df.spent %>% mutate(Date = format(Date, "%Y-S%V")) %>% summarise_by_(.cat = c('Date', 'DSP')) %>%
    mutate(Date = as.character(Date)) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.hebdo, start.col = "DSP")
  
  # Ecriture du tableau par Size
  start.row <- df.ad %>% mutate(Date = format(Date, "%Y-S%V")) %>% summarise_by_(.cat = c('Date', 'Size')) %>%
    mutate(Date = as.character(Date)) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.hebdo, start.col = "Size")
  
  # Ecriture du tableau par Ad
  start.row <- df.ad %>% mutate(Date = format(Date, "%Y-S%V")) %>% summarise_by_(.cat = c('Date', 'Ad')) %>%
    mutate(Date = as.character(Date)) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.hebdo, start.col = "Ad")
  
  ############################################
  #### Ecriture des tableaux dans graphe.mens
  ############################################
  cat('   - ecriture de l\'onglet \'Graphe Mensuel\' \n')
  # Initialisation
  df.tmp <- df.spent %>% mutate(Date = factor(format(Date, "%B-%y"),
                                              levels = c(paste(month.name, 16, sep = "-"), paste(month.name, 17, sep = "-")))) %>%
    summarise_by_(.cat = 'Date') %>%
    select(-Advertiser, -IO, -DSP, -Type, -Campaign, -`Ad Group`, -Size, -Ad, -Redirect)
  start.row <- write.header(head = head, dfs.names = names(df.tmp), wb = wb, sheet = sheets$`Graphe Mensuel`)
  
  # Ecriture du tableau le plus général
  start.row <- df.tmp %>%
    write.df(start.row = start.row, wb = wb, sheet = sheets$`Graphe Mensuel`)
  
  ############################################
  #### Ecriture des tableaux dans suivi.mens
  ############################################
  cat('   - ecriture de l\'onglet \'Suivi Mensuel\' \n')
  # Initialisation
  start.row <- write.header(head = head, wb = wb, sheet = suivi.mens, dfs.names = dfs.names)
  
  # Ecriture du tableau le plus général
  start.row <- df.spent %>% mutate(Date = factor(format(Date, "%B-%y"),
                                                 levels = c(paste(month.name, 16, sep = "-"), paste(month.name, 17, sep = "-")))) %>%
    summarise_by_(.cat = 'Date') %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.mens)
  
  # Ecriture du tableau par DSP
  start.row <- df.spent %>% mutate(Date = factor(format(Date, "%B-%y"),
                                                 levels = c(paste(month.name, 16, sep = "-"), paste(month.name, 17, sep = "-")))) %>%
    summarise_by_(.cat = c('Date', 'DSP')) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.mens, start.col = "DSP")
  
  # Ecriture du tableau par Size
  start.row <- df.ad %>% mutate(Date = factor(format(Date, "%B-%y"),
                                              levels = c(paste(month.name, 16, sep = "-"), paste(month.name, 17, sep = "-")))) %>%
    summarise_by_(.cat = c('Date', 'Size')) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.mens, start.col = "Size")
  
  # Ecriture du tableau par Ad
  start.row <- df.ad %>% mutate(Date = factor(format(Date, "%B-%y"),
                                              levels = c(paste(month.name, 16, sep = "-"), paste(month.name, 17, sep = "-")))) %>%
    summarise_by_(.cat = c('Date', 'Ad')) %>%
    write.df(start.row = start.row, wb = wb, sheet = suivi.mens, start.col = "Ad")
  
  ############################################
  #### Enregister le fichier
  ############################################
  cat('   - enregistrement du fichier \n')
  xlsx::saveWorkbook(wb, file = save.dir)
  
}