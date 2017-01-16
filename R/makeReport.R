#' makeReport
#'
#' This function creates the report based on the data found in the folders Weborama and Mediarithmics
#' 
#' @import tidyr
#' @import dplyr
#' 
#' @export
makeReport <- function(client,
                       #' @param client 
                       dfs.names,
                       #' @param dfs.names the header of the data
                       ...
                       #' @param ... other parameters given to writeExcel
) {
  
  if(missing(client)) client <- tail(strsplit(getwd(), split = "/")[[1]], 1)
  
  if(missing(dfs.names)){
    # Define variables of interest
    dfs.names <- c('Date', # la date de l'événement
                   'Advertiser', # le client
                   'IO', # Insertion Order
                   'DSP', # pour le moment Weboram ou Mediarithmics
                   'Type', # Prospecting ou Retargeting
                   'Campaign', # la campagne considérée
                   'Ad Group', # le Ad group
                   'Size', # la taille de l'ad, cad les dimensions
                   'Ad', # la créa utilisée
                   'Redirect', # la page vers laquelle redirige l'ad
                   'Impressions', # le nombre d'impressions
                   'Clicks', # le nombre de clicks
                   'Goals', # le nombre de ventes/formulaires, etc.
                   'Goals PV', # le nombre de goals "première visite" ?
                   'Goals PC', # le nombre de goals post clique
                   'Spent', # l'argent dépensé
                   'CPM', # le Cost-Per-Mille, le prix payé pour 1000 impressions
                   'CTR', # le Click-Through-Rate, le taux de cliques par affichage, en pourcents
                   'CPC', # le Cost-Per-Click, le coût par clique, ie spent/clicks
                   'CPA') # le Cost-Per-Acquisition, le coût par concrétisation, is spent/goals
  }
  
  # Load Weborama ad data
  cat(' * Chargement des fichiers Weborama \n')
  webo.ad <- tryCatch({
    loadWeborama.ad(client = client, dfs.names = dfs.names)
  }, error = function(cond){
    # message("Error in loading Weborama data, perhaps no data")
    message(cond);cat('\n')
  })
  
  # Load Mediarithmics ad data
  cat(' * Chargement des fichiers Mediarithmics \n')
  media.ad <- tryCatch({
    loadMediarithmics.ad(client = client, dfs.names = dfs.names)
  }, error = function(cond){
    # message("Error in loading Mediarithmics data, perhaps no data")
    message(cond);cat('\n')
  })
  
  #######################################
  ### Ajout données manquante "à la main"
  # Load Weborama goals data
  cat(' * Chargements des goals Weborama\n')
  webo.goals <- tryCatch({
    loadWeborama.goals(client = client, datadir = list.files(pattern = "^(Weborama)(.*)(.xlsx)$"), dfs.names = dfs.names)
  }, error = function(cond){
    # message("Error in loading Weborama data, perhaps no data")
    message(cond);cat('\n')
  })
  
  # Load Mediarithmics goals data
  cat(' * Chargements des goals Mediarithmics\n')
  media.goals <- tryCatch({
    loadMediarithmics.goals(datadir = list.files(pattern = "^(Media)(.*)(.xlsx)$")) %>% filter(!is.na(Date))
  }, error = function(cond){
    # message("Error in loading Weborama data, perhaps no data")
    message(cond);cat('\n')
  })
  
  #######################################
  
  # Write Excel
  cat(' * Ecriture du fichier rapport excel \n')
  args <- c(list(webo.ad = webo.ad,
                 media.ad = media.ad,
                 webo.goals = webo.goals,
                 media.goals = media.goals,
                 dfs.names = dfs.names), list(...))
  do.call(writeExcel, args)
  
}