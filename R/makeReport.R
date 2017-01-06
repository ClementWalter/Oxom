#' makeReport
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
  webo.ad <- loadWeborama.ad(client = client, dfs.names = dfs.names)
  
  # Load Mediarithmics ad data
  media.ad <- loadMediarithmics.ad(client = client, dfs.names = dfs.names)
  
  #######################################
  ### Ajout données manquante "à la main"
  # Load Weborama goals data
  webo.goals <- loadWeborama.goals(client = client, datadir = list.files(pattern = "^(Weborama)(.*)(.xlsx)$"), dfs.names = dfs.names)
  
  # Load Mediarithmics goals data
  media.goals <- loadMediarithmics.goals(datadir = list.files(pattern = "^(Media)(.*)(.xlsx)$")) %>% filter(!is.na(Date))
  #######################################
  
  # Write Excel
  writeExcel(webo.ad,
             media.ad,
             webo.goals,
             media.goals,
             dfs.names = dfs.names,
             ...
  )
  
}