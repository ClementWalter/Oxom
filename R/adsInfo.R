# Some functions for getting info from ad names

# modifier fonction avec info
getAdGroup <- function(ad){
  if(length(ad)>1){
    ad_group <- sapply(ad, getAdGroup)
  } else {
    ad_group <- switch(ad,
                       'ad_group')
  }
  return(ad_group)
}

# modifier fonction avec info
getRedirect <- function(ad){
  if(length(ad)>1){
    redirect <- sapply(ad, getRedirect)
  } else {
    redirect <- switch(ad,
                       "Noel" = "idee-cadeau-noel.html",
                       "Promos" = "Promos.html",
                       "Home" = "home",
                       "2" = "idee-cadeau-noel.html",
                       "1" = "home",
                       "3" = "Promos.html",
                       'redirect')
  }
  return(redirect)
}