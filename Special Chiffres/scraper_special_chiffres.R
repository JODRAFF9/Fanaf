# Collecte des "Special Chiffres FANAF" (PDF) depuis fanaf.org
#
# Sources explorees :
#   1. la mediatheque WordPress du site (API wp-json) ;
#   2. les pages Statistiques et Documents, avec leurs pages suivantes ;
#   3. les anciennes adresses article_ressources, testees annee par annee ;
#   4. la Wayback Machine, pour les fichiers retires du site.
# Les PDF sont enregistres dans "Special Chiffres/documents", avec un inventaire.csv.
#
# Lancement depuis la racine du depot :  Rscript "Special Chiffres/scraper_special_chiffres.R"
# Paquets : install.packages(c("httr2", "rvest", "xml2", "stringr", "purrr"))

library(httr2)
library(rvest)
library(xml2)
library(stringr)
library(purrr)

base    <- "https://fanaf.org"
dossier <- "Special Chiffres/documents"
dir.create(dossier, recursive = TRUE, showWarnings = FALSE)

# Reconnait "special chiffres", "Special-Chiffres", "special%20chiffres", "spécial_chiffres"...
# Pour ratisser plus large : motif <- regex("chiffres", ignore_case = TRUE)
motif <- regex("sp(e|%C3%A9|é)cial([-_ ]|%20)*chiffres", ignore_case = TRUE)

`%||%` <- function(a, b) if (is.null(a)) b else a

requete <- function(url) {
  request(url) |>
    req_user_agent("Mozilla/5.0 (collecte statistiques FANAF)") |>
    req_timeout(60) |>
    req_retry(max_tries = 3) |>
    req_error(is_error = function(r) FALSE)
}

# Execute une requete ; NULL en cas d'erreur reseau
executer <- function(req, ...) tryCatch(req_perform(req, ...), error = function(e) NULL)

ok <- function(rep) !is.null(rep) && resp_status(rep) == 200

# ---------------------------------------------------------------------------
# 1. Mediatheque WordPress
# ---------------------------------------------------------------------------
media_wp <- function(terme) {
  urls <- character(); page <- 1
  repeat {
    rep <- executer(requete(paste0(base, "/wp-json/wp/v2/media")) |>
      req_url_query(search = terme, per_page = 100, page = page,
                    `_fields` = "source_url,mime_type"))
    if (!ok(rep)) break
    d <- resp_body_json(rep, simplifyVector = TRUE)
    if (length(d) == 0) break
    urls <- c(urls, d$source_url[d$mime_type == "application/pdf"])
    total <- as.integer(resp_header(rep, "X-WP-TotalPages") %||% "1")
    if (is.na(total) || page >= total) break
    page <- page + 1
  }
  urls
}

message("1. Mediatheque WordPress")
urls_api <- unique(c(media_wp("chiffres"), media_wp("special")))

# ---------------------------------------------------------------------------
# 2. Parcours des pages de listes de documents
# ---------------------------------------------------------------------------
liens <- function(url) {
  rep <- executer(requete(url))
  if (!ok(rep)) return(character())
  h <- read_html(resp_body_string(rep)) |> html_elements("a") |> html_attr("href")
  h <- h[!is.na(h)]
  unique(str_remove(url_absolute(h, url), "#.*$"))
}

message("2. Parcours du site")
a_visiter <- paste0(base, c("/cat_doc/statistiques/", "/ova_doc/",
                            "/marche-de-lassurance-en-afrique-2/", "/"))
vues <- character(); pdf_pages <- character()

while (length(a_visiter) > 0 && length(vues) < 300) {
  url <- a_visiter[1]; a_visiter <- a_visiter[-1]
  if (url %in% vues) next
  vues <- c(vues, url)
  l <- liens(url)
  pdf_pages <- c(pdf_pages, l[str_detect(l, regex("\\.pdf($|\\?)", ignore_case = TRUE))])
  suivre <- l[str_starts(l, base) &
              str_detect(l, "/(cat_doc|ova_doc)/|/page/\\d+/?$") &
              !str_detect(l, regex("\\.(pdf|jpe?g|png|zip)($|\\?)", ignore_case = TRUE))]
  a_visiter <- unique(c(a_visiter, setdiff(suivre, vues)))
  Sys.sleep(0.5)
}
message("   ", length(vues), " pages visitees")

# ---------------------------------------------------------------------------
# 3. Anciennes adresses, testees annee par annee
# ---------------------------------------------------------------------------
message("3. Anciennes adresses")
annees <- 2008:as.integer(format(Sys.Date(), "%Y"))
candidats <- c(
  paste0(base, "/article_ressources/file/", annees, "_special_chiffres_fanaf.pdf"),
  paste0(base, "/article_ressources/file/Documents%20fanaf%20", annees,
         "/special%20chiffres%20FANAF%20", annees, ".pdf")
)
# Ne demande que les 4 premiers octets (certains serveurs refusent HEAD)
existe <- function(url) {
  rep <- executer(requete(url) |> req_headers(Range = "bytes=0-3"))
  !is.null(rep) && resp_status(rep) %in% c(200, 206)
}
urls_anciennes <- keep(candidats, existe)

# ---------------------------------------------------------------------------
# 4. Wayback Machine, captures reussies seulement
# ---------------------------------------------------------------------------
message("4. Wayback Machine")
archives <- NULL
rep <- executer(requete("https://web.archive.org/cdx/search/cdx") |>
  req_url_query(url = "fanaf.org/*", output = "json", collapse = "urlkey",
                fl = "timestamp,original,mimetype,statuscode"))
if (ok(rep)) {
  d <- resp_body_json(rep, simplifyVector = TRUE)
  if (is.matrix(d) && nrow(d) > 1) {
    d <- as.data.frame(d[-1, , drop = FALSE], stringsAsFactors = FALSE)
    names(d) <- c("timestamp", "original", "mimetype", "statuscode")
    archives <- d[d$mimetype == "application/pdf" & d$statuscode == "200" &
                  str_detect(d$original, motif), c("timestamp", "original")]
  }
}

# ---------------------------------------------------------------------------
# Inventaire et telechargement
# ---------------------------------------------------------------------------
urls_site <- unique(c(urls_api, pdf_pages, urls_anciennes))
urls_site <- urls_site[str_detect(urls_site, motif)]

# Prefixe annee-mois quand l'adresse vient de wp-content/uploads/AAAA/MM/
nom_fichier <- function(url) {
  u  <- URLdecode(str_remove(url, "\\?.*$"))
  am <- str_match(u, "/uploads/(\\d{4})/(\\d{2})/")[1, 2:3]
  f  <- basename(u)
  if (!anyNA(am)) paste0(am[1], "-", am[2], "_", f) else f
}

inventaire <- data.frame(source  = rep("site", length(urls_site)),
                         url     = urls_site,
                         fichier = map_chr(urls_site, nom_fichier),
                         stringsAsFactors = FALSE)

if (!is.null(archives) && nrow(archives) > 0) {
  archives$fichier <- map_chr(archives$original, nom_fichier)
  archives <- archives[!archives$fichier %in% inventaire$fichier, ]
  if (nrow(archives) > 0) {
    inventaire <- rbind(inventaire, data.frame(
      source  = "wayback",
      url     = paste0("https://web.archive.org/web/", archives$timestamp, "id_/", archives$original),
      fichier = archives$fichier,
      stringsAsFactors = FALSE))
  }
}

telecharger <- function(url, dest) {
  if (file.exists(dest)) return("deja present")
  rep <- executer(requete(url), path = dest)
  valide <- ok(rep) && file.exists(dest) &&
            identical(readBin(dest, "raw", 4), charToRaw("%PDF"))
  if (!valide) { unlink(dest); return("echec") }
  "telecharge"
}

message("Telechargement de ", nrow(inventaire), " fichiers")
inventaire$etat <- map2_chr(inventaire$url, file.path(dossier, inventaire$fichier), telecharger)

print(inventaire[, c("source", "fichier", "etat")])
write.csv(inventaire, file.path(dossier, "inventaire.csv"), row.names = FALSE)
