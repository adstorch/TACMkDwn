library(RDCOMClient)

rmarkdown::render("TACAnnRepMarkDwn.Rmd")
system2("open","TACAnnRepMarkDwn.docx")
system2("HeaderFooterVBA.vb")

devtools::install_github("omegahat/RDCOMClient")


pathofvbscript = "H:\\Projects\\TAC\\TAC Annual Report Markdown\\HeaderFooterVBA.vbs"
shell(shQuote(normalizePath(pathofvbscript)))
