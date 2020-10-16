# Loading packages
library(devtools)
library(roxygen2)

# Creating package directory
setwd("c:/Users/H52Z/Desktop/")
create("sharepointr")

# Creating documentation
setwd("./sharepointr")
document()

# Creating vignette
usethis::use_vignette("introduction")

# Installing package
setwd("..")
install("sharepointr")
