rm(list=ls()) #Clear all


elspot_prices_2021_hourly_eur <- read_excel("elspot-prices_2021_hourly_eur.xlsx")
elspot_prices_2021_daily_eur <- read_excel("elspot-prices_2021_daily_eur.xlsx")


#Seasonaity of same hour
elspot = as.double(elspot_prices_2021_hourly_eur$...9[-1]) #DK1 is row 9 
plot((24*(1:31)-21),elspot[(24*(1:31)-21)],col="deepskyblue", ylim=c(9,120), xlab = "Hour", ylab = "Price in Eur/h", main = "Seasonality of hourly prices in January 2021")
lines((24*(1:31)-21),elspot[(24*(1:31)-21)],col="deepskyblue")
points((24*(1:31)-6),elspot[(24*(1:31)-6)],col="lightcoral")
lines((24*(1:31)-6),elspot[(24*(1:31)-6)],col="lightcoral")
legend("topleft", legend = c("3AM", "6AM"), col = 6:3, pch = 19, bty = "n")


date_daily = as.Date(elspot_prices_2021_daily_eur$`Elspot Prices in EUR/MWh`, format="%m/%d/%y %H")[-1]
hours = c("01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24")
date_hourly = paste(hours[(0:8759)%%24+1],date_daily[(0:8759)%/%24+1])
date_hourly = as_datetime(date_hourly, format="%H %Y-%m-%d")

elspot

#Seasonaity of same hour
elspot = elspot_prices_2021_daily_eur$...9 #DK1 is row 9 
elspot$Date <- as.Date(elspot$Date, "%m/%d/%Y") #add dates
plot(elspot, xlab = "", ylab = "Price in Eur/h", main = "Seasonality of hourly prices of 2021")
points((24*(1:30)-21),elspot[(24*(1:30)-21)],col="deepskyblue")
lines((24*(1:30)-21),elspot[(24*(1:30)-21)],col="deepskyblue")
points((24*(1:30)-6),elspot[(24*(1:30)-6)],col="lightcoral")
lines((24*(1:30)-6),elspot[(24*(1:30)-6)],col="lightcoral")
legend("topleft", legend = c("3AM", "6AM"), col = 6:3, pch = 19, bty = "n")

