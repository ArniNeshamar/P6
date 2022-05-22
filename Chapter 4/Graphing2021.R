
rm(list=ls()) #Clear all

library("ggplot2")
library("readxl")


#Load all .xlxs data sheets
elspot_prices_2021_hourly_eur <- read_excel("elspot-prices_2021_hourly_eur.xlsx")
elspot_prices_2022_hourly_eur <- read_excel("elspot-prices_2022_hourly_eur.xlsx")
elspot_prices_2021_daily_eur <- read_excel("elspot-prices_2021_daily_eur.xlsx")

#Consumption
consumption_dk_areas_2021_hourly <- read_excel("consumption-dk-areas_2021_hourly.xlsx")
consumption_dk_areas_2022_hourly <- read_excel("consumption-dk-areas_2022_hourly.xlsx")
consumption_prognosis_dk_2021_hourly <- read_excel("consumption-prognosis-dk_2021_hourly.xlsx")
consumption_prognosis_dk_2022_hourly <- read_excel("consumption-prognosis-dk_2022_hourly.xlsx")

#Production
production_dk_areas_2021_hourly <- read_excel("production-dk-areas_2021_hourly.xlsx")
production_dk_areas_2022_hourly <- read_excel("production-dk-areas_2022_hourly.xlsx")
production_prognosis_dk_areas_2021_hourly <- read_excel("production-prognosis_2021_hourly.xlsx")
production_prognosis_2022_hourly_ <- read_excel("production-prognosis_2022_hourly .xlsx")

#Windpower
wind_power_dk_2021_hourly <- read_excel("wind-power-dk_2021_hourly.xlsx")
wind_power_dk_2022_hourly <- read_excel("wind-power-dk_2022_hourly.xlsx")
wind_power_dk_prognosis_2022_hourly <- read_excel("wind-power-dk-prognosis_2022_hourly.xlsx")
wind_power_dk_prognosis_2021_hourly <- read_excel("wind-power-dk-prognosis_2021_hourly.xlsx")

#Rainfall
NedbrDK5zoner_1_ <- read_excel("NedbrDK5zoner (1).xlsx")

#outside air temperature
TemperatureDanmark2021_1_ <- read_excel("TemperatureDanmark2021 (1).xlsx")

#Oilprices reintroduced post interpolation
OilDailyComplete <- read_excel("OilDailyComplete.xlsx")
OilHourlyComplete <- read_excel("OilHourlyComplete.xlsx")

#Oilprices
Oil <- read_excel("Oil.xlsx")

str(data)

LinearInterpolation = function(x){
  x_ind = c()
  for (i in 1:length(x)){
    if (!is.na(x[i])){
      x_ind = c(x_ind,i)
    }
  }
  j = 0
  for (i in 1:length(x)){
    if (is.na(x[i])){
      if (j == 0){ x[i] = x[x_ind[j+1]] }
      else if (j == length(x_ind)){ x[i] = x[x_ind[j]] }
      else{
        x[i] = ((x_ind[j+1]-i)*x[x_ind[j]] + (i-x_ind[j])*x[x_ind[j+1]])/(x_ind[j+1]-x_ind[j])
      }
    }
    else { j = j + 1 }
  }
  return(x)
}

Oil$...2 = LinearInterpolation(Oil$...2)

#Conversion
Oil = Oil[(0:8759)%/%24+1,]


HtoD = function(X) {
  #browser()
  Y = X
  n = nrow(Y)
  if (is.null(n)){n=length(Y)}
  n = n/24
  m = ncol(Y)
  if (is.null(m)){
    for(i in 1:n){
      Y[i] = sum(Y[(i:(i-1+24))])
      Y = Y[-((i+1):(i+24-1))]
    }
  }
  else{
    for(i in 1:n){
      for(j in 1:m){
        Y[i,j] =  sum(Y[(i:(i-1+24)),j])
      }
      Y = Y[-((i+1):(i+24-1)),]
    }
  }
  return(Y)
}
replacee_na = function(x,stand=TRUE,hour=FALSE){
  #if (hour) { x = HtoD(x) } #HoursToDays
  mu=mean(x[!is.na(x)])
  if (stand){
    x[!is.na(x)] = x[!is.na(x)] - mu
    x[!is.na(x)] = x[!is.na(x)]/sqrt(var(x[!is.na(x)]))
    mu=mean(x[!is.na(x)])
  }
  return(replace(x,is.na(x),mu))
}


date_daily = as.Date(elspot_prices_2021_daily_eur$`Elspot Prices in EUR/MWh`, format="%m/%d/%y %H")[-1]
hours = c("01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24")
date_hourly = paste(hours[(0:8759)%%24+1],date_daily[(0:8759)%/%24+1])
date_hourly = as_datetime(date_hourly, format="%H %Y-%m-%d")
#Graphing variables

#consumption 2021
#hourly
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(consumption_dk_areas_2021_hourly$...3[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly consumption of Mwh") +
  scale_color_manual(values = c("black", "red" ))
#daily
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(consumption_dk_areas_2021_hourly$...3[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly consumption of Mwh") +
  scale_color_manual(values = c("black", "red" ))


#consumption prognosis 2021
#hourly
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(consumption_prognosis_dk_2021_hourly$...3[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly consumption prognosis of Mwh") +
  scale_color_manual(values = c("black", "red" ))
#daily
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(consumption_prognosis_dk_2021_hourly$...3[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Daily consumption prognosis of Mwh") +
  scale_color_manual(values = c("black", "red" ))



#daily windpower 2021
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(wind_power_dk_2021_hourly$...3[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Daily Windpower production in Mwh") +
  scale_color_manual(values = c("black", "red" ))

#hourly windpower 2021
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(wind_power_dk_2021_hourly$...3[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly Windpower production in Mwh") +
  scale_color_manual(values = c("black", "red" ))



#daily windpower prognosis 2021
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(wind_power_dk_prognosis_2021_hourly$...3[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Daily Windpower production prognosis in Mwh") +
  scale_color_manual(values = c("black", "red" ))

#hourly windpower prognosis 2021
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(wind_power_dk_prognosis_2021_hourly$...3[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly Windpower production prognosis in Mwh") +
  scale_color_manual(values = c("black", "red" ))


#daily production prognosis 2021
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(production_prognosis_dk_areas_2021_hourly$...13[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Total daily production prognosis in Mwh") +
  scale_color_manual(values = c("black", "red" ))

#hourly production prognosis 2021
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(production_prognosis_dk_areas_2021_hourly$...13[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Total hourly production prognosis in Mwh") +
  scale_color_manual(values = c("black", "red" ))


#Rainfall 2021
#daily
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(NedbrDK5zoner_1_$...3[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Daily Rainfall in millimeters") +
  scale_color_manual(values = c("black", "red" ))
#hourly
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(NedbrDK5zoner_1_$...3[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly Rainfall in millimeters") +
  scale_color_manual(values = c("black", "red" ))


#Oilprices 2021 defective

ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(OilDailyComplete$...2[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Total daily production in Mwh") +
  scale_color_manual(values = c("black", "red" ))

ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(OilHourlyComplete$...2[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly Rainfall in millimeters") +
  scale_color_manual(values = c("black", "red" ))



#daily production 2021
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(production_dk_areas_2021_hourly$...3[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Total daily production in Mwh") +
  scale_color_manual(values = c("black", "red" ))

#hourly production 2021
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(production_dk_areas_2021_hourly$...3[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Total hourly production in Mwh") +
  scale_color_manual(values = c("black", "red" ))


#Temperature 2021
#daily
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(TemperatureDanmark2021_1_$...9[-1]),stand=FALSE,hour=TRUE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Daily weighted average temperature") +
  scale_color_manual(values = c("black", "red" ))
#hourly
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(TemperatureDanmark2021_1_$...9[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly weighted average temperature") +
  scale_color_manual(values = c("black", "red" ))

