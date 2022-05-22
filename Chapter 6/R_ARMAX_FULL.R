
rm(list=ls()) #Clear all

library("writexl")
library("sysid")
library("rugarch")
library("ggplot2")
library("astsa")
library("readxl")
library("fable")
library("scales")
library("lubridate")

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



x_windprog = replacee_na(as.double(wind_power_dk_prognosis_2021_hourly$...3[-1])) #strtoi converts strings to integers
x_temp = replacee_na(as.double(TemperatureDanmark2021_1_$...9[-1])) #strtoi converts strings to integers
x_rainfall = replacee_na(as.double(NedbrDK5zoner_1_$...10[-1])) #strtoi converts strings to integers
x_prodprog = replacee_na(as.double(production_prognosis_dk_areas_2021_hourly$...3[-1])) #new
x_consprog = replacee_na(as.double(consumption_prognosis_dk_2021_hourly$...3[-1]))
x_windprod = replacee_na(as.double(wind_power_dk_2021_hourly$...3[-1])) #is.na eliminates NA values from dataset
x_windprod = x_windprod[-length(x_windprod)]
x_production = replacee_na(as.double(production_dk_areas_2021_hourly$...3[-1]))
x_production = x_production[-length(x_production)]
x_consumption = replacee_na(as.double(consumption_dk_areas_2021_hourly$...3[-1]))
x_consumption = x_consumption[-length(x_consumption)]
Y = replacee_na(as.double(elspot_prices_2021_hourly_eur$...9[-1]),stand=FALSE) #Converting integers to double class 
Y = Y[-1]
x_oilprices = replacee_na(Oil$...2[-8760])
x_temp = x_temp[-1]
x_rainfall = x_rainfall[-1]
x_windprog = x_windprog[!is.na(x_windprog)]
x_windprog = x_windprog[-1]
x_prodprog = x_prodprog[!is.na(x_prodprog)]
x_prodprog = x_prodprog[-1]
x_consprog = x_consprog[!is.na(x_consprog)]
x_consprog = x_consprog[-1]
x_hour = 1*(matrix((0:(8783*23-1)%%24+1),8783,23) == 2)[1:8759,]

X = cbind(x_windprog,x_windprod,x_consprog,x_consumption,x_temp,x_rainfall,x_oilprices, x_hour) #combining vectors


#Q1
x = X[1:2184,]
y = Y[1:2184]


#Q2
x = X[2180:4380,]
y = Y[2180:4380]

#prewhite
x = cbind(x,x_windprod[2179:4379],x_windprod[2176:4376])

#Q3
x = X[4369:6552,]
y = Y[4369:6552]


#Q4
x = X[6553:8760,]
y = Y[6553:8760]


#Limit data, other
x = X[1:720,]
y = Y[1:720]

X = cbind(x_windprog,x_windprod,x_production,x_prodprog,x_consprog,x_consumption,x_temp,x_rainfall,x_oilprices,x_hour) #combining vectors
t(x)%*%x/sqrt(diag(t(x)%*%x)%*%t(diag(t(x)%*%x))) #correlation matrix


#ARMA
Aic = Inf
order = c(0,0)
for (i in 1:5){
  for (j in 1:5){
    model = arima(y , order = c(i,0,j) , include.mean = TRUE ,method = "ML") 
    if (model$aic < Aic){Aic = model$aic
    order = c(i,j)
    }
  }
}
order #print order
model$aic
order

#ARMAX
Aic = Inf
order = c(0,0)
for (i in 1:5){
  for (j in 1:5){
    model = arima(y , order = c(i,0,j) , include.mean = TRUE , xreg = x ,method = "ML") 
    if (model$aic < Aic){Aic = model$aic
    order = c(i,j)
    }
  }
}
order #print order
model$aic
order
order = c(5,4)
X = cbind(x_windprod, x_windprog, x_prodprog, x_production, x_consprog, x_oilprices, x_consumption, x_temp, x_rainfall)
X = cbind(x_windprod, x_windprog, x_consprog, x_oilprices, x_consumption, x_temp, x_rainfall)
X = cbind(x_prodprog, x_production, x_consprog, x_oilprices, x_consumption, x_temp, x_rainfall)

x = X[2180:4380,]
y = Y[2180:4380]

#ARMAX
model = arima(y , order = c(order[1],0,order[2]) , include.mean = TRUE , xreg = x ,method = "ML")
spec = ugarchspec(mean.model = list(armaOrder = order,include.mean = TRUE,
                                    external.regressors = as.matrix(x)), 
                  distribution = "norm",) 

#ARMA
model = arima(y , order = c(order[1],0,order[2]) , include.mean = TRUE, method = "ML") 
spec = ugarchspec(mean.model = list(armaOrder = order,include.mean = TRUE
                                  ), 
                  distribution = "norm",) 



ARMAXGARCH = ugarchfit(spec,y,solver = "hybrid") 
infocriteria(ARMAXGARCH)
fit = ARMAXGARCH@fit
shape = as.numeric(fit$coef[length(fit$coef)])
fit$coef
model$coef 



ugarchforecast(ARMAXGARCH,n.ahead=1,n.roll=0)
ARMAXGARCH_tgarch = function(X,Y, ar = 5, ma = 4, xreg, xreg_out, dist = "norm", B = 100, P = NULL, Debug = F, ARMA = FALSE){
  if(Debug){
    browser()  
  }
  library(rugarch);library(forecast);library(tseries)
  Mean_Prediction = list()
  Sigma_Prediction = list()
  Upper = list()
  Lower = list()
  Upper80 = list()
  Lower80 = list()
  
  # initialization 
  xreg = as.matrix(xreg)
  xreg_out = as.matrix(xreg_out)
  
  #Length of in and out sample
  n = length(X)
  m = length(Y)
  
  #redefinition 
  series = c(X,Y)
  U = as.matrix(rbind(xreg,xreg_out))
  rownames(U) = 1:length(U[,1])
  
  xreg_out = as.data.frame(xreg_out)
  
  # Constructing Models
  spec = ugarchspec(mean.model = list(armaOrder = c(ar,ma),include.mean = FALSE,
                                      external.regressors = U), 
                    distribution = dist)
  if (ARMA) {spec = ugarchspec(mean.model = list(armaOrder = c(ar,ma),include.mean = FALSE), 
                    distribution = dist)}
  
  for (i in 1:m){
    t_start = Sys.time()
    X_i = series[-c(1:i,((n+i):(n+m)))]
    U_i = U[-c(1:i,((n+i):(n+m))),]
    
    object = ugarchfit(spec,series, solver = "hybrid",out.sample = m-i+1)
    
    fit = object@fit
    coef = fit$coef 
    std_res = as.vector(fit$res/fit$sigma)
    forecast = ugarchforecast(object,n.ahead = 1,
                              n.roll = 0)
    
    arma_fit = Arima(X_i,order = c(ar,0,ma),include.mean = F, 
                     xreg = U_i,
                     fixed = c(coef[1:(ar+ma+length(xreg[1,]))]))
    Mean_Prediction[[i]] = forecast(arma_fit, h = 1, xreg = as.matrix(xreg_out[i,]))$mean
    if(ARMA){arma_fit = Arima(X_i,order = c(ar,0,ma),include.mean = F, 
                              fixed = c(coef[1:(ar+ma)]))
    Mean_Prediction[[i]] = forecast(arma_fit, h = 1)$mean}
    
    Sigma_Prediction[[i]] = forecast@forecast$sigmaFor
    
    
    
    
    Upper[[i]] = quantile(drop(as.vector(Sigma_Prediction[[i]]))*std_res + drop(as.vector(Mean_Prediction[[i]])),0.975)
    Lower[[i]] = quantile(drop(as.vector(Sigma_Prediction[[i]]))*std_res + drop(as.vector(Mean_Prediction[[i]])),0.025)
    
    Upper80[[i]] = quantile(drop(as.vector(Sigma_Prediction[[i]]))*std_res + drop(as.vector(Mean_Prediction[[i]])),0.9)
    Lower80[[i]] = quantile(drop(as.vector(Sigma_Prediction[[i]]))*std_res + drop(as.vector(Mean_Prediction[[i]])),0.1)
    
    t_start = Sys.time() - t_start
    
    print(paste("Prediction", i, "of",m,"took", round(t_start,3),"seconds",sep = " "))
    
  }
  return(list("Prediction" = Mean_Prediction, "Sigma" = Sigma_Prediction, "Upper" = Upper, "Lower" = Lower, "Upper80" = Upper80, "Lower80" = Lower80))
}




#Multiplot
pacf(x)

Ses = sarima(x,0,0,0,P=0,D=0,Q=0,S=24)

Ses = sarima(y,1,0,1,P=1,D=0,Q=0,S=24)

#Prewhitning 
Ses1 = sarima(x_windprog,1,0,1,P=1,D=,Q=0,S=24)
x_windprog = Ses1$fit$residuals

Ses3 = sarima(x_production,2,0,0,P=0,D=0,Q=0,S=24)

x_production = Ses3$fit$residuals

Ses4 = sarima(x_windprod,3,0,0,P=0,D=0,Q=0,S=24)
x_windprod = Ses4$fit$residuals

y = as.vector(matrix(c(Y[4:length(Y)],Y[3:(length(Y)-1)],Y[2:(length(Y)-2)],Y[1:(length(Y)-3)]),ncol = 4)%*%c(1,-Ses4$fit$coef[1],-Ses4$fit$coef[2],-Ses4$fit$coef[3]))

ccf(x_windprod[4:length(x_windprod)],y,main="CCF for elspot price and previous hour wind production")

Ses5 = sarima(x_consprog,2,0,1,P=1,D=0,Q=0,S=24)
x_consprog = Ses5$fit$residuals

Ses6 = sarima(x_consumption,2,0,1,P=1,D=0,Q=0,S=24)
x_consumption = Ses6$fit$residuals

Ses7 = sarima(x_temp,3,0,1,P=1,D=0,Q=0,S=24)
x_temp = Ses7$fit$residuals

Ses8 = sarima(x_rainfall,2,0,1,P=0,D=0,Q=0,S=24)
x_rainfall = Ses8$fit$residuals

Ses9 = sarima(x_oilprices,1,0,0,P=0,D=0,Q=0,S=24)
x_oilprices = Ses9$fit$residuals

Ses10 = sarima(x_prodprog,3,0,1,P=,D=0,Q=0,S=24)
x_prodprog = Ses10$fit$residuals

X = cbind(x_prodprog, x_production, x_consprog, x_oilprices, x_consumption, x_temp, x_rainfall)

X = cbind(x_temp, x_rainfall, x_hour) #combining vectors



  
#Q1
FT = 1
FS = 2020 
FE = 2184

#Q2
FT = 2185 
FS = 4200
FE = 4368

#Q3
FT = 4369  
FS = 6384
FE = 6552

#Q4
FT = 6553   
FS = 8592
FE = 8758


x = X[1:FE,]
y = Y[1:FE]

x = cbind(x,c(0,x_windprod[1:4367]),c(0,0,0,0,x_windprod[1:4364]))

Re = ARMAXGARCH_tgarch(X=y[FT:(FS-1)], Y=y[FS:FE], ar = 5,ma = 4, xreg = x[FT:(FS-1),], xreg_out = x[FS:FE,],ARMA = FALSE)
Re$Prediction[[1]][1]

Predict = rep(0,(FE-FS+1))
 Upper = rep(0,(FE-FS+1))
Lower = rep(0,(FE-FS+1))
Upper80 = rep(0,(FE-FS+1))
Lower80 = rep(0,(FE-FS+1))
Re$Lower[[1]][[1]]
for (i in 1:(FE-FS+1)){
  Predict[i] = Re$Prediction[[i]][1]
  Upper[i] = Re$Upper[[i]][[1]]
  Lower[i] = Re$Lower[[i]][[1]]
  Upper80[i] = Re$Upper80[[i]][[1]]
  Lower80[i] = Re$Lower80[[i]][[1]]
}

date_daily = as.Date(elspot_prices_2021_daily_eur$`Elspot Prices in EUR/MWh`, format="%m/%d/%y %H")[-1]
hours = c("01","02","03","04","05","06","07","08","09","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24")
date_hourly = paste(hours[(0:8759)%%24+1],date_daily[(0:8759)%/%24+1])
date_hourly = as_datetime(date_hourly, format="%H %Y-%m-%d")

#Forecasting
ggplot() +
  geom_line(aes( x = date_hourly[(FS-1):FE] , y = ts(y[(FS-1):FE] , start = (FS-1) ) , colour = "Observed Values"))  +
  geom_line(aes( x = date_hourly[(FS-1):FE] , y = ts ( c(y[(FS-1)], Predict) , start = (FS-1) ) , colour = "Predictions" ) ) +
  geom_ribbon(aes(x = date_hourly[FS:FE], ymin = ts(Lower,start=FS) , ymax = ts(Upper,start=FS)) , alpha = 0.2 , fill = "red") +
  geom_ribbon(aes(x = date_hourly[FS:FE], ymin = ts(Lower80,start=FS) , ymax = ts(Upper80,start=FS)) , alpha = 0.2 , fill = "red") +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Daily Prices in Eur/Mwh") +
  scale_color_manual(values = c("black", "red" ))
Lower


#Fitness

#ARMA
spec = ugarchspec(mean.model = list(armaOrder = c(5,4),include.mean = FALSE), 
                  distribution = "norm")

#ARMAX
spec = ugarchspec(mean.model = list(armaOrder = c(5,4),include.mean = FALSE,
                                    external.regressors = x[FT:FS,]), 
                  distribution = "norm")


ARMAXGARCH = ugarchfit(spec,y[FT:FS],solver = "hybrid")
fit = ARMAXGARCH@fit

fit$fitted.values

#Overlap
ggplot() +
  geom_line(aes( x = date_hourly[(FT):FS] , y = ts(y[(FT):FS] , start = (FS-1) ) , colour = "Observed Values"))  +
  geom_line(aes( x = date_hourly[(FT):FS] , y = ts (fit$fitted.values , start = (FS-1) ) , colour = "Predictions" ) ) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Daily Prices in Eur/Mwh") +
  scale_color_manual(values = c("black", "red" ))


#Performance
MAE = sum(abs(y[FS:FE]-Predict))/(FE-FS+1)

RMSE = sqrt(sum((y[FS:FE]-Predict)^2)/(FE-FS+1))

sum(Upper > y[FS:FE] & Lower < y[FS:FE])/length(Lower)

sum(Upper80 > y[FS:FE] & Lower80 < y[FS:FE])/length(Lower80)


RMSE
MAE

ARMAXGARCH = ugarchfit(spec,y,solver = "hybrid") 
infocriteria(ARMAXGARCH)

#DIAG
Ses = sarima(y[FT:(FS)],1,0,1,P=2,D=0,Q=0,S=24)

adf.test(Ses$fit$residuals) #Testing for stationarity. 

jarque.bera.test(y[FT:(FS)]-fit$fitted.values)

hist(y[FT:(FS)]-fit$fitted.values)



#Graphing variables

#prices 2021
ggplot() +
  geom_line(aes( x = date_hourly, y = ts(replacee_na(as.double(elspot_prices_2021_hourly_eur$...9[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Hourly Prices in Eur/Mwh") +
  scale_color_manual(values = c("black", "red" ))



#Daily plot with quarterly coloring 
ggplot() +
  geom_line(aes( x = date_daily, y = ts(replacee_na(as.double(elspot_prices_2021_daily_eur$...9[-1]),stand=FALSE), start = 1 ) , colour = "Observed Values")) +
  geom_ribbon(aes(x = date_daily[1:91], ymin = ts(rep(0,91),start=FS) , ymax = ts(rep(450,91),start=FS)) , alpha = 0.2 , fill = "blue") +
  geom_ribbon(aes(x = date_daily[91:183], ymin = ts(rep(0,93),start=FS) , ymax = ts(rep(450,93),start=FS)) , alpha = 0.2 , fill = "green") +
  geom_ribbon(aes(x = date_daily[183:276], ymin = ts(rep(0,94),start=FS) , ymax = ts(rep(450,94),start=FS)) , alpha = 0.2 , fill = "yellow") +
  geom_ribbon(aes(x = date_daily[276:365], ymin = ts(rep(0,90),start=FS) , ymax = ts(rep(450,90),start=FS)) , alpha = 0.2 , fill = "red") +
  geom_ribbon(aes(x = date_daily[85:91], ymin = ts(rep(0,7),start=FS) , ymax = ts(rep(450,7),start=FS)) , alpha = 0.2 , fill = "blue") +
  geom_ribbon(aes(x = date_daily[177:183], ymin = ts(rep(0,7),start=FS) , ymax = ts(rep(450,7),start=FS)) , alpha = 0.2 , fill = "green") +
  geom_ribbon(aes(x = date_daily[270:276], ymin = ts(rep(0,7),start=FS) , ymax = ts(rep(450,7),start=FS)) , alpha = 0.2 , fill = "yellow") +
  geom_ribbon(aes(x = date_daily[359:365], ymin = ts(rep(0,7),start=FS) , ymax = ts(rep(450,7),start=FS)) , alpha = 0.2 , fill = "red") +
  theme(legend.title = element_blank() ) + xlab("") + ylab("Daily Prices in Eur/Mwh") +
  scale_color_manual(values = c("black", "red" ))





