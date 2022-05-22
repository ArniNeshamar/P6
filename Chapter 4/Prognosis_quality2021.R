rm(list=ls()) #Clear all


#Consumption
consumption_dk_areas_2021_hourly <- read_excel("consumption-dk-areas_2021_hourly.xlsx")
consumption_prognosis_dk_2021_hourly <- read_excel("consumption-prognosis-dk_2021_hourly.xlsx")

#Production
production_dk_areas_2021_hourly <- read_excel("production-dk-areas_2021_hourly.xlsx")
production_prognosis_dk_areas_2021_hourly <- read_excel("production-prognosis_2021_hourly.xlsx")

#Windpower production
wind_power_dk_2021_hourly <- read_excel("wind-power-dk_2021_hourly.xlsx")
wind_power_dk_prognosis_2021_hourly <- read_excel("wind-power-dk-prognosis_2021_hourly.xlsx")



PlotFit = function(X,Y){
  y = (strtoi(Y))
  x = (strtoi(X))
  Ind = !is.na(y)&!is.na(x)
  y = y[Ind]
  x = x[Ind]
  beta = solve(t(x)%*%x)%*%t(x)%*%y
  plot(x,y)
  lines(as.vector(0:(2*max(y))),beta[1]*as.vector(0:(2*max(y))),col="red")
  return(beta)}

#par(mfrow=c(1,3))


#Windpronosis vs realized wind power (...3 for DK1, ...4 for DK2 etc.)
PlotFit(wind_power_dk_prognosis_2021_hourly$...3,wind_power_dk_2021_hourly$...3)

#Production prognosis vs realized production 
PlotFit(production_prognosis_dk_areas_2021_hourly$...13,production_dk_areas_2021_hourly$...3)

#Consumption prognosis vs realized consumption 
PlotFit(consumption_prognosis_dk_2021_hourly$...3,consumption_dk_areas_2021_hourly$...3)




