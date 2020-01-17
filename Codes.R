rm(list=ls())
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

library(xlsx)
library(plyr)
library(dplyr)
data <- read.xlsx("Data.xlsx",sheetIndex = 1,stringsAsFactors=F)
colnames(data) <- c("Date","Speed","Direction","Energy")
data$Date <- as.Date(data$Date,"%d/%m/%Y")

data$Speed <- as.numeric(data$Speed)
data$Direction <- as.numeric(data$Direction)
data$Energy <- as.numeric(data$Energy)
# Fill NA
data$Energy <- outlier.IQR(data$Energy,replace = T) 

data$Speed <- fill_na(data$Speed)
data$Direction <- fill_na(data$Direction)
data$Energy <- fill_na(data$Energy)
#================
library(ggplot2)

Fig.1 <- data %>% ggplot(aes(x=Date,y=Energy))+geom_line() + labs(y="Energy (kWh)")+
  theme_bw()
tiff("Fig.1 rawdata of energy.tiff",units = "in",height = 4,width = 6.6,res=300)
print(Fig.1)
dev.off()
#Many Outliers, remove, filled by NA
data$Energy <- outlier.IQR(data$Energy,replace=T)[[1]]
data$Energy <- fill_na(data$Energy)
Fig.2 <- data %>% ggplot(aes(x=Date,y=Energy))+geom_line() + labs(y="Energy (kWh)")+
  theme_bw()
tiff("Fig.2 clean of energy.tiff",units = "in",height = 4,width = 6.6,res=300)
print(Fig.2)
dev.off()


#Wind Speed
Fig.3 <- data %>% ggplot(aes(x=Date,y=Speed))+geom_line() + labs(y="Wind speed (m/s)")+
  theme_bw()
tiff("Fig.3 Wind speed.tiff",units = "in",height = 4,width = 6.6,res=300)
print(Fig.3)
dev.off()

#Wind direction
Fig.4 <- data %>% ggplot(aes(x=Date,y=Direction))+geom_line() + labs(y="Direction")+
  theme_bw()
tiff("Fig.4 Wind direction.tiff",units = "in",height = 4,width = 6.6,res=300)
print(Fig.4)
dev.off()

brea <- c(-1, 11.25 + (22.5*0:16))
winddir <- cut(data$Direction, brea, labels = c("N", "NNE", "NE", "ENE", "E", "ESE",
                                               "SE", "SSE", "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW", "N2"))
levels(winddir)[17] = "N"
windspeed <- cut(data$Speed, breaks = c(0, 5, 10, 15, 20))
df.dir  <- data.frame(winddir, windspeed)
rosebrea <- 1:5*max(table(winddir))/5

p3  <- ggplot(df.dir, aes(x = winddir, fill = windspeed))
p3 <- p3 + labs(x = NULL, y = NULL) + geom_bar(aes(y = ..count..))+coord_polar(start = -pi/16) + 
  scale_y_continuous(breaks = rosebrea, labels = paste(round(rosebrea/length(winddir), 2)*100, "%"))
tiff("Fig.4.1 Rose.tiff",units = "in",height = 4,width = 6.6,res=300)
print(p3)
dev.off()

#Energy vs wind speed
Fig.5 <- data %>% ggplot(aes(x=Speed,y=Energy))+geom_point() + labs(x="Wind speed (m/s)",y="Energy (kWh)")+
  theme_bw()
tiff("Fig.5 energy vs wind speed.tiff",units = "in",height = 4,width = 6.6,res=300)
print(Fig.5)
dev.off()

#Energy vs wind speed
Fig.5 <- data %>% ggplot(aes(x=Speed,y=Energy))+geom_point() +geom_smooth() + labs(x="Wind speed (m/s)",y="Energy (kWh)")+
  theme_bw()
tiff("Fig.5 energy vs wind speed.tiff",units = "in",height = 4,width = 6.6,res=300)
print(Fig.5)
dev.off()

#Energy vs wind direction
Fig.6 <- data %>% ggplot(aes(x=Speed,y=Energy))+geom_point()+geom_smooth() + labs(x="Wind direction",y="Energy (kWh)")+
  theme_bw()
tiff("Fig.6 energy vs wind direction.tiff",units = "in",height = 4,width = 6.6,res=300)
print(Fig.6)
dev.off()

#============
train_dataset <- data[1:(nrow(data)*0.7),]
test_dataset <- data[(nrow(data)*0.7+1):nrow(data),]


#=====glm: gaussian====
glm.model.train <- glm(Energy~Speed+Direction,data=train_dataset)
glm.prediction.test <- predict(glm.model.train,test_dataset)

R_square(glm.model.train$model$Energy,glm.model.train$fitted.values)#0.6238792
R_square(test_dataset$Energy,glm.prediction.test)#0.5770964

MAE(glm.model.train$model$Energy,glm.model.train$fitted.values) #26.07703
MAE(test_dataset$Energy,glm.prediction.test) #25.20899

RMSE(glm.model.train$model$Energy,glm.model.train$fitted.values) #36.65904
RMSE(test_dataset$Energy,glm.prediction.test) #38.57369

par(mfrow=c(2,2))
plot(glm.model.train)
#======Random Forest
library(randomForest)
rf.model.train <- randomForest(Energy~Speed+Direction,data=train_dataset,
                         ntree=1000,importance=TRUE,na.action=na.roughfix)
rf.prediction.test <- predict(rf.model,test_dataset)
R_square(train_dataset$Energy,rf.model.train$predicted)#0.6852775
R_square(test_dataset$Energy,rf.prediction.test)#0.6071667

MAE(train_dataset$Energy,rf.model.train$predicted) #22.10128
MAE(test_dataset$Energy,rf.prediction.test) #23.72595

RMSE(train_dataset$Energy,rf.model.train$predicted) #30.67478
RMSE(test_dataset$Energy,rf.prediction.test) #35.83092


#=======
library(caret)
xgboost.model.train <- train(
  Energy ~Speed+Direction, data = train_dataset, method = "xgbTree",
  trControl = trainControl("cv", number = 10)
)
xgboost.prediction.train <- predict(xgboost.model.train,newdata=train_dataset)
xgboost.prediction.test <- predict(xgboost.model.train,newdata=test_dataset)
R_square(train_dataset$Energy,xgboost.prediction.train)#0.6972889
R_square(test_dataset$Energy,xgboost.prediction.test)#0.6271492

MAE(train_dataset$Energy,xgboost.prediction.train) #21.42739
MAE(test_dataset$Energy,xgboost.prediction.test) #23.08835

RMSE(train_dataset$Energy,xgboost.prediction.train) #29.50408
RMSE(test_dataset$Energy,xgboost.prediction.test) #34.00829


#==Time series==
library(tseries)
library(forecast)
Energy <- ts(data$Energy,frequency = 365,start = c(2010,1,1))
plot(Energy,ylab="Energy (kWh)")

tsdisplay(Energy)
dc<-decompose(Energy)

par(mfrow=c(4,2))
plot(ma(Energy,7),main='Simple Moving Averages (k=7)')
plot(ma(Energy,30),main='Simple Moving Averages (k=30)')
plot(ma(Energy,120),main='Simple Moving Averages (k=120)')
plot(ma(Energy,365),main='Simple Moving Averages (k=365)')

Energy.train <- ts(train_dataset$Energy,frequency = 365,start = c(2010,1,1))

auto.ar <- auto.arima(Energy.train)
plot(forecast(auto.ar,1000))
#===========Functions=============#
# Fill NA
fill_na <- function(x)
{
  na.id <- which(is.na(x))
  for(i in na.id)
  {
    x[i] <- mean(c(x[i-3],x[i-2],x[i-1],x[i+1],x[i+2],x[i+3]),na.rm=T)
  }
  return(x)
}

outlier.IQR <- function(x, multiple = 1.5, replace = FALSE, revalue = NA) { 
  q <- quantile(x, na.rm = TRUE)
  IQR <- q[4] - q[2]
  x1 <- which(x < q[2] - multiple * IQR | x > q[4] + multiple * IQR)
  x2 <- x[x1]
  if (length(x2) > 0) outlier <- data.frame(location = x1, value = x2)
  else outlier <- data.frame(location = 0, value = 0)
  if (replace == TRUE) {
    x[x1] <- revalue
  }
  return(list(new.value = x, outlier = outlier))
}

R_square <- function(y_act,y_pre)
{
  Q=sum((y_act-y_pre)^2)
  V=sum(y_act^2)
  R_squ <- 1-sqrt(Q/V)
  return(R_squ)
}

MAE <- function(y_act,y_pre)
{
 APE <- abs(y_act-y_pre)
 MAE <- mean(APE)
 return(MAE)
}

RMSE <- function(y_act,y_pre)
{
  X <- (y_act-y_pre)^2
  RMSE <- sqrt(mean(X))
  return(RMSE)
}

