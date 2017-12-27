#Installing REQUIRED Packages
install.packages("pacman")
library(pacman)
pacman::p_load(dplyr,tidyr,ggplot2,tidyverse,forecast,timeSeries,forecastHybrid,readxl,rpart,fpp,
               rpart.plot,RColorBrewer,party, partykit,caret,lubridate,zoo,
               TTR,sweep, prophet,timetk, reshape, reshape2, readxl, matrixStats, RODBC ,odbc)


############################################# 1.1.   Training Data Preparation
#Importing UNIT LEVEL SKILL DATA
channel <- odbcDriverConnect("")

UNIT_FTE <- sqlQuery(channel,paste("SELECT * 
                                 FROM [SkillDemand].[dbo].[Unit_Skill_concat_Month]
                                 "))

#Subsetting CIS Unit - HAS TO BE CHANGED IN SQL - CURRENTLY HAS WHITESPACE
UNIT_FTE_CIS <- subset(UNIT_FTE,unit == "CIS                 ")

#Creating Year Month Flags - For aggregation
UNIT_FTE_CIS$year_month<-paste0(UNIT_FTE_CIS$iyear,"-", UNIT_FTE_CIS$txtmonth,"-01")
UNIT_FTE_CIS$year_month <- as.Date(UNIT_FTE_CIS$year_month,format = "%Y-%h-%d")

#subsetting only required months - Removing extra months. "2017-12-01" should be changed to "2018-01-01"
#once december data is available "2017-12-01" should be changed to "2018-01-01"
UNIT_FTE_CIS<-subset(UNIT_FTE_CIS, year_month >= "2014-10-01" &  year_month < "2017-12-01")

#Creating month year grouping totals - pivoting the data with skill id in rows and month lvel total in columns
UNIT_FTE_CIS_my<- UNIT_FTE_CIS %>% group_by(skillid,year_month)%>%
                                     summarise(sum_fte = sum(Sum_fte)) %>%
                                     spread(year_month,sum_fte)


#Creating year totals
UNIT_FTE_CIS_y<- UNIT_FTE_CIS %>% group_by(skillid,iyear)%>%
                                    summarise(sum_fte = sum(Sum_fte)) %>%
                                    spread(iyear,sum_fte)

#Joining Year and month year
UNIT_FTE_CIS_Final<- left_join(UNIT_FTE_CIS_y, UNIT_FTE_CIS_my, by = "skillid")

#Replacing NA with 0
UNIT_FTE_CIS_Final[is.na(UNIT_FTE_CIS_Final)]<-0

#last 12 month total
UNIT_FTE_CIS_Final$Total_12Months <-((rowSums(UNIT_FTE_CIS_Final[,(ncol(UNIT_FTE_CIS_Final)-11):
                                                                       (ncol(UNIT_FTE_CIS_Final))]))) 


#24-13 month total
UNIT_FTE_CIS_Final$Total_12Months_2 <-((rowSums(UNIT_FTE_CIS_Final[,(ncol(UNIT_FTE_CIS_Final)-24):
                                                                       (ncol(UNIT_FTE_CIS_Final)-13)]))) 

#last 12 month avg
UNIT_FTE_CIS_Final$Avg_12Months <- UNIT_FTE_CIS_Final$Total_12Months / 12

#UNIT LEVEL FTE DATA OUTPUT
write.csv(UNIT_FTE_CIS_Final,"Module1.1.csv")


######################1.2 Skill Classification

#Volume Percentage                          
UNIT_FTE_CIS_Final$Volume_Percentage<-UNIT_FTE_CIS_Final$Total_12Months/sum(UNIT_FTE_CIS_Final$Total_12Months)

#Sort Function
sort_df_de<-function(data,vars = names(data), decreasing = F){
  if (length(vars) == 0 || is.null(vars))
  return(data)
  data[do.call("order", list(what = data[,vars, drop = FALSE],decreasing= decreasing)),,drop = FALSE]
}

###Volume Sort
UNIT_FTE_CIS_Final<- sort_df_de(as.data.frame(UNIT_FTE_CIS_Final),vars = "Volume_Percentage",decreasing = T)


##Volume %Calculation
UNIT_FTE_CIS_Final$Cumulative_Volume_Percentage<-cumsum(UNIT_FTE_CIS_Final$Volume_Percentage)

#Volume Flag Calculation - Currently set at 0.60 for CIS.
UNIT_FTE_CIS_Final$Volume_Flag<- ifelse(UNIT_FTE_CIS_Final$Cumulative_Volume_Percentage < 0.60, "L","S")

##SLope calc
for (i in 1:nrow(UNIT_FTE_CIS_Final)){
  x<-as.factor(month(colnames(UNIT_FTE_CIS_Final[c("2016-12-01","2017-01-01","2017-02-01",
                                                   "2017-03-01","2017-04-01","2017-05-01",
                                                   "2017-06-01","2017-07-01","2017-08-01",
                                                   "2017-09-01","2017-10-01","2017-11-01")])))
  
  y<-as.numeric((UNIT_FTE_CIS_Final[i,c("2016-12-01","2017-01-01","2017-02-01",
                                        "2017-03-01","2017-04-01","2017-05-01",
                                        "2017-06-01","2017-07-01","2017-08-01",
                                        "2017-09-01","2017-10-01","2017-11-01")]))
  lm_s<-lm(y~x)
  UNIT_FTE_CIS_Final$Slope<-UNIT_FTE_CIS_Final$Volume
  UNIT_FTE_CIS_Final[i,"Slope"]<-lm_s$coefficients[1]
}
UNIT_FTE_CIS_Final$Slope_Flag<- ifelse(UNIT_FTE_CIS_Final$Slope < 0 ,"N","P")

#growth flag
UNIT_FTE_CIS_Final$Growth <- (UNIT_FTE_CIS_Final$Total_12Months/ UNIT_FTE_CIS_Final$Total_12Months_2)-1
UNIT_FTE_CIS_Final$Growth[is.infinite(UNIT_FTE_CIS_Final$Growth)]<-1
UNIT_FTE_CIS_Final$Growth[is.nan(UNIT_FTE_CIS_Final$Growth)]<-0
UNIT_FTE_CIS_Final$Growth_Flag<-ifelse(UNIT_FTE_CIS_Final$Growth > 0.02, "P","N")


#volatality flag - Volatality cutoff set at 0.5. Can be customzied below.
UNIT_FTE_CIS_Final$Volatality <- rowSds(as.matrix(UNIT_FTE_CIS_Final[,c("2016-12-01","2017-01-01","2017-02-01",
                                                                        "2017-03-01","2017-04-01","2017-05-01",
                                                                        "2017-06-01","2017-07-01","2017-08-01",
                                                                        "2017-09-01","2017-10-01","2017-11-01")]))/rowMedians(as.matrix(UNIT_FTE_CIS_Final[,c("2016-12-01","2017-01-01","2017-02-01",
                                                                                                                                                              "2017-03-01","2017-04-01","2017-05-01",
                                                                                                                                                              "2017-06-01","2017-07-01","2017-08-01",
                                                                                                                                                              "2017-09-01","2017-10-01","2017-11-01")]))
UNIT_FTE_CIS_Final$Volatality[is.infinite(UNIT_FTE_CIS_Final$Volatality)]<-0
UNIT_FTE_CIS_Final$Volatality[is.nan(UNIT_FTE_CIS_Final$Volatality)]<-0
UNIT_FTE_CIS_Final$Volatality_Flag <- ifelse(UNIT_FTE_CIS_Final$Volatality >= 0.5, "U","S")



###Ranking by volume %
UNIT_FTE_CIS_Final$SKILL_RANK<-rank((-UNIT_FTE_CIS_Final$Volume_Percentage),
                                      ties.method = "first")
###Adding forecast flag based on rank
UNIT_FTE_CIS_Final$Forecast_Flag <-paste0("LS",UNIT_FTE_CIS_Final$SKILL_RANK)

#UNIT LEVEL FTE DATA OUTPUT WITH FLAGS
write.csv(UNIT_FTE_CIS_Final,"Module1.2.csv")

###editing forecast flag such that LARGE Volume have ranked forecast,
#Small vol but stable under "SS, SMALL VOL BUT UNSTABLE under "SU",Large Unstable under "LU",
UNIT_FTE_CIS_Final$Forecast_Flag<-ifelse(UNIT_FTE_CIS_Final$Volume_Flag == "L",            
                                         UNIT_FTE_CIS_Final$Forecast_Flag,
                                         ifelse(UNIT_FTE_CIS_Final$Volatality_Flag == "U","SU","SS"))

UNIT_FTE_CIS_Final$Forecast_Flag<-ifelse((UNIT_FTE_CIS_Final$Volume_Flag == "L" & UNIT_FTE_CIS_Final$Volatality_Flag == "U"),            
                                         "LU",
                                         UNIT_FTE_CIS_Final$Forecast_Flag)


###REFLAGGING LS Skills other than top50 as LSG AND LSD based on growth
UNIT_FTE_CIS_Final$Forecast_Flag<-ifelse((UNIT_FTE_CIS_Final$SKILL_RANK > 50 & 
                                            UNIT_FTE_CIS_Final$Growth_Flag == "P" &
                                             UNIT_FTE_CIS_Final$Volume_Flag == "L"),            
                                         "LSG",
                                         UNIT_FTE_CIS_Final$Forecast_Flag)

UNIT_FTE_CIS_Final$Forecast_Flag<-ifelse((UNIT_FTE_CIS_Final$SKILL_RANK > 50 & 
                                            UNIT_FTE_CIS_Final$Growth_Flag == "N" &
                                            UNIT_FTE_CIS_Final$Volume_Flag == "L"),            
                                         "LSD",
                                         UNIT_FTE_CIS_Final$Forecast_Flag)


#UNIT LEVEL FTE DATA OUTPUT WITH CLASSIFIED SKILLS
write.csv(UNIT_FTE_CIS_Final,"Module1.3.csv")




############################1.4 Summarising at model input level
###new table with forecast flags and skill demands
UNIT_FTE_CIS_Final_2<-UNIT_FTE_CIS_Final[,c("Forecast_Flag",names(UNIT_FTE_CIS_Final[grepl("-01",names(UNIT_FTE_CIS_Final))]))]  

##Skill table for reference
SKILL_TABLE<-UNIT_FTE_CIS_Final[,c("skillid","Forecast_Flag")]

#summing at forecast flag level
UNIT_FTE_CIS_Final_3<- data.frame(UNIT_FTE_CIS_Final_2 %>%
                           group_by(Forecast_Flag) %>%
                           summarise_all(funs(sum)))

UNIT_FTE_CIS_Final_3<-unique(UNIT_FTE_CIS_Final_3)


###creating overall demand
Data_PU <- data.frame(t(data.frame(c(Forecast_Flag = "ALL",
                                     (colSums(UNIT_FTE_CIS_Final_3[,-c(1)]))))))

#converting factor to numeric for demands , factor to character for forecast flag
Data_PU[,-c(1)] <- lapply(Data_PU[,-c(1)], function(x) as.numeric(as.character(x)))
Data_PU[,c(1)] <- lapply(Data_PU[,c(1)], function(x) as.character(x))

#renaming it to base table names for rbind
colnames(Data_PU)<-names(UNIT_FTE_CIS_Final_3)

#rbinding overall demand and Forecast flag demand
UNIT_FTE_CIS_Final_5<-rbind(Data_PU ,UNIT_FTE_CIS_Final_3)


#Time series data - Final
write.csv(UNIT_FTE_CIS_Final_5,"Module1.4.csv")




############################################## 1.5 Independant Variables 
###Demand Count
channel <- odbcDriverConnect("")

DemandCount<- sqlQuery(channel,paste("SELECT * 
                                   FROM [SkillDemand].[dbo].[PU_Skill_Month]
                                   "))

###Unit subset - WHitespace to be removed
DemandCount<-subset(DemandCount,unit == 'CIS                 ')

##creating year month flag for joining with the base table
DemandCount$Year_Month<-paste0(DemandCount$iyear,"-", DemandCount$txtmonth,"-01")
DemandCount$Year_Month <- as.Date(DemandCount$Year_Month,format = "%Y-%h-%d")

##summing at yearmonth level 
DemandCount_Grouped<- DemandCount %>% group_by(intskillid,Year_Month)%>%
                      summarise(Demand = sum(count_demands)) %>%
                      spread(Year_Month,Demand)

### skillid level demand count at month level
DemandCount_Grouped <- cbind(skillid = DemandCount_Grouped$intskillid,
                             DemandCount_Grouped[,names(UNIT_FTE_CIS_Final[grepl("-01",names(UNIT_FTE_CIS_Final))])])

###replacing na with 0
DemandCount_Grouped[is.na(DemandCount_Grouped)]<-0

#creating overall demand for that unit
Data_PU_Demand <- data.frame(t(data.frame(c(skillid = "ALL",
                                            (colSums(DemandCount_Grouped[,names(DemandCount_Grouped[grepl("-01",names(DemandCount_Grouped))])]))))))


#converting factor to numeric for demands , factor to character for forecast flag
Data_PU_Demand[-1] <- lapply(Data_PU_Demand[-1], function(x) as.numeric(as.character(x)))
Data_PU_Demand[1] <- lapply(Data_PU_Demand[1], function(x) as.character(x))

#renaming it to base table names for rbind
colnames(Data_PU_Demand)<-names(DemandCount_Grouped)

###rbding it with base demand column
DemandCount_Grouped<-rbind(Data_PU_Demand,DemandCount_Grouped)



DemandCount_Grouped2<-left_join(DemandCount_Grouped,SKILL_TABLE,by = "skillid")

####creating final demand count with the overall demand as IV for all skills  - CIS
DemandCount_Final<-data.frame(Month = colnames(DemandCount_Grouped[,names(DemandCount_Grouped[grepl("-01",names(DemandCount_Grouped))])]),
                                              (t(DemandCount_Grouped[1,names(DemandCount_Grouped[grepl("-01",names(DemandCount_Grouped))])])))
colnames(DemandCount_Final)[2]<-"Demand"
DemandCount_Final$Month<-as.Date(DemandCount_Final$Month)


                                
####Creating SKILL LEVEL DF

#Removing year level totals
#UNIT_FTE_CIS_Final_5<-UNIT_FTE_CIS_Final_5[,-c(2,3,4,5)]

#pivoting it to forecasting flag, month , employee level
UNIT_FTE_CIS_Final_6<-data.frame(UNIT_FTE_CIS_Final_5 %>% gather(Forecast_Flag))

#renaming columns
colnames(UNIT_FTE_CIS_Final_6)<-c("Forecast_Flag","Month","FTE")
#converting month to date
UNIT_FTE_CIS_Final_6$Month<-as.Date(UNIT_FTE_CIS_Final_6$Month,"X%Y.%m.%d")

#joining demand count and fte table by month  - TBC for IVS AND OTHERS
UNIT_FTE_CIS_Final_7<-left_join(UNIT_FTE_CIS_Final_6,DemandCount_Final,by = "Month")

#######SPLITTING DATA by list of dataframes for each forecast flag
UNIT_FTE_CIS_Final_8<-split(UNIT_FTE_CIS_Final_7,f = UNIT_FTE_CIS_Final_7$Forecast_Flag)





###Contract Type
channel <- odbcDriverConnect("")

Contract_Type <-  sqlQuery(channel,paste("SELECT * 
                                         FROM [SkillDemand].[dbo].[PU_unit_contractcode_view]
                                         "))
##whitespace to be removed
Contract_Type<-subset(Contract_Type,unit == 'CIS                 ')


####Creating year month flag
Contract_Type$Year_Month<- paste0(Contract_Type$iyear,"-",Contract_Type$txtmonth,"-01")

##converting to date format
Contract_Type$Year_Month<-as.Date(Contract_Type$Year_Month,format = "%Y-%h-%d")

###summing contract type at unit, year month level
Contract_Type_Grouped<-Contract_Type %>%
  group_by(unit,Year_Month,contractcode)%>%
  summarise(Contract_Count = sum(Count_project)) %>% 
  spread(contractcode,sum(Contract_Count))

Contract_Type_IV<-as.data.frame(Contract_Type_Grouped)

###replacing na with 0
Contract_Type_IV[is.na(Contract_Type_IV)]<-0


##creating CTM%
Contract_Type_IV$CTM_Perct<-Contract_Type_Grouped$`CTM  `/(Contract_Type_Grouped$`CTM  ` + 
                                                             Contract_Type_Grouped$`T&M  ` + 
                                                             Contract_Type_Grouped$`FP   `)
##CREATING fp%
Contract_Type_IV$FP_Perct<-Contract_Type_Grouped$`FP   `/(Contract_Type_Grouped$`CTM  ` + 
                                                            Contract_Type_Grouped$`T&M  `+ 
                                                            Contract_Type_Grouped$`FP   `)


##creating T AND M %
Contract_Type_IV$T_AND_M_Perct<-Contract_Type_Grouped$`T&M  `/(Contract_Type_Grouped$`CTM  ` + 
                                                                 Contract_Type_Grouped$`FP   ` + 
                                                                 Contract_Type_Grouped$`T&M  `)



##MAPPING THE contract type % to all units
for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]]<-left_join(UNIT_FTE_CIS_Final_8[[i]],Contract_Type_IV[c("Year_Month","CTM_Perct",
                                                                                    "FP_Perct","T_AND_M_Perct")],
                                       by = c("Month" ="Year_Month"))
}



###Revenue and RTBR - Similar steps as contract type
channel <- odbcDriverConnect("")

Revenue<- sqlQuery(channel,paste("SELECT * 
                                 FROM [SkillDemand].[dbo].[UNIT_REVENUE]
                                 "))

Revenue$Year_Month<-paste0(Revenue$Year,"-",Revenue$Month, "-01")

Revenue$Year_Month<-as.Date(Revenue$Year_Month,format = "%Y-%h-%d")

Revenue_Grouped<-Revenue %>% group_by(Unit,Year_Month)%>%
  summarise(Revenue = sum(Revenue))

Revenue_Grouped<-subset(Revenue_Grouped,Unit  == "CIS")

RTBR<-sqlQuery(channel,paste("SELECT * 
                             FROM [SkillDemand].[dbo].[UNIT_RTBR]
                             "))
RTBR$Year_Month<-paste0(RTBR$Year,"-",RTBR$Month,"-01")

RTBR$Year_Month<-as.Date(RTBR$Year_Month,format = "%Y-%m-%d")

RTBR_Grouped<-RTBR %>% group_by(Unit,Year_Month)%>%
  summarise(Revenue = sum(RTBR)/1000)

RTBR_Grouped<-subset(RTBR_Grouped,Unit  == "CIS")

Revenue_RTBR<- rbind(Revenue_Grouped, RTBR_Grouped)

###replacing na with 0
Revenue_RTBR[is.na(Revenue_RTBR)]<-0


for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]]<-left_join(UNIT_FTE_CIS_Final_8[[i]],Revenue_RTBR[,c("Year_Month","Revenue")],
                                       by = c("Month" ="Year_Month"))
}

######creating revenue revised column
for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]]$Revenue_Revised<-UNIT_FTE_CIS_Final_8[[i]]$Revenue
}







####Opportunity  - Same as revene variable
channel <- odbcDriverConnect("")

Opportunity<- sqlQuery(channel,paste("SELECT * 
                                 FROM [SkillDemand].[dbo].[UNIT_OPPORTUNITY]
                                 "))

Opportunity<-subset(Opportunity,Unit == "CIS")

Opportunity$Year_Month<- paste0(Opportunity$Year,"-",Opportunity$Month,"-01")


Opportunity$Year_Month<-as.Date(Opportunity$Year_Month,format = "%Y-%h-%d")

Opportunity_Grouped<-Opportunity %>% group_by(Year_Month)%>%
  summarise(Opportunity = sum(OppValue))

###creating 3 month opportunity flag
Opportunity_MA3<- c(rollmean(Opportunity_Grouped$Opportunity,3),0,0)

Opportunity_Grouped$Opportunity_MA3<-Opportunity_MA3


for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]]<-left_join(UNIT_FTE_CIS_Final_8[[i]],
                                       Opportunity_Grouped,by = c("Month" ="Year_Month"))
}



#######creating rampupdownabs variable using FTE
for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]]$Rampupdownabs<-c(mean(diff(UNIT_FTE_CIS_Final_8[[i]]$FTE,lag = 1)),
                                             diff(UNIT_FTE_CIS_Final_8[[i]]$FTE,lag = 1))
}




#######creating rampupdowndel variable using FTE DELTA
for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]]$Rampupdowndel<-UNIT_FTE_CIS_Final_8[[i]]$Rampupdownabs/UNIT_FTE_CIS_Final_8[[i]]$FTE
}




#######creating demand  variable using demnand
for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]]$Demand_del<-UNIT_FTE_CIS_Final_8[[i]]$Demand/UNIT_FTE_CIS_Final_8[[i]]$FTE
}

#####creatign month column
for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]]$Mx<-month(UNIT_FTE_CIS_Final_8[[i]]$Month)
}


##############Combine all data frames adn print o/p
UNIT_FTE_CIS_Final_8_Combined<-do.call("rbind",UNIT_FTE_CIS_Final_8)
write.csv(UNIT_FTE_CIS_Final_8,"Module1.5.csv")

########################outlier module


###test_data  -- next 3 month variables for regression, arimax models
for (i in 1 : length(UNIT_FTE_CIS_Final_8)){
  UNIT_FTE_CIS_Final_8[[i]] <- rbind(UNIT_FTE_CIS_Final_8[[i]],data.frame(Forecast_Flag = UNIT_FTE_CIS_Final_8[[i]][1:3,1],
                                                        Month = seq((UNIT_FTE_CIS_Final_8[[i]][nrow(UNIT_FTE_CIS_Final_8[[i]]),2]+30),by = "months",length.out = 3),
                                                        FTE = NA,
                                                        Demand = NA,
                                                        CTM_Perct = NA,
                                                        FP_Perct = NA,
                                                        T_AND_M_Perct = NA,
                                                        Revenue = NA,
                                                        Revenue_Revised = NA,
                                                        Opportunity = NA,
                                                        Opportunity_MA3 = NA,
                                                        Rampupdownabs = NA,
                                                        Rampupdowndel = NA,
                                                        Demand_del = NA,
                                                        Mx = month(UNIT_FTE_CIS_Final_8[[i]][(nrow(UNIT_FTE_CIS_Final_8[[i]])-2):nrow(UNIT_FTE_CIS_Final_8[[i]]),2])))
}


#Adding test variable values
for(i in 1:length(UNIT_FTE_CIS_Final_8)){
  for (j in 39:41){ 
    for (k in c(4:14)){ 
      test_cr <- UNIT_FTE_CIS_Final_8[[i]]
      test_cr[j,k] <- mean(test_cr[((j-13):(j-2)),k])
      UNIT_FTE_CIS_Final_8[[i]] <- test_cr
    }
  }
}




set.seed(1000)


#Creating Empty Data Frames
Future_Preds_Best_Fit_Final_Cons<-data.frame(Skill = character(),Month = as.Date(character()),
                                             Fitted = integer(),Fitted_Low = integer(),Fitted_High = integer())

metrics_consolidated_outtime_2<-data.frame(Model = character(),
                                           SKILL = as.character(),
                                           MAPE = numeric(),
                                           RMSE = numeric(),
                                           MAPE_1stMonth = numeric(),
                                           MAPE_2ndMonth = numeric(),
                                           MAPE_3rdMonth = numeric(),
                                           MAPE_1stMonth_Rank = numeric(),
                                           MAPE_2ndMonth_Rank = numeric(),
                                           MAPE_3rdMonth_Rank = numeric())



Outtime_Preds_Metrics_Avg_Cons<-data.frame(SKILL = as.character(),
                                           MAPE = numeric(),
                                           RMSE = numeric(),
                                           MAPE_1stMonth = numeric(),
                                           MAPE_2ndMonth = numeric(),
                                           MAPE_3rdMonth = numeric())

Last_Month_Actuals_Cons <-data.frame(SKILL = as.character(),
                                     Actual_1 = integer(),
                                     Actual_2 = integer(),
                                     Actual_3 = integer())


#Model Loop Starts
for (i in 1:length(UNIT_FTE_CIS_Final_8)){
  p_0<-as.data.frame(UNIT_FTE_CIS_Final_8[[i]])
  p_0$Rampupdowndel[is.infinite(p_0$Rampupdowndel)]<-1
  p_0$Rampupdowndel[is.nan(p_0$Rampupdowndel)]<-0
  p_0$Demand_del[is.infinite(p_0$Demand_del)]<-1
  p_0$Demand_del[is.nan(p_0$Demand_del)]<-0
  p_0$Opportunity[is.na(p_0$Opportunity)]<-mean(p_0$Opportunity,na.rm = TRUE)
  p_0$Opportunity_MA3[is.na(p_0$Opportunity_MA3)]<-mean(p_0$Opportunity_MA3, na.rm = TRUE)
  
  #######Creating ts data for selecting best models - last 3 months kept as outtime sample
  ts_data<-ts(p_0[,3],frequency = 12,start = c(2014,10), end = c(2017,8))
  
  #######Creating ts data for predicting future 3 months using whole dataset
  ts_data_2 <-ts(p_0[,3],frequency = 12,start = c(2014,10), end = c(2017,11))
  
  ts_data_components <- decompose(ts_data)
  ts_data_sa<- ts_data - ts_data_components$seasonal  
  ts_data_components_2 <- decompose(ts_data_2)
  ts_data_sa_2<- ts_data_2 - ts_data_components_2$seasonal  
  
  
  #SES - ETS Model - Best of 30 Combination
  lambdas <- BoxCox.lambda(ts_data_sa)
  
  #Train
  ts_ets<-forecast(ts_data_sa,lambda = lambdas)
  
  #Finind in time forecasts
  intime_sample_ets<-data.frame(Model = "ETS",
                                Skill = rep(p_0[1,1],length(ts_data)),
                                Month =  p_0[1:length(ts_data),2],
                                Actual = p_0[1:length(ts_data),3],
                                Fitted = ts_ets$fitted)
  
  #renaming column names
  colnames(intime_sample_ets)<-c('Model',"Skill","Month","Actual","Fitted")
  
  #Finding intime MAPE
  metrics_intime_ets<-data.frame(Model = "ETS",
                                 MAPE = MAPE(intime_sample_ets$Actual,intime_sample_ets$Fitted,percent = FALSE),
                                 RMSE = RMSE(intime_sample_ets$Actual,intime_sample_ets$Fitted))
  
  #Forecasting outtime samples
  forecasted_values_ets<-data.frame(forecast(ts_ets,h = 3))
  
  #Calculating outtime MAPE
  outtime_sample_ETS<-data.frame(Model = "ETS",
                                 Skill = rep(p_0[1,1],3),
                                 Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                 Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                 Fitted = forecasted_values_ets[1:3,1],
                                 Fitted_Low = forecasted_values_ets[1:3,2],
                                 Fitted_High = forecasted_values_ets[1:3,3])
  
  #Renaming the columns
  colnames(outtime_sample_ETS)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  #calculating outtime metrics
  metrics_outtime_ets<-data.frame(Model = "ETS",
                                  SKILL = p_0[1,1],
                                  MAPE = MAPE(outtime_sample_ETS$Actual,outtime_sample_ETS$Fitted,percent = FALSE),
                                  RMSE = RMSE(outtime_sample_ETS$Actual,outtime_sample_ETS$Fitted),
                                  MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ETS$Actual,outtime_sample_ETS$Fitted)[1],
                                  MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ETS$Actual,outtime_sample_ETS$Fitted)[2],
                                  MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ETS$Actual,outtime_sample_ETS$Fitted)[3])
  
  ##Future predictions model for full training data
  lambdas <- BoxCox.lambda(ts_data_sa_2)
  
  ts_ets_2<-forecast(ts_data_sa_2,lambda = lambdas)
  
  forecasted_values_ets_2<-data.frame(forecast(ts_ets_2,h = 3))
  
  Future_predictions_ETS<-data.frame(Model = "ETS",
                                     Skill = rep(p_0[1,1],3),
                                     Month = rownames(forecasted_values_ets_2[1:3,]),
                                     Fitted = forecasted_values_ets_2[1:3,1],
                                     Fitted_Low = forecasted_values_ets_2[1:3,2],
                                     Fitted_High = forecasted_values_ets_2[1:3,3])
  
  
  #arima
  lambdas <- BoxCox.lambda(ts_data)
  
  ts_arima<-auto.arima(ts_data,max.p = 10, max.q = 10, max.order = 2, 
                       max.d = 10, start.p = 0,
                       start.q = 0, 
                       stationary = FALSE,
                       seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                       stepwise = TRUE,
                       trace = TRUE,
                       truncate = NULL, test = c("kpss", "adf", "pp"),
                       seasonal.test = c("ocsb", "ch"), 
                       allowdrift = TRUE, allowmean = TRUE,
                       lambda = lambdas, biasadj = FALSE,
                       parallel = FALSE, num.cores = NULL)
  
  
  intime_sample_arima<-data.frame(Model = "ARIMA",
                                  Skill = rep(p_0[1,1],length(ts_data)),
                                  Month =  p_0[1:length(ts_data),2],
                                  Actual = p_0[1:length(ts_data),3],
                                  Fitted = ts_arima$fitted)
  
  
  colnames(intime_sample_arima)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_arima<-data.frame(Model = "ARIMA",
                                   MAPE = MAPE(intime_sample_arima$Actual,intime_sample_arima$Fitted,percent = FALSE),
                                   RMSE = RMSE(intime_sample_arima$Actual,intime_sample_arima$Fitted))
  
  
  forecasted_values_arima<-data.frame(forecast(ts_arima,h = 3))
  
  outtime_sample_ARIMA<-data.frame(Model = "ARIMA",
                                   Skill = rep(p_0[1,1],3),
                                   Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                   Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                   Fitted = forecasted_values_arima[1:3,1],
                                   Fitted_Low = forecasted_values_arima[1:3,2],
                                   Fitted_High = forecasted_values_arima[1:3,3])
  
  colnames(outtime_sample_ARIMA)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_arima<-data.frame(Model = "ARIMA",
                                    SKILL = p_0[1,1],
                                    MAPE = MAPE(outtime_sample_ARIMA$Actual,outtime_sample_ARIMA$Fitted,percent = FALSE),
                                    RMSE = RMSE(outtime_sample_ARIMA$Actual,outtime_sample_ARIMA$Fitted),
                                    MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ARIMA$Actual,outtime_sample_ARIMA$Fitted)[1],
                                    MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ARIMA$Actual,outtime_sample_ARIMA$Fitted)[2],
                                    MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ARIMA$Actual,outtime_sample_ARIMA$Fitted)[3])
  
  lambdas <- BoxCox.lambda(ts_data_sa_2)
  
  ts_arima_2<-auto.arima(ts_data_2,max.p = 10, max.q = 10, max.order = 2, 
                         max.d = 10, start.p = 0,
                         start.q = 0, 
                         stationary = FALSE,
                         seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                         stepwise = TRUE,
                         trace = TRUE,
                         truncate = NULL, test = c("kpss", "adf", "pp"),
                         seasonal.test = c("ocsb", "ch"), 
                         allowdrift = TRUE, allowmean = TRUE,
                         lambda = lambdas, biasadj = FALSE,
                         parallel = FALSE, num.cores = NULL)
  
  
  forecasted_values_arima_2<-data.frame(forecast(ts_arima_2,h = 3))
  
  
  Future_predictions_ARIMA<-data.frame(Model = "ARIMA",
                                       Skill = rep(p_0[1,1],3),
                                       Month = rownames(forecasted_values_arima_2[1:3,]),
                                       Fitted = forecasted_values_arima_2[1:3,1],
                                       Fitted_Low = forecasted_values_arima_2[1:3,2],
                                       Fitted_High = forecasted_values_arima_2[1:3,3])
  #Hybrid
  lambdas <- BoxCox.lambda(ts_data_sa)
  
  ts_hybrid<-hybridModel(ts_data_sa,models = "aefst", lambda = lambdas)
  
  
  intime_sample_hybrid<-data.frame(Model = "HYBRID",
                                  Skill = rep(p_0[1,1],length(ts_data)),
                                  Month =  p_0[1:length(ts_data),2],
                                  Actual = p_0[1:length(ts_data),3],
                                  Fitted = ts_hybrid$fitted)
  
  
  colnames(intime_sample_hybrid)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_hybrid<-data.frame(Model = "HYBRID",
                                   MAPE = MAPE(intime_sample_hybrid$Actual,intime_sample_hybrid$Fitted,percent = FALSE),
                                   RMSE = RMSE(intime_sample_hybrid$Actual,intime_sample_hybrid$Fitted))
  
  
  forecasted_values_hybird<-data.frame(forecast(ts_hybrid,h = 3))
  
  outtime_sample_HYBRID<-data.frame(Model = "HYBRID",
                                   Skill = rep(p_0[1,1],3),
                                   Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                   Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                   Fitted = forecasted_values_hybird[1:3,1],
                                   Fitted_Low = forecasted_values_hybird[1:3,2],
                                   Fitted_High = forecasted_values_hybird[1:3,3])
  
  colnames(outtime_sample_HYBRID)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_HYBRID<-data.frame(Model = "HYBRID",
                                    SKILL = p_0[1,1],
                                    MAPE = MAPE(outtime_sample_HYBRID$Actual,outtime_sample_HYBRID$Fitted,percent = FALSE),
                                    RMSE = RMSE(outtime_sample_HYBRID$Actual,outtime_sample_HYBRID$Fitted),
                                    MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_HYBRID$Actual,outtime_sample_HYBRID$Fitted)[1],
                                    MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_HYBRID$Actual,outtime_sample_HYBRID$Fitted)[2],
                                    MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_HYBRID$Actual,outtime_sample_HYBRID$Fitted)[3])
  
  lambdas <- BoxCox.lambda(ts_data_sa_2)
  
  
  ts_hybrid_2<-hybridModel(ts_data_sa_2,models = "aefst", lambda = lambdas)
  
  
  
  forecasted_values_hybrid_2<-data.frame(forecast(ts_hybrid_2,h = 3))
  
  
  Future_predictions_HYBRID<-data.frame(Model = "HYRBID",
                                       Skill = rep(p_0[1,1],3),
                                       Month = rownames(forecasted_values_hybrid_2[1:3,]),
                                       Fitted = forecasted_values_hybrid_2[1:3,1],
                                       Fitted_Low = forecasted_values_hybrid_2[1:3,2],
                                       Fitted_High = forecasted_values_hybrid_2[1:3,3])
  
  
  ####Arimax - Demandxreg
  xreg_1_train<- as.matrix(p_0[1:length(ts_data),c("Demand")])
  lambdas <- BoxCox.lambda(ts_data)
  ts_arimax_1<-auto.arima(ts_data,max.p = 10, max.q = 10, max.order = 2, 
                          max.d = 10, start.p = 0,
                          start.q = 0, xreg = xreg_1_train,
                          stationary = FALSE,
                          seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                          stepwise = TRUE,
                          trace = TRUE,
                          truncate = NULL, test = c("kpss", "adf", "pp"),
                          seasonal.test = c("ocsb", "ch"), 
                          allowdrift = TRUE, allowmean = TRUE,
                          lambda = lambdas, biasadj = FALSE,
                          parallel = FALSE, num.cores = NULL)
  
  intime_sample_arimax_1<-data.frame(Model = "ARIMAX_1",
                                     Skill = rep(p_0[1,1],length(ts_data)),
                                     Month =  p_0[1:length(ts_data),2],
                                     Actual = p_0[1:length(ts_data),3],
                                     Fitted = ts_arimax_1$fitted)
  
  
  colnames(intime_sample_arimax_1)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_arimax_1<-data.frame(Model = "ARIMAX_1",
                                      MAPE = MAPE(intime_sample_arimax_1$Actual,intime_sample_arimax_1$Fitted,percent = FALSE),
                                      RMSE = RMSE(intime_sample_arimax_1$Actual,intime_sample_arimax_1$Fitted))
  
  xreg_1_test<- as.matrix(p_0[(length(ts_data)+1):(length(ts_data)+3),c("Demand")])
  forecasted_values_arimax_1<-data.frame(forecast(ts_arimax_1,h = 3, xreg = xreg_1_test))
  
  outtime_sample_ARIMAX_1<-data.frame(Model = "ARIMAX_1",
                                      Skill = rep(p_0[1,1],3),
                                      Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                      Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                      Fitted = forecasted_values_arimax_1[1:3,1],
                                      Fitted_Low = forecasted_values_arimax_1[1:3,2],
                                      Fitted_High = forecasted_values_arimax_1[1:3,3])
  
  colnames(outtime_sample_ARIMAX_1)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_ARIMAX_1<-data.frame(Model = "ARIMAX_1",
                                       SKILL = p_0[1,1],
                                       MAPE = MAPE(outtime_sample_ARIMAX_1$Actual,outtime_sample_ARIMAX_1$Fitted,percent = FALSE),
                                       RMSE = RMSE(outtime_sample_ARIMAX_1$Actual,outtime_sample_ARIMAX_1$Fitted),
                                       MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_1$Actual,outtime_sample_ARIMAX_1$Fitted)[1],
                                       MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_1$Actual,outtime_sample_ARIMAX_1$Fitted)[2],
                                       MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_1$Actual,outtime_sample_ARIMAX_1$Fitted)[3])
  
  xreg_1_train_2<- as.matrix(p_0[1:length(ts_data_2),c("Demand")])
  lambdas <- BoxCox.lambda(ts_data_2)
  ts_arimax_1_2<-auto.arima(ts_data_2,max.p = 10, max.q = 10, max.order = 2, 
                            max.d = 10, start.p = 0,
                            start.q = 0, xreg = xreg_1_train_2,
                            stationary = FALSE,
                            seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                            stepwise = TRUE,
                            trace = TRUE,
                            truncate = NULL, test = c("kpss", "adf", "pp"),
                            seasonal.test = c("ocsb", "ch"), 
                            allowdrift = TRUE, allowmean = TRUE,
                            lambda = lambdas, biasadj = FALSE,
                            parallel = FALSE, num.cores = NULL)
  
  xreg_1_test_2<- as.matrix(p_0[(length(ts_data_2)+1):(length(ts_data_2)+3),c("Demand")])
  forecasted_values_arimax_1_2<-data.frame(forecast(ts_arimax_1_2,h = 3, xreg = xreg_1_test_2))
  
  
  Future_predictions_ARIMAX_1<-data.frame(Model = "ARIMAX_1",
                                          Skill = rep(p_0[1,1],3),
                                          Month = rownames(forecasted_values_arimax_1_2[1:3,]),
                                          Fitted = forecasted_values_arimax_1_2[1:3,1],
                                          Fitted_Low = forecasted_values_arimax_1_2[1:3,2],
                                          Fitted_High = forecasted_values_arimax_1_2[1:3,3])
  
  
  ####Arimax_2 - OPPORTUNITYxreg
  xreg_2_train<- as.matrix(p_0[1:length(ts_data),c("Opportunity")])
  lambdas <- BoxCox.lambda(ts_data)
  ts_arimax_2<-auto.arima(ts_data,max.p = 10, max.q = 10, max.order = 2, 
                          max.d = 10, start.p = 0,
                          start.q = 0, xreg = xreg_2_train,
                          stationary = FALSE,
                          seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                          stepwise = TRUE,
                          trace = TRUE,
                          truncate = NULL, test = c("kpss", "adf", "pp"),
                          seasonal.test = c("ocsb", "ch"), 
                          allowdrift = TRUE, allowmean = TRUE,
                          lambda = lambdas, biasadj = FALSE,
                          parallel = FALSE, num.cores = NULL)
  
  intime_sample_arimax_2<-data.frame(Model = "ARIMAX_2",
                                     Skill = rep(p_0[1,1],length(ts_data)),
                                     Month =  p_0[1:length(ts_data),2],
                                     Actual = p_0[1:length(ts_data),3],
                                     Fitted = ts_arimax_2$fitted)
  
  
  colnames(intime_sample_arimax_2)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_arimax_2<-data.frame(Model = "ARIMAX_2",
                                      MAPE = MAPE(intime_sample_arimax_1$Actual,intime_sample_arimax_1$Fitted,percent = FALSE),
                                      RMSE = RMSE(intime_sample_arimax_1$Actual,intime_sample_arimax_1$Fitted))
  
  xreg_2_test<- as.matrix(p_0[(length(ts_data)+1):(length(ts_data)+3),c("Opportunity")])
  forecasted_values_arimax_2<-data.frame(forecast(ts_arimax_2,h = 3, xreg = xreg_2_test))
  
  outtime_sample_ARIMAX_2<-data.frame(Model = "ARIMAX_2",
                                      Skill = rep(p_0[1,1],3),
                                      Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                      Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                      Fitted = forecasted_values_arimax_2[1:3,1],
                                      Fitted_Low = forecasted_values_arimax_2[1:3,2],
                                      Fitted_High = forecasted_values_arimax_2[1:3,3])
  
  colnames(outtime_sample_ARIMAX_2)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_ARIMAX_2<-data.frame(Model = "ARIMAX_2",
                                       SKILL = p_0[1,1],
                                       MAPE = MAPE(outtime_sample_ARIMAX_2$Actual,outtime_sample_ARIMAX_2$Fitted,percent = FALSE),
                                       RMSE = RMSE(outtime_sample_ARIMAX_2$Actual,outtime_sample_ARIMAX_2$Fitted),
                                       MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_2$Actual,outtime_sample_ARIMAX_2$Fitted)[1],
                                       MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_2$Actual,outtime_sample_ARIMAX_2$Fitted)[2],
                                       MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_2$Actual,outtime_sample_ARIMAX_2$Fitted)[3])
  
  
  xreg_2_train_2<- as.matrix(p_0[1:length(ts_data_2),c("Opportunity")])
  lambdas <- BoxCox.lambda(ts_data_2)
  ts_arimax_2_2<-auto.arima(ts_data_2,max.p = 10, max.q = 10, max.order = 2, 
                            max.d = 10, start.p = 0,
                            start.q = 0, xreg = xreg_2_train_2,
                            stationary = FALSE,
                            seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                            stepwise = TRUE,
                            trace = TRUE,
                            truncate = NULL, test = c("kpss", "adf", "pp"),
                            seasonal.test = c("ocsb", "ch"), 
                            allowdrift = TRUE, allowmean = TRUE,
                            lambda = lambdas, biasadj = FALSE,
                            parallel = FALSE, num.cores = NULL)
  
  xreg_2_test_2<- as.matrix(p_0[(length(ts_data)+1):(length(ts_data)+3),c("Opportunity")])
  forecasted_values_arimax_2_2<-data.frame(forecast(ts_arimax_2_2,h = 3, xreg = xreg_2_test_2))
  
  
  
  
  Future_predictions_ARIMAX_2<-data.frame(Model = "ARIMAX_2",
                                          Skill = rep(p_0[1,1],3),
                                          Month = rownames(forecasted_values_arimax_2_2[1:3,]),
                                          Fitted = forecasted_values_arimax_2_2[1:3,1],
                                          Fitted_Low = forecasted_values_arimax_2_2[1:3,2],
                                          Fitted_High = forecasted_values_arimax_2_2[1:3,3])
  
  
  
  ####Arimax_3 - OPPORTUNITY_MA3xreg
  xreg_3_train<- as.matrix(p_0[1:length(ts_data),c("Opportunity_MA3")])
  lambdas <- BoxCox.lambda(ts_data)
  ts_arimax_3<-auto.arima(ts_data,max.p = 10, max.q = 10, max.order = 2, 
                          max.d = 10, start.p = 0,
                          start.q = 0, xreg = xreg_3_train,
                          stationary = FALSE,
                          seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                          stepwise = TRUE,
                          trace = TRUE,
                          truncate = NULL, test = c("kpss", "adf", "pp"),
                          seasonal.test = c("ocsb", "ch"), 
                          allowdrift = TRUE, allowmean = TRUE,
                          lambda = lambdas, biasadj = FALSE,
                          parallel = FALSE, num.cores = NULL)
  
  intime_sample_arimax_3<-data.frame(Model = "ARIMAX_3",
                                     Skill = rep(p_0[1,1],length(ts_data)),
                                     Month =  p_0[1:length(ts_data),2],
                                     Actual = p_0[1:length(ts_data),3],
                                     Fitted = ts_arimax_3$fitted)
  
  
  colnames(intime_sample_arimax_3)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_arimax_3<-data.frame(Model = "ARIMAX_3",
                                      MAPE = MAPE(intime_sample_arimax_3$Actual,intime_sample_arimax_3$Fitted,percent = FALSE),
                                      RMSE = RMSE(intime_sample_arimax_3$Actual,intime_sample_arimax_3$Fitted))
  
  xreg_3_test<- as.matrix(p_0[(length(ts_data)+1):(length(ts_data)+3),c("Opportunity_MA3")])
  forecasted_values_arimax_3<-data.frame(forecast(ts_arimax_3,h = 3, xreg = xreg_3_test))
  
  outtime_sample_ARIMAX_3<-data.frame(Model = "ARIMAX_3",
                                      Skill = rep(p_0[1,1],3),
                                      Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                      Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                      Fitted = forecasted_values_arimax_3[1:3,1],
                                      Fitted_Low = forecasted_values_arimax_3[1:3,2],
                                      Fitted_High = forecasted_values_arimax_3[1:3,3])
  
  colnames(outtime_sample_ARIMAX_3)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_ARIMAX_3<-data.frame(Model = "ARIMAX_3",
                                       SKILL = p_0[1,1],
                                       MAPE = MAPE(outtime_sample_ARIMAX_3$Actual,outtime_sample_ARIMAX_3$Fitted,percent = FALSE),
                                       RMSE = RMSE(outtime_sample_ARIMAX_3$Actual,outtime_sample_ARIMAX_3$Fitted),
                                       MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_3$Actual,outtime_sample_ARIMAX_3$Fitted)[1],
                                       MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_3$Actual,outtime_sample_ARIMAX_3$Fitted)[2],
                                       MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_3$Actual,outtime_sample_ARIMAX_3$Fitted)[3])
  
  xreg_3_train_2<- as.matrix(p_0[1:length(ts_data_2),c("Opportunity_MA3")])
  lambdas <- BoxCox.lambda(ts_data_2)
  ts_arimax_3_2<-auto.arima(ts_data_2,max.p = 10, max.q = 10, max.order = 2, 
                            max.d = 10, start.p = 0,
                            start.q = 0, xreg = xreg_3_train_2,
                            stationary = FALSE,
                            seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                            stepwise = TRUE,
                            trace = TRUE,
                            truncate = NULL, test = c("kpss", "adf", "pp"),
                            seasonal.test = c("ocsb", "ch"), 
                            allowdrift = TRUE, allowmean = TRUE,
                            lambda = lambdas, biasadj = FALSE,
                            parallel = FALSE, num.cores = NULL)
  
  xreg_3_test_2<- as.matrix(p_0[(length(ts_data_2)+1):(length(ts_data_2)+3),c("Opportunity_MA3")])
  forecasted_values_arimax_3_2<-data.frame(forecast(ts_arimax_3_2,h = 3, xreg = xreg_3_test_2))
  
  Future_predictions_ARIMAX_3<-data.frame(Model = "ARIMAX_3",
                                          Skill = rep(p_0[1,1],3),
                                          Month = rownames(forecasted_values_arimax_3_2[1:3,]),
                                          Fitted = forecasted_values_arimax_3_2[1:3,1],
                                          Fitted_Low = forecasted_values_arimax_3_2[1:3,2],
                                          Fitted_High = forecasted_values_arimax_3_2[1:3,3])
  
  
  
  
  ####Arimax_4 - RevenueActual
  xreg_4_train<- as.matrix(p_0[1:length(ts_data),c("Revenue")])
  lambdas <- BoxCox.lambda(ts_data)
  ts_arimax_4<-auto.arima(ts_data,max.p = 10, max.q = 10, max.order = 2, 
                          max.d = 10, start.p = 0,
                          start.q = 0, xreg = xreg_4_train,
                          stationary = FALSE,
                          seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                          stepwise = TRUE,
                          trace = TRUE,
                          truncate = NULL, test = c("kpss", "adf", "pp"),
                          seasonal.test = c("ocsb", "ch"), 
                          allowdrift = TRUE, allowmean = TRUE,
                          lambda = lambdas, biasadj = FALSE,
                          parallel = FALSE, num.cores = NULL)
  
  intime_sample_arimax_4<-data.frame(Model = "ARIMAX_4",
                                     Skill = rep(p_0[1,1],length(ts_data)),
                                     Month =  p_0[1:length(ts_data),2],
                                     Actual = p_0[1:length(ts_data),3],
                                     Fitted = ts_arimax_4$fitted)
  
  
  colnames(intime_sample_arimax_4)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_arimax_4<-data.frame(Model = "ARIMAX_4",
                                      MAPE = MAPE(intime_sample_arimax_4$Actual,intime_sample_arimax_4$Fitted,percent = FALSE),
                                      RMSE = RMSE(intime_sample_arimax_4$Actual,intime_sample_arimax_4$Fitted))
  
  xreg_4_test<- as.matrix(p_0[(length(ts_data)+1):(length(ts_data)+3),c("Revenue")])
  forecasted_values_arimax_4<-data.frame(forecast(ts_arimax_4,h = 3, xreg = xreg_4_test))
  
  outtime_sample_ARIMAX_4<-data.frame(Model = "ARIMAX_4",
                                      Skill = rep(p_0[1,1],3),
                                      Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                      Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                      Fitted = forecasted_values_arimax_4[1:3,1],
                                      Fitted_Low = forecasted_values_arimax_4[1:3,2],
                                      Fitted_High = forecasted_values_arimax_4[1:3,3])
  
  colnames(outtime_sample_ARIMAX_4)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_ARIMAX_4<-data.frame(Model = "ARIMAX_4",
                                       SKILL = p_0[1,1],
                                       MAPE = MAPE(outtime_sample_ARIMAX_4$Actual,outtime_sample_ARIMAX_4$Fitted,percent = FALSE),
                                       RMSE = RMSE(outtime_sample_ARIMAX_4$Actual,outtime_sample_ARIMAX_4$Fitted),
                                       MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_4$Actual,outtime_sample_ARIMAX_4$Fitted)[1],
                                       MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_4$Actual,outtime_sample_ARIMAX_4$Fitted)[2],
                                       MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_4$Actual,outtime_sample_ARIMAX_4$Fitted)[3])
  
  xreg_4_train_2<- as.matrix(p_0[1:length(ts_data_2),c("Revenue")])
  lambdas <- BoxCox.lambda(ts_data_2)
  ts_arimax_4_2<-auto.arima(ts_data_2,max.p = 10, max.q = 10, max.order = 2, 
                            max.d = 10, start.p = 0,
                            start.q = 0, xreg = xreg_4_train_2,
                            stationary = FALSE,
                            seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                            stepwise = TRUE,
                            trace = TRUE,
                            truncate = NULL, test = c("kpss", "adf", "pp"),
                            seasonal.test = c("ocsb", "ch"), 
                            allowdrift = TRUE, allowmean = TRUE,
                            lambda = lambdas, biasadj = FALSE,
                            parallel = FALSE, num.cores = NULL)
  
  
  xreg_4_test_2<- as.matrix(p_0[(length(ts_data_2)+1):(length(ts_data_2)+3),c("Revenue")])
  forecasted_values_arimax_4_2<-data.frame(forecast(ts_arimax_4_2,h = 3, xreg = xreg_4_test_2))
  
  
  Future_predictions_ARIMAX_4<-data.frame(Model = "ARIMAX_4",
                                          Skill = rep(p_0[1,1],3),
                                          Month = rownames(forecasted_values_arimax_4_2[1:3,]),
                                          Fitted = forecasted_values_arimax_4_2[1:3,1],
                                          Fitted_Low = forecasted_values_arimax_4_2[1:3,2],
                                          Fitted_High = forecasted_values_arimax_4_2[1:3,3])
  
  
  
  
  
  ####Arimax_5 - Revenue Revised
  xreg_5_train<- as.matrix(p_0[1:length(ts_data),c("Revenue_Revised")])
  lambdas <- BoxCox.lambda(ts_data)
  ts_arimax_5<-auto.arima(ts_data,max.p = 10, max.q = 10, max.order = 2, 
                          max.d = 10, start.p = 0,
                          start.q = 0, xreg = xreg_5_train,
                          stationary = FALSE,
                          seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                          stepwise = TRUE,
                          trace = TRUE,
                          truncate = NULL, test = c("kpss", "adf", "pp"),
                          seasonal.test = c("ocsb", "ch"), 
                          allowdrift = TRUE, allowmean = TRUE,
                          lambda = lambdas, biasadj = FALSE,
                          parallel = FALSE, num.cores = NULL)
  
  intime_sample_arimax_5<-data.frame(Model = "ARIMAX_5",
                                     Skill = rep(p_0[1,1],length(ts_data)),
                                     Month =  p_0[1:length(ts_data),2],
                                     Actual = p_0[1:length(ts_data),3],
                                     Fitted = ts_arimax_5$fitted)
  
  
  colnames(intime_sample_arimax_5)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_arimax_5<-data.frame(Model = "ARIMAX_5",
                                      MAPE = MAPE(intime_sample_arimax_5$Actual,intime_sample_arimax_5$Fitted,percent = FALSE),
                                      RMSE = RMSE(intime_sample_arimax_5$Actual,intime_sample_arimax_5$Fitted))
  
  xreg_5_test<- as.matrix(p_0[(length(ts_data)+1):(length(ts_data)+3),c("Revenue_Revised")])
  forecasted_values_arimax_5<-data.frame(forecast(ts_arimax_5,h = 3, xreg = xreg_5_test))
  
  outtime_sample_ARIMAX_5<-data.frame(Model = "ARIMAX_5",
                                      Skill = rep(p_0[1,1],3),
                                      Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                      Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                      Fitted = forecasted_values_arimax_5[1:3,1],
                                      Fitted_Low = forecasted_values_arimax_5[1:3,2],
                                      Fitted_High = forecasted_values_arimax_5[1:3,3])
  
  colnames(outtime_sample_ARIMAX_5)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_ARIMAX_5<-data.frame(Model = "ARIMAX_5",
                                       SKILL = p_0[1,1],
                                       MAPE = MAPE(outtime_sample_ARIMAX_5$Actual,outtime_sample_ARIMAX_5$Fitted,percent = FALSE),
                                       RMSE = RMSE(outtime_sample_ARIMAX_5$Actual,outtime_sample_ARIMAX_5$Fitted),
                                       MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_5$Actual,outtime_sample_ARIMAX_5$Fitted)[1],
                                       MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_5$Actual,outtime_sample_ARIMAX_5$Fitted)[2],
                                       MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_5$Actual,outtime_sample_ARIMAX_5$Fitted)[3])
  
  xreg_5_train_2<- as.matrix(p_0[1:length(ts_data_2),c("Revenue_Revised")])
  lambdas <- BoxCox.lambda(ts_data_2)
  ts_arimax_5_2<-auto.arima(ts_data_2,max.p = 10, max.q = 10, max.order = 2, 
                            max.d = 10, start.p = 0,
                            start.q = 0, xreg = xreg_5_train_2,
                            stationary = FALSE,
                            seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                            stepwise = TRUE,
                            trace = TRUE,
                            truncate = NULL, test = c("kpss", "adf", "pp"),
                            seasonal.test = c("ocsb", "ch"), 
                            allowdrift = TRUE, allowmean = TRUE,
                            lambda = lambdas, biasadj = FALSE,
                            parallel = FALSE, num.cores = NULL)
  
  
  xreg_5_test_2<- as.matrix(p_0[(length(ts_data_2)+1):(length(ts_data_2)+3),c("Revenue_Revised")])
  forecasted_values_arimax_5_2<-data.frame(forecast(ts_arimax_5_2,h = 3, xreg = xreg_5_test_2))
  
  Future_predictions_ARIMAX_5<-data.frame(Model = "ARIMAX_5",
                                          Skill = rep(p_0[1,1],3),
                                          Month = rownames(forecasted_values_arimax_5_2[1:3,]),
                                          Fitted = forecasted_values_arimax_5_2[1:3,1],
                                          Fitted_Low = forecasted_values_arimax_5_2[1:3,2],
                                          Fitted_High = forecasted_values_arimax_5_2[1:3,3])
  
  
  
  ####Arimax_6 - Revenue and Demand
  xreg_6_train<- as.matrix(p_0[1:length(ts_data),c("Revenue" , "Demand")])
  lambdas <- BoxCox.lambda(ts_data)
  ts_arimax_6<-auto.arima(ts_data,max.p = 10, max.q = 10, max.order = 2, 
                          max.d = 10, start.p = 0,
                          start.q = 0, xreg = xreg_6_train,
                          stationary = FALSE,
                          seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                          stepwise = TRUE,
                          trace = TRUE,
                          truncate = NULL, test = c("kpss", "adf", "pp"),
                          seasonal.test = c("ocsb", "ch"), 
                          allowdrift = TRUE, allowmean = TRUE,
                          lambda = lambdas, biasadj = FALSE,
                          parallel = FALSE, num.cores = NULL)
  
  intime_sample_arimax_6<-data.frame(Model = "ARIMAX_6",
                                     Skill = rep(p_0[1,1],length(ts_data)),
                                     Month =  p_0[1:length(ts_data),2],
                                     Actual = p_0[1:length(ts_data),3],
                                     Fitted = ts_arimax_6$fitted)
  
  
  colnames(intime_sample_arimax_6)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_arimax_6<-data.frame(Model = "ARIMAX_6",
                                      MAPE = MAPE(intime_sample_arimax_6$Actual,intime_sample_arimax_6$Fitted,percent = FALSE),
                                      RMSE = RMSE(intime_sample_arimax_6$Actual,intime_sample_arimax_6$Fitted))
  
  xreg_6_test<- as.matrix(p_0[(length(ts_data)+1):(length(ts_data)+3),c("Revenue" ,"Demand")])
  forecasted_values_arimax_6<-data.frame(forecast(ts_arimax_6,h = 3, xreg = xreg_6_test))
  
  outtime_sample_ARIMAX_6<-data.frame(Model = "ARIMAX_6",
                                      Skill = rep(p_0[1,1],3),
                                      Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                      Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                      Fitted = forecasted_values_arimax_6[1:3,1],
                                      Fitted_Low = forecasted_values_arimax_6[1:3,2],
                                      Fitted_High = forecasted_values_arimax_6[1:3,3])
  
  colnames(outtime_sample_ARIMAX_6)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_ARIMAX_6<-data.frame(Model = "ARIMAX_6",
                                       SKILL = p_0[1,1],
                                       MAPE = MAPE(outtime_sample_ARIMAX_6$Actual,outtime_sample_ARIMAX_6$Fitted,percent = FALSE),
                                       RMSE = RMSE(outtime_sample_ARIMAX_6$Actual,outtime_sample_ARIMAX_6$Fitted),
                                       MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_6$Actual,outtime_sample_ARIMAX_6$Fitted)[1],
                                       MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_6$Actual,outtime_sample_ARIMAX_6$Fitted)[2],
                                       MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_6$Actual,outtime_sample_ARIMAX_6$Fitted)[3])
  
  
  xreg_6_train_2<- as.matrix(p_0[1:length(ts_data_2),c("Revenue" , "Demand")])
  lambdas <- BoxCox.lambda(ts_data_2)
  ts_arimax_6_2<-auto.arima(ts_data_2,max.p = 10, max.q = 10, max.order = 2, 
                            max.d = 10, start.p = 0,
                            start.q = 0, xreg = xreg_6_train_2,
                            stationary = FALSE,
                            seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                            stepwise = TRUE,
                            trace = TRUE,
                            truncate = NULL, test = c("kpss", "adf", "pp"),
                            seasonal.test = c("ocsb", "ch"), 
                            allowdrift = TRUE, allowmean = TRUE,
                            lambda = lambdas, biasadj = FALSE,
                            parallel = FALSE, num.cores = NULL)
  
  
  
  xreg_6_test_2<- as.matrix(p_0[(length(ts_data_2)+1):(length(ts_data_2)+3),c("Revenue" ,"Demand")])
  forecasted_values_arimax_6_2<-data.frame(forecast(ts_arimax_6_2,h = 3, xreg = xreg_6_test_2))
  
  Future_predictions_ARIMAX_6<-data.frame(Model = "ARIMAX_6",
                                          Skill = rep(p_0[1,1],3),
                                          Month = rownames(forecasted_values_arimax_6_2[1:3,]),
                                          Fitted = forecasted_values_arimax_6_2[1:3,1],
                                          Fitted_Low = forecasted_values_arimax_6_2[1:3,2],
                                          Fitted_High = forecasted_values_arimax_6_2[1:3,3])
  
  
  ####Arimax_7 - Opportunity and Demand
  xreg_7_train<- as.matrix(p_0[1:length(ts_data),c("Opportunity","Demand")])
  lambdas <- BoxCox.lambda(ts_data)
  ts_arimax_7<-auto.arima(ts_data,max.p = 10, max.q = 10, max.order = 2, 
                          max.d = 10, start.p = 0,
                          start.q = 0, xreg = xreg_7_train,
                          stationary = FALSE,
                          seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                          stepwise = TRUE,
                          trace = TRUE,
                          truncate = NULL, test = c("kpss", "adf", "pp"),
                          seasonal.test = c("ocsb", "ch"), 
                          allowdrift = TRUE, allowmean = TRUE,
                          lambda = lambdas, biasadj = FALSE,
                          parallel = FALSE, num.cores = NULL)
  
  intime_sample_arimax_7<-data.frame(Model = "ARIMAX_7",
                                     Skill = rep(p_0[1,1],length(ts_data)),
                                     Month =  p_0[1:length(ts_data),2],
                                     Actual = p_0[1:length(ts_data),3],
                                     Fitted = ts_arimax_7$fitted)
  
  
  colnames(intime_sample_arimax_7)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_arimax_7<-data.frame(Model = "ARIMAX_7",
                                      MAPE = MAPE(intime_sample_arimax_7$Actual,intime_sample_arimax_7$Fitted,percent = FALSE),
                                      RMSE = RMSE(intime_sample_arimax_7$Actual,intime_sample_arimax_7$Fitted))
  
  xreg_7_test<- as.matrix(p_0[(length(ts_data)+1):(length(ts_data)+3),c("Opportunity","Demand")])
  forecasted_values_arimax_7<-data.frame(forecast(ts_arimax_7,h = 3, xreg = xreg_7_test))
  
  outtime_sample_ARIMAX_7<-data.frame(Model = "ARIMAX_7",
                                      Skill = rep(p_0[1,1],3),
                                      Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                      Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                      Fitted = forecasted_values_arimax_7[1:3,1],
                                      Fitted_Low = forecasted_values_arimax_7[1:3,2],
                                      Fitted_High = forecasted_values_arimax_7[1:3,3])
  
  colnames(outtime_sample_ARIMAX_7)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_ARIMAX_7<-data.frame(Model = "ARIMAX_7",
                                       SKILL = p_0[1,1],
                                       MAPE = MAPE(outtime_sample_ARIMAX_7$Actual,outtime_sample_ARIMAX_7$Fitted,percent = FALSE),
                                       RMSE = RMSE(outtime_sample_ARIMAX_7$Actual,outtime_sample_ARIMAX_7$Fitted),
                                       MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_7$Actual,outtime_sample_ARIMAX_7$Fitted)[1],
                                       MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_7$Actual,outtime_sample_ARIMAX_7$Fitted)[2],
                                       MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_ARIMAX_7$Actual,outtime_sample_ARIMAX_7$Fitted)[3])
  
  xreg_7_train_2<- as.matrix(p_0[1:length(ts_data_2),c("Opportunity","Demand")])
  lambdas <- BoxCox.lambda(ts_data_2)
  ts_arimax_7_2<-auto.arima(ts_data_2,max.p = 10, max.q = 10, max.order = 2, 
                            max.d = 10, start.p = 0,
                            start.q = 0, xreg = xreg_7_train_2,
                            stationary = FALSE,
                            seasonal = TRUE, ic = c("aicc", "aic", "bic"),
                            stepwise = TRUE,
                            trace = TRUE,
                            truncate = NULL, test = c("kpss", "adf", "pp"),
                            seasonal.test = c("ocsb", "ch"), 
                            allowdrift = TRUE, allowmean = TRUE,
                            lambda = lambdas, biasadj = FALSE,
                            parallel = FALSE, num.cores = NULL)
  
  
  xreg_7_test_2<- as.matrix(p_0[(length(ts_data_2)+1):(length(ts_data_2)+3),c("Opportunity","Demand")])
  forecasted_values_arimax_7_2<-data.frame(forecast(ts_arimax_7_2,h = 3, xreg = xreg_7_test_2))
  
  
  
  Future_predictions_ARIMAX_7<-data.frame(Model = "ARIMAX_7",
                                          Skill = rep(p_0[1,1],3),
                                          Month = rownames(forecasted_values_arimax_7_2[1:3,]),
                                          Fitted = forecasted_values_arimax_7_2[1:3,1],
                                          Fitted_Low = forecasted_values_arimax_7_2[1:3,2],
                                          Fitted_High = forecasted_values_arimax_7_2[1:3,3])
  #
  #
  ####Poisson Regression - Absolute
  Poisson_Regression_1 <- step(glm( p_0[c(1:length(ts_data)),3] ~ .,
                                    family="poisson",
                                    data= p_0[c(1:length(ts_data)),
                                              c("Demand","CTM_Perct","FP_Perct",
                                                "T_AND_M_Perct","Revenue","Opportunity",
                                                "Rampupdownabs","Mx")]),
                               direction = "both")
  
  intime_sample_poisson_1<-data.frame(Model = "POISSON_1",
                                      Skill = rep(p_0[1,1],length(ts_data)),
                                      Month =  p_0[1:length(ts_data),2],
                                      Actual = p_0[1:length(ts_data),3],
                                      Fitted = Poisson_Regression_1$fitted)
  
  
  colnames(intime_sample_poisson_1)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_poisson_1<-data.frame(Model = "POISSON_1",
                                       MAPE = MAPE(intime_sample_poisson_1$Actual,intime_sample_poisson_1$Fitted,percent = FALSE),
                                       RMSE = RMSE(intime_sample_poisson_1$Actual,intime_sample_poisson_1$Fitted))
  
  forecasted_values_poisson_1<-data.frame(predicted = predict(Poisson_Regression_1,
                                                              p_0[(length(ts_data)+1):(length(ts_data)+3),],
                                                              type = "response",se.fit = TRUE))
  
  outtime_sample_POISSON_1<-data.frame(Model = "POISSON_1",
                                       Skill = rep(p_0[1,1],1),
                                       Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                       Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                       Fitted = forecasted_values_poisson_1[1:3,1],
                                       Fitted_Low = forecasted_values_poisson_1[1:3,1]-(1.96 *forecasted_values_poisson_1[1:3,2]),
                                       Fitted_High = forecasted_values_poisson_1[1:3,1]+(1.96 *forecasted_values_poisson_1[1:3,2]))
  
  colnames(outtime_sample_POISSON_1)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_POISSON_1<-data.frame(Model = "POISSON_1",
                                        SKILL = p_0[1,1],
                                        MAPE = MAPE(outtime_sample_POISSON_1$Actual,outtime_sample_POISSON_1$Fitted,percent = FALSE),
                                        RMSE = RMSE(outtime_sample_POISSON_1$Actual,outtime_sample_POISSON_1$Fitted),
                                        MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_POISSON_1$Actual,outtime_sample_POISSON_1$Fitted)[1],
                                        MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_POISSON_1$Actual,outtime_sample_POISSON_1$Fitted)[2],
                                        MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_POISSON_1$Actual,outtime_sample_POISSON_1$Fitted)[3])
  
  Poisson_Regression_1_2 <- step(glm( p_0[c(1:length(ts_data_2)),3] ~ .,
                                      family="poisson",
                                      data= p_0[c(1:length(ts_data_2)),
                                                c("Demand","CTM_Perct","FP_Perct",
                                                  "T_AND_M_Perct","Revenue","Opportunity",
                                                  "Rampupdownabs","Mx")]),
                                 direction = "both")
  
  forecasted_values_poisson_1_2<-data.frame(predicted = predict(Poisson_Regression_1_2,
                                                                p_0[(length(ts_data_2)+1):(length(ts_data_2)+3),],
                                                                type = "response",se.fit = TRUE))
  
  
  
  Future_predictions_POISSON_1<-data.frame(Model = "POISSON_1",
                                           Skill = rep(p_0[1,1],3),
                                           Month = Future_predictions_ARIMAX_1$Month,
                                           Fitted = forecasted_values_poisson_1_2[1:3,1],
                                           Fitted_Low = forecasted_values_poisson_1_2[1:3,1]-(1.96 *forecasted_values_poisson_1_2[1:3,2]),
                                           Fitted_High = forecasted_values_poisson_1_2[1:3,1]+(1.96 *forecasted_values_poisson_1_2[1:3,2]))
  
  #
  #
  ####Poisson Regression - Delta
  Poisson_Regression_2 <- step(glm(  p_0[c(1:length(ts_data)),3] ~ . ,
                                     family="poisson",
                                     data= p_0[c(1:length(ts_data)),c("CTM_Perct","FP_Perct","T_AND_M_Perct",
                                                                      "Rampupdowndel","Demand_del","Mx")]),
                               direction = "both")
  
  
  intime_sample_poisson_2<-data.frame(Model = "POISSON_2",
                                      Skill = rep(p_0[1,1],length(ts_data)),
                                      Month =  p_0[1:length(ts_data),2],
                                      Actual = p_0[1:length(ts_data),3],
                                      Fitted = Poisson_Regression_2$fitted)
  
  
  colnames(intime_sample_poisson_2)<-c('Model',"Skill","Month","Actual","Fitted")
  
  
  metrics_intime_poisson_2<-data.frame(Model = "POISSON_2",
                                       MAPE = MAPE(intime_sample_poisson_2$Actual,intime_sample_poisson_2$Fitted,percent = FALSE),
                                       RMSE = RMSE(intime_sample_poisson_2$Actual,intime_sample_poisson_2$Fitted))
  
  forecasted_values_poisson_2<-data.frame(predicted = predict(Poisson_Regression_2,
                                                              newdata = p_0[(length(ts_data)+1):(length(ts_data)+6),3:ncol(p_0)],
                                                              type = "response",se.fit = TRUE))
  
  outtime_sample_POISSON_2<-data.frame(Model = "POISSON_2",
                                       Skill = rep(p_0[1,1],3),
                                       Month =  p_0[(length(ts_data)+1):(length(ts_data)+3),2],
                                       Actual = p_0[(length(ts_data)+1):(length(ts_data)+3),3],
                                       Fitted = forecasted_values_poisson_2[1:3,1],
                                       Fitted_Low = forecasted_values_poisson_2[1:3,1]-(1.96 *forecasted_values_poisson_2[1:3,2]),
                                       Fitted_High = forecasted_values_poisson_2[1:3,1]+(1.96 *forecasted_values_poisson_2[1:3,2]))
  
  
  colnames(outtime_sample_POISSON_2)<-c('Model',"Skill","Month","Actual","Fitted","Low_Forecast","High_Forecast")
  
  
  metrics_outtime_poisson_2<-data.frame(Model = "POISSON_2",
                                        SKILL = p_0[1,1],
                                        MAPE = MAPE(outtime_sample_POISSON_2$Actual,outtime_sample_POISSON_2$Fitted,percent = FALSE),
                                        RMSE = RMSE(outtime_sample_POISSON_2$Actual,outtime_sample_POISSON_2$Fitted),
                                        MAPE_1stMonth= WEIGHTED_MAPE(outtime_sample_POISSON_2$Actual,outtime_sample_POISSON_2$Fitted)[1],
                                        MAPE_2ndMonth= WEIGHTED_MAPE(outtime_sample_POISSON_2$Actual,outtime_sample_POISSON_2$Fitted)[2],
                                        MAPE_3rdMonth= WEIGHTED_MAPE(outtime_sample_POISSON_2$Actual,outtime_sample_POISSON_2$Fitted)[3])
  
  
  Poisson_Regression_2_2 <- step(glm(  p_0[c(1:length(ts_data_2)),3] ~ . ,
                                       family="poisson",
                                       data= p_0[c(1:length(ts_data_2)),c("CTM_Perct","FP_Perct","T_AND_M_Perct",
                                                                          "Rampupdowndel","Demand_del","Mx")]),
                                 direction = "both")
  
  forecasted_values_poisson_2_2<-data.frame(predicted = predict(Poisson_Regression_2_2,
                                                                newdata = p_0[(length(ts_data_2)+1):(length(ts_data_2)+3),3:ncol(p_0)],
                                                                type = "response",se.fit = TRUE))
  
  
  
  Future_predictions_POISSON_2<-data.frame(Model = "POISSON_2",
                                           Skill = rep(p_0[1,1],3),
                                           Month = Future_predictions_ARIMA$Month,
                                           Fitted = forecasted_values_poisson_2_2[1:3,1],
                                           Fitted_Low = forecasted_values_poisson_2_2[1:3,1]-(1.96 *forecasted_values_poisson_2_2[1:3,2]),
                                           Fitted_High = forecasted_values_poisson_2_2[1:3,1]+(1.96 *forecasted_values_poisson_2_2[1:3,2]))
  
  #intime Metrics Consolidation
  metrics_list_intime<-mget(ls(pattern = "metrics_intime"))
  metrics_consolidated_intime<-as.data.frame(bind_rows(metrics_list_intime))

  
  #outtime Metrics Consolidation
  metrics_list_outtime<-mget(ls(pattern = "metrics_outtime"))
  metrics_consolidated_outtime<-as.data.frame(bind_rows(metrics_list_outtime))
  metrics_consolidated_outtime$MAPE_1stMonth[is.na(metrics_consolidated_outtime$MAPE_1stMonth)]<-1
  metrics_consolidated_outtime$MAPE_2ndMonth[is.na(metrics_consolidated_outtime$MAPE_2ndMonth)]<-1
  metrics_consolidated_outtime$MAPE_3rdMonth[is.na(metrics_consolidated_outtime$MAPE_3rdMonth)]<-1
  metrics_consolidated_outtime$MAPE_1stMonth[is.infinite(metrics_consolidated_outtime$MAPE_1stMonth)]<-1
  metrics_consolidated_outtime$MAPE_2ndMonth[is.infinite(metrics_consolidated_outtime$MAPE_2ndMonth)]<-1
  metrics_consolidated_outtime$MAPE_3rdMonth[is.infinite(metrics_consolidated_outtime$MAPE_3rdMonth)]<-1

  
  #Rank_Cutoff<- 0.25
  #Ranking
  metrics_consolidated_outtime$MAPE_1stMonth_Rank<-rank(metrics_consolidated_outtime$MAPE_1stMonth,ties.method = "random")
  metrics_consolidated_outtime$MAPE_2ndMonth_Rank<-rank(metrics_consolidated_outtime$MAPE_2ndMonth,ties.method = "random")
  metrics_consolidated_outtime$MAPE_3rdMonth_Rank<-rank(metrics_consolidated_outtime$MAPE_3rdMonth,ties.method = "random")
  
  MAPE_1st<-subset(metrics_consolidated_outtime,MAPE_1stMonth_Rank == 1)
  MAPE_2nd<-subset(metrics_consolidated_outtime,MAPE_2ndMonth_Rank == 1)
  MAPE_3rd<-subset(metrics_consolidated_outtime,MAPE_3rdMonth_Rank == 1)
  MAPE_Best_Models<-as.list(rbind(MAPE_1st$Model,MAPE_2nd$Model,MAPE_3rd$Model))
  MAPE_Best_Models<-unique(MAPE_Best_Models)
  MAPE_Best_Models_outtime<-MAPE_Best_Models
  
  
  #Consoldating the Future predictions from the top model ranked in previous step
  for (i in 1:length(MAPE_Best_Models)){
    MAPE_Best_Models[i]<-paste0("Future_predictions_",MAPE_Best_Models[i])
  }
  
  
  if (length(MAPE_Best_Models) == 1){
    Future_Preds_Best<-bind_rows(get(as.vector(MAPE_Best_Models[[1]])))
  }else if (length(MAPE_Best_Models) == 2){
    Future_Preds_Best<-bind_rows(list(get(as.vector(MAPE_Best_Models[[1]]))),
                                 get(as.vector(MAPE_Best_Models[[2]])))
  } else if (length(MAPE_Best_Models) == 3){
    Future_Preds_Best<-
      Future_Preds_Best<-bind_rows(list(get(as.vector(MAPE_Best_Models[[1]]))),
                                   get(as.vector(MAPE_Best_Models[[2]])),
                                   get(as.vector(MAPE_Best_Models[[3]])))
  }
  
  #Consolidating outtime sample forecasts based on top model ranking
  for (i in 1:length(MAPE_Best_Models_outtime)){
    MAPE_Best_Models_outtime[i]<-paste0("outtime_sample_",MAPE_Best_Models_outtime[i])
  }
  
  if (length(MAPE_Best_Models_outtime) == 1){
    Future_Preds_Best_OUTTIME<-bind_rows(get(as.vector(MAPE_Best_Models_outtime[[1]])))
  }else if (length(MAPE_Best_Models_outtime) == 2){
    Future_Preds_Best_OUTTIME<-bind_rows(list(get(as.vector(MAPE_Best_Models_outtime[[1]]))),
                                         get(as.vector(MAPE_Best_Models_outtime[[2]])))
  } else if (length(MAPE_Best_Models_outtime) == 3){
    Future_Preds_Best_OUTTIME<-
      Future_Preds_Best_OUTTIME<-bind_rows(list(get(as.vector(MAPE_Best_Models_outtime[[1]]))),
                                           get(as.vector(MAPE_Best_Models_outtime[[2]])),
                                           get(as.vector(MAPE_Best_Models_outtime[[3]])))
  }
  

  
  
  #Creating Year Month Flags
  Future_Preds_Best$Month<-paste0("01 ",Future_Preds_Best$Month)
  Future_Preds_Best$Month <- as.Date(Future_Preds_Best$Month,format = "%d %h %Y")
  
  #averaging top model predictions - Future Predictions
  Future_Preds_Best_Fit_Final<- data.frame(Future_Preds_Best%>% group_by(Skill,Month)%>%
                                             summarise( Actual = 0,
                                                        Fitted = mean(Fitted),
                                                        Fitted_Low =mean(Fitted_Low), 
                                                        Fitted_High = mean(Fitted_High)))
  
  #averaging top model predictions - outtime sample Predictions
  Future_Preds_Best_Fit_Final_OUTTIME<- data.frame(Future_Preds_Best_OUTTIME%>% group_by(Skill,Month)%>%
                                                     summarise( Actual = mean(Actual),
                                                                Fitted = mean(Fitted),
                                                                Fitted_Low =mean(Low_Forecast), 
                                                                Fitted_High = mean(High_Forecast)))
  
  #Combining Outime and Future predictions
  Future_Preds_Best_Fit_Final<-rbind(Future_Preds_Best_Fit_Final_OUTTIME,Future_Preds_Best_Fit_Final)
  
  #Calculating performance metrics for the averaged forecasts
  Future_Preds_Best_Fit_Final_OUTTIME_AVG<-data.frame(SKILL = p_0[1,1],
                                                      MAPE = MAPE(Future_Preds_Best_Fit_Final_OUTTIME$Actual,Future_Preds_Best_Fit_Final_OUTTIME$Fitted,percent = FALSE),
                                                      RMSE = RMSE(Future_Preds_Best_Fit_Final_OUTTIME$Actual,Future_Preds_Best_Fit_Final_OUTTIME$Fitted),
                                                      MAPE_1stMonth= WEIGHTED_MAPE(Future_Preds_Best_Fit_Final_OUTTIME$Actual,Future_Preds_Best_Fit_Final_OUTTIME$Fitted)[1],
                                                      MAPE_2ndMonth= WEIGHTED_MAPE(Future_Preds_Best_Fit_Final_OUTTIME$Actual,Future_Preds_Best_Fit_Final_OUTTIME$Fitted)[2],
                                                      MAPE_3rdMonth= WEIGHTED_MAPE(Future_Preds_Best_Fit_Final_OUTTIME$Actual,Future_Preds_Best_Fit_Final_OUTTIME$Fitted)[3])
  
  #Consolidating Actual Values
  Last_Month_Actuals <- data.frame(Skill = rep(p_0[1,1]),
                                   Actual_1 = p_0[(length(ts_data)+1),3],
                                   Actual_2 = p_0[(length(ts_data)+2),3],
                                   Actual_3 = p_0[(length(ts_data)+3),3])
  
  
  Future_Preds_Best_Fit_Final_Cons<-rbind(Future_Preds_Best_Fit_Final_Cons,data.frame(Future_Preds_Best_Fit_Final))
  
  metrics_consolidated_outtime_2<-rbind(metrics_consolidated_outtime_2,metrics_consolidated_outtime)
  
  Outtime_Preds_Metrics_Avg_Cons<-rbind(Outtime_Preds_Metrics_Avg_Cons,data.frame(Future_Preds_Best_Fit_Final_OUTTIME_AVG))
  
  Last_Month_Actuals_Cons<- rbind(Last_Month_Actuals_Cons,data.frame(Last_Month_Actuals))
  
  rm(list = setdiff(ls(),c("Future_Preds_Best_Fit_Final_Cons","data_list","data_list_9","SKILL_TABLE",
                           "Outtime_Preds_Metrics_Avg_Cons","Last_Month_Actuals_Cons","UNIT_FTE_CIS_Final_8",
                           "i","MAPE","RMSE","WEIGHTED_MAPE","read_excel_all",
                           "metrics_consolidated_outtime_2")))
  
}


Future_Preds_Best_Fit_Final_Cons_Fitted <-  data.frame(Future_Preds_Best_Fit_Final_Cons[,c(1,2,4)] %>% 
                                                         group_by(Skill,Month) %>%
                                                         summarise(Fitted = mean(Fitted)) %>%
                                                         spread(Month,Fitted))
colnames(Future_Preds_Best_Fit_Final_Cons_Fitted)<-paste0("Fitted_",colnames(Future_Preds_Best_Fit_Final_Cons_Fitted))


Future_Preds_Best_Fit_Final_Cons_Fitted_Low <-  data.frame(Future_Preds_Best_Fit_Final_Cons[,c(1,2,5)] %>% 
                                                             group_by(Skill,Month) %>%
                                                             summarise(Fitted_Low = mean(Fitted_Low)) %>%
                                                             spread(Month,Fitted_Low))

colnames(Future_Preds_Best_Fit_Final_Cons_Fitted_Low)<-paste0("Fitted_Low_",colnames(Future_Preds_Best_Fit_Final_Cons_Fitted_Low))


Future_Preds_Best_Fit_Final_Cons_Fitted_High <-  data.frame(Future_Preds_Best_Fit_Final_Cons[,c(1,2,6)] %>% 
                                                              group_by(Skill,Month) %>%
                                                              summarise(Fitted_High = mean(Fitted_High)) %>%
                                                              spread(Month,Fitted_High))

colnames(Future_Preds_Best_Fit_Final_Cons_Fitted_High)<-paste0("Fitted_High",colnames(Future_Preds_Best_Fit_Final_Cons_Fitted_High))


Future_Preds_Best_Fit_Final_Summary<-cbind(Last_Month_Actuals_Cons,
                                           Future_Preds_Best_Fit_Final_Cons_Fitted[-1],
                                           Future_Preds_Best_Fit_Final_Cons_Fitted_Low[-1],
                                           Future_Preds_Best_Fit_Final_Cons_Fitted_High[-1],
                                           Outtime_Preds_Metrics_Avg_Cons[-1])

Future_Preds_Best_Fit_Final_Summary$Accuracy<- (1-(Future_Preds_Best_Fit_Final_Summary$MAPE)) 

Future_Preds_Best_Fit_Final_Summary$Fitted_90Day_Avg<-rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(8,9,10)])
Future_Preds_Best_Fit_Final_Summary$Fitted_Low_90Day_Avg<-rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(14,15,16)])
Future_Preds_Best_Fit_Final_Summary$Fitted_High_90Day_Avg<-rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(20,21,22)])


Future_Preds_Best_Fit_Final_Summary$Fitted_Growth<- (Future_Preds_Best_Fit_Final_Summary$Fitted_90Day_Avg -
                                                       rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(2,3,4)]))/rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(2,3,4)])


Future_Preds_Best_Fit_Final_Summary$Fitted_Low_Growth<- (Future_Preds_Best_Fit_Final_Summary$Fitted_Low_90Day_Avg -
                                                           rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(2,3,4)]))/rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(2,3,4)])

Future_Preds_Best_Fit_Final_Summary$Fitted_High_Growth<- (Future_Preds_Best_Fit_Final_Summary$Fitted_High_90Day_Avg -
                                                            rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(2,3,4)]))/rowMeans(Future_Preds_Best_Fit_Final_Summary[,c(2,3,4)])



Future_Preds_Best_Fit_Final_Summary2<-left_join(Future_Preds_Best_Fit_Final_Summary,
                                                SKILL_TABLE, by = c("Skill" = "Forecast_Flag"))

write.csv(Future_Preds_Best_Fit_Final_Summary,"IVS_Summary.csv")
write.csv(Future_Preds_Best_Fit_Final_Summary2,"IVS_Summary_With_Skillid.csv")


#####################
Future_Preds_Best_Fit_Final_Summary_2<-cbind(Future_Preds_Best_Fit_Final_Summary[1],
                                             cSplit(Future_Preds_Best_Fit_Final_Summary[1],"Skill",sep = ","))

Future_Preds_Best_Fit_Final_Summary_2<-Future_Preds_Best_Fit_Final_Summary_2 %>%
  mutate_all(as.character)

Taxonomy <- read_excel("")

Taxonomy_L2<-Taxonomy[,c(1,6)]

Taxonomy_L2<-Taxonomy_L2 %>%
  mutate_all(as.character)

for ( i in 2:ncol(Future_Preds_Best_Fit_Final_Summary_2)){
  name<-  interp(colnames(Future_Preds_Best_Fit_Final_Summary_2[i]))
  Future_Preds_Best_Fit_Final_Summary_2<- merge(Future_Preds_Best_Fit_Final_Summary_2,
                                                Taxonomy_L2,
                                                by.x = name ,
                                                by.y = "TaxonomyID",all.x = TRUE)
}

Future_Preds_Best_Fit_Final_Summary_2<-cbind(Skill = Future_Preds_Best_Fit_Final_Summary_2[,"Skill"],
                                             Future_Preds_Best_Fit_Final_Summary_2[,grepl("Skill_Combined", colnames(Future_Preds_Best_Fit_Final_Summary_2))])

Future_Preds_Best_Fit_Final_Summary_2[is.na(Future_Preds_Best_Fit_Final_Summary_2)]<-""

cols<-colnames(Future_Preds_Best_Fit_Final_Summary_2[grepl("Skill_Combined", colnames(Future_Preds_Best_Fit_Final_Summary_2))])

Future_Preds_Best_Fit_Final_Summary_2 <- data.frame( Skill = Future_Preds_Best_Fit_Final_Summary_2$Skill,
                                                     Level_2 = apply( Future_Preds_Best_Fit_Final_Summary_2[ , cols ] , 1 , paste0 , collapse = "-" ))


Future_Preds_Best_Fit_Final_Summary_2<- left_join(Future_Preds_Best_Fit_Final_Summary,
                                                  Future_Preds_Best_Fit_Final_Summary_2,
                                                  by = "Skill")


Future_Preds_Best_Fit_Final_Summary_2<-cbind(Future_Preds_Best_Fit_Final_Summary[1],
                                             Level_2 = Future_Preds_Best_Fit_Final_Summary_2[,ncol(Future_Preds_Best_Fit_Final_Summary_2)],
                                             Future_Preds_Best_Fit_Final_Summary[,-c(1,ncol(Future_Preds_Best_Fit_Final_Summary_2))])

Future_Preds_Best_Fit_Final_Summary_3<-left_join(Future_Preds_Best_Fit_Final_Summary_2,
                                                 skill_table,
                                                 by = c("Skill" = "Skill ID"))

Future_Preds_Best_Fit_Final_Summary_3<-Future_Preds_Best_Fit_Final_Summary_3[!duplicated(Future_Preds_Best_Fit_Final_Summary_3[,33]),]
Future_Preds_Best_Fit_Final_Summary_3 <- Future_Preds_Best_Fit_Final_Summary_3[,c(33,1:32)]

Taxonomy_L3<-Taxonomy[c(1,5)]
Future_Preds_Best_Fit_Final_Summary_4<-merge(Future_Preds_Best_Fit_Final_Summary_3,
                                             Taxonomy_L3,
                                             by.x = "Skill"  ,
                                             by.y = "TaxonomyID",all.x = TRUE)

Future_Preds_Best_Fit_Final_Summary_4 <- Future_Preds_Best_Fit_Final_Summary_4[,c(1,2,3,4:34)]

Future_Preds_Best_Fit_Final_Summary_4<-Future_Preds_Best_Fit_Final_Summary_4[order(Future_Preds_Best_Fit_Final_Summary_4$Actual_1),]
Future_Preds_Best_Fit_Final_Summary_4$Level_2<-as.character(Future_Preds_Best_Fit_Final_Summary_4$Level_2)
Future_Preds_Best_Fit_Final_Summary_4$Level_2[Future_Preds_Best_Fit_Final_Summary_4$Forecast_Flag == "PR_U"]<-"-"
Future_Preds_Best_Fit_Final_Summary_4$Level_2[Future_Preds_Best_Fit_Final_Summary_4$Forecast_Flag == "PR_S"]<-"-"

write.csv(Future_Preds_Best_Fit_Final_Summary_4,paste0("Forecast_Summary.csv"),
          row.names = FALSE)


