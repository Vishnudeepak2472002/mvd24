library(readxl)
library(stringr)
library(purrr)
library(dplyr)
library(gtools)
# taking excel sheets as input and forming combined matrices --------------
list_path = list(mixedsort(list.files(path = "Type 3/2017/", pattern = ".xlsx", full.names = T),), mixedsort(list.files(path = "Type 3/2018/", pattern = ".xlsx", full.names = T),), mixedsort(list.files(path = "Type 3/2019/", pattern = ".xlsx", full.names = T),), mixedsort(list.files(path = "Type 3/2021/", pattern = ".xlsx", full.names = T),))
combined_list = list()
#excel files which are taken as input should solely contain the table of electricity consumption for each month containing (qtrno, previous reading, previous date, present reading, present date) (only one sheet with required apartment) and column names should be same for all months of a given year.
for(x in 1:4)
{
  files = unlist(list_path[x])                                                        # excel sheets having electricity consumption data for each month for 2017 is read through the following command.
  tbl = sapply(files, read_excel, simplify = FALSE)%>%bind_rows(.id = "id")
  incr = nrow(tbl)/12                                                           # no of apartments 
  tbl[is.na(tbl)] = 0                                                           # all NA values are assigned 0 as calculations are further simplified.
  combined = tbl[c(1:incr),-1]                                                  # combined matrix is initially assigned the monthly data of January and further rest of the months are merged with it.
  j = 5
  count = 0
  for(i in seq(from = incr + 1, to = nrow(tbl), by = incr))
  {
    colnames(combined)[j] = 'Prst Rdng_temp'
    temp = tbl[c(i:(i+(incr - 1))),-c(1,2)]                                     
    for(a in 1:incr)
    {
      temp_na = sum((temp[,2]) == 0)                                            # used to calculate number of NA's 
      combined_na = sum((combined[,j]) == 0)
    }
    if(temp_na>0 | combined_na>0)                                               # filtering of rows is done
    {                                                                           # each and every value of (present reading of combined matrix) is compared with each and every value of (previous reading of temp matrix).
      for (k in 1:incr)
      {
        if (!(combined[k, j] %in% temp$Prv.Rdng))
          combined[k, c(j - 2, j)] = 0
      }
      
      for (m in 1:incr) 
      {
        if (!(temp[m, 2] %in% combined$`Prst Rdng_temp`))
          temp[m, 2] = 0
      }
    }
    colnames(combined)[j] = 'Prst.Rdng'
    temp = temp[order(temp[, 2]), ]
    combined = combined[order(combined[, j]), ]
    combined = cbind(combined,temp)
    j = j + 4
  }
  colnames(combined) = c('Qtr No','Prvs Date','Prv Rdng dec','Prst Date','Prs Rdng dec','Prvs Date','Prv Rdng jan','Prst Date','Prs Rdng jan','Prvs Date','Prv Rdng feb','Prst Date','Prs Rdng feb','Prvs Date','Prv Rdng mar','Prst Date','Prs Rdng mar','Prvs Date','Prv Rdng apr','Prst Date','Prs Rdng apr','Prvs Date','Prv Rdng may','Prst Date','Prs Rdng may','Prvs Date','Prv Rdng jun','Prst Date','Prs Rdng jun','Prvs Date','Prv Rdng jul','Prst Date','Prs Rdng jul','Prvs Date','Prv Rdng aug','Prst Date','Prs Rdng aug','Prvs Date','Prv Rdng sep','Prst Date','Prs Rdng sep','Prvs Date','Prv Rdng oct','Prst Date','Prs Rdng oct','Prvs Date','Prv Rdng nov','Prst Date','Prs Rdng nov')
  combined[combined == 0] <- NA
  temp[temp == 0] <- NA
  combined_list[[x]] = combined
}
rm(combined)
rm(temp)
rm(tbl)
rm(list_path)
# calculating net consumption and error matrix from combined data ---------
temp_data <- read_excel("C:/Users/User/Desktop/IITK internship/OG data/temp_data.xlsx")    # reads the temperature data in the form of an excel sheet having temperatures for each day
net_consump_list = list()
errormat_list = list()
data_new_list = list()
colnames_list <- list(c('Dec 16', 'Jan 17', 'Feb 17', 'Mar 17', 'Apr 17', 'May 17', 'Jun 17', 'Jul 17', 'Aug 17', 'Sep 17', 'Oct 17', 'Nov 17'), c('Dec 17', 'Jan 18', 'Feb 18', 'Mar 18', 'Apr 18', 'May 18', 'Jun 18', 'Jul 18', 'Aug 18', 'Sep 18', 'Oct 18', 'Nov 18'), c('Dec 18', 'Jan 19', 'Feb 19', 'Mar 19', 'Apr 19', 'May 19', 'Jun 19', 'Jul 19', 'Aug 19', 'Sep 19', 'Oct 19', 'Nov 19'), c('Dec 20', 'Jan 21', 'Feb 21', 'Mar 21', 'Apr 21', 'May 21', 'Jun 21', 'Jul 21', 'Aug 21', 'Sep 21', 'Oct 21', 'Nov 21'))
for(x in 1:4)                                                                   # a loop is run to assign values of combined_data according to the years.
{
  combined_data = combined_list[[x]]
  temp_2 = combined_data[, seq(from = 5, by = 2,  to = 47)]                     # a matrix is formed so as to find the error matrix
  errormat = temp_2                                                             # error matrix is created using the criteria given
  for(i in 1:nrow(combined_data))
  {
    for(j in seq(from = 1, to = 21, by = 2))
    {
      if(is.na(temp_2[i,j]) | is.na(temp_2[i,j+1]))                             # is both the given value and succeeding values are NA, then NA is assigned to error matrix.
      {
        errormat[i,j+1] = NA
        errormat[i,j] = NA
      }
      else if(temp_2[i,j] == temp_2[i,j+1])                                     # 0 is assigned to error matrix if values are equal.
      {
        errormat[i,j+1] = 0
        errormat[i,j] = 0
      }
      else if(temp_2[i,j] != temp_2[i,j+1])                                     # error is calculated if values are not equal.
      {
        errormat[i,j+1] = abs(temp_2[i,j+1] - temp_2[i,j])
        errormat[i,j] = abs(temp_2[i,j+1] - temp_2[i,j])
      }
    }
  }
  data_new = temp_2                                                             # a new updated matrix is created that removes the errors.
  for (i in 1:nrow(combined_data))
  {
    for(j in 1:22)
    {
      if(is.na(errormat[i,j]))                                                  
        data_new[i,j] = NA                                                      # NA is assigned to element of updated matrix if element of error matrix has NA.
      
      else if(errormat[i,j] > 10)                                               # NA is assigned to element of updated matrix if element of error matrix has error greater than 10.
      {
        data_new[i,j] = NA
      }
      else if(errormat[i,j] > 0 & errormat[i,j] < 10)                           # average value is assigned to element of updated matrix if element of error matrix has an error between 0 to 10.
      {
        data_new[i,j] = mean(data_new[i,j], data_new[i,j+1])
        data_new[i,j+1] = mean(data_new[i,j], data_new[i,j+1])
      }
    }
  }
  combined_data_1 = combined_data                                               # the updated combined data(after getting rid of errors) is assigned to combined_data_1
  combined_data_1[, seq(from = 5, by = 2,  to = 47)] = data_new
  temp_2 = combined_data_1[, c(seq(from = 3, by = 4, to = 49),49)]              # net consumption is further calculated using updated data matrix.
  net_consump = temp_2[, (1:12)]                                                # arbitary net consumption matrix is created so as to store the values after running the loop.
  {
    for(i in 1:nrow(combined_data))
    {
      for(j in 1:12)
      {
        if(is.na(temp_2[i,j+1]))                                                # NA is assigned to net consumption if both present and succeeding values are NA
        {
          net_consump[i,j] = NA
        }
        else if(is.na(temp_2[i,j]))
        {
          net_consump[i,j] = NA
        }
        else
        {
          net_consump[i,j] = temp_2[i,j+1] - temp_2[i,j]                        # difference is taken otherwise. 
        }
      }
    }
  }                                                                             # assigning arbitary value to median
  net_consump = net_consump[rowSums(is.na(net_consump)) != ncol(net_consump), ] # removing rows having all NA values
  net_consump_list[[x]] = net_consump
  data_new_list[[x]] = combined_data_1
  errormat_list[[x]] = errormat
  colnames(net_consump_list[[x]]) = colnames_list[[x]]
}
rm(colnames_list)
rm(combined_data_1)
rm(data_new)
rm(combined_data)
rm(errormat)
rm(net_consump)
rm(temp_2)
# calculating median electricity consumption ------------------------------
median_list = list()
for(x in 1:4)
{
  median_list[[x]] = apply(net_consump_list[[x]], 2, median, na.rm = TRUE)      # used to calculate med EC for each month by excluding the 'NA' Values
}
# calculating degree days with base temp as 18 ----------------------------
DDandEC_list = list()
degreedays_2_list = list()
years = c(2017, 2018, 2019, 2021)
bt = 18                                                                         # value of base temperature 18 is assigned to bt
for(x in 1:4)                                                                   # a loop is run to calculate the degree days.
{
  net_consump = net_consump_list[[x]]
  med = median_list[[x]]
  degreedays_1 = temp_data[temp_data$year == years[x], ]                            # degreedays_1 is created so as to seperate out temp data to different years
  degreedays_1 = rbind(temp_data[temp_data$year == years[x] - 1 & temp_data$month == 12, ], degreedays_1)     # as we want the data from the span of December of previous year to November of given year, hence December of previous year is added and December of given year is removed.
  degreedays_1 = degreedays_1[1:365, ]
  degreedays_1 = cbind(degreedays_1, matrix((1:730), nrow = 365, ncol = 2, byrow = FALSE))
  colnames(degreedays_1)[5] = 'HDD'
  colnames(degreedays_1)[6] = 'CDD'
  
  for(i in 1:365)                                                               # below code is used to calculate degree days and assigns them to degreedays_2
  {
    if(degreedays_1[i,4] > bt)
    {
      degreedays_1[i,6] = (degreedays_1[i,4] - bt)
      degreedays_1[i,5] = 0
    }
    else if(degreedays_1[i,4] < bt)
    {
      degreedays_1[i,5] = (bt - degreedays_1[i,4])
      degreedays_1[i,6] = 0
    }
    else if(degreedays_1[i,4] == bt)
    {
      degreedays_1[i,5] = 0
      degreedays_1[i,6] = 0 
    }
  }
  #this is done to make an arbitary table for storing degree days which should have row names as months hence transpose of net_consumption is taken as it has column names as months.
  #hence 2 rows from net consumption are taken and further they are transposed.
  degreedays_2 = t(net_consump[c(1,2), ])
  for(i in 1:11)                                                                # a loop is run to sum the values of degree days from Jan to Nov of current year 
  {
    temp_3 = degreedays_1[degreedays_1$year == years[x] & degreedays_1$month == i, ]
    degreedays_2[i, ] = apply(temp_3[, c(5,6)], 2, sum)
  }
  
  temp_4 = degreedays_1[degreedays_1$year == years[x] - 1 & degreedays_1$month == 12, ]    # as we also want degree days of Dec of prev year hence its further added and the extra row present is removed.
  temp_5 = apply(temp_4[, c(5,6)], 2, sum)
  degreedays_2 = rbind(temp_5, degreedays_2)
  degreedays_2 = degreedays_2[-13, ]
  rownames(degreedays_2) = colnames(net_consump)
  degreedays_2 = as.data.frame(degreedays_2)                                    
  DDandEC = cbind(degreedays_2, matrix(med, nrow = 12, ncol = 1, byrow = FALSE))
  rownames(DDandEC) = rownames(degreedays_2)
  degreedays_2_list[[x]] = degreedays_2
  colnames(DDandEC)[3] = 'med EC'
  DDandEC_list[[x]] = DDandEC
}
rm(years)
rm(DDandEC)
rm(degreedays_1)
rm(degreedays_2)
rm(temp_3)
rm(temp_4)
rm(net_consump)
# boxplots are plotted using net electricity consumption and degree days --------
par(mfrow = c(2,2))
boxplot(net_consump_list[[1]], outline = FALSE, xlab = 'months', ylab = 'electricity consumption (KW) for 2017', ylim = c(0,2000))
for(i in 1:12)
{
  if(i == 5)
  {
    text(x = i, y = median_list[[1]][i] + 50, labels = median_list[[1]][i], cex = 0.75)
  }
  else
    text(x = i, y = median_list[[1]][i] + 100, labels = median_list[[1]][i], cex = 0.85)
}
points(degreedays_2_list[[1]]$CDD, type = 'l', col = "blue")
points(degreedays_2_list[[1]]$HDD, type = 'l', col = "red")
boxplot(net_consump_list[[2]], outline = FALSE, xlab = 'months', ylab = 'electricity consumption (KW) for 2018', ylim = c(0,2000))
for(i in 1:12)
{
  if(i == 4 | i == 11)
  {
    text(x = i, y = median_list[[2]][i]  + 111, labels = median_list[[2]][i], cex = 0.85)
  }
  else
    text(x = i, y = median_list[[2]][i] + 50, labels = median_list[[2]][i], cex = 0.85)
}
points(degreedays_2_list[[2]]$CDD, type = 'l', col = "blue")
points(degreedays_2_list[[2]]$HDD, type = 'l', col = "red")
boxplot(net_consump_list[[3]], outline = FALSE, xlab = 'months', ylab = 'electricity consumption (KW) for 2019', ylim = c(0,2000))
for(i in 1:12)
{
  text(x = i, y = median_list[[3]][i] + 100, labels = median_list[[3]][i], cex = 0.85)
}
points(degreedays_2_list[[3]]$CDD, type = 'l', col = "blue")
points(degreedays_2_list[[3]]$HDD, type = 'l', col = "red")
boxplot(net_consump_list[[4]], outline = FALSE, xlab = 'months', ylab = 'electricity consumption (KW) for 2021', ylim = c(0,2000))
for(i in 1:12)
{
  text(x = i, y = median_list[[4]][i] + 80, labels = median_list[[4]][i], cex = 0.85)
}
points(degreedays_2_list[[4]]$CDD, type = 'l', col = "blue")
points(degreedays_2_list[[4]]$HDD, type = 'l', col = "red")
DDandEC_combined = rbind(DDandEC_list[[1]],DDandEC_list[[2]], DDandEC_list[[3]], DDandEC_list[[4]])
attach(DDandEC_combined)
names(combined_list) = c('combined_2017', 'combined_2018', 'combined_2019', 'combined_2021')
names(errormat_list) = c('errormat_2017', 'errormat_2018', 'errormat_2019', 'errormat_2021')
names(data_new_list) = c('data_new_2017', 'data_new_2018', 'data_new_2019', 'data_new_2021')
names(DDandEC_list) = c('DDandEC_2017', 'DDandEC_2018', 'DDandEC_2019', 'DDandEC_2021')
names(degreedays_2_list) = c('degreedays_2017', 'degreedays_2018', 'degreedays_2019', 'degreedays_2021')
names(net_consump_list) = c('net_consump_2017', 'net_consump_2018', 'net_consump_2019', 'net_consump_2021')
names(median_list) = c('median_2017', 'median_2018', 'median_2019', 'median_2021')
# multiparameter regression -----------------------------------------------
model <- lm(`med EC` ~ (HDD + CDD), data = DDandEC_combined)                    # model is formed having med EC as dependant variable and degree days as independant variables.                 
print(summary(model))
parameters_list = list(round(as.numeric(model$coefficients[2]),2), round(as.numeric(model$coefficients[3]),2), round(as.numeric(model$coefficients[1],2)))
names(parameters_list) = c('HDD', 'CDD', 'base line consump')
sprintf('Hence the equation is obtained as - med EC = %f*HDD + %f*CDD + %f', parameters_list$HDD, parameters_list$CDD, parameters_list$`base line consump`)
