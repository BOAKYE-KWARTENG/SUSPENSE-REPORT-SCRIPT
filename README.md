# SUSPENSE-REPORT
R Script to clean suspense data and prepare a report for it in Excel. I lead a project in my office which requires me to wranggle an extracted data from our database to bring solutions, and report to managers on the results, progress and the way forward.



Below is how I automatically clean the data and get the required figures to populate my Excel Dashboard week on week using R.



#### Load the required libraries for the wrangling

```{r LOAD LIBRARIES}
library(tidyverse)
library(readxl)
library(stringr)
library(dplyr)
library(openxlsx)
library(svglite)
```


#### Load the data
use the *read_excel* function to load or read the .xlsx excel files for the wranggling

```{r LOAD LIBRARIES}
Active <- read_excel("Global Unallocated.Active.22.09.2022.xlsx")
Refund <- read_excel("Global Unallocated.Refunded.22.09.2022.xlsx")
Allocated <- read_excel("Global Unallocated.Allocated.22.09.2022.xlsx")
```

#### Get the variables or the columns in the right data types
All the columns that will be needed for the analysis should be in the appropriate data type for the right handling
```{r}
Active$`Allocated Amount` <- as.numeric(Active$`Allocated Amount`)
Refund$`Allocated Amount` <- as.numeric(Refund$`Allocated Amount`)
Allocated$`Allocated Amount` <- as.numeric(Allocated$`Allocated Amount`)
Allocated$`Bank Code` <- as.factor(Allocated$`Bank Code`)


Active$Name <- as.factor(Active$Name)
Active$`Bank Code` <- as.factor(Active$`Bank Code`)
Active$created_by <- as.factor(Active$created_by)


Active$created_on <- as.POSIXct(format(Active$created_on), tz = "UTC")
```


#### Summary of the data
Have a summary look at the data, especially, the interested columns

```{r}
summary(Active)
summary(Refund)
summary(Allocated)
```


#### Count and Amount of Active.Refund.Allocated
To record the count an total amount of the various data loaded

```{r}
Active %>% 
  summarise(count = n(), ActiveAmount = sum(`Allocated Amount`))

Refund %>% 
  summarise(count = n(), RefundAmount = sum(`Allocated Amount`))

Allocated %>% 
  summarise(count = n(), AllocatedAmount = sum(`Allocated Amount`))

```



#### Filter out data of the various channels (MoMo, Bank and Worksite)
```{r}


#filter out the MoMo ebtries (AIRTEL, MTN and VODAFONE) in name variable

MoMo <- Active %>% 
  filter(Name == "AIRTEL TIGO MOBILE MONEY" | 
           Name == "MTN MOBILE MONEY" |
           Name == "VODAFONE MOBILE MONEY" )



#subset non-empty variables in the Bank code: which are all bank observations
Bank <- Active[!(is.na(Active$`Bank Code`) | Active$`Bank Code`==""), ]



# subset from the name variable where the observation is not empty or
# AIRTEL.MTN.VODAFONE

Worksite <- Active[!(is.na(Active$Name) 
                     | Active$Name=="AIRTEL TIGO MOBILE MONEY" 
                     | Active$Name=="MTN MOBILE MONEY" 
                     | Active$Name=="VODAFONE MOBILE MONEY"), ]


```




#### Get the count and amount recorded in suspense per channel
```{r}
MoMo %>% 
  summarise(count = n(), MoMoAmount = sum(`Allocated Amount`))

Bank %>% 
  summarise(count = n(), BankAmount = sum(`Allocated Amount`))

Worksite %>% 
  summarise(count = n(), AllocatedAmount = sum(`Allocated Amount`))

```



#### Channel Breakdown
This shows a breakdown of the suspense entries per channel and give the result as data frame

```{r}


MoMoBreakdown <- data.frame(MoMo %>% 
                              group_by(Name) %>% 
                              summarise_at(vars(Amount = `Allocated Amount`),
                                           funs(sum(.,na.rm=TRUE))) %>% 
                              arrange(desc(Amount)))




BanksBreakdown <- data.frame(Bank %>% 
                               group_by(`Bank Code`) %>% 
                               summarise_at(vars(Amount = `Allocated Amount`),
                                            funs(sum(.,na.rm=TRUE))) %>% 
                               arrange(desc(Amount)))




WorksitesBreakdown <- data.frame(Worksite %>% 
                                   group_by(Name) %>% 
                                   summarise_at(vars(Amount = `Allocated Amount`),
                                                funs(sum(.,na.rm=TRUE))) %>% 
                                   arrange(desc(Amount)))


```



#### Add-on Suspense
This is to track all the newly added entries within the week

```{r}
# Give the lower-bound of the created_on variable to fetch out the newly created suspense

Active$created_on <- as.POSIXct(format(Active$created_on), tz = "UTC")

Add.on <- Active[Active$created_on >= "2022-09-16" , ] # Plus seven(7) days for next weeks analysis


# summarise the Add.on entries by count and amount
Add.on %>% 
  summarise(count = n(), Amount = sum(`Allocated Amount`))

```


#### Extracted the added entries for each channel during the week
* MoMo.Add.on
* Bank.Add.on
* Worksite.Add.on


```{r}
MoMo.Add.on <- Add.on %>% 
  filter(Name == "AIRTEL TIGO MOBILE MONEY" | 
           Name == "MTN MOBILE MONEY" |
           Name == "VODAFONE MOBILE MONEY" )





Bank.Add.on <- Add.on[!(is.na(Add.on$`Bank Code`) | Add.on$`Bank Code`==""), ]



Worksite.Add.on <- Add.on[!(is.na(Add.on$Name) 
                            | Add.on$Name=="AIRTEL TIGO MOBILE MONEY" 
                            | Add.on$Name=="MTN MOBILE MONEY" 
                            | Add.on$Name=="VODAFONE MOBILE MONEY"), ]

```


#### Add.on per channel summary by count and amount
```{r}
MoMo.Add.on %>%
  summarise(count = n(), MoMoAmount = sum(`Allocated Amount`))

Bank.Add.on %>%
  summarise(count = n(), BankAmount = sum(`Allocated Amount`))

Worksite.Add.on %>%
  summarise(count = n(), AllocatedAmount = sum(`Allocated Amount`))
```



#### Officers Suspense for the week
```{r}
Add.on %>% 
  group_by(Add.on$created_by) %>% 
  summarise(count = n(), AllocatedAmount = sum(`Allocated Amount`))
```



#### save the data frames in one excel workbook
```{r}
#create a workbook
work_book <- createWorkbook()


#And then add three work sheets with different sheet names.
addWorksheet(work_book, sheetName="Active")
addWorksheet(work_book, sheetName="MoMo")
addWorksheet(work_book, sheetName="Bank")
addWorksheet(work_book,sheetName="Worksite")
addWorksheet(work_book, sheetName="Add.on")
addWorksheet(work_book, sheetName="MoMo.Add.on")
addWorksheet(work_book, sheetName="Bank.Add.on")
addWorksheet(work_book,sheetName="Worksite.Add.on")



#Now we can write multiple dataframes one by one using writeData() function with the sheet name we assigned before.
writeData(work_book, "Active", Active)
writeData(work_book, "MoMo", MoMo)
writeData(work_book, "Bank", Bank)
writeData(work_book, "Worksite", Worksite)
writeData(work_book, "Add.on", Add.on)
writeData(work_book, "MoMo.Add.on", MoMo.Add.on)
writeData(work_book, "Bank.Add.on", Bank.Add.on)
writeData(work_book, "Worksite.Add.on", Worksite.Add.on)




#Finally, we write to excel file using saveWorkbook() with overwrite=TRUE.
saveWorkbook(work_book,
             file= "ActiveSuspenseBreakdown.xlsx",
             overwrite = TRUE)
             
```


#### SUSPENSE CREATED_BY CHART 
```{r}

library(ggplot2)

Add.On.Graph <- Add.on %>% 
  group_by(created_by) %>% 
  count() %>% 
  ungroup() %>% 
  mutate(perc = `n` / sum(`n`)) %>% 
  arrange(perc) %>%
  mutate(labels = scales::percent(perc))




B <- ggplot(Add.On.Graph, aes(x = "", y = perc, fill = created_by)) +
  geom_col() +
  geom_text(aes(label = labels),
            position = position_stack(vjust = 0.5)) +
  coord_polar(theta = "y")



graph1 <- B + ggtitle("SUSPENSE CREATED BY OFFICERS DURING THE WEEK") + 
  theme(plot.title = element_text(lineheight=.8, color="black", size=17))


ggsave(file="AddOnGraph.svg", plot=graph1, width=10, height=10)
```



#### Clear Environment
```{r}
# rm(list = ls())
# unlink("ActiveSuspenseBreakdown.xlsx")
# unlink("AddOnGraph.svg")

```
