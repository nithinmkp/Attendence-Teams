#Setup
source("functions.R")

#Packages
packages<-c("readxl","rio","tidyverse","fs","here","janitor","zoo","writexl","openxlsx")
package_fn(packages)

#Files
files<-list.files(path = "./Data/", full.names = T)
data_set<-import_list(file = files)

#Roll list
roll<-read_excel("student_rollList_monetary_spring2021.xlsx")
roll<- roll %>% select(1,2,3) 
#names(roll)<-roll[2,]
#roll<-roll[-c(1,2),]

#Attendence data-wrangling
attend<-map_dfr(data_set,rbind) # binding all sheets together for relevant info
attend<-attend[,c(1,3)] #select relevant columns
attend<-attend %>% separate(Timestamp,into = c("Timestamp",NA),
                            sep = ",",remove = T) %>%  #separate columns to get only date
        mutate(Timestamp=parse_date(Timestamp,"%m/%d/%Y")) %>% # convert date to R format 
        clean_names() %>%  #standardize names
        group_by(timestamp) %>% dplyr::count(full_name) %>% #counting by date
        pivot_wider(names_from = timestamp,values_from=n) #pivoting
attend[,-1]<-lapply(attend[,-1],na.fill,0) #replace NA values with zeros
colnames(attend[,-1])<-lapply(attend[,-1],as.Date) #make colum names in date format
days<-c("Weight",weekdays(as.Date(colnames(attend[,-1]))))
attend<-rbind(attend,days)
attend<-attend %>% slice(c(nrow(attend),1:nrow(attend)-1)) #reorder rows

#for loop to change weight of days
for(i in 2:(ncol(attend)-1)){
        if(attend[1,i]=="Tuesday"){
                attend[1,i]<-"2"      
        }else{
                attend[1,i]<-"1"    
        }
}

attend[,-1]<-lapply(attend[,-1],as.numeric) #convert weights to numeric

#for loop to change attendence values to match weights
for(i in 2:ncol(attend)){
        for(j in 2:nrow(attend)){
                if(attend[j,i]!=0){
                        attend[j,i]<-attend[1,i]
                }    
        }
        
}

#Filter Nithin and Siddhartha (T.As)
attend<-subset(attend,attend$full_name!="Nithin M" &
                       attend$full_name!="siddhartha"&
                       attend$full_name!="Christopher Kuruvilla Mathen" &
                       attend$full_name!="Anwesha Das" &
                       attend$full_name != "Arpan Chakraborty" &
                       attend$full_name != "Neha Kumari Mishra")


colnames(attend)[1]<-"Name"

final<-full_join(roll,attend)


write.xlsx(final,"attendence.xlsx")
