########################## function to pull data from prism ###########################
collect_prism_data <- function(local_path, variable, data_type, year_range, month_range, keep_zip){
  options(prism.path = local_path)
  if(data_type =="monthlys"){
    get_prism_monthlys(type = variable, year = year_range, mon = month_range, keepZip = keep_zip) 
  } else if (data_type == "annual") {
    get_prism_annual()
  } else if (data_type == "dailys"){
    get_prism_dailystype(type = variable, year = year_range, mon = month_range, keepZip = keep_zip)
  } else if (data_type == "normals"){
    get_prism_normals(type = variable, year = year_range, mon = month_range, keepZip = keep_zip)
  } 
  files <- ls_prism_data(name= T)
  
  return(files)
}


########### function to create dataframe of precipitation and summary stats ###########

create_prism_dataframe <- function(local_path, month){
  options(prism_path = local_path)
  files <- ls_prism_data(name = T)
  month <- toString(month)
  mon_index_pad <- str_pad(month, 2, pad = "0")
  mon_indexes <- str_split(mon_index_pad, "", simplify=F)
  mon_index1 <- mon_indexes[[1]][1]
  mon_index2 <- mon_indexes[[1]][2]
  month_files <- grep(paste("_[0-9]{4}[",mon_index1,"][",mon_index2,"]", sep = ""),ls_prism_data()[,1],value=T)
  RS <- prism_stack(month_files)
  proj4string(RS) <- CRS("+proj=longlat +ellps=WGS84 +towgs84=0,0,0,0,0,0,0 +units=m +no_defs")
  RS_DF <- data.frame(rasterToPoints(RS))
  names(RS_DF)[1:2] <- c("lon", "lat")
  options(scipen = 999)
  precip_df2 <- RS_DF[3:40]*0.0393700787
  precip_df2$lon <- RS_DF$lon
  precip_df2$lat <- RS_DF$lat
  precip_df2 <- precip_df2[, c(39, 40,1:38)]
  precip_df_final <- precip_df2 %>% 
    mutate(Precip_Mean = rowMeans(dplyr::select(precip_df2, -lon, -lat)),
           Precip_Min = apply(dplyr::select(precip_df2, -lon, -lat), 1, FUN = min),
           Precip_Max = apply(dplyr::select(precip_df2, -lon, -lat), 1, FUN = max))
  
  return(precip_df_final)
}

############# function to create plots of precipitation and summary stats #############

