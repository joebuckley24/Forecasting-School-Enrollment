####################################################################

# created by Joe Buckley on Aug 3 2018
# Monte Carlo simulation of school enrollments for Statistical Forecasting LLC
# last modified Aug 23 2018

####################################################################

# this opens up a window to select your "working directory"
# make sure to choose the folder with District Confidence Intervals.xlsx in it
setwd(choose.dir())
# clear variables from any previous R sessions
rm(list=ls())

# this package "readxl" allows you to read excel files
# if not installed, type the following: install.packages("readxl")
library(readxl) # loads package
set.seed(2018) # make replicable, current year at the time of writing chosen as the seed
n = 5000 # run 5000 samples for Monte Carlo simulation
sheetname <- "District Confidence Intervals.xlsx"

# do this entire process 3 times, once per district, using this for-loop
# each district will have one csv as output
# each csv will contain 5 quantiles (5%, 16.5%, 50%, 83.5%, 95%) 
# for 5 years (2018-19, 2019-20, ..., 2022-23)
# for the following 17 groups: K, 1, ..., 12, elem (K-5), middle (6-8),
# high (9-12), and total (K-12)
districts <- c("Elizabeth", "Elizabeth_nooutlier", "Hackensack", "Passaic")
for (district in districts) {
  
  ####################################################################
  ## Inputs from spreadsheet
  ####################################################################
  if (district == "Elizabeth_nooutlier") {
  } else {
    # read in 5 year gpr average for year ending in 2018
    five_yr_roll_ave <- as.matrix(
      read_excel(sheetname, range = paste(district, "!X38:AJ38", sep = ""), col_names = F)
    )
    
    # read in matrix of deltas in gpr
    deltas_matrix <- as.matrix(
      read_excel(sheetname, range = paste(district, "!AL4:AX22", sep = ""), col_names = F)
    )
    
    # read in 5 year lagged births for each district
    # (actual for 2018-19 through 2021-22, projected 2022-23)
    lagged_births <- as.matrix(
      read_excel(sheetname, range = paste(district, "!B23:B27", sep = ""), col_names = F)
    )
    
    # read in most recent enrollment statistics for 2017-2018
    enrollment_17_18 <- as.matrix(
      read_excel(sheetname, range = paste(district, "!C22:O22", sep = ""), col_names = F)
    )
  }
  
  ####################################################################
  ## create data structures to store results
  ####################################################################
  
  # create blank 5 x 5000 x 17 array to store Monte Carlo enrollment results by grade
  # dimensions represent 5 years, 5000 trials, for 17 groups
  # 17 groups are K, 1, ..., 12, elem, midl, high, total
  enrollment_results <- array(0, dim = c(5, 17, n))
  
  # create 17 blank 5 x 5 matrices to hold quantiles data, years as rows, for each district
  # dimensions represent 5 years, 5 quantiles (5%, 16.5%, 50%, 83.5%, 95%)
  # 17 groups are K, 1, ..., 12, elem, midl, high, total
  # initialize empty list of 17 matrices
  quantiles_by_grade <- rep(list(matrix(NA, 5, 5)), 17)
  
  # create vector of strings called group_names to help name matrices
  groups <- vector(mode = "character", length = 17)
  groups[1] = "K"
  for (i in 1:12)
    groups[i + 1] <- as.character(i)
  groups[14:17] <- c("elem", "midl", "high", "total")
  group_names <- paste(rep("quantiles", 17), groups, sep = ".")
  # assign names to quantile matrices in list using vector of strings created above
  names(quantiles_by_grade) <- group_names
  for (i in 1:17) {
    colnames(quantiles_by_grade[[i]]) <- c("5%", "16.5%", "50%", "83.5%", "95%")
    row.names(quantiles_by_grade[[i]]) <- 
      c("2018-19", "2019-20", "2020-21", "2021-22", "2022-23")
  }
  
  ##########################
  ## outlier removal code ##
  ##########################
  
  if (district == "Elizabeth_nooutlier") {
    # read in kindergarten/birth ratios from spreadsheet
    m_b_k <- as.matrix(
      read_excel(sheetname, range = "Elizabeth!X3:X22", col_names = F)
    )
    # create new column for k/b ratios
    m_b_k.outlier_removed <- rep(NA, 19)
    # add values (excluding outlier) to this column
    m_b_k.outlier_removed[1:3] <- m_b_k[1:3]
    m_b_k.outlier_removed[4:19] <- m_b_k[5:20]
    # create new column for deltas from k/b ratios
    deltas.outlier_removed <- rep(NA, 18)
    # populate deltas column
    for(i in 1:18)
      deltas.outlier_removed[i] <- m_b_k.outlier_removed[i + 1] - m_b_k.outlier_removed[i]
    # replace col of m b-k deltas with modified column (outlier removed)
    deltas_matrix[, 1] <- c(deltas.outlier_removed, NA)
  }
  
  ##########################
  ##########################
  
  gpr_projections <- matrix(NA, 5, 13) # blank matrix for projected gpr
  enroll_projections <- matrix(NA, 6, 13) # blank matrix for projected student enrollment
  enroll_projections[1, ] <- enrollment_17_18 # first row is actual 2017-18 enrollment
  
  ####################################################################
  ## run 5000 trials to get enrollment projects for next 5 years
  ####################################################################
  
  for(i in 1:n) {
    
    if (district == "Elizabeth_nooutlier") {
      # 18 k/b deltas from which to choose if outlier removed
      sampl <- sample(1:18, 5, replace = T)
    } else {
      # 19 kindergarten/birth deltas from which to choose
      # only 18 for the other years
      sampl <- sample(1:19, 5, replace = T)
    }
    
    # add deltas to 5 year average for projected k/b ratios
    gpr_projections[, 1] <- deltas_matrix[sampl, 1] + five_yr_roll_ave[, 1]
    # recover projected enrollment by multiplying births by gpr projections
    enroll_projections[2:6, 1] <- lagged_births*gpr_projections[, 1]
    for(j in 2:13) {
      sampl <- sample(2:19, 5, replace = T) # 18 gpr deltas from which to choose
      # add deltas to 5 year average for projected gpr for k->1 through 11->12
      gpr_projections[, j] <- deltas_matrix[sampl, j] + five_yr_roll_ave[, j]
      # recover projected enrollment by multiplying student enrollment by gpr projections
      enroll_projections[2:6, j] <- enroll_projections[1:5, j - 1] * gpr_projections[, j]
    }
    # put enrollment projections into results matrix
    enrollment_results[ , 1:13, i] <- enroll_projections[2:6, ]
    for(r in 1:5) {
      # sum enrollments to get elementary, middle, high, and total school enrollment
      enrollment_results[r, 14, i] <- sum(enrollment_results[r, 1:6, i])
      enrollment_results[r, 15, i] <- sum(enrollment_results[r, 7:9, i])
      enrollment_results[r, 16, i] <- sum(enrollment_results[r, 10:13, i])
      enrollment_results[r, 17, i] <- sum(enrollment_results[r, 14:16, i])
    }
  }
  
  ####################################################################
  ## Extract quantile measurements from Monte Carlo simulation
  ####################################################################
  
  for (g in 1:17) {# for each group (K, 1, ..., 12, elem, ..., total)
    for (r in 1:5) {# for each year (2018-19, ..., 2022-23)
      # calculate quantiles for that particular year (r) for that particular group (g)
      # enrollment_results[r, g, ] is vector of 5000 trials, extract quantiles from this
      quantiles_by_grade[[g]][r, ] <- quantile(enrollment_results[r, g, ],
                                               prob = c(.05, .165, .5, .835, .95))
    }
  }
  
  ####################################################################
  ## Output results to csv file for each district
  ####################################################################
  
  # create matrix to hold all data for each district together
  output <- matrix(NA, 1, 6)
  # bind each group k, ..., 12, elem, ..., total successively to output
  for(i in 1:17) {
    output <- rbind(output,
                    cbind(names(quantiles_by_grade)[i],
                          matrix(colnames(quantiles_by_grade[[i]]), 1, 5)),
                    cbind(matrix(NA, 5, 1),
                          quantiles_by_grade[[i]]),
                    rep(NA, 6))
  }
  # write output matrix to csv
  write.csv(output, paste(district, "output.csv", sep = "_"), na = "")
  
}
