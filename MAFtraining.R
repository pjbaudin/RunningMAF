# Load library packages
library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)

# Set working directory
setwd("C:/Users/PB/SkyDrive/DataScience/MAF - Marathon Training")

# Import sheet names
FileName <- "MAFtrackrecord.xlsx" 

SheetNames <- excel_sheets(FileName)

# The palette with grey:
cbPalette <- c("#999999", "#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2", "#D55E00", "#CC79A7")

# Plot function for each dataset (long run and daily run)

FileNameXL <- "MAFtrackrecord.xlsx" 

plotMAF <- function(i) {
        
        # Import original dataset from Excel spreadsheet
        df <- read_excel(FileNameXL, sheet = i)
        
        # Format the duration and pace time
        df$Duration <- format(df$Duration,"%H:%M")
        df$Pace <- format(df$Pace, "%H:%M")
        
        # Process data to add week, month, duration object for Pace and time
        # and calculate the average speed
        df <- df %>%
                mutate(Month = month(Date, label = TRUE)) %>%
                mutate(Week = as.factor(week(Date))) %>%
                mutate(AverageSpeed = (Distance * as.numeric(duration(minutes = 60)))
                       / as.numeric(hm(Duration))) %>%
                mutate(AveragePace = as.duration(ms(Pace))) %>%
                mutate(DurationTime = as.duration(hm(Duration)))
        
        # saving dataset in csv
        FileName <- paste(i, "_dataset.csv", sep = "")
        write.csv(df, file = FileName)
        
        # Extract summary of dataset and export in csv format
        MAFsum <- summary(df)
        FileName <- paste(i, "_Summary.csv", sep = "")
        write.csv(MAFsum, file = FileName)
        
        # Plotting exploration
        MAFplot <- ggplot(df, aes(Date, as.factor(Pace), colour = as.factor(Duration))) +
                geom_point(size = 3, alpha = 0.5) +
                scale_color_manual(values = cbPalette) +
                geom_vline(xintercept = c(5, 6, 7, 8),
                           colour = "steelblue", lty = 4, lwd = 1) +
                ylab("Pace in min/km")
        
        FileName <- paste(i, "_DatevsPace.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(Date, as.numeric(ms(Pace)) / 60, colour = as.factor(Duration))) +
                geom_point(size = 3, alpha = 0.5) +
                scale_color_manual(values = cbPalette) +
                geom_smooth(method = "lm", se = FALSE) +
                geom_hline(yintercept = c(5, 6, 7, 8), colour = "steelblue", lty = 4, lwd = 1) +
                ylab("Pace in min/km")
        
        FileName <- paste(i, "_DatevsPaceCalc.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(as.factor(Pace), AverageHR, colour = as.factor(Duration))) +
                geom_point(size = 3, alpha = 0.5) +
                geom_smooth(method = "lm", se = FALSE) +
                geom_hline(yintercept = c(135, 145), colour = "red", lty = 3) +
                scale_color_manual(values = cbPalette) +
                xlab("Pace in min/km")
        
        FileName <- paste(i, "_AverageHRvsPace.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(as.numeric(ms(Pace)) / 60, AverageHR, colour = as.factor(Duration))) +
                geom_point(size = 3, alpha = 0.5) +
                geom_hline(yintercept = c(135, 145), colour = "red", lty = 3) +
                scale_color_manual(values = cbPalette) +
                geom_smooth(se = FALSE, method = "lm") +
                xlab("Pace in min/km")
        
        FileName <- paste(i, "_AverageHRvsPaceCalc.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(y = as.numeric(ms(Pace)) / 60,
                          x = AverageHR,
                          colour = as.factor(Duration))) +
                geom_point(size = 3, alpha = 0.5) +
                geom_vline(xintercept = c(135, 145), colour = "red", lty = 3) +
                geom_hline(yintercept = c(5, 6, 7, 8), colour = "steelblue", lty = 4, lwd = 1) +
                scale_color_manual(values = cbPalette) +
                geom_smooth(se = FALSE, method = "lm") +
                ylab("Pace in min/km")
        
        FileName <- paste(i, "_PaceCalcvsAverageHR.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(Date, AverageSpeed, colour = as.factor(Duration))) +
                geom_point(size = 3, alpha = 0.5) +
                scale_color_manual(values = cbPalette) +
                geom_smooth(method = "lm", se = FALSE) +
                geom_hline(yintercept = c(6, 7, 8, 9, 10, 11),
                           colour = "steelblue", lty = 4, lwd = 1) +
                ylab("Average Speed in Km/h")
        
        FileName <- paste(i, "_DatevsAverageSpeed.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(y = as.numeric(ms(Pace)) / 60,
                          x = Distance,
                          colour = as.factor(Duration))) +
                geom_point(size = 3, alpha = 0.5) +
                geom_vline(xintercept = 5, colour = "red", lty = 3) +
                geom_hline(yintercept = c(5, 6, 7, 8), colour = "steelblue", lty = 4, lwd = 1) +
                scale_color_manual(values = cbPalette) +
                geom_smooth(se = FALSE, method = "lm") +
                ylab("Pace in min/km")
        
        FileName <- paste(i, "_DistancevsPaceCalc.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(Date, Distance, colour = as.factor(Duration))) +
                geom_point(size = 3, alpha = 0.5) +
                scale_color_manual(values = cbPalette) +
                geom_smooth(method = "lm", se = FALSE) +
                geom_hline(yintercept = c(0, 5, 10), colour = "steelblue", lty = 4, lwd = 1) +
                ylab("Distance in km")
        
        FileName <- paste(i, "_DatevsDistance.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(Date, Distance, fill = as.factor(Duration))) +
                geom_col(colour = "darkgrey", alpha = 0.5) +
                geom_smooth(se = FALSE, method = "lm") +
                scale_fill_manual(values = cbPalette) +
                geom_hline(yintercept = c(0, 5, 10),
                           colour = "steelblue", lty = 4, lwd = 1) +
                ylab("Distance in km")
        
        FileName <- paste(i, "_DatevsDistanceCol.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(Month, Distance, fill = as.factor(Duration))) +
                geom_col(colour = "darkgrey", alpha = 0.5) +
                scale_fill_manual(values = cbPalette) +
                geom_hline(yintercept = c(0, 10, 20, 30),
                           colour = "steelblue", lty = 4, lwd = 1) +
                ylab("Distance in km")
        
        FileName <- paste(i, "_MonthvsDistanceCol.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(df, aes(Week, Distance, fill = as.factor(Duration))) +
                geom_col(colour = "darkgrey", alpha = 0.5) +
                scale_fill_manual(values = cbPalette) +
                geom_hline(yintercept = c(0, 10, 20),
                           colour = "steelblue", lty = 4, lwd = 1) +
                ylab("Distance in km")
        
        FileName <- paste(i, "_WeekvsDistanceCol.png", sep = "")
        ggsave(FileName, MAFplot)
        
        # Compute monthly stat
        MAFdfmonthstat <- df %>%
                group_by(Month) %>%
                summarize(MeanDuration = as.duration(mean(DurationTime)),
                          MeanPace = as.duration(mean(AveragePace)),
                          MeanSpeed = round(mean(AverageSpeed), digits = 2),
                          MeanDistance = round(mean(Distance), digits = 2),
                          MeanHR = round(mean(AverageHR), digits = 2),
                          MeanMAxHR = round(mean(MaxHR), digits = 2))
        
        # Export monthly stat
        FileName <- paste(i, "_Monthly Summary.csv", sep = "")
        write.csv(MAFdfmonthstat, file = FileName)
        
        # Compute weekly stat
        MAFdfweekstat <- df %>%
                group_by(Week) %>%
                summarize(MeanDuration = as.duration(mean(DurationTime)),
                          MeanPace = as.duration(round(mean(AveragePace), digits = 0)),
                          MeanSpeed = round(mean(AverageSpeed), digits = 2),
                          MeanDistance = round(mean(Distance), digits = 2),
                          MeanHR = round(mean(AverageHR), digits = 2),
                          MeanMAxHR = round(mean(MaxHR), digits = 2))
        
        # Export weekly stat
        FileName <- paste(i, "_Weekly Summary.csv", sep = "")
        write.csv(MAFdfweekstat, file = FileName)
        
        # Additional plots for weekly stat
        MAFplot <- ggplot(MAFdfweekstat, aes(as.numeric(Week), as.numeric(MeanPace) / 60)) +
                geom_point(size = 4, alpha = 0.5, fill = cbPalette[1]) +
                geom_path(colour = cbPalette[3], alpha = 0.5) +
                geom_hline(yintercept = c(5, 6, 7, 8),
                           colour = "steelblue", lty = 4, lwd = 1) +
                geom_smooth(method = "lm", se = FALSE) +
                ylab("Mean Pace in min/km")
        
        FileName <- paste(i, "_WeeklyvsMeanPace.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(MAFdfweekstat, aes(as.numeric(Week), MeanDistance)) +
                geom_point(size = 4, alpha = 0.5, fill = cbPalette[1]) +
                geom_path(colour = cbPalette[3], alpha = 0.5) +
                geom_hline(yintercept = c(5, 10),
                           colour = "steelblue", lty = 4, lwd = 1) +
                geom_smooth(method = "lm", se = FALSE) +
                ylab("Mean Distance in km")
        
        FileName <- paste(i, "_WeeklyvsMeanDistance.png", sep = "")
        ggsave(FileName, MAFplot)
        
        # Additional plots for monthly stat
        MAFplot <- ggplot(MAFdfmonthstat, aes(as.numeric(Month), as.numeric(MeanPace) / 60)) +
                geom_point(size = 4, alpha = 0.5, fill = cbPalette[1]) +
                geom_path(colour = cbPalette[3], alpha = 0.5) +
                geom_hline(yintercept = c(5, 6, 7, 8),
                           colour = "steelblue", lty = 4, lwd = 1) +
                geom_smooth(method = "lm", se = FALSE) +
                ylab("Mean Pace in min/km")
        
        FileName <- paste(i, "_WeeklyvsMeanPaceCalc.png", sep = "")
        ggsave(FileName, MAFplot)
        
        MAFplot <- ggplot(MAFdfmonthstat, aes(as.numeric(Month), MeanDistance)) +
                geom_point(size = 4, alpha = 0.5, fill = cbPalette[1]) +
                geom_path(colour = cbPalette[3], alpha = 0.5) +
                geom_hline(yintercept = c(5, 10),
                           colour = "steelblue", lty = 4, lwd = 1) +
                geom_smooth(method = "lm", se = FALSE) +
                ylab("Mean Distance in km")
        
        FileName <- paste(i, "_MonthlyvsMeanDistance.png", sep = "")
        ggsave(FileName, MAFplot)
        
        # Additionnal plot for prediction exploration
        MAFplot <- ggplot(df, aes(Date, Distance)) +
                geom_point(size = 4, alpha = 0.5, colour = "steelblue") +
                stat_smooth(method = "lm", fullrange = TRUE) +
                scale_x_datetime(date_breaks = "1 week", date_labels = "%W",
                                 limits = c(min(df$Date), max(df$Date) + days(10)))
        
        FileName <- paste(i, "_DatevsDistance_prediction.png", sep = "")
        ggsave(FileName, MAFplot)
        
        # smoothing method (function) to use, eg. "lm", "glm", "gam", "loess", "rlm".
        MAFplot <- ggplot(df, aes(Date, as.numeric(ms(Pace)) / 60)) +
                geom_point(size = 4, alpha = 0.5, colour = "steelblue") +
                stat_smooth(method = "gam", fullrange = TRUE) +
                scale_x_datetime(date_breaks = "1 week", date_labels = "%W",
                                 limits = c(min(df$Date), max(df$Date) + days(10)))
        
        FileName <- paste(i, "_DatevsPace_prediction.png", sep = "")
        ggsave(FileName, MAFplot)
        
}

map(SheetNames, plotMAF)
