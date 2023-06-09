# Stock Analysis - Module 2 Challenge
<sub>UNC Chapel Hill Data Analytics Bootcamp</sub>

## Overview of Project
For Module 2, VBA was used with Microsoft Excel to calculate and display stock data from 2017 and 2018 for 12 companies. The total daily volume and return were calculated for 12 companies to compare total daily volume and return percentages for the years 2017 and 2018.

## Results & Analysis
### Stock Performance
#### 2017
In the year 2017, FSLR had the highest total daily volume at 684,181,400 units. DQ had the lowest total daily volume to 35,796,200. Across the companies in the dataset, the average total daily volume was 263,886,592, and the median was 199,258,700. All companies being tracked had a positive return except for TERP, which had a return of -7.2%. The top three companies with the highest returns were DQ, SEDG, and ENPH with returns of 199.4%, 184.5%, and 129.5% respectively. The average return over all companies was 67.3%, and the median was 41.5%.

#### 2018
For 2018, ENPH had the highest total daily volume at 607,473,500 units, and the lowest was AY at 83,079,900 units. The average total daily volume was 275,503,183, and the median was 179,594,450 units. The only companies in 2018 with a positive return were ENPH and RUN at 81.9% and 84.0%, respectively. The company with the worst return was DQ with a return of 62.6%. Across the companies being tracked, the average return was -8.5% and the median was -12.0%.

![2017 Stock Performance](/images/VBA_Challenge_2017.png) ![2018 Stock Performance](/images/VBA_Challenge_2018.png)
#### Comparison
When comparing the data gathered from the two years, 2017 had the highest and lowest total daily volumes, and the median of total daily volumes was nearly 20 million units higher than that in 2018. However, 2018 had an average total daily volume 11.6 million units higher than that of 2017. 2018 also had a combined total daily volume of 3.31 billion units, compared to 2017's 3.17 billion units. Of the 12 companies, AY, CSIQ, FSLR, JKS, and SPWR had higher total daily volumes in 2017 than 2018. In terms of stock return, only RUN and TERP had higher returns in 2018 than in 2017. Even though TERP's return increased, it was from -7.2% to -5.0%, remaining a negative value.

RUN showed the highest overall rate of growth, nearly doubling its total daily volume and increasing its return from 5.5% to 84.0%. When looking solely at 2018 data, ENPH appears to be highly successful, with the highest total daily volume of the listed companies and a positive return. Although its total daily volume increased more than twofold, 2018's return was down to 81.9% compared to 2017's 129.5%, showing decreased rate of return compared to its total daily volume.

### Execution Times
The original code took 0.508 and 0.633 seconds to analyze the data for 2017 and 2018, respectively.
![2017 Non-Refactored Time](/images/VBA_Challenge_2017_Time_Not_Refactored.png) ![2017 Refactored Time](/images/VBA_Challenge_2018_Time_Not_Refactored.png)

The refactored code took 0.086 and 0.055 seconds to analyze the data for 2017 and 2018, respectively.
![2018 Non-Refactored Time](/images/VBA_Challenge_2017_Time.png) ![2018 Refactored Time](/images/VBA_Challenge_2018_Time.png)

## Summary
Through refactoring, code can be made clearer, more consistent, and more efficient. However, refining the code may also introduce bugs, which can be difficult to resolve. For the macro used in this module, refactoring was far more useful than harmful. The original script contained a single large `for` loop to initialize and calculate the stock volume, compute starting and ending prices, and output the results on a separate worksheet. Through refactoring, the macro was streamlined by separating processes into smaller loops, working on one worksheet at a time, and storing values in arrays rather than immediately writing them to the analysis output worksheet. `Option Explicit` was also implemented - forcing all variables to be declared before use - to ensure all variables are spelled and used correctly. These changes resulted in greatly reduced computing times and increased readability of the macro's code.
