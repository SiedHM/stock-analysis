
# Report for Module-2 Challenge

# 1.  Overview of Project.
Steeve, a fresh graduate in finance, wanted to help his parents who are interested to make an investment on green energy producing companies. Initially they wanted to make all their investment on one green energy company DAQO New Energy Corporation (DQ), but he is concerned diversifying their funds and hence he wanted to analyses a handful of green energy stocks in addition to DQ stocks. He made an excel data of the different green energy stocks with their prices and trade volumes. Steve wanted the stock return and trade volumes to be analyzed using VISUAL BUSINESS APPLICATION(VBA). As an expert in VBA, he asked me to help him out to do the analysis. His initial request was to do the analysis for the year 2018. I gave him the VBA codes and analytical results for the 2018. After looking the 2018 result, he wanted to include additional years and also refactor the VBA scripts made for the 2018.
# The purpose of the report is
i.	Explore green energy stock performance by analyzing the well street financial data using Visual Business Application (VBA) . Specifically,
•	Analyze the return of investment on green energy stocks over time
•	Analyze the total daily trading volume of green energy stocks over time
ii.	Refactor VBA scripts made for the year 2018 and discuss the advantages and disadvantages of refactoring in general and in particular related to the current code 
## 2.	Data and Method of analysis
 I used Wall Street Green Energy Stock data for the year 2017 and 2018.  It contains stocks of 12 green energy companies. The data contains company ticker, starting date, ending data, closing prices and trading volumes.  The data was obtained from Bootcamp for this exercise. I used the Visual Business Application (VBA) to generate Tabular results shown below. 



# 3.	Results
# 3.1.	  Analysis of stock investment return
Table-1 shows green energy stock returns for the years 2017 and 2018. Two colors are used to identify positive and negative returns. Green color shows positive return and red color shows negative returns. I used the VBA codes shown at the end of this subtopic to color cells in the table. Further, Initially the analysis was done only for 2018. However, using a ‘yearValue’ variable in the ‘InputBox’ syntax, we can produce results for both years 2017 and 2018(see sample codes at the end of this subtopic). 
Table-1: Returns (%) of green energy stock investment for the years 2017 and 2018
Ticker	Return (2017)	Return (2018)

![Table1](https://github.com/nebil2016/stock-analysis/blob/main/Table-1.png)
		
Results from Table-1 shows that most companies had positive return for the year 2017 except for TERP. The DAQO New Energy Corporation (DQ), which Steve parents were initially interested in to invest  all their money on, had in fact, the highest return(about 199%).  SEDG and ENPH were the second and third companies with highest return on their stock investment.   However, in 2018, most companies have negative stock investment return except EPNH and RUN.  EPNH had consistently positive stock investment return. DQ return had about -63% return. SEDG which had second highest positive return in 2017, had about -8% return, which is much higer than DQ’s return in 2018.  Given the stock return variability, I agree with Steeve that it is not good idea to make all money invested in one company. Given the results for 2017 and 2018, I advise his parents first to consider EPNH and add one or two additional companies like RUN, SEDG. 
  

# 3.2.	Analysis of Stock trade volumes
Stock trade volume measures the quantity(number) of shares traded in a given period of time, in the current case, daily. Volume can indicate market strength, as rising markets on increasing volume are typically viewed as strong and healthy (Cory Mitchell,2022) . In particular, stocks with positive change over time are an indication of rising markets and negative changes are indication of declining markets. 



Table-2 shows the total daily volume of stock trades in 2017 and 2018 for 12 companies. Like Table-1 above, the change in trade volume is colored with green and red colors. Green color is for positive changes in total trade volume between 2017 and 2018 and red for negative changes. From this table it can be shown that EPNH has large positive change in the volume of trade which may viewed as strong and healthy. DQ, which Steeve parents planned to make an investment on, had also positive change. 

![Table1](https://github.com/nebil2016/stock-analysis/blob/main/Table-2.png)


Given the results in Tables-1 and 2, we can now advice Steeve parents not to make all their investment in one company, DQ. In fact, there are other companies like EPNH which may be viewed as strong and healthy by financial analyst.  I advise his parents to consider companies with positive return and positive trading volumes. These are companies which can be predictable and viewed as strong and healthy. 
 




# 4.	Conclusions
The findings of this project imply that Steeve parents need to diversify their investment rather than investing all their money in one company. EPNH is the green energy company which has positive return in both years and also with largest positive total daily stock trading volume. Steeve parents can be benefited if they make investment on EPNH and additional companies with positive return in both return and change in trading volume. 
This project is based on refactoring of an original code which were made only one year. Refactoring is a restructuring an existing body of code, altering its internal structure without changing its external behavior (Martin Fowler's definition).  I learned from the internet that Refactoring has pros and cons in general. 


# The advantages of refactoring a code in general are
•	to make the code more efficient and maintainable
•	to improves readability
•	helps program run faster
•	to make the debugging process to go much more smoothly
•	fixes or prevents decay from occurring


The disadvantages of Code Refactoring: Time Consuming. One may have no idea how much time it may take to complete the process. Because of this it is advised not to refactor if one does not have time or is working near deadline. Stable codes are not also advised to be refactored.
In this particular assignment, the first advantage of refactoring is it generates results of two years simultaneously instead of having two separate codes for each   year. It also reduces the program processing time. Table-3 below shows the processing time for the year 2018 before and after refactoring. Processing time is reduced significantly. 

![Table1](https://github.com/nebil2016/stock-analysis/blob/main/Table-3.png)

The second advantage is that it adds new features like coloring of the results which increases the readability and detectability of results very easily. 
For beginner like me the guidance on refactoring help me to learn coding better and know that long codes can be shorten. As a beginner, I have not identified   disadvantage of refactoring in this exercise as the instructions and examples were clear. I find it very helpful. 


