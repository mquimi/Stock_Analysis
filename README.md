# Stock_Analysis with VBA


### 1. Overview of Project: Explain the purpose of this analysis.

While working on this project, I was able to incoporate my excel skills, along with my new VBA coding skills. The purpose of this project was to help Steve create a report that will help analyze new stocks for his parents to invest. While running the yearly stock reports of 2017 and 2018, we are also coding the time it takes to loop through all the data once. Then we will be comparing the refactored code to the original code to determine which analysis took longer.

### 2. Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
###### Stock Performance
 If you take a look at the 2017 and 2018 stocks, you can tell determine that 2018 was not a good year for any stock, eventhough ENPH and RUN are still highlighted green. Based on the images below, you can also make that conclusion. Image A is 2017 stocks analysis and Image B is 2018 stock analysis.
![alt text](https://github.com/mquimi/Stock_Analysis/blob/main/Resources/2017_Stocks.png) 
![alt text](https://github.com/mquimi/Stock_Analysis/blob/main/Resources/2018_Stocks.png)


###### Stock Report Execution Time
For the original code, we were using coding that included nested looping. The nested loop coding seem to have taken longer than the code with individualizing each loop. For example, Image A and C shows the original script report time, while Image B and D shows the refactored script report time. 

![alt text](https://github.com/mquimi/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017_GS.png)
![alt text](https://github.com/mquimi/Stock_Analysis/blob/main/Resources/VBA_Challenge_2017.png)

![alt text](https://github.com/mquimi/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018_GS.png)
![alt text](https://github.com/mquimi/Stock_Analysis/blob/main/Resources/VBA_Challenge_2018.png)

### 3. Summary: In a summary statement, address the following questions.
- What are the advantages or disadvantages of refactoring code?

Aside from cutting the time of how long the script runs, other advantages for refactoring code are that it makes it easier to read and understand, and you can "modifiy without changing the observable behavior by changing the internal structure." [(StackOverFlow)](https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software)

Accodring to the site [Rotate](https://rotate.cc/should-you-refactor-or-rewrite-your-code/), when you are refactoring, you can't add new functionality. Also, when refactoring, some of the complex functions that were being used may become difficult to manage.

- How do these pros and cons apply to refactoring the original VBA script?

As mentioned in the results section, one of the pros in applying refactoring to the original vba script was that it cut down on the execution time and eliminated unecessary code. 

