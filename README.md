# Stock-Analysis
## Overview of project:
Stock Analysis project involves playing Steve to analyze the entire data set and expanding the data set to include entire stock market analysis over last few years. Steve’s parents are passionate about green energy and decided to invest all their money in DAQO new energy corporation. Steve thinks his parent’s funds should be more diversified, so he analyzes several green energy stocks in addition to DAQO.

### Purpose:
Purpose of this project is to utilize the VBA scripting and github skills learned in module 2 for analyzing the data set and editing refractor code to see if VBA script runs faster and write a report to include the findings from data analyzed.

## Results
Steve analyzed that DAQO stocks had gain of 200% in 2017 however in 2018 had loss of 63%. Since other stocks were analyzed too, it gave clear picture of which stocks the money can be invested.

Refactoring code helped in running the script fast. The reason behind the script running fast is due to the ways loops were nested or used initially versus non-nested loop in later.
Original 2017 code took 0.7 sec whereas refactor code ran in 0.1 sec. original 2018 code ran in 0.7 sec whereas refactored ran in 0.09sec.

### Original 2017

![Original 2017](Resources/Original 2017.png)
[https://github.com/JyotsnaMuppasani/Stocks-Analysis/blob/573b035a898c355906df5fcd808ee9100c4886c0/Resources/Original%202017.png]

### Refactored 2017

![Resources/VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)
[https://github.com/JyotsnaMuppasani/Stocks-Analysis/blob/573b035a898c355906df5fcd808ee9100c4886c0/Resources/VBA_Challenge_2017.png]

### Original 2018

![Resources/Original 2018](Resources/Original 2018.png)
[https://github.com/JyotsnaMuppasani/Stocks-Analysis/blob/573b035a898c355906df5fcd808ee9100c4886c0/Resources/Original%202018.png]

### Refactored 2018

![Resources/VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)
[https://github.com/JyotsnaMuppasani/Stocks-Analysis/blob/573b035a898c355906df5fcd808ee9100c4886c0/Resources/VBA_Challenge_2018.png]

## Summary: 
### Advantages of refactoring code:
Better comprehensibility facilitates maintenance and the extendibility of the software
Restructuring the source code is possible without altering the functionality
Improved legibility improves the comprehensibility of the code for other programmers
Removed redundancies and duplications improve the effectiveness of the code
Self-contained methods prevent local changes from having an effect on other parts of the code
Clean code with shorter, self-contained methods and classes is characterized by better testability

### Disadvantages of refactoring code:
Imprecise refactoring could introduce new bugs and errors into the code
There is no clear definition of “neat code”
An improved code is often difficult for the customer to recognize, since the functionality stays the same, i.e. the benefit is not self-evident
In the case of larger teams working on refactoring, the coordination effort required could be surprisingly high
(Reference: https://www.ionos.com/digitalguide/websites/web-development/what-is-refactoring/)
### Pros and cons of refactoring the original VBA script
Refactor code successfully loop through all the data one time in order to collect the same information thus increasing speed of the code run. If code is not refactored properly then new bugs and errors are introduced in the code.


