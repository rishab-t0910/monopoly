I'll be using this README file to track each project, and a short summary of it. These projects have come about during the partial lockdown. With nothing to do, some topics have sparked my interest and I've been creating projects about them. 

1. Run Tracker

This project came about because I wanted to learn how to combine python and excel (using openpyxl). As I run a frequent amount, I thought this would be the best way to combine the two as I track my stats on an excel sheet. 
The project takes in the date, duration, distance of the run and creates an excel sheet with this informatio, as well as overall analysis of pace, distance, time, and individual run pace. 

Edit on 19/05/21: Sorted out some bugs with the program. Solved (a lot of) bugs! Added the option for the user to query (find) values, as well as some usability stuff (allow the user to re-confirm their run, exit at any point). In the coming days, I'll look to improve the quality of the code, as right now it's a bit clunky and a tad unreadable. 

Edit on 21/05/21: Found out more ways that the program did not function. Added a check for invalid input (eg: words for distance). Formatted the code to look cleaner. 

Edit on 26/05/21: It seems as though everytime I use the code, I find more bugs. Added two new methods, floatCheck and intCheck using ValueError exception handling to weed out any false inputs, such as alphabets where there should be numbers. Trying to make the code cleaner on behest of a mate, but there are so many cases for each input it is tough. Sad that there isn't a switch statement in python, would've been easier to sort it out.

2. Master Schedule

This project tries to schedule a 8, 16, and 32 person squash draw. While 8 is straightforward in how to schedule, 16 and 32 are not. I tried to find the best way to schedule 16 and 32 without hardcoding it, but it is quite tough. Still trying to think of a solution. Maybe recrusive, and call the smaller method?

3. Monopoly

Created a simulation for a monopoly game to figure out which properties are most landed on. I adjusted the way I calculated some of the values: the chance/community chest values are the entire chance/community chest percentage, not the specific property. Meaning, the final value is the overall chance of landing on any of the community chests. Might include some ROI information if I get the opportunity. Will need the specifics on houses/hotels cost, mortgage, rent etc. Good project addition. 

Edit on 01/06/21: Created some nicer formatting for the excel spreadsheet, inlcluding colour and, percentages. I thought about the ROI information and realised there are far too many variations to accurately create the best information. I believe that the percentage information will provide enough information about which properties should be purchased. Evidently, jail is the most visited space as there are many ways of getting there. 

Edit on 04/06/21: Cleaned up more of the code whilst brainstroming how to include the ROI. Considering aborting python for that and using a purely theoretical calculation using Markov Chain and Matrix Representations. TBC.

Edit on 02/08/21: Been almost a month looking at this. I realised there was always a round-off error in the excel sheet that came from the program. Have solved the issue. Still learning about Markov Chain in an attempt to integrate ROI for the properties. 
