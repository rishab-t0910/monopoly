Edit on 27/02/24: I am revamping this project as it is outdated and I have more understanding of Data Analysis in Python. I have remade the simulation portion of the project. Here are my future goals: 
- Plot the most frequented squares
- Plot the most frequented group of squares - colour of property, type
- Construct a ROI plot based on the following conditions:
  - If someone has not landed on the property yet, it will be bought
  - Every subsequent landing by the buyer on the property will force the player to improve on the house (+1 house or hotel)
  - Every player has infinite money
- Summarise the above 3 points to make an optimal playing strategy

Created a simulation for a monopoly game to figure out which properties are most landed on. I adjusted the way I calculated some of the values: the chance/community chest values are the entire chance/community chest percentage, not the specific property. Meaning, the final value is the overall chance of landing on any of the community chests. Might include some ROI information if I get the opportunity. Will need the specifics on houses/hotels cost, mortgage, rent etc. Good project addition. 

Edit on 01/06/21: Created some nicer formatting for the excel spreadsheet, inlcluding colour and, percentages. I thought about the ROI information and realised there are far too many variations to accurately create the best information. I believe that the percentage information will provide enough information about which properties should be purchased. Evidently, jail is the most visited space as there are many ways of getting there. 

Edit on 04/06/21: Cleaned up more of the code whilst brainstroming how to include the ROI. Considering aborting python for that and using a purely theoretical calculation using Markov Chain and Matrix Representations. TBC.

Edit on 02/08/21: Been almost a month looking at this. I realised there was always a round-off error in the excel sheet that came from the program. Have solved the issue. Still learning about Markov Chain in an attempt to integrate ROI for the properties. 
