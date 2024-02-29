**Edit on 28/02/24**: Finished the data analysis portion using Matplotlib and NumPy. Here are the findings
- Chance, Community Chest, and Jail are the most landed on properties
- Other than brown and blue properties,  all other properties have roughly the same proportion of landings
- Green properties make the most profit, and brown the least

My next steps is to simulate the game, and run the same analysis with different buying strategies
- Aggressive: Buy every property landed on
- Average: Buy 1 in 2 properties landed on
- Conservative: Buy 1 in 4 properties landed on

After this, I want to run the game with the strategies, and giving each player a finite amount of money before showing the visualizations. 

**Edit on 27/02/24**: I am revamping this project as it is outdated and I have a better understanding of Data Analysis in Python. I will no longer using OpenPyxl, but instead moving into Matplotlib for my EDA. I have remade the simulation portion of the project. Here are my future goals: 
- Plot the most frequented squares
- Plot the most frequented group of squares - colour of property, type
- Construct a ROI plot based on the following conditions:
  - If someone has not landed on the property yet, it will be bought
  - Every subsequent landing by the buyer on the property will force the player to improve on the house (+1 house or hotel)
  - Every player has infinite money
- Summarise the above 3 points to make an optimal playing strategy

Created a simulation for a monopoly game to figure out which properties are most landed on. I adjusted the way I calculated some of the values: the chance/community chest values are the entire chance/community chest percentage, not the specific property. Meaning, the final value is the overall chance of landing on any of the community chests. 
