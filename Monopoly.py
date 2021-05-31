#player positions run from 0 - 39 (need to mod 40 on every roll)
#0 - go, 1 = medeter ave, 2 = comm chest ... 39 = boardwalk
#0 = go, 10 = jail, 20 = free parking, 30 = go to jail
#railroads in the middle of each lane @ 5, 15, 25, 35
#utilities 2 away from jails @ 12 & 28
#tax @ 4 & 38
#chance at 8, 22, 26.
#comm chest @ 2, 17, 33

'''Results of this that can be interpreted: which property is most likely to be landed on'''

from openpyxl.styles import Font, Alignment
from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime
import openpyxl
import random

def shuffle(value): #chance and comm chest card only [1,16]

    result = {}

    for i in range(1, 17):
        rand_val = random.randint(1, 16) #16 cards for chance/comm chest
        
        while (value[rand_val] is None): #Ensure no repeats in the result
            rand_val = random.randint(1, 16)
            
        result[i] = value[rand_val]
        value[rand_val] = None
        
    return result

def newPos(curr, card_val): 
    ''' Check based on current position if there is a needed move. 
        Eg: Chance card, comm chest card
    '''

    if (0 <= card_val <= 39): #move to point on board
        curr = card_val
    
    elif (0 > card_val): #back 3 spaces
        curr += card_val

    elif (card_val == 41): #utility
        if (13 <= curr <= 28):
            curr = 28
            
        else:
            curr = 12

    elif (card_val == 42): #railroad
        if (6 <= curr <= 15):
            curr = 15

        elif (16 <= curr <= 25):
            curr = 25

        elif (26 <= curr <= 35):
            curr = 35

        else:
            curr = 5

    return curr % 40
    
def clean(value): #print the dictionary nicely

    length = len(value) - 1
    
    first_key = list(value.keys())[0]
    last_key = list(value.keys())[length]
    
    for i in range(first_key, last_key+1):
        print('{}: {}'.format(i, round(value[i],2)))

    return ''
    
def rollDice(): #individual dice to allow for doubles
    
    return random.randint(1,6)

def community(comm_bank): #increment the deck

    first_value = comm_bank[1]

    for i in range (1, 16):
        comm_bank[i] = comm_bank[i+1]

    comm_bank[16] = first_value

    return comm_bank

def chance(chance_cards): #increment the deck
    
    first_value = chance_cards[1]

    for i in range (1, 16):
        chance_cards[i] = chance_cards[i+1]

    chance_cards[16] = first_value

    return chance_cards

'''-----------------------------------------------------------------------------------'''
#wb = Workbook()
#wb.save(filename = 'Monopoly.xlsx')
wb = load_workbook('Monopoly.xlsx')
sheet = wb.active

head_1 = sheet.cell(row = 1, column = 1)
head_1.value = 'Location'
head_1.font = Font(name = 'Times New Roman', size = 12, bold = True)
head_1.alignment = Alignment(horizontal="center", vertical="center")

head_2 = sheet.cell(row = 1, column = 2)
head_2.value = 'Count'
head_2.font = Font(name = 'Times New Roman', size = 12, bold = True)
head_2.alignment = Alignment(horizontal="center", vertical="center")

head_3 = sheet.cell(row = 1, column = 3)
head_3.value = 'Percentage'
head_3.font = Font(name = 'Times New Roman', size = 12, bold = True)
head_3.alignment = Alignment(horizontal="center", vertical="center")

sheet.column_dimensions['A'].width = 19.5
sheet.column_dimensions['B'].width = 19.5
sheet.column_dimensions['C'].width = 19.5

boardName = ['Go', 'Mediterranean Avenue', 'Community Chest', 'Baltic Avenue', 'Income Tax', 'Reading Railroad', 'Oriental Avenue', 'Chance', 'Vermont Avenue', 'Connecticut Avenue', 'Jail', 'St. Charles Place', 'Electric Company', 'States Avenue', 'Virginia Avenue', 'Pennsylvania Railroad', 'St. James Place', 'Tennessee Avenue', 'New York Avenue', 'Free Parking', 'Kentucky Avenue', 'Indiana Avenue', 'Illinois Avenue', 'B & O Railroad', 'Atlantic Avenue', 'Ventnor Avenue', 'Water Works', 'Marvin Gardens', 'Go To Jail', 'Pacific Avenue', 'North Carolina Avenue', 'Pennsylvania Avenue','Short Line', 'Park Place', 'Luxury Tax', 'Boardwalk']

for i in range(0, len(boardName)):
    sheet.cell(row = i+2, column = 1).value = boardName[i]
    sheet.cell(row = i+2, column = 1).font = Font(name = 'Times New Roman', size = 12)
    sheet.cell(row = i+2, column = 1).alignment = Alignment(horizontal="center", vertical="center")
    
    sheet.cell(row = i+2, column = 2).font = Font(name = 'Times New Roman', size = 12)
    sheet.cell(row = i+2, column = 2).alignment = Alignment(horizontal="center", vertical="center")
    sheet.cell(row = i+2, column = 2).value = None
    
    sheet.cell(row = i+2, column = 3).font = Font(name = 'Times New Roman', size = 12)
    sheet.cell(row = i+2, column = 3).alignment = Alignment(horizontal="center", vertical="center")
    sheet.cell(row = i+2, column = 3).value = None
    
comm_chest = {1: 0, 2: 10}
chance_card = {1: 39, 2: 0, 3: 30, 4: 24, 5: 11, 6: 5, 7: 42, 8: 42, 9: 41, 10: -3, 11: 99, 12: 99, 13: 99, 14: 99, 15: 99, 16: 99}
jail_turn = {}
player_pos = {}
loc_count = {}
    
for i in range (3, 17):
    comm_chest[i] = 99

comm_chest = shuffle(shuffle(comm_chest))
chance_card = shuffle(shuffle(chance_card))

for i in range(1, 5):
    player_pos[i] = 0
    jail_turn[i] = 0
    
for i in range(0, 40):
    loc_count[i] = 0

for a in range(0, 1000000):
    for i in range (1, 5):
        double_count = 0

        while(double_count < 3):
            roll_1 = rollDice()
            roll_2 = rollDice()
            dice_roll = roll_1+roll_2
            isDouble = False

            if(roll_1 == roll_2):
                double_count += 1
                isDouble = True
            
            if(double_count == 3):
                player_pos[i] = 10
                loc_count[player_pos[i]] += 1
                jail_turn[i] = 3
                isDouble = False
                break
                
            if(jail_turn[i] == 0): #if they are not in jail
                
                player_pos[i] = (player_pos[i]+dice_roll) % 40
                loc_count[player_pos[i]] += 1

                if(player_pos[i] == (2 or 17 or 33)): #comm chest
                    player_pos[i] = newPos(player_pos[i], comm_chest[1])
                    comm_chest = community(comm_chest)

                elif(player_pos[i] == (8 or 22 or 36)): #chance
                    player_pos[i] = newPos(player_pos[i], chance_card[1])
                    chance_card = chance(chance_card)

                if(player_pos[i] == 30):
                    jail_turn[i] = 3
                    player_pos[i] = 10
                
            else:
                if(isDouble is True):
                    jail_turn[i] == 0

                else:
                    jail_turn[i] += -1
                    
            loc_count[player_pos[i]] += 1

            if(not isDouble):
                break

loc_count [8] = loc_count[8] + loc_count[22] + loc_count[36] #combine all the chance
loc_count[22] = -1
loc_count[36] = -1
loc_count [2] = loc_count[2] + loc_count[17] + loc_count[33] #combine all the comm chest
loc_count[17] = -1
loc_count[33] = -1

sum_ = 0
for a in range(1, 40):
    if (loc_count[a] >= 0):
        sum_ += loc_count[a]

#print(clean(player_pos))
print(clean(loc_count))
print(sum_, '\n')

ave_count = {}
for a in range(0, 40):
    if (loc_count[a] >= 0):
        ave_count[a] = round(100*(loc_count[a]/sum_),2)
    else:
        ave_count[a] = -1

for a in range(2, 38):
    if (a <= 18):
        sheet.cell(row = a, column = 2).value = loc_count[a-2]
        sheet.cell(row = a, column = 3).value = ave_count[a-2]

    elif (19 <= a <= 22):
        sheet.cell(row = a, column = 2).value = loc_count[a-1]
        sheet.cell(row = a, column = 3).value = ave_count[a-1]

    elif (23 <= a <= 32):
        sheet.cell(row = a, column = 2).value = loc_count[a]
        sheet.cell(row = a, column = 3).value = ave_count[a]

    elif (33 <= a <= 34):
        sheet.cell(row = a, column = 2).value = loc_count[a+1]
        sheet.cell(row = a, column = 3).value = ave_count[a+1]

    elif (35 <= a <= 37):
        sheet.cell(row = a, column = 2).value = loc_count[a+2]
        sheet.cell(row = a, column = 3).value = ave_count[a+2]

print(clean(ave_count))
wb.save("Monopoly.xlsx")
