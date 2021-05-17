import math

def printSeed(seeds): #get dict of seeds
    print("")
    print("\u0332".join("Seeding"), end = "\n\n")
    for i in range (1, len(seeds)+1):
        print('{}.'.format(i), seeds[i])
        
    print("\n")

def eight(grand):

    size = 8
    seeds = {}

    for i in range (1, size+1):
        if(i<=grand):
            print("Enter seed #", i, "name: ", end = "")
            seeds[i] = input("")
        else:
            seeds[i] = 'Bye'

    printSeed(seeds)

    matchup = {} #creates player pairing (A vs B)
    rank = {} #contains the seed pairing (1 vs 16)
    fin_matchup = {}
    size = len(seeds)
    back_counter = size
    
    print("\u0332".join("Matchup"), end = "\n\n")
    for i in range (1, int(size/2)+1):
        matchup[i] = seeds[i]+' vs '+seeds[back_counter]
        back_counter = back_counter-1
        rank[i] = str(i)+' vs '+str(size-i+1)
        print('{}:'.format(rank[i]), matchup[i])

    print("")
    print("\u0332".join("Final Draw"), end = "\n\n")
    print('{}:'.format(rank[1]), matchup[1])
    print('{}:'.format(rank[4]), matchup[4], '\n--------------------------------')
    print('{}:'.format(rank[3]), matchup[3])
    print('{}:'.format(rank[2]), matchup[2], '\n--------------------------------')
    
def sixt(grand): #grand tells me how many players will be playing
    
    size = 16
    seeds = {}

    for i in range (1, size+1):
        if(i<=grand):
            print("Enter seed #", i, "name: ", end = "")
            seeds[i] = input("")
        else:
            seeds[i] = 'Bye'

    printSeed(seeds)

    matchup = {} #creates player pairing (A vs B)
    rank = {} #contains the seed pairing (1 vs 16)
    fin_matchup = {}
    size = len(seeds)
    back_counter = size
    
    print("\u0332".join("Matchup"), end = "\n\n")
    for i in range (1, int(size/2)+1):
        matchup[i] = seeds[i]+' vs '+seeds[back_counter]
        back_counter = back_counter-1
        rank[i] = str(i)+' vs '+str(size-i+1)
        print('{}:'.format(rank[i]), matchup[i])

    print("") 
    print("\u0332".join("Final Draw"), end = "\n\n")
    print('{}:'.format(rank[1]), matchup[1])
    print('{}:'.format(rank[8]), matchup[8], '\n--------------------------------')
    print('{}:'.format(rank[4]), matchup[4])
    print('{}:'.format(rank[5]), matchup[5], '\n--------------------------------')
    print('{}:'.format(rank[3]), matchup[3])
    print('{}:'.format(rank[6]), matchup[6], '\n--------------------------------')
    print('{}:'.format(rank[7]), matchup[7])
    print('{}:'.format(rank[2]), matchup[2], '\n--------------------------------')


def thre(grand):
    
    size = 32
    seeds = {}

    for i in range (1, size+1):
        if(i<=grand):
            print("Enter seed #", i, "name: ", end = "")
            seeds[i] = input("")
        else:
            seeds[i] = 'Bye'

    printSeed(seeds)

    matchup = {} #creates player pairing (A vs B)
    rank = {} #contains the seed pairing (1 vs 16)
    fin_matchup = {}
    size = len(seeds)
    back_counter = size
    
    print("\u0332".join("Matchup"), end = "\n\n")
    for i in range (1, int(size/2)+1):
        matchup[i] = seeds[i]+' vs '+seeds[back_counter]
        back_counter = back_counter-1
        rank[i] = str(i)+' vs '+str(size-i+1)
        print('{}:'.format(rank[i]), matchup[i])

    print("")
    
    print("\u0332".join("Final Draw"), end = "\n\n")
    print('{}:'.format(rank[1]), matchup[1])
    print('{}:'.format(rank[16]), matchup[16], '\n--------------------------------')
    print('{}:'.format(rank[8]), matchup[8])
    print('{}:'.format(rank[9]), matchup[9], '\n--------------------------------')
    print('{}:'.format(rank[4]), matchup[4])
    print('{}:'.format(rank[13]), matchup[13], '\n--------------------------------')
    print('{}:'.format(rank[5]), matchup[5])
    print('{}:'.format(rank[12]), matchup[12], '\n--------------------------------')
    print('{}:'.format(rank[3]), matchup[3])
    print('{}:'.format(rank[14]), matchup[14], '\n--------------------------------')
    print('{}:'.format(rank[6]), matchup[6])
    print('{}:'.format(rank[11]), matchup[11], '\n--------------------------------')
    print('{}:'.format(rank[7]), matchup[7])
    print('{}:'.format(rank[10]), matchup[10], '\n--------------------------------')
    print('{}:'.format(rank[15]), matchup[15])
    print('{}:'.format(rank[2]), matchup[2], '\n--------------------------------')


size = -1
count = 1

while(size<=0 and count<=3):
    size = int(input('How many registered players?: '))
    count = count+1

if(size <=0):
    print('Goodbye, have a nice day!')
elif(size<=8):
    eight(size)
elif(size<=16):
    sixt(size)
elif(size<=32):
    thre(size)
else:
    print("No capacity for {} players.".format(size))
