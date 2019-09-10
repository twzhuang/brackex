# MASTERLIST MAKER PROGRAM
# 
# Workbook imported from Tournament software must be named "players.xlsx"
# 
# This program will create and format the masterlist to include all flights that each player is playing and their mixed partner and doubles partner
# 
# If a partner is jumping flights or signed up with mulitple partners, you can find that in masterErrors.txt


from openpyxl import load_workbook, Workbook
from openpyxl.styles.borders import Border, Side

# import Font function from openpyxl 
from openpyxl.styles import Font 

# function to parse through player list from tournament software
def parse():
    mp = open("multiplePartners.txt", "w+")
    jf = open("jumpFlight.txt", "w+")
    f = open("partnerNotFound.txt", "w+")
    wb = load_workbook('players.xlsx')
    sheet = wb['Players']

    # delete top 3 rows of unuseful info
    sheet.delete_rows(1,3)
    sheet[1][0].value = "Last Name"
    sheet[1][1].value = "First Name"

    # insert columns and set titles for columns
    addCols(sheet)
    # change column widths
    setColWidths(sheet)

    # fill in columns with data
    addPlayerInfo(sheet, mp, jf, f)

    wb.save("masterlist.xlsx")

def addPlayerInfo(sheet,mp, jf, f):
    players = {}
    for row in sheet:
        if row[0].value == "Last Name":
            continue
        playerName = row[1].value + " " + row[0].value
        # store mixed and doubles partner for each player
        players[playerName] = {"X": None, "D": None}

        # get list of events that player is playing
        events = row[11].value.split(", ")
        # get list of events and partners that the player signed up with
        entryInfo = row[12].value.split("\n")
        entryInfo.pop()

        for entry in entryInfo:
            # remove all withdrawn entries from entry info
            if "[Withdrawn]" in entry:
                entryInfo.remove(entry)
            else:
                # mixed event
                if entry[2] == "X":
                    checkPartners("X", entry, players, playerName, mp, f)
                # doubles event
                if entry[2] == "D":
                    checkPartners("D", entry, players, playerName, mp, f)

        writePartnersToMasterlist(players, row, playerName)

        playerEvents = {}
        playerFlights = getPlayerFlights(playerEvents, events)

        print(row[1].value, row[0].value)
        print("Entry Info:", entryInfo)
        # print("Player events:", playerEvents)
        
        # checks if players are jumping flights
        # players who jump flights can be found in masterErrors.txt
        checkJumpFlight(jf, playerFlights, row)

        # add event flights on excel sheet
        addFlightsToEventCols(playerEvents, row)
    print(players)

def writePartnersToMasterlist(players, row, playerName):
    # loop through all rows and fill in the respective column if they have a partner listed
    # does not account for partners that did not sign up with each other
    for partner in players[playerName]:
        if partner == "X" and players[playerName]["X"] != None:
            row[8].value = players[playerName]["X"]
        elif partner == "D" and players[playerName]["D"] != None:
            row[9].value = players[playerName]["D"]

def checkPartners(flight, entry, players, playerName, mp, f):
    # check if a player's partner already exists in player dictionary
    # if not, add the player's partner in their dictonary
    if players[playerName][flight] == None or players[playerName][flight] == "":
        players[playerName][flight] = entry[4:].split(" (")[0].title()
    # if a player's partner already exists for another flight,
    # make sure that the partner listed for the other flight is the same
    else:
        if players[playerName][flight] != entry[4:].split(" (")[0].title():
            mp.write(playerName + " has multiple partners\n")
    partnerName = players[playerName][flight]
    # check if the partner they signed up with also signed up with that person
    # find the "matching pair"
    if partnerName in players:
        if players[partnerName][flight] != playerName:
            f.write(entry[:3] + " " + playerName + " listed " + partnerName + " as their partner, but could not be found as " + partnerName + "'s partner \n")

def addFlightsToEventCols(playerEvents, row):
    for event in playerEvents:
        if event == "MS":
            row[3].value = playerEvents[event]
        elif event == "WS":
            row[4].value = playerEvents[event]
        elif event == "MX":
            row[5].value = playerEvents[event]
        elif event == "MD":
            row[6].value = playerEvents[event]
        elif event == "WD":
            row[7].value = playerEvents[event]

def checkJumpFlight(jf, playerFlights, row):
    if len(playerFlights) > 1:
        # print(ord(playerFlights[-1]) - ord(playerFlights[0]))
        if ord(playerFlights[-1]) - ord(playerFlights[0]) > 1:
            jf.write(row[1].value + " " + row[0].value + " is jumping flights\n")

def getPlayerFlights(playerEvents, events):
    playerFlights = []
    # check if player did not withdraw from all events
    if events != ['']:
        for event in events:
            if event[1:] in playerEvents:
                playerEvents[event[1:]] += event[0]
            else:
                playerEvents[event[1:]] = event[0]
            if event[0] not in playerFlights:
                playerFlights.append(event[0])
    return playerFlights

def addCols(sheet):
    # insert columns for events and flights: MS, WS, MX, MD, WD
    sheet.insert_cols(4, amount=7)
    
    colTitles = ["MS", "WS", "MX", "MD", "WD", "Mixed Partner", "Doubles Partner"]
    bottomBorder = Border(bottom=Side(style='thin'))
    col = 3
    for title in colTitles:
        sheet[1][col].value = title
        sheet[1][col].font = Font(bold = True)
        sheet[1][col].border = bottomBorder
        col += 1

def setColWidths(sheet):
    # change width of columns
    for i in range(ord('D'),ord('I')):
        sheet.column_dimensions[chr(i)].width = 5
    
    sheet.column_dimensions['I'].width = 15
    sheet.column_dimensions['J'].width = 15
    sheet.column_dimensions['K'].width = 20
    sheet.column_dimensions['L'].width = 20

def main():
    parse()

if __name__ == "__main__":
    main()