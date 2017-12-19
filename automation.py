#Author: David Gong
#Contributor: Jonathan Lian

import xlsxwriter

def groupsInfo():
    while True:
        try:
            numberOfGroups = int(input('Number of groups: '))
        except ValueError:
            print('Please input an integer.')
            continue
        if numberOfGroups > 4:
            print('There cannot be more than four groups. Try again.')
            continue
        else:
            break

    
    perGroupNumber = []
    groupNumber = 1

    for x in range(0, numberOfGroups):
        while True:
            try:
                numberOfPeople = int(input('How many people were in Group ' + str(x + 1) + '? '))
            except ValueError:
                print('Please input an integer.')
                continue
            if numberOfPeople < 4:
                print('There has to be at least four people in a group. Try again.')
                continue
            if numberOfPeople > 7:
                print('There cannot be more than seven people in a group. Try again.')
            else:
                break
        perGroupNumber.append(numberOfPeople)

    return perGroupNumber



def groupOneInfo(x):
    print("\n\nPlease refer to the current league ratings as you fill this part out.")
    print("\nYou are currently putting in information for Group 1.")
    playerInfo = {}
    
    for i in range(0, x[0]):
        playerName = input("Name of person in Group 1: ")
        playerRating = int(input("Rating of person in Group 1: "))
        playerInfo[playerName] = playerRating
        print('\n' + str(playerName) + "(" + str(playerRating) + ") has been added to Group 1.")
        print('\n')

    return playerInfo



def groupTwoInfo(x):
    playerInfo = {}
    if (len(x) < 2):
        return None
    else:
        print("\nYou are currently putting in information for Group 2.")
        for i in range(0, x[1]):
            playerName = input("Name of person in Group 2: ")
            playerRating = int(input("Rating of person in Group 2: "))
            playerInfo[playerName] = playerRating
            print('\n' + str(playerName) + "(" + str(playerRating) + ") has been added to Group 2.")
            print('\n')
        return playerInfo


        
def groupThreeInfo(x):
    playerInfo = {}
    if (len(x) < 3):
        return None
    else:
        print("\nYou are currently putting in informatin for Group 3.")
        for i in range(0, x[2]):
            playerName = input("Name of person in Group 3: ")
            playerRating = int(input("Rating of person in Group 3: "))
            playerInfo[playerName] = playerRating
            print('\n' + str(playerName) + "(" + str(playerRating) + ") has been added to Group 3.")
            print('\n')
        return playerInfo



def groupFourInfo(x):
    playerInfo = {}
    if (len(x) < 4):
        return None
    else:
        print("\nYou are currently putting in informatin for Group 4.")
        for i in range(0, x[3]):
            playerName = input("Name of person in Group 4: ")
            playerRating = int(input("Rating of person in Group 4: "))
            playerInfo[playerName] = playerRating
            print('\n' + str(playerName) + "(" + str(playerRating) + ") has been added to Group 4.")
            print('\n')
        return playerInfo

def swapper(group):
    for i in range(0, len(group)):
        if (i == 1):
            if (group[i] < group[i + 2]):
                swapRating = group[i]
                swapName = group[i - 1]

                group[i - 1] = group[i + 1]
                group[i] = group[i + 2]

                group[i + 1] = swapName
                group[i + 2] = swapRating
        elif (i == len(group) - 2):
            break
        elif (i > 1 and i % 2 == 1):
            if (group[i] < group[i + 2]):
                swapRating = group[i]
                swapName = group[i - 1]

                group[i - 1] = group[i + 1]
                group[i] = group[i + 2]

                group[i + 1] = swapName
                group[i + 2] = swapRating
                if (group[i - 2] < group[i]):
                    swapRating2 = group[i - 2]
                    swapName2 = group[i - 3]

                    group[i - 3] = group[i - 1]
                    group[i - 2] = group[i]

                    group[i - 1] = swapName2
                    group[i] = swapRating2
        else:
            continue

def swapChecker(group):
    if (group[1] < group[3]):
        swapRating = group[1]
        swapName = group[0]

        group[0] = group[2]
        group[1] = group[3]

        group[2] = swapName
        group[3] = swapRating


def ratingCalc(higherRating, lowerRating, higherRatingWins):
    difference = higherRating - lowerRating
    pointChange = 0
    if higherRatingWins == True:
        if difference < 13:
            pointChange = 8
            higherRating += 8
            lowerRating -= 8
        if (difference >= 13 and difference < 38):
            pointChange = 7
            higherRating += 7
            lowerRating -= 7
        if (difference >= 38 and difference < 63):
            pointChange = 6
            higherRating += 6
            lowerRating -= 6
        if (difference >= 63 and difference < 88):
            pointChange = 5
            higherRating += 5
            lowerRating -= 5
        if (difference >= 88 and difference < 113):
            pointChange = 4
            higherRating += 4
            lowerRating -= 4
        if (difference >= 113 and difference < 138):
            pointChange = 3
            higherRating += 3
            lowerRating -= 3
        if (difference >= 138 and difference < 163):
            pointChange = 2
            higherRating += 2
            lowerRating -= 2
        if (difference >= 163 and difference < 188):
            pointChange = 2
            higherRating += 2
            lowerRating -= 2
        if (difference >= 188 and difference < 213):
            pointChange = 1
            higherRating += 1
            lowerRating -= 1
        if (difference >= 213 and difference < 238):
            pointChange = 1
            higherRating += 1
            lowerRating -= 1
        if (difference >= 238):
            pointChange = 0
            higherRating += 0
            lowerRating -= 0
    if higherRatingWins == False:
        if (difference < 13):
            pointChange = -8
            higherRating -= 8
            lowerRating += 8
        if (difference >= 13 and difference < 38):
            pointChange = -10
            higherRating -= 10
            lowerRating += 10
        if (difference >= 38 and difference < 63):
            pointChange = -13
            higherRating -= 13
            lowerRating += 13
        if (difference >= 63 and difference < 88):
            pointChange = -16
            higherRating -= 16
            lowerRating += 16
        if (difference >= 88 and difference < 113):
            pointChange = -20
            higherRating -= 20
            lowerRating += 20
        if (difference >= 113 and difference < 138):
            pointChange = -25
            higherRating -= 25
            lowerRating += 25
        if (difference >= 138 and difference < 163):
            pointChange = -30
            higherRating -= 30
            lowerRating += 30
        if (difference >= 163 and difference < 188):
            pointChange = -35
            higherRating -= 35
            lowerRating += 35
        if (difference >= 188 and difference < 213):
            pointChange = -40
            higherRating -= 40
            lowerRating += 40
        if (difference >= 213 and difference < 238):
            pointChange = -45
            higherRating -= 45
            lowerRating += 45
        if (difference >= 238):
            pointChange = -50
            higherRating -= 50
            lowerRating += 50

    return pointChange

def higherRatingWins(gameScore):
    if int(gameScore) < 3:
        return False
    else:
        return True

def headerWriter(sheet):
    sheet.write('A2', 'Match')
    sheet.write('B2', 'Players')
    sheet.write('C2', 'Score')
    sheet.write('D2', 'Rating Before')
    sheet.write('E2', 'Point Change')

def fourPersonResultFormat(sheet, groupPlayers):
    
    sheet.add_table('A2:E20',{'autofilter': False})
    sheet.set_column(1, 1, 15)
    sheet.set_column(3, 3, len('Rating Before') - 1)
    sheet.set_column(4, 4, len('Point Change'))

    merge_format1 = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color' : '#C0C0C0'})
    merge_format1.set_font_size(15)
    

    merge_format2 = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color' : '#C0C0C0'})

    sheetMerger(groupPlayers, sheet, merge_format1, merge_format2)
    headerWriter(sheet)

    sheet.write('A4', 'B')
    sheet.write('A5', 'D')
    sheet.write('A7', 'A')
    sheet.write('A8', 'C')
    sheet.write('A10', 'B')
    sheet.write('A11', 'C')
    sheet.write('A13', 'A')
    sheet.write('A14', 'D')
    sheet.write('A16', 'C')
    sheet.write('A17', 'D')
    sheet.write('A19', 'A')
    sheet.write('A20', 'B')

    playerA = groupPlayers[0]
    playerB = groupPlayers[2]
    playerC = groupPlayers[4]
    playerD = groupPlayers[6]
    playerARating = groupPlayers[1]
    playerBRating = groupPlayers[3]
    playerCRating = groupPlayers[5]
    playerDRating = groupPlayers[7]
    
    sheet.write('B4', playerB)
    sheet.write('B5', playerD)
    sheet.write('B7', playerA)
    sheet.write('B8', playerC)
    sheet.write('B10', playerB)
    sheet.write('B11', playerC)
    sheet.write('B13', playerA)
    sheet.write('B14', playerD)
    sheet.write('B16', playerC)
    sheet.write('B17', playerD)
    sheet.write('B19', playerA)
    sheet.write('B20', playerB)

    print('\nPlease input the game scores for the matches.')
    print('\nFor example, assuming B won 3 - 2 for Match B vs D, (3:2) should be inputted, excluding parentheses and spaces.')
    print('In the case of B losing to D 2-3, input (2:3) excluding the parentheses and spaces.\n')
    matchResultList = []
    match1 = input('B vs D: ')
    match2 = input('A vs C: ')
    match3 = input('B vs C: ')
    match4 = input('A vs D: ')
    match5 = input('C vs D: ')
    match6 = input('A vs B: ')
    matchResultList.extend((match1, match2, match3, match4, match5, match6))
    resultWriter(matchResultList, sheet)

    sheet.write('D4', playerBRating)
    sheet.write('D5', playerDRating)
    sheet.write('D7', playerARating)
    sheet.write('D8', playerCRating)
    sheet.write('D10', playerBRating)
    sheet.write('D11', playerCRating)
    sheet.write('D13', playerARating)
    sheet.write('D14', playerDRating)
    sheet.write('D16', playerCRating)
    sheet.write('D17', playerDRating)
    sheet.write('D19', playerARating)
    sheet.write('D20', playerBRating)

    pointChangeList = []
    pointChange1 = ratingCalc(playerBRating, playerDRating, higherRatingWins(match1[0]))
    pointChange2 = ratingCalc(playerARating, playerCRating, higherRatingWins(match2[0]))
    pointChange3 = ratingCalc(playerBRating, playerCRating, higherRatingWins(match3[0]))
    pointChange4 = ratingCalc(playerARating, playerDRating, higherRatingWins(match4[0]))
    pointChange5 = ratingCalc(playerCRating, playerDRating, higherRatingWins(match5[0]))
    pointChange6 = ratingCalc(playerARating, playerBRating, higherRatingWins(match6[0]))
    pointChangeList.extend((pointChange1, pointChange2, pointChange3,
                            pointChange4, pointChange5, pointChange6))
    pointChangeWriter(pointChangeList, sheet)

    playersTotalChange = []
    playerAChange = pointChange2 + pointChange4 + pointChange6
    playerBChange = pointChange1 + pointChange3 + -(pointChange6)
    playerCChange = -(pointChange2) + -(pointChange3) + pointChange5
    playerDChange = -(pointChange1) + -(pointChange4) + -(pointChange5)
    playersTotalChange.extend((playerAChange, playerBChange, playerCChange, playerDChange))

    return playersTotalChange

def fivePersonResultFormat(sheet, groupPlayers):
    
    sheet.add_table('A2:E33',{'autofilter': False})
    sheet.set_column(1, 1, 15)
    sheet.set_column(3, 3, len('Rating Before') - 1)
    sheet.set_column(4, 4, len('Point Change'))

    merge_format1 = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color' : '#C0C0C0'})
    merge_format1.set_font_size(15)
    

    merge_format2 = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color' : '#C0C0C0'})

    sheetMerger(groupPlayers, sheet, merge_format1, merge_format2)

    headerWriter(sheet)

    sheet.write('A4', 'A')
    sheet.write('A5', 'D')
    sheet.write('A7', 'B')
    sheet.write('A8', 'C')
    sheet.write('A10', 'B')
    sheet.write('A11', 'E')
    sheet.write('A13', 'C')
    sheet.write('A14', 'D')
    sheet.write('A16', 'A')
    sheet.write('A17', 'E')
    sheet.write('A19', 'B')
    sheet.write('A20', 'D')
    sheet.write('A22', 'A')
    sheet.write('A23', 'C')
    sheet.write('A25', 'D')
    sheet.write('A26', 'E')
    sheet.write('A28', 'C')
    sheet.write('A29', 'E')
    sheet.write('A31', 'A')
    sheet.write('A32', 'B')

    playerA = groupPlayers[0]
    playerB = groupPlayers[2]
    playerC = groupPlayers[4]
    playerD = groupPlayers[6]
    playerE = groupPlayers[8]
    playerARating = groupPlayers[1]
    playerBRating = groupPlayers[3]
    playerCRating = groupPlayers[5]
    playerDRating = groupPlayers[7]
    playerERating = groupPlayers[9]
    
    sheet.write('B4', playerA)
    sheet.write('B5', playerD)
    sheet.write('B7', playerB)
    sheet.write('B8', playerC)
    sheet.write('B10', playerB)
    sheet.write('B11', playerE)
    sheet.write('B13', playerC)
    sheet.write('B14', playerD)
    sheet.write('B16', playerA)
    sheet.write('B17', playerE)
    sheet.write('B19', playerB)
    sheet.write('B20', playerD)
    sheet.write('B22', playerA)
    sheet.write('B23', playerC)
    sheet.write('B25', playerD)
    sheet.write('B26', playerE)
    sheet.write('B28', playerC)
    sheet.write('B29', playerE)
    sheet.write('B31', playerA)
    sheet.write('B32', playerB)

    sheet.write('D4', playerARating)
    sheet.write('D5', playerDRating)
    sheet.write('D7', playerBRating)
    sheet.write('D8', playerCRating)
    sheet.write('D10', playerBRating)
    sheet.write('D11', playerERating)
    sheet.write('D13', playerCRating)
    sheet.write('D14', playerDRating)
    sheet.write('D16', playerARating)
    sheet.write('D17', playerERating)
    sheet.write('D19', playerBRating)
    sheet.write('D20', playerDRating)
    sheet.write('D22', playerARating)
    sheet.write('D23', playerCRating)
    sheet.write('D25', playerDRating)
    sheet.write('D26', playerERating)
    sheet.write('D28', playerCRating)
    sheet.write('D29', playerERating)
    sheet.write('D31', playerARating)
    sheet.write('D32', playerBRating)

    print('\nPlease input the game scores for the matches.')
    print('\nFor example, assuming B won 3 - 2 for Match B vs D, 3:2 or 3-2 should be inputted.')
    print('In the case of B losing to D 2 - 3, input 2:3 or 2-3.\n')
    matchResultList = []
    match1 = input('A vs D: ')
    match2 = input('B vs C: ')
    match3 = input('B vs E: ')
    match4 = input('C vs D: ')
    match5 = input('A vs E: ')
    match6 = input('B vs D: ')
    match7 = input('A vs C: ')
    match8 = input('D vs E: ')
    match9 = input('C vs E: ')
    match10 = input('A vs B: ')
    matchResultList.extend((match1, match2, match3, match4, match5,
                            match6, match7, match8, match9, match10))
    resultWriter(matchResultList, sheet)


    pointChangeList = []
    pointChange1 = ratingCalc(playerARating, playerDRating, higherRatingWins(match1[0]))
    pointChange2 = ratingCalc(playerBRating, playerCRating, higherRatingWins(match2[0]))
    pointChange3 = ratingCalc(playerBRating, playerERating, higherRatingWins(match3[0]))
    pointChange4 = ratingCalc(playerCRating, playerDRating, higherRatingWins(match4[0]))
    pointChange5 = ratingCalc(playerARating, playerERating, higherRatingWins(match5[0]))
    pointChange6 = ratingCalc(playerBRating, playerDRating, higherRatingWins(match6[0]))
    pointChange7 = ratingCalc(playerARating, playerCRating, higherRatingWins(match7[0]))
    pointChange8 = ratingCalc(playerDRating, playerERating, higherRatingWins(match8[0]))
    pointChange9 = ratingCalc(playerCRating, playerERating, higherRatingWins(match9[0]))
    pointChange10 = ratingCalc(playerARating, playerBRating, higherRatingWins(match10[0]))
    pointChangeList.extend((pointChange1, pointChange2, pointChange3, pointChange4, pointChange5,
                        pointChange6, pointChange7, pointChange8, pointChange9, pointChange10))
    pointChangeWriter(pointChangeList, sheet)


    playersTotalChange = []
    playerAChange = pointChange1 + pointChange5 + pointChange7 + pointChange10
    playerBChange = pointChange2 + pointChange3 + pointChange6 + -(pointChange10)
    playerCChange = -(pointChange2) + pointChange4 + -(pointChange7) + pointChange9
    playerDChange = -(pointChange1) + -(pointChange4) + -(pointChange6) + pointChange8
    playerEChange = -(pointChange3) + -(pointChange5) + -(pointChange8) + -(pointChange9)
    playersTotalChange.extend((playerAChange, playerBChange, playerCChange, playerDChange, playerEChange))

    return playersTotalChange

def sixPersonResultFormat(sheet, groupPlayers):
    
    sheet.add_table('A2:E48',{'autofilter': False})
    sheet.set_column(1, 1, 15)
    sheet.set_column(3, 3, len('Rating Before') - 1)
    sheet.set_column(4, 4, len('Point Change'))

    merge_format1 = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color' : '#C0C0C0'})
    merge_format1.set_font_size(15)
    

    merge_format2 = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color' : '#C0C0C0'})

    sheetMerger(groupPlayers, sheet, merge_format1, merge_format2)
    headerWriter(sheet)

    sheet.write('A4', 'A')
    sheet.write('A5', 'D')
    sheet.write('A7', 'B')
    sheet.write('A8', 'C')
    sheet.write('A10', 'E')
    sheet.write('A11', 'F')
    sheet.write('A13', 'A')
    sheet.write('A14', 'E')
    sheet.write('A16', 'B')
    sheet.write('A17', 'D')
    sheet.write('A19', 'C')
    sheet.write('A20', 'F')
    sheet.write('A22', 'B')
    sheet.write('A23', 'F')
    sheet.write('A25', 'D')
    sheet.write('A26', 'E')
    sheet.write('A28', 'A')
    sheet.write('A29', 'C')
    sheet.write('A31', 'A')
    sheet.write('A32', 'F')
    sheet.write('A34', 'B')
    sheet.write('A35', 'E')
    sheet.write('A37', 'C')
    sheet.write('A38', 'D')
    sheet.write('A40', 'C')
    sheet.write('A41', 'E')
    sheet.write('A43', 'D')
    sheet.write('A44', 'F')
    sheet.write('A46', 'A')
    sheet.write('A47', 'B')

    playerA = groupPlayers[0]
    playerB = groupPlayers[2]
    playerC = groupPlayers[4]
    playerD = groupPlayers[6]
    playerE = groupPlayers[8]
    playerF = groupPlayers[10]
    playerARating = groupPlayers[1]
    playerBRating = groupPlayers[3]
    playerCRating = groupPlayers[5]
    playerDRating = groupPlayers[7]
    playerERating = groupPlayers[9]
    playerFRating = groupPlayers[11]
    
    sheet.write('B4', playerA)
    sheet.write('B5', playerD)
    sheet.write('B7', playerB)
    sheet.write('B8', playerC)
    sheet.write('B10', playerE)
    sheet.write('B11', playerF)
    sheet.write('B13', playerA)
    sheet.write('B14', playerE)
    sheet.write('B16', playerB)
    sheet.write('B17', playerD)
    sheet.write('B19', playerC)
    sheet.write('B20', playerF)
    sheet.write('B22', playerB)
    sheet.write('B23', playerF)
    sheet.write('B25', playerD)
    sheet.write('B26', playerE)
    sheet.write('B28', playerA)
    sheet.write('B29', playerC)
    sheet.write('B31', playerA)
    sheet.write('B32', playerF)
    sheet.write('B34', playerB)
    sheet.write('B35', playerE)
    sheet.write('B37', playerC)
    sheet.write('B38', playerD)
    sheet.write('B40', playerC)
    sheet.write('B41', playerE)
    sheet.write('B43', playerD)
    sheet.write('B44', playerF)
    sheet.write('B46', playerA)
    sheet.write('B47', playerB)

    sheet.write('D4', playerARating)
    sheet.write('D5', playerDRating)
    sheet.write('D7', playerBRating)
    sheet.write('D8', playerCRating)
    sheet.write('D10', playerERating)
    sheet.write('D11', playerFRating)
    sheet.write('D13', playerARating)
    sheet.write('D14', playerERating)
    sheet.write('D16', playerBRating)
    sheet.write('D17', playerDRating)
    sheet.write('D19', playerCRating)
    sheet.write('D20', playerFRating)
    sheet.write('D22', playerBRating)
    sheet.write('D23', playerFRating)
    sheet.write('D25', playerDRating)
    sheet.write('D26', playerERating)
    sheet.write('D28', playerARating)
    sheet.write('D29', playerCRating)
    sheet.write('D31', playerARating)
    sheet.write('D32', playerFRating)
    sheet.write('D34', playerBRating)
    sheet.write('D35', playerERating)
    sheet.write('D37', playerCRating)
    sheet.write('D38', playerDRating)
    sheet.write('D40', playerCRating)
    sheet.write('D41', playerERating)
    sheet.write('D43', playerDRating)
    sheet.write('D44', playerFRating)
    sheet.write('D46', playerARating)
    sheet.write('D47', playerBRating)
    

    print('\nPlease input the game scores for the matches.')
    print('\nFor example, assuming B won 3 - 2 for Match B vs D, (3:2) should be inputted, excluding parentheses and spaces.')
    print('In the case of B losing to D 2-3, input (2:3) excluding the parentheses and spaces.\n')
    matchResultList = []
    match1 = input('A vs D: ')
    match2 = input('B vs C: ')
    match3 = input('E vs F: ')
    match4 = input('A vs E: ')
    match5 = input('B vs D: ')
    match6 = input('C vs F: ')
    match7 = input('B vs F: ')
    match8 = input('D vs E: ')
    match9 = input('A vs C: ')
    match10 = input('A vs F: ')
    match11 = input('B vs E: ')
    match12 = input('C vs D: ')
    match13 = input('C vs E: ')
    match14 = input('D vs F: ')
    match15 = input('A vs B: ')
    matchResultList.extend((match1, match2, match3, match4, match5,
                            match6, match7, match8, match9, match10,
                            match11, match12, match13, match14, match15))
    resultWriter(matchResultList, sheet)


    pointChangeList = []
    pointChange1 = ratingCalc(playerARating, playerDRating, higherRatingWins(match1[0]))
    pointChange2 = ratingCalc(playerBRating, playerCRating, higherRatingWins(match2[0]))
    pointChange3 = ratingCalc(playerERating, playerFRating, higherRatingWins(match3[0]))
    pointChange4 = ratingCalc(playerARating, playerERating, higherRatingWins(match4[0]))
    pointChange5 = ratingCalc(playerBRating, playerDRating, higherRatingWins(match5[0]))
    pointChange6 = ratingCalc(playerCRating, playerFRating, higherRatingWins(match6[0]))
    pointChange7 = ratingCalc(playerBRating, playerERating, higherRatingWins(match7[0]))
    pointChange8 = ratingCalc(playerDRating, playerERating, higherRatingWins(match8[0]))
    pointChange9 = ratingCalc(playerARating, playerCRating, higherRatingWins(match9[0]))
    pointChange10 = ratingCalc(playerARating, playerFRating, higherRatingWins(match10[0]))
    pointChange11 = ratingCalc(playerBRating, playerERating, higherRatingWins(match11[0]))
    pointChange12 = ratingCalc(playerCRating, playerDRating, higherRatingWins(match12[0]))
    pointChange13 = ratingCalc(playerCRating, playerERating, higherRatingWins(match13[0]))
    pointChange14 = ratingCalc(playerDRating, playerFRating, higherRatingWins(match14[0]))
    pointChange15 = ratingCalc(playerARating, playerBRating, higherRatingWins(match15[0]))
    pointChangeList.extend((pointChange1, pointChange2, pointChange3, pointChange4, pointChange5,
                        pointChange6, pointChange7, pointChange8, pointChange9, pointChange10,
                        pointChange11, pointChange12, pointChange13, pointChange14, pointChange15,))
    pointChangeWriter(pointChangeList, sheet)


    playersTotalChange = []
    playerAChange = pointChange1 + pointChange4 + pointChange9 + pointChange10 + pointChange15
    playerBChange = pointChange2 + pointChange5 + pointChange7 + pointChange11 + -(pointChange15)
    playerCChange = -(pointChange2) + pointChange6 + -(pointChange9) + pointChange12 + pointChange13
    playerDChange = -(pointChange1) + -(pointChange5) + pointChange8 + -(pointChange12) + pointChange14
    playerEChange = pointChange3 + -(pointChange4) + -(pointChange8) + -(pointChange11) + -(pointChange13)
    playerFChange = -(pointChange3) + -(pointChange6) + -(pointChange7) + -(pointChange10) + -(pointChange14)
    playersTotalChange.extend((playerAChange, playerBChange, playerCChange, playerDChange, playerEChange, playerFChange))

    return playersTotalChange

def sevenPersonResultFormat(sheet, groupPlayers):
    
    sheet.add_table('A2:E66',{'autofilter': False})
    sheet.set_column(1, 1, 15)
    sheet.set_column(3, 3, len('Rating Before') - 1)
    sheet.set_column(4, 4, len('Point Change'))

    merge_format1 = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color' : '#C0C0C0'})
    merge_format1.set_font_size(15)
    

    merge_format2 = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color' : '#C0C0C0'})

    sheetMerger(groupPlayers, sheet, merge_format1, merge_format2)
    headerWriter(sheet)

    sheet.write('A4', 'A')
    sheet.write('A5', 'F')
    sheet.write('A7', 'B')
    sheet.write('A8', 'E')
    sheet.write('A10', 'C')
    sheet.write('A11', 'D')
    sheet.write('A13', 'B')
    sheet.write('A14', 'G')
    sheet.write('A16', 'C')
    sheet.write('A17', 'F')
    sheet.write('A19', 'D')
    sheet.write('A20', 'E')
    sheet.write('A22', 'A')
    sheet.write('A23', 'E')
    sheet.write('A25', 'B')
    sheet.write('A26', 'D')
    sheet.write('A28', 'C')
    sheet.write('A29', 'G')
    sheet.write('A31', 'A')
    sheet.write('A32', 'C')
    sheet.write('A34', 'D')
    sheet.write('A35', 'F')
    sheet.write('A37', 'E')
    sheet.write('A38', 'G')
    sheet.write('A40', 'F')
    sheet.write('A41', 'G')
    sheet.write('A43', 'A')
    sheet.write('A44', 'D')
    sheet.write('A46', 'B')
    sheet.write('A47', 'C')
    sheet.write('A49', 'A')
    sheet.write('A50', 'B')
    sheet.write('A52', 'E')
    sheet.write('A53', 'F')
    sheet.write('A55', 'D')
    sheet.write('A56', 'G')
    sheet.write('A58', 'A')
    sheet.write('A59', 'G')
    sheet.write('A61', 'B')
    sheet.write('A62', 'F')
    sheet.write('A64', 'C')
    sheet.write('A65', 'E')

    playerA = groupPlayers[0]
    playerB = groupPlayers[2]
    playerC = groupPlayers[4]
    playerD = groupPlayers[6]
    playerE = groupPlayers[8]
    playerF = groupPlayers[10]
    playerG = groupPlayers[12]
    playerARating = groupPlayers[1]
    playerBRating = groupPlayers[3]
    playerCRating = groupPlayers[5]
    playerDRating = groupPlayers[7]
    playerERating = groupPlayers[9]
    playerFRating = groupPlayers[11]
    playerGRating = groupPlayers[13]
    
    sheet.write('B4', playerA)
    sheet.write('B5', playerF)
    sheet.write('B7', playerB)
    sheet.write('B8', playerE)
    sheet.write('B10', playerC)
    sheet.write('B11', playerD)
    sheet.write('B13', playerB)
    sheet.write('B14', playerG)
    sheet.write('B16', playerC)
    sheet.write('B17', playerF)
    sheet.write('B19', playerD)
    sheet.write('B20', playerE)
    sheet.write('B22', playerA)
    sheet.write('B23', playerE)
    sheet.write('B25', playerB)
    sheet.write('B26', playerD)
    sheet.write('B28', playerC)
    sheet.write('B29', playerG)
    sheet.write('B31', playerA)
    sheet.write('B32', playerC)
    sheet.write('B34', playerD)
    sheet.write('B35', playerF)
    sheet.write('B37', playerE)
    sheet.write('B38', playerG)
    sheet.write('B40', playerF)
    sheet.write('B41', playerG)
    sheet.write('B43', playerA)
    sheet.write('B44', playerD)
    sheet.write('B46', playerB)
    sheet.write('B47', playerC)
    sheet.write('B49', playerA)
    sheet.write('B50', playerB)
    sheet.write('B52', playerE)
    sheet.write('B53', playerF)
    sheet.write('B55', playerD)
    sheet.write('B56', playerG)
    sheet.write('B58', playerA)
    sheet.write('B59', playerG)
    sheet.write('B61', playerB)
    sheet.write('B62', playerF)
    sheet.write('B64', playerC)
    sheet.write('B65', playerE)

    sheet.write('D4', playerARating)
    sheet.write('D5', playerFRating)
    sheet.write('D7', playerBRating)
    sheet.write('D8', playerERating)
    sheet.write('D10', playerCRating)
    sheet.write('D11', playerDRating)
    sheet.write('D13', playerBRating)
    sheet.write('D14', playerGRating)
    sheet.write('D16', playerCRating)
    sheet.write('D17', playerFRating)
    sheet.write('D19', playerDRating)
    sheet.write('D20', playerERating)
    sheet.write('D22', playerARating)
    sheet.write('D23', playerERating)
    sheet.write('D25', playerBRating)
    sheet.write('D26', playerDRating)
    sheet.write('D28', playerCRating)
    sheet.write('D29', playerGRating)
    sheet.write('D31', playerARating)
    sheet.write('D32', playerCRating)
    sheet.write('D34', playerDRating)
    sheet.write('D35', playerFRating)
    sheet.write('D37', playerERating)
    sheet.write('D38', playerGRating)
    sheet.write('D40', playerFRating)
    sheet.write('D41', playerGRating)
    sheet.write('D43', playerARating)
    sheet.write('D44', playerDRating)
    sheet.write('D46', playerBRating)
    sheet.write('D47', playerCRating)
    sheet.write('D49', playerARating)
    sheet.write('D50', playerBRating)
    sheet.write('D52', playerERating)
    sheet.write('D53', playerFRating)
    sheet.write('D55', playerDRating)
    sheet.write('D56', playerGRating)
    sheet.write('D58', playerARating)
    sheet.write('D59', playerGRating)
    sheet.write('D61', playerBRating)
    sheet.write('D62', playerFRating)
    sheet.write('D64', playerCRating)
    sheet.write('D65', playerERating)
    

    print('\nPlease input the game scores for the matches.')
    print('\nFor example, assuming B won 3 - 2 for Match B vs D, (3:2) should be inputted, excluding parentheses and spaces.')
    print('In the case of B losing to D 2-3, input (2:3) excluding the parentheses and spaces.\n')
    matchResultList = []
    match1 = input('A vs F: ')
    match2 = input('B vs E: ')
    match3 = input('C vs D: ')
    match4 = input('B vs G: ')
    match5 = input('C vs F: ')
    match6 = input('D vs E: ')
    match7 = input('A vs E: ')
    match8 = input('B vs D: ')
    match9 = input('C vs G: ')
    match10 = input('A vs C: ')
    match11 = input('D vs F: ')
    match12 = input('E vs G: ')
    match13 = input('F vs G: ')
    match14 = input('A vs D: ')
    match15 = input('B vs C: ')
    match16 = input('A vs B: ')
    match17 = input('E vs F: ')
    match18 = input('D vs G: ')
    match19 = input('A vs G: ')
    match20 = input('B vs F: ')
    match21 = input('C vs E: ')
    matchResultList.extend((match1, match2, match3, match4, match5,
                            match6, match7, match8, match9, match10,
                            match11, match12, match13, match14, match15,
                            match16, match17, match18, match19, match20, match21))
    resultWriter(matchResultList, sheet)


    pointChangeList = []
    pointChange1 = ratingCalc(playerARating, playerFRating, higherRatingWins(match1[0]))
    pointChange2 = ratingCalc(playerBRating, playerERating, higherRatingWins(match2[0]))
    pointChange3 = ratingCalc(playerCRating, playerDRating, higherRatingWins(match3[0]))
    pointChange4 = ratingCalc(playerBRating, playerGRating, higherRatingWins(match4[0]))
    pointChange5 = ratingCalc(playerCRating, playerFRating, higherRatingWins(match5[0]))
    pointChange6 = ratingCalc(playerDRating, playerERating, higherRatingWins(match6[0]))
    pointChange7 = ratingCalc(playerARating, playerERating, higherRatingWins(match7[0]))
    pointChange8 = ratingCalc(playerBRating, playerDRating, higherRatingWins(match8[0]))
    pointChange9 = ratingCalc(playerCRating, playerGRating, higherRatingWins(match9[0]))
    pointChange10 = ratingCalc(playerARating, playerCRating, higherRatingWins(match10[0]))
    pointChange11 = ratingCalc(playerDRating, playerFRating, higherRatingWins(match11[0]))
    pointChange12 = ratingCalc(playerERating, playerGRating, higherRatingWins(match12[0]))
    pointChange13 = ratingCalc(playerFRating, playerGRating, higherRatingWins(match13[0]))
    pointChange14 = ratingCalc(playerARating, playerDRating, higherRatingWins(match14[0]))
    pointChange15 = ratingCalc(playerBRating, playerCRating, higherRatingWins(match15[0]))
    pointChange16 = ratingCalc(playerARating, playerBRating, higherRatingWins(match16[0]))
    pointChange17 = ratingCalc(playerERating, playerFRating, higherRatingWins(match17[0]))
    pointChange18 = ratingCalc(playerDRating, playerGRating, higherRatingWins(match18[0]))
    pointChange19 = ratingCalc(playerARating, playerGRating, higherRatingWins(match19[0]))
    pointChange20 = ratingCalc(playerBRating, playerFRating, higherRatingWins(match20[0]))
    pointChange21 = ratingCalc(playerCRating, playerERating, higherRatingWins(match21[0]))
    pointChangeList.extend((pointChange1, pointChange2, pointChange3, pointChange4, pointChange5,
                        pointChange6, pointChange7, pointChange8, pointChange9, pointChange10,
                        pointChange11, pointChange12, pointChange13, pointChange14, pointChange15,
                        pointChange16, pointChange17, pointChange18, pointChange19, pointChange20,
                        pointChange21))
    pointChangeWriter(pointChangeList, sheet)


    playersTotalChange = []
    playerAChange = pointChange1 + pointChange7 + pointChange10 + pointChange14 + pointChange16 + pointChange19
    playerBChange = pointChange2 + pointChange4 + pointChange8 + pointChange15 + -(pointChange16) + pointChange20
    playerCChange = pointChange3 + pointChange5 + pointChange9 + -(pointChange10) + -(pointChange15) + pointChange21
    playerDChange = -(pointChange2) + -(pointChange6) + -(pointChange7) + pointChange12 + pointChange17 + -(pointChange21)
    playerEChange = -(pointChange1) + -(pointChange5) + -(pointChange11) + pointChange13 + -(pointChange17) + -(pointChange20)
    playerFChange = -(pointChange4) + -(pointChange9) + -(pointChange12) + -(pointChange13) + -(pointChange18) + -(pointChange19)
    playersTotalChange.extend((playerAChange, playerBChange, playerCChange, playerDChange, playerEChange, playerFChange))

    return playersTotalChange

def resultWriter(matchResultList, sheet):
    counter = 4
    for i in range(0, len(matchResultList)):
        sheet.write('C' + str(counter), int(matchResultList[i][0]))
        sheet.write('C' + str(counter + 1), int(matchResultList[i][2]))
        counter += 3

def pointChangeWriter(pointChangeList, sheet):
    counter = 4
    for i in range(0, len(pointChangeList)):
        sheet.write('E' + str(counter), pointChangeList[i])
        sheet.write('E' + str(counter + 1), -(pointChangeList[i]))
        counter += 3

def sheetMerger(groupPlayers, sheet, merge_format1, merge_format2):
    sheet.merge_range('A1:E1', 'Group {Replace This} - Match Record', merge_format1)
    amountOfMerges = 6
    merge = 3
    amountOfPlayers = len(groupPlayers) / 2
    if (amountOfPlayers != 4):
        start = 4
        while (amountOfPlayers > start):
            amountOfMerges += 5
            start += 1
    for _ in range(amountOfMerges):
        sheet.merge_range('A' + str(merge) + ':E' + str(merge), ' ', merge_format2)
        merge += 3

def tableWriter(worksheet, groupSize, groupPlayers, overallRatingChange, seedLetterList, spacer):
    seedCounter = 0
    groupNameCounter = 0
    groupRatingCounter = 1
    ratingChangeCounter = 0

    regular_fill = workbook.add_format()
    regular_fill.set_pattern(1)
    regular_fill.set_bg_color('white')

    for i in range(2, groupSize + 2):
        ratingAfter = groupPlayers[groupRatingCounter] + overallRatingChange[ratingChangeCounter]
        worksheet.write(i + spacer, 1, seedLetterList[seedCounter], regular_fill)
        worksheet.write(i + spacer, 2, groupPlayers[groupNameCounter], regular_fill)
        worksheet.write(i + spacer, 3, groupPlayers[groupRatingCounter], regular_fill)
        worksheet.write(i + spacer, 4, ' ', regular_fill)
        worksheet.write(i + spacer, 5, overallRatingChange[ratingChangeCounter], regular_fill)
        worksheet.write(i + spacer, 6, ratingAfter, regular_fill)

        seedCounter += 1
        groupNameCounter += 2
        groupRatingCounter += 2
        ratingChangeCounter += 1


def tableMaker(worksheet, spacer, spacerGroupSize, groupNumber):
    group_title = workbook.add_format({
        'border' : 1,
        'align' : 'center',
        'valign' : 'vcenter',
        'fg_color': '#99CCFF'})
    header_fill = workbook.add_format()
    header_fill.set_pattern(1)
    header_fill.set_bg_color('gray')
    
    group_title.set_font_size(17)
    worksheet.add_table(spacer, 1, spacerGroupSize, 6, {'style' : 'Table Style Light 9'})
    worksheet.merge_range(spacer - 1, 1, spacer - 1, 6, 'Group ' + str(groupNumber), group_title)
    worksheet.write(spacer, 1, 'Seed', header_fill)
    worksheet.write(spacer, 2, 'Player', header_fill)
    worksheet.write(spacer, 3, 'Rating Before', header_fill)
    worksheet.write(spacer, 4, 'Matches Won', header_fill)
    worksheet.write(spacer, 5, 'Rating Change', header_fill)
    worksheet.write(spacer, 6, 'Rating After', header_fill)

def determineFormat(matchRecord, groupPlayers):
    if (len(groupPlayers) == 8):
        return fourPersonResultFormat(matchRecord, groupPlayers)
    if (len(groupPlayers) == 10):
        return fivePersonResultFormat(matchRecord, groupPlayers)
    if (len(groupPlayers) == 12):
        return sixPersonResultFormat(matchRecord, groupPlayers)

def groupMaker(groupInfo):
    groupList = []
    for key, value in groupInfo.items():
        groupList.append(key)
        groupList.append(value)
    return groupList
        
print("Basic rules for league at GTTTA:\n")
print("There can be no more than four groups.")
print("There can be no more than seven players in any group.")
print("There can be no less than four people per group.\n")



groupsInfo = groupsInfo()
groupOneInfo = groupOneInfo(groupsInfo)
groupTwoInfo = groupTwoInfo(groupsInfo)
groupThreeInfo = groupThreeInfo(groupsInfo)
groupFourInfo = groupFourInfo(groupsInfo)


print('\nWhen asked to trust the source of the workbook, click TRUST.')
path = input('\nBefore continuing, please specify the path directory to which you would like to save your file. ')

workbook = xlsxwriter.Workbook(path)
worksheet = workbook.add_worksheet('Summary')




worksheet.set_column(1, 1, len('Seed') + 1)
worksheet.set_column(2, 2, 15)
worksheet.set_column(3, 3, len('Rating Before') - 1)
worksheet.set_column(4, 4, len('Matches Won') + 1)
worksheet.set_column(5, 5, len('Rating Change'))
worksheet.set_column(6, 6, len('Rating After') - 1)
seedLetter = ['A','B','C','D','E','F','G']



if (len(groupsInfo) == 1): ##Makes table just for group one

    groupOnePlayers = groupMaker(groupOneInfo)
    swapper(groupOnePlayers)
    swapChecker(groupOnePlayers)

    matchRecord1 = workbook.add_worksheet('Group 1')
    overallRatingChange = determineFormat(matchRecord1, groupOnePlayers)

    tableMaker(worksheet, 1, groupsInfo[0] + 1, 1)
    tableWriter(worksheet, groupsInfo[0], groupOnePlayers, overallRatingChange, seedLetter, 0)


if (len(groupsInfo) == 2): ##Makes tables just for groups one and two

    groupOnePlayers = groupMaker(groupOneInfo)
    swapper(groupOnePlayers)
    swapChecker(groupOnePlayers)

    groupTwoPlayers = groupMaker(groupTwoInfo)
    swapper(groupTwoPlayers)
    swapChecker(groupTwoPlayers)

    matchRecord1 = workbook.add_worksheet('Group 1')
    overallRatingChange = determineFormat(matchRecord1, groupOnePlayers)

    matchRecord2 = workbook.add_worksheet('Group 2')
    overallRatingChange2 = determineFormat(matchRecord2, groupTwoPlayers)

    group1Size = groupsInfo[0]
    group2Size = groupsInfo[1]
    groupSizeDifference = abs(group1Size - group2Size)

    #Makes 1st table and inputs basic info
    tableMaker(worksheet, 1, group1Size + 1, 1)
    tableWriter(worksheet, group1Size, groupOnePlayers, overallRatingChange, seedLetter, 0)
    
    #Makes 2nd table and inputs basic info
    spacer = group1Size - groupSizeDifference

    tableMaker(worksheet, spacer + 5, spacer + group2Size + 5, 2)
    tableWriter(worksheet, group2Size, groupTwoPlayers, overallRatingChange2, seedLetter, spacer + 4)
    

if (len(groupsInfo) == 3): ##Makes tables for groups 1, 2, and 3

    groupOnePlayers = groupMaker(groupOneInfo)
    swapper(groupOnePlayers)
    swapChecker(groupOnePlayers)

    groupTwoPlayers = groupMaker(groupTwoInfo)
    swapper(groupTwoPlayers)
    swapChecker(groupTwoPlayers)

    groupThreePlayers = groupMaker(groupThreeInfo)
    swapper(groupThreePlayers)
    swapChecker(groupThreePlayers)

    matchRecord1 = workbook.add_worksheet('Group 1')
    overallRatingChange = determineFormat(matchRecord1, groupOnePlayers)

    matchRecord2 = workbook.add_worksheet('Group 2')
    overallRatingChange2 = determineFormat(matchRecord2, groupTwoPlayers)

    matchRecord3 = workbook.add_worksheet('Group 3')
    overallRatingChange3 = determineFormat(matchRecord1, groupThreePlayers)
    
    group1Size = groupsInfo[0]
    group2Size = groupsInfo[1]
    group3Size = groupsInfo[2]
    groupSizeDifference = abs(group1Size - group2Size)

    #Makes 1st table and inputs basic info
    tableMaker(worksheet, 1, group1Size + 1, 1)
    tableWriter(worksheet, group1Size, groupOnePlayers, overallRatingChange, seedLetter, 0)
    
    #Makes 2nd table and inputs basic info
    spacer = group1Size - groupSizeDifference
    
    tableMaker(worksheet, spacer + 5, spacer + group2Size + 5, 2)
    tableWriter(worksheet, group2Size, groupTwoPlayers, overallRatingChange2, seedLetter, spacer + 4)

    #Makes 3rd table and inputs basic info
    spacer3 = group3Size + spacer + group2Size + 2

    tableMaker(worksheet, spacer3 + 4, spacer3 + group3Size + 4, 3)
    tableWriter(worksheet, group3Size, groupThreePlayers, overallRatingChange3, seedLetter, spacer3 + 3)


if (len(groupsInfo) == 4): ##Makes tables for all 4 groups

    groupOnePlayers = groupMaker(groupOneInfo)
    swapper(groupOnePlayers)
    swapChecker(groupOnePlayers)

    groupTwoPlayers = groupMaker(groupTwoInfo)
    swapper(groupTwoPlayers)
    swapChecker(groupTwoPlayers)

    groupThreePlayers = groupMaker(groupThreeInfo)
    swapper(groupThreePlayers)
    swapChecker(groupThreePlayers)

    groupFourPlayers = groupMaker(groupFourInfo)
    swapper(groupFourPlayers)
    swapChecker(groupFourPlayers)
    
    group1Size = groupsInfo[0]
    group2Size = groupsInfo[1]
    group3Size = groupsInfo[2]
    group4Size = groupsInfo[3]
    groupSizeDifference = abs(group1Size - group2Size)

    matchRecord1 = workbook.add_worksheet('Group 1')
    overallRatingChange = determineFormat(matchRecord1, groupOnePlayers)
    
    matchRecord2 = workbook.add_worksheet('Group 2')
    overallRatingChange2 = determineFormat(matchRecord2, groupTwoPlayers)

    matchRecord3 = workbook.add_worksheet('Group 3')
    overallRatingChange3 = determineFormat(matchRecord3, groupThreePlayers)

    matchRecord4 = workbook.add_worksheet('Group 4')
    overallRatingChange4 = determineFormat(matchRecord4, groupFourPlayers)

    #Makes 1st table and inputs basic info
    tableMaker(worksheet, 1, group1Size + 1, 1)
    tableWriter(worksheet, group1Size, groupOnePlayers, overallRatingChange, seedLetter, 0)
    
    #Makes 2nd table and inputs basic info
    spacer = group1Size - groupSizeDifference
    
    tableMaker(worksheet, spacer + 5, spacer + group2Size + 5, 2)
    tableWriter(worksheet, group2Size, groupTwoPlayers, overallRatingChange2, seedLetter, spacer + 4)


    #Makes 3rd table and inputs basic info
    spacer3 = group3Size + spacer + group2Size + 2

    tableMaker(worksheet, spacer3 + 4, spacer3 + group3Size + 4, 3)
    tableWriter(worksheet, group3Size, groupThreePlayers, overallRatingChange3, seedLetter, spacer3 + 3)
    

    #Makes 4th table and inputs basic info
    spacer4 = group4Size + spacer3 + group3Size + 2

    tableMaker(worksheet, spacer4 + 4, spacer4 + group4Size + 4, 4)
    tableWriter(worksheet, group4Size, groupFourPlayers, overallRatingChange4, seedLetter, spacer4 + 4)


print('\n\n')
print('NOTE: This only creates an excel file on your computer, when using Google Sheets, you must create a new xlsx file within the directory and import the file.')
print('\nTake NOTE OF THIS AS WELL! Since the tables do not keep the same format as excel when imported into Google Sheets, you must copy and paste the tables to replace the cells.')
print('\nDo not forget to also fill out the (Matches Won) category as well, as that is not filled out automatically by this script.')
print('\nLastly, do not forget to also update the "Current League Ratings" document.')
workbook.close()
