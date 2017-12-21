# Author: David Gong
# Contributor: Jonathan Lian

import xlsxwriter

class Player:
    def __init__(self, player_name, player_rating):
        self.player_name = player_name
        self.player_rating = player_rating
        self.matches_won = 0
        self.games_won = 0
        self.rating_change = 0
        self.final_rating = player_rating

    def __str__(self):
        return str(self.player_name) + " (" + str(self.player_rating) + ")"

class Group:
    def __init__(self, group_num, num_players):
        self.group_num = group_num
        self.num_players = num_players
        self.players = []

    def get_info(self):
        for i in range(1, self.num_players + 1):
            player_name = input("Name of person {} in Group {}: ".format(i, self.group_num))
            player_rating = int(input("Rating of person {} in Group {}: ".format(i, self.group_num)))
            player_info = Player(player_name=player_name, player_rating=player_rating)
            self.players.append(player_info)
            print(str(player_info) + " has been added to Group {}.\n".format(self.group_num))

        self.sort_ratings()

    def sort_ratings(self):
        self.players = sorted(self.players, key=lambda player_info: player_info.player_rating, reverse=True)

class Groups:
    def __init__(self, num_groups=0, group_list=[]):
        self.num_groups = num_groups
        self.group_list = group_list

    def add_group(self, group):
        self.group_list.append(group)

def construct_groups():
    while True:
        try:
            num_groups = int(input('Number of groups: '))
        except ValueError:
            print('Please input an integer.')
            continue
        if num_groups > 4:
            print('There cannot be more than four groups. Try again.')
            continue
        else:
            break

    groups = Groups(num_groups=num_groups)

    for x in range(num_groups):
        group_num = x + 1

        while True:
            try:
                num_players = int(input('How many people were in Group ' + str(group_num) + '? '))
                print('\n')
            except ValueError:
                print('Please input an integer.')
                continue
            if num_players < 4:
                print('There has to be at least four people in a group. Try again.')
                continue
            if num_players > 7:
                print('There cannot be more than seven people in a group. Try again.')
            else:
                break

        groups.add_group(Group(group_num=group_num, num_players=num_players))

    return groups

def groupOneInfo(x):
    print("\n\nPlease refer to the current league ratings as you fill this part out.")
    print("\nYou are currently putting in information for Group 1.")
    playerInfo = {}
    for i in range(0, x[0]):
        playerName = input("Name of person in Group 1: ")
        playerRating = int(input("Rating of person in Group 1: "))
        playerInfo[playerName] = playerRating
        print(str(playerName) + "(" + str(playerRating) + ") has been added to Group 1.\n")

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
        print("\nYou are currently putting in information for Group 3.")
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
        print("\nYou are currently putting in information for Group 4.")
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

class ResultFormat:
    def __init__(self, sheet, group):
        self.sheet = sheet
        self.group = group
        self.num_players = group.num_players
        self.first_row = 4
        self.last_row_selection = {4: 20, 5: 33, 6: 48, 7: 66}
        self.match_ordering_selection = {4: ['B:D', 'A:C', 'B:C', 'A:D', 'C:D', 'A:B'],
                                         5: ['A:D', 'B:C', 'B:E', 'C:D', 'A:E', 'B:D', 'A:C', 'D:E', 'C:E', 'A:B'],
                                         6: ['A:D', 'B:C', 'E:F', 'A:E', 'B:D', 'C:F', 'B:F', 'D:E', 'A:C', 'A:F',
                                             'B:E', 'C:D', 'C:E', 'D:F', 'A:B'],
                                         7: ['A:F', 'B:E', 'C:D', 'B:G', 'C:F', 'D:E', 'A:E', 'B:D', 'C:G', 'A:C',
                                             'D:F', 'E:G', 'F:G', 'A:D', 'B:C', 'A:B', 'E:F', 'D:G', 'A:G', 'B:F',
                                             'C:E']}
        self.letter_dict = {'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6}
        self.last_row = self.last_row_selection[self.num_players]
        self.match_ordering = self.match_ordering_selection[self.num_players]

    def higher_rating_is_winner(self, match):
        games_won = match[0]
        games_lost = match[2]

        return games_won > games_lost

    def rating_calc(self, higher_rating, lower_rating, higher_rating_is_winner):
        difference = higher_rating - lower_rating
        rating_increment = 25
        min_rating_threshold = 13
        max_rating_threshold = min_rating_threshold + rating_increment * 9 + 1

        if higher_rating_is_winner:
            point_change = 8
            if 138 <= difference < 188:
                return 2
            elif 188 <= difference < 238:
                return 1
            elif difference >= 238:
                return 0
            for difference_threshold in range(min_rating_threshold, 139, rating_increment):
                if difference < difference_threshold:
                    return point_change
                else:
                    point_change -= 1
        else:
            point_change = -20
            if difference < 13:
                return -8
            elif 13 <= difference < 37:
                return -10
            elif 38 <= difference < 63:
                return -13
            elif 63 <= difference < 88:
                return -16
            for difference_threshold in range(113, max_rating_threshold, rating_increment):
                if difference < difference_threshold:
                    return point_change
                else:
                    point_change -= 5

        return point_change

    def sheet_merger(self):
        merge_format1 = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#C0C0C0'})
        merge_format1.set_font_size(15)

        merge_format2 = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#C0C0C0'})

        self.sheet.merge_range('A1:E1', 'Group {Replace This} - Match Record', merge_format1)

        num_merges = 6
        merge = 3
        if self.num_players != 4:
            start = 4
            while self.num_players > start:
                num_merges += 5
                start += 1
        for _ in range(num_merges):
            self.sheet.merge_range('A' + str(merge) + ':E' + str(merge), ' ', merge_format2)
            merge += 3

    def header_writer(self):
        self.sheet.write('A2', 'Match')
        self.sheet.write('B2', 'Players')
        self.sheet.write('C2', 'Score')
        self.sheet.write('D2', 'Rating Before')
        self.sheet.write('E2', 'Point Change')

    def construct_sheet(self):
        self.sheet_merger()
        self.header_writer()

        print('\nPlease input the game scores for the matches.')
        print('\nFor example, assuming B won 3 - 2 for Match B vs D, input 3:2')
        print('In the case of B losing to D 2-3, input 2:3\n')

        for index, row_num in enumerate(range(self.first_row, self.last_row, 3)):
            player_one_letter = self.match_ordering[index][0]
            player_two_letter = self.match_ordering[index][2]
            player_one = self.group.players[self.letter_dict[player_one_letter]]
            player_two = self.group.players[self.letter_dict[player_two_letter]]
            match = input(self.match_ordering[index][0] + " versus " + self.match_ordering[index][2] + ": ")
            point_change = self.rating_calc(player_one.player_rating,
                                            player_two.player_rating,
                                            self.higher_rating_is_winner(match))

            self.sheet.write('A' + str(row_num), player_one_letter)
            self.sheet.write('B' + str(row_num), player_one.player_name)
            self.sheet.write('C' + str(row_num), match[0])
            self.sheet.write('D' + str(row_num), player_one.player_rating)
            self.sheet.write('E' + str(row_num), point_change)

            self.sheet.write('A' + str(row_num + 1), player_two_letter)
            self.sheet.write('B' + str(row_num + 1), player_two.player_name)
            self.sheet.write('C' + str(row_num + 1), match[2])
            self.sheet.write('D' + str(row_num + 1), player_two.player_rating)
            self.sheet.write('E' + str(row_num + 1), -point_change)

            player_one.final_rating += point_change
            player_one.rating_change += point_change
            player_two.final_rating -= point_change
            player_two.rating_change -= point_change

def tableWriter(worksheet, groupSize, groupPlayers, seedLetterList, spacer):
    regular_fill = workbook.add_format()
    regular_fill.set_pattern(1)
    regular_fill.set_bg_color('white')

    for i in range(0, groupSize):
        row_num = i + spacer + 2
        ratingAfter = groupPlayers.players[i].final_rating
        worksheet.write(row_num, 1, seedLetterList[i], regular_fill)
        worksheet.write(row_num, 2, groupPlayers.players[i].player_name, regular_fill)
        worksheet.write(row_num, 3, groupPlayers.players[i].player_rating, regular_fill)
        worksheet.write(row_num, 4, ' ', regular_fill)
        worksheet.write(row_num, 5, groupPlayers.players[i].rating_change, regular_fill)
        worksheet.write(row_num, 6, ratingAfter, regular_fill)

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

if __name__ == "__main__":
    print("Basic rules for league at GTTTA:\n")
    print("There can be no more than four groups.")
    print("There can be no more than seven players in any group.")
    print("There can be no less than four people per group.\n")

    groups = construct_groups()
    group_list = groups.group_list
    for group in group_list:
        group.get_info()

    print('\nWhen asked to trust the source of the workbook, click TRUST.')
    path = input('\nBefore continuing, please input the name of this file. ')

    workbook = xlsxwriter.Workbook(path)
    worksheet = workbook.add_worksheet('Summary')
    worksheet.set_column(1, 1, len('Seed') + 1)
    worksheet.set_column(2, 2, 15)
    worksheet.set_column(3, 3, len('Rating Before') - 1)
    worksheet.set_column(4, 4, len('Matches Won') + 1)
    worksheet.set_column(5, 5, len('Rating Change'))
    worksheet.set_column(6, 6, len('Rating After') - 1)
    seedLetter = ['A', 'B', 'C', 'D', 'E', 'F', 'G']

    if (len(group_list) == 1): # Makes table just for group one
        group_one = group_list[0]
        sheet = workbook.add_worksheet('Group 1')
        result_format = ResultFormat(sheet, group_one)
        result_format.construct_sheet()

        tableMaker(worksheet, 1, group_one.num_players + 1, group_one.group_num)
        tableWriter(worksheet, group_one.num_players, group_one, seedLetter, 0)

    # if (len(group_list) == 2): ##Makes tables just for groups one and two
    #
    #     groupOnePlayers = groupMaker(groupOneInfo)
    #     swapper(groupOnePlayers)
    #     swapChecker(groupOnePlayers)
    #
    #     groupTwoPlayers = groupMaker(groupTwoInfo)
    #     swapper(groupTwoPlayers)
    #     swapChecker(groupTwoPlayers)
    #
    #     matchRecord1 = workbook.add_worksheet('Group 1')
    #     overallRatingChange = determineFormat(matchRecord1, groupOnePlayers)
    #
    #     matchRecord2 = workbook.add_worksheet('Group 2')
    #     overallRatingChange2 = determineFormat(matchRecord2, groupTwoPlayers)
    #
    #     group1Size = group_list[0]
    #     group2Size = group_list[1]
    #     groupSizeDifference = abs(group1Size - group2Size)
    #
    #     #Makes 1st table and inputs basic info
    #     tableMaker(worksheet, 1, group1Size + 1, 1)
    #     tableWriter(worksheet, group1Size, groupOnePlayers, overallRatingChange, seedLetter, 0)
    #
    #     #Makes 2nd table and inputs basic info
    #     spacer = group1Size - groupSizeDifference
    #
    #     tableMaker(worksheet, spacer + 5, spacer + group2Size + 5, 2)
    #     tableWriter(worksheet, group2Size, groupTwoPlayers, overallRatingChange2, seedLetter, spacer + 4)
    #
    #
    # if (len(group_list) == 3): ##Makes tables for groups 1, 2, and 3
    #
    #     groupOnePlayers = groupMaker(groupOneInfo)
    #     swapper(groupOnePlayers)
    #     swapChecker(groupOnePlayers)
    #
    #     groupTwoPlayers = groupMaker(groupTwoInfo)
    #     swapper(groupTwoPlayers)
    #     swapChecker(groupTwoPlayers)
    #
    #     groupThreePlayers = groupMaker(groupThreeInfo)
    #     swapper(groupThreePlayers)
    #     swapChecker(groupThreePlayers)
    #
    #     matchRecord1 = workbook.add_worksheet('Group 1')
    #     overallRatingChange = determineFormat(matchRecord1, groupOnePlayers)
    #
    #     matchRecord2 = workbook.add_worksheet('Group 2')
    #     overallRatingChange2 = determineFormat(matchRecord2, groupTwoPlayers)
    #
    #     matchRecord3 = workbook.add_worksheet('Group 3')
    #     overallRatingChange3 = determineFormat(matchRecord1, groupThreePlayers)
    #
    #     group1Size = group_list[0]
    #     group2Size = group_list[1]
    #     group3Size = group_list[2]
    #     groupSizeDifference = abs(group1Size - group2Size)
    #
    #     #Makes 1st table and inputs basic info
    #     tableMaker(worksheet, 1, group1Size + 1, 1)
    #     tableWriter(worksheet, group1Size, groupOnePlayers, overallRatingChange, seedLetter, 0)
    #
    #     #Makes 2nd table and inputs basic info
    #     spacer = group1Size - groupSizeDifference
    #
    #     tableMaker(worksheet, spacer + 5, spacer + group2Size + 5, 2)
    #     tableWriter(worksheet, group2Size, groupTwoPlayers, overallRatingChange2, seedLetter, spacer + 4)
    #
    #     #Makes 3rd table and inputs basic info
    #     spacer3 = group3Size + spacer + group2Size + 2
    #
    #     tableMaker(worksheet, spacer3 + 4, spacer3 + group3Size + 4, 3)
    #     tableWriter(worksheet, group3Size, groupThreePlayers, overallRatingChange3, seedLetter, spacer3 + 3)
    #
    #
    # if (len(group_list) == 4): ##Makes tables for all 4 groups
    #
    #     groupOnePlayers = groupMaker(groupOneInfo)
    #     swapper(groupOnePlayers)
    #     swapChecker(groupOnePlayers)
    #
    #     groupTwoPlayers = groupMaker(groupTwoInfo)
    #     swapper(groupTwoPlayers)
    #     swapChecker(groupTwoPlayers)
    #
    #     groupThreePlayers = groupMaker(groupThreeInfo)
    #     swapper(groupThreePlayers)
    #     swapChecker(groupThreePlayers)
    #
    #     groupFourPlayers = groupMaker(groupFourInfo)
    #     swapper(groupFourPlayers)
    #     swapChecker(groupFourPlayers)
    #
    #     group1Size = group_list[0]
    #     group2Size = group_list[1]
    #     group3Size = group_list[2]
    #     group4Size = group_list[3]
    #     groupSizeDifference = abs(group1Size - group2Size)
    #
    #     matchRecord1 = workbook.add_worksheet('Group 1')
    #     overallRatingChange = determineFormat(matchRecord1, groupOnePlayers)
    #
    #     matchRecord2 = workbook.add_worksheet('Group 2')
    #     overallRatingChange2 = determineFormat(matchRecord2, groupTwoPlayers)
    #
    #     matchRecord3 = workbook.add_worksheet('Group 3')
    #     overallRatingChange3 = determineFormat(matchRecord3, groupThreePlayers)
    #
    #     matchRecord4 = workbook.add_worksheet('Group 4')
    #     overallRatingChange4 = determineFormat(matchRecord4, groupFourPlayers)
    #
    #     #Makes 1st table and inputs basic info
    #     tableMaker(worksheet, 1, group1Size + 1, 1)
    #     tableWriter(worksheet, group1Size, groupOnePlayers, overallRatingChange, seedLetter, 0)
    #
    #     #Makes 2nd table and inputs basic info
    #     spacer = group1Size - groupSizeDifference
    #
    #     tableMaker(worksheet, spacer + 5, spacer + group2Size + 5, 2)
    #     tableWriter(worksheet, group2Size, groupTwoPlayers, overallRatingChange2, seedLetter, spacer + 4)
    #
    #
    #     #Makes 3rd table and inputs basic info
    #     spacer3 = group3Size + spacer + group2Size + 2
    #
    #     tableMaker(worksheet, spacer3 + 4, spacer3 + group3Size + 4, 3)
    #     tableWriter(worksheet, group3Size, groupThreePlayers, overallRatingChange3, seedLetter, spacer3 + 3)
    #
    #
    #     #Makes 4th table and inputs basic info
    #     spacer4 = group4Size + spacer3 + group3Size + 2
    #
    #     tableMaker(worksheet, spacer4 + 4, spacer4 + group4Size + 4, 4)
    #     tableWriter(worksheet, group4Size, groupFourPlayers, overallRatingChange4, seedLetter, spacer4 + 4)


    print('\n\n')
    print('NOTE: This only creates an excel file on your computer, when using Google Sheets, you must create a new xlsx file within the directory and import the file.')
    print('\nTake NOTE OF THIS AS WELL! Since the tables do not keep the same format as excel when imported into Google Sheets, you must copy and paste the tables to replace the cells.')
    print('\nDo not forget to also fill out the (Matches Won) category as well, as that is not filled out automatically by this script.')
    print('\nLastly, do not forget to also update the "Current League Ratings" document.')
    workbook.close()
