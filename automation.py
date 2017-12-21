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
        self.group_name = "Group {}".format(self.group_num)
        self.num_players = num_players
        self.players = []

    def get_info(self):
        for i in range(1, self.num_players + 1):
            player_name = input("Name of person {} in {}: ".format(i, self.group_name))
            player_rating = int(input("Rating of person {} in {}: ".format(i, self.group_name)))
            player_info = Player(player_name=player_name, player_rating=player_rating)
            self.players.append(player_info)
            print(str(player_info) + " has been added to {}.\n".format(self.group_name))

        self.sort_ratings()

    def sort_ratings(self):
        self.players = sorted(self.players, key=lambda player_info: player_info.player_rating, reverse=True)


class Groups:
    def __init__(self, num_groups=0, group_list=[]):
        self.num_groups = num_groups
        self.group_list = group_list

    def add_group(self, group):
        self.group_list.append(group)

    def construct_groups(self):
        while True:
            try:
                self.num_groups = int(input('Number of groups: '))
            except ValueError:
                print('Please input an integer.')
                continue
            if self.num_groups > 4:
                print('There cannot be more than four groups. Try again.')
                continue
            else:
                break

        for x in range(self.num_groups):
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
            self.add_group(Group(group_num=group_num, num_players=num_players))


class ResultSheet:
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
        return int(games_won) > int(games_lost)

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

        self.sheet.merge_range('A1:E1', '{} - Match Record'.format(group.group_name), merge_format1)

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

        print('\nPlease input the game scores for {}.'.format(self.group.group_name))
        print('\nFor example, assuming B won 3 - 2 for Match B vs D, input 3:2')
        print('In the case of B losing to D 2-3, input 2:3\n')

        for index, row_num in enumerate(range(self.first_row, self.last_row, 3)):
            player_one_letter = self.match_ordering[index][0]
            player_two_letter = self.match_ordering[index][2]
            player_one = self.group.players[self.letter_dict[player_one_letter]]
            player_two = self.group.players[self.letter_dict[player_two_letter]]

            match = input(self.match_ordering[index][0] + " versus " + self.match_ordering[index][2] + ": ")
            player_one.games_won += int(match[0])
            player_two.games_won += int(match[2])
            if int(match[0]) > int(match[2]):
                player_one.matches_won += 1
            else:
                player_two.matches_won += 1

            point_change = self.rating_calc(player_one.player_rating, player_two.player_rating,
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


class SummarySheet:
    def __init__(self, worksheet, group_title_format, header_fill, regular_fill):
        self.worksheet = worksheet
        self.group_title_format = group_title_format
        self.header_fill = header_fill
        self.regular_fill = regular_fill
        self.seed_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G']

    def set_columns(self):
        self.worksheet.set_column(1, 1, len('Seed') + 1)
        self.worksheet.set_column(2, 2, 15)
        self.worksheet.set_column(3, 3, len('Rating Before') - 1)
        self.worksheet.set_column(4, 4, len('Matches Won') + 1)
        self.worksheet.set_column(5, 5, len('Rating Change'))
        self.worksheet.set_column(6, 6, len('Rating After') - 1)

    def make_table(self, title_row_num, header_row_num, last_row_num, group_num):
        self.worksheet.merge_range(first_row=title_row_num, first_col=1, last_row=title_row_num, last_col=6,
                                   data='Group ' + str(group_num), cell_format=self.group_title_format)
        self.worksheet.add_table(first_row=header_row_num, first_col=1, last_row=last_row_num, last_col=6,
                                 options={'style': 'Table Style Light 9'})
        self.worksheet.write(header_row_num, 1, 'Seed', self.header_fill)
        self.worksheet.write(header_row_num, 2, 'Player', self.header_fill)
        self.worksheet.write(header_row_num, 3, 'Rating Before', self.header_fill)
        self.worksheet.write(header_row_num, 4, 'Matches Won', self.header_fill)
        self.worksheet.write(header_row_num, 5, 'Rating Change', self.header_fill)
        self.worksheet.write(header_row_num, 6, 'Rating After', self.header_fill)

    def write_to_table(self, group_size, group, first_data_row_num):
        for i in range(0, group_size):
            row_num = i + first_data_row_num
            self.worksheet.write(row_num, 1, self.seed_letters[i], self.regular_fill)
            self.worksheet.write(row_num, 2, group.players[i].player_name, self.regular_fill)
            self.worksheet.write(row_num, 3, group.players[i].player_rating, self.regular_fill)
            self.worksheet.write(row_num, 4, group.players[i].matches_won, self.regular_fill)
            self.worksheet.write(row_num, 5, group.players[i].rating_change, self.regular_fill)
            self.worksheet.write(row_num, 6, group.players[i].final_rating, self.regular_fill)


def set_up_workbook():
    print('\nWhen asked to trust the source of the workbook, click TRUST.')
    name = input('\nBefore continuing, please input the name of this file. ')
    if '.xlsx' not in name:
        name += '.xlsx'
    workbook = xlsxwriter.Workbook(name)

    group_title_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#99CCFF'})
    group_title_format.set_font_size(17)

    header_fill = workbook.add_format()
    header_fill.set_pattern(1)
    header_fill.set_bg_color('gray')

    regular_fill = workbook.add_format()
    regular_fill.set_pattern(1)
    regular_fill.set_bg_color('white')

    return workbook, group_title_format, header_fill, regular_fill


def set_up_summary_sheet(workbook, group_title_format, header_fill, regular_fill):
    worksheet = workbook.add_worksheet('Summary')
    summary_sheet = SummarySheet(worksheet, group_title_format, header_fill, regular_fill)
    summary_sheet.set_columns()

    return summary_sheet


if __name__ == "__main__":
    print("Basic rules for league at GTTTA:\n")
    print("There can be no more than four groups.")
    print("There can be no more than seven players in any group.")
    print("There can be no less than four people per group.\n")

    workbook, group_title_format, header_fill, regular_fill = summary_params = set_up_workbook()
    summary_sheet = set_up_summary_sheet(*summary_params)
    title_row_num = 0

    groups = Groups()
    groups.construct_groups()
    group_list = groups.group_list

    for group in group_list:
        group.get_info()
        sheet = workbook.add_worksheet(group.group_name)
        result_sheet = ResultSheet(sheet, group)
        result_sheet.construct_sheet()
        header_row_num = title_row_num + 1
        first_data_row_num = header_row_num + 1
        last_row_num = header_row_num + group.num_players
        summary_sheet.make_table(title_row_num=title_row_num, header_row_num=header_row_num,
                                 last_row_num=last_row_num, group_num=group.group_num)
        summary_sheet.write_to_table(group_size=group.num_players, group=group, first_data_row_num=first_data_row_num)
        title_row_num = last_row_num + 2

    print('\n\n')
    print(
        'NOTE: This only creates an excel file on your computer, when using Google Sheets, you must create a new xlsx file within the directory and import the file.')
    print(
        '\nTake NOTE OF THIS AS WELL! Since the tables do not keep the same format as excel when imported into Google Sheets, you must copy and paste the tables to replace the cells.')
    print('\nLastly, do not forget to also update the "Current League Ratings" document.')
    workbook.close()