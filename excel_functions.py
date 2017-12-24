# Author: David Gong
# Contributor: Jonathan Lian

import xlsxwriter

class Player:
    def __init__(self, player_name, player_rating):
        self.player_name = player_name
        self.player_rating = player_rating
        self.final_rating = player_rating
        self.matches_won = 0
        self.games_won = 0
        self.rating_change = 0

    def __str__(self):
        return str(self.player_name) + " (" + str(self.player_rating) + ")"


class Group:
    def __init__(self, group_num, num_players):
        self.group_num = group_num
        self.group_name = "Group {}".format(self.group_num)
        self.num_players = num_players
        self.players = []
        self.group_winner = None

    def get_info(self):
        for i in range(1, self.num_players + 1):
            player_name = input("Name of person {} in {}: ".format(i, self.group_name))
            player_rating = int(input("Rating of person {} in {}: ".format(i, self.group_name)))
            player_info = Player(player_name=player_name, player_rating=player_rating)
            self.players.append(player_info)
            print(str(player_info) + " has been added to {}.\n".format(self.group_name))
        self.sort_ratings()

    def sort_ratings(self):
        self.players = sorted(self.players, key=lambda player: player.player_rating, reverse=True)


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
    def __init__(self, sheet, group, results_regular_format, results_group_title_format, results_merge_format,
                 results_header_format):
        self.sheet = sheet
        self.group = group
        self.results_regular_format = results_regular_format
        self.results_group_title_format = results_group_title_format
        self.results_merge_format = results_merge_format
        self.results_header_format = results_header_format
        self.num_players = group.num_players
        self.player_name_col_len = 10
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
        self.match_winner = None

    def get_match_winner(self):
        most_matches_won = max(self.group.players, key=lambda player: player.matches_won).matches_won
        players_most_matches = list(filter(lambda player: player.matches_won == most_matches_won,
                                               self.group.players))
        most_games_won = max(players_most_matches, key=lambda player: player.games_won).games_won
        players_most_matches_games = list(filter(lambda player: player.games_won == most_games_won,
                                               players_most_matches))
        match_winner = sorted(players_most_matches_games, key=lambda player: player.final_rating)[-1]
        return match_winner

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
        self.sheet.merge_range('A1:E1', '{} - Match Record'.format(self.group.group_name),
                               self.results_group_title_format)

        merge_row_num = 3
        for _ in range(len(self.match_ordering)):
            self.sheet.merge_range('A' + str(merge_row_num) + ':E' + str(merge_row_num), ' ', self.results_merge_format)
            merge_row_num += 3

    def header_writer(self):
        self.sheet.set_column(0, 0, len_longest_substring('Match') + 1)
        self.sheet.set_column(1, 1, self.player_name_col_len + 1)
        self.sheet.set_column(2, 2, len_longest_substring('Score') + 1)
        self.sheet.set_column(3, 3, len_longest_substring('Rating Before') + 1)
        self.sheet.set_column(4, 4, len_longest_substring('Point Change') + 1)

        for player in self.group.players:
            len_player_name = len(player.player_name)
            if len_player_name > self.player_name_col_len:
                self.player_name_col_len = len_player_name + 2
                self.sheet.set_column(1, 1, self.player_name_col_len)

        self.sheet.write('A2', 'Match', self.results_header_format)
        self.sheet.write('B2', 'Players', self.results_header_format)
        self.sheet.write('C2', 'Score', self.results_header_format)
        self.sheet.write('D2', 'Rating Before', self.results_header_format)
        self.sheet.write('E2', 'Point Change', self.results_header_format)

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
            match = input(self.match_ordering[index][0] + " versus " + self.match_ordering[index][2] + ": ").strip()

            if len(match) == 0 or (int(match[0]) == int(match[2]) == 0):
                point_change = 0
                self.sheet.write('C' + str(row_num), 0, self.results_regular_format)
                self.sheet.write('C' + str(row_num + 1), 0, self.results_regular_format)
            else:
                point_change = self.rating_calc(player_one.player_rating, player_two.player_rating,
                                                self.higher_rating_is_winner(match))
                if int(match[0]) > int(match[2]):
                    player_one.matches_won += 1
                else:
                    player_two.matches_won += 1

                player_one.games_won += int(match[0])
                player_two.games_won += int(match[2])
                self.sheet.write('C' + str(row_num), int(match[0]), self.results_regular_format)
                self.sheet.write('C' + str(row_num + 1), int(match[2]), self.results_regular_format)

            self.sheet.write('A' + str(row_num), player_one_letter, self.results_regular_format)
            self.sheet.write('B' + str(row_num), player_one.player_name, self.results_regular_format)
            self.sheet.write('D' + str(row_num), player_one.player_rating, self.results_regular_format)
            self.sheet.write('E' + str(row_num), point_change, self.results_regular_format)
            self.sheet.write('A' + str(row_num + 1), player_two_letter, self.results_regular_format)
            self.sheet.write('B' + str(row_num + 1), player_two.player_name, self.results_regular_format)
            self.sheet.write('D' + str(row_num + 1), player_two.player_rating, self.results_regular_format)
            self.sheet.write('E' + str(row_num + 1), -point_change, self.results_regular_format)

            player_one.final_rating += point_change
            player_one.rating_change += point_change
            player_two.final_rating -= point_change
            player_two.rating_change -= point_change

        print('\n')
        self.match_winner = self.get_match_winner()

class SummarySheet:
    def __init__(self, worksheet, summary_main_title_format, summary_description_format, summary_group_title_format,
                 summary_header_format, summary_regular_format, summary_bold_format, name):
        self.worksheet = worksheet
        self.summary_main_title_format = summary_main_title_format
        self.summary_description_format = summary_description_format
        self.summary_group_title_format = summary_group_title_format
        self.summary_header_format = summary_header_format
        self.summary_regular_format = summary_regular_format
        self.summary_bold_format = summary_bold_format
        self.name = name
        self.date = name
        self.player_name_col_len = 15
        self.seed_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G']

    def set_columns(self):
        self.worksheet.set_column(0, 0, 5)
        self.worksheet.set_column(1, 1, len_longest_substring('Seed') + 1)
        self.worksheet.set_column(2, 2, self.player_name_col_len)
        self.worksheet.set_column(3, 3, len_longest_substring('Rating Before') + 1)
        self.worksheet.set_column(4, 4, len_longest_substring('Matches Won') + 1)
        self.worksheet.set_column(5, 5, len_longest_substring('Rating Change') + 1)
        self.worksheet.set_column(6, 6, len_longest_substring('Rating After') + 1)
        self.worksheet.set_column(7, 7, 5)

    def make_table(self, title_row_num, header_row_num, last_row_num, group_num):
        self.worksheet.merge_range(first_row=title_row_num, first_col=1, last_row=title_row_num, last_col=6,
                                   data='Group ' + str(group_num), cell_format=self.summary_group_title_format)
        self.worksheet.write(header_row_num, 1, 'Seed', self.summary_header_format)
        self.worksheet.write(header_row_num, 2, 'Player', self.summary_header_format)
        self.worksheet.write(header_row_num, 3, 'Rating Before', self.summary_header_format)
        self.worksheet.write(header_row_num, 4, 'Matches Won', self.summary_header_format)
        self.worksheet.write(header_row_num, 5, 'Rating Change', self.summary_header_format)
        self.worksheet.write(header_row_num, 6, 'Rating After', self.summary_header_format)

    def write_to_table(self, group_size, group, first_data_row_num, match_winner):
        description = 'Group winners (denoted by **) are promoted to the next higher table during the next week' \
                      ' if they are present.'
        self.worksheet.merge_range(first_row=0, first_col=0, last_row=0, last_col=7,
                                   data='League Summary - {}'.format(self.name),
                                   cell_format=self.summary_main_title_format)
        self.worksheet.merge_range(first_row=2, first_col=1, last_row=2, last_col=6, data=description,
                                   cell_format=self.summary_description_format)

        for i in range(0, group_size):
            row_num = i + first_data_row_num
            player_name = group.players[i].player_name
            if group.players[i] is match_winner:
                player_name += "**"

            self.worksheet.write(row_num, 1, self.seed_letters[i], self.summary_regular_format)
            self.worksheet.write(row_num, 2, player_name, self.summary_bold_format)
            self.worksheet.write(row_num, 3, group.players[i].player_rating, self.summary_regular_format)
            self.worksheet.write(row_num, 4, group.players[i].matches_won, self.summary_regular_format)
            self.worksheet.write(row_num, 5, group.players[i].rating_change, self.summary_regular_format)
            self.worksheet.write(row_num, 6, group.players[i].final_rating, self.summary_bold_format)

            len_longest_name = len_longest_substring(group.players[i].player_name)
            if len_longest_name > self.player_name_col_len:
                self.player_name_col_len = len_longest_name + 1
                self.worksheet.set_column(2, 2, self.player_name_col_len)

def len_longest_substring(string):
    return len(max(string.split(' '), key=len))

def set_up_workbook():
    name = input("\nPlease input the date this league took place in MM-DD-YY format. If you are inputting results for "
                 "tryouts, input 'MM-DD-YY Tryouts': ").strip().replace('\'', '')
    if '.xlsx' not in name:
        file_name = name + '.xlsx'
    workbook = xlsxwriter.Workbook(file_name)
    name = name.replace('-', '/')

    summary_main_title_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 18,
        'bg_color': '#ffed67',
        'bold': True
    })

    summary_description_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'italic': True,
        'font_size': 11
    })

    summary_group_title_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 14,
        'bg_color': '#c6d9f0'
    })

    summary_header_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'bold': True,
        'valign': 'vcenter',
        'text_wrap': True,
        'pattern': 1,
        'font_color': '#3f3f3f',
        'bg_color': '#f2f2f2'
    })

    summary_regular_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'pattern': 1,
        'bg_color': 'white'
    })

    summary_bold_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bold': True,
        'text_wrap': True,
        'pattern': 1,
        'bg_color': 'white'
    })

    results_regular_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11
    })

    results_group_title_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 15,
        'bg_color': '#c6d9f0'
    })

    results_merge_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#c6d9f0'
    })

    results_header_format = workbook.add_format({
        'font_color': '#3f3f3f',
        'bg_color': '#f2f2f2',
        'border': 1,
        'bold': True,
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter'
    })

    all_info = {'summary_info': (workbook, summary_main_title_format, summary_description_format,
                                 summary_group_title_format, summary_header_format, summary_regular_format,
                                 summary_bold_format, name),
                'results_info': (results_regular_format, results_group_title_format, results_merge_format,
                                 results_header_format)}

    return all_info, workbook, file_name


def set_up_summary_sheet(workbook, summary_main_title_format, summary_description_format, summary_group_title_format,
                         summary_header_format, summary_regular_format, summary_bold_format, name):
    worksheet = workbook.add_worksheet('Summary')
    summary_sheet = SummarySheet(worksheet, summary_main_title_format, summary_description_format,
                                 summary_group_title_format, summary_header_format, summary_regular_format,
                                 summary_bold_format, name)
    summary_sheet.set_columns()

    return summary_sheet

def generate_workbook():
    print("Basic rules for league at GTTTA:\n")
    print("There can be no more than four groups.")
    print("There can be no more than seven players in any group.")
    print("There can be no less than four people per group.\n")

    all_info, workbook, file_name = set_up_workbook()
    summary_sheet = set_up_summary_sheet(*all_info['summary_info'])
    title_row_num = 4

    groups = Groups()
    groups.construct_groups()
    group_list = groups.group_list

    for group in group_list:
        group.get_info()
        sheet = workbook.add_worksheet(group.group_name)
        result_sheet = ResultSheet(sheet, group, *all_info['results_info'])
        result_sheet.construct_sheet()
        header_row_num = title_row_num + 1
        first_data_row_num = header_row_num + 1
        last_row_num = header_row_num + group.num_players
        summary_sheet.make_table(title_row_num=title_row_num, header_row_num=header_row_num,
                                 last_row_num=last_row_num, group_num=group.group_num)
        summary_sheet.write_to_table(group_size=group.num_players, group=group,
                                     first_data_row_num=first_data_row_num,
                                     match_winner=result_sheet.match_winner)
        title_row_num = last_row_num + 2

    workbook.close()
    return file_name

if __name__ == "__main__":
    generate_workbook()