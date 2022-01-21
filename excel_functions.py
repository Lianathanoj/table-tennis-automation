# Author: David Gong
# Contributor: Jonathan Lian

import xlsxwriter
import shared_functions
import google_sheets_functions
import operator
import readline
import logging
import datetime
import copy
import sys
import os
from collections import defaultdict
import math

LOG_FILENAME = 'completer.log'
logging.basicConfig(
    format='%(message)s',
    filename=LOG_FILENAME,
    level=logging.DEBUG,
)

class Completer:
    def __init__(self, options):
        self.options = sorted(options)

    def complete(self, text, state):
        response = None
        if state == 0:
            if text:
                self.matches = [s for s in self.options if s and s.lower().startswith(text.lower())]
                logging.debug('%s matches: %s', repr(text), self.matches)
            else:
                self.matches = self.options[:]
                logging.debug('(empty input) matches: %s', self.matches)

        try:
            response = self.matches[state]
        except IndexError:
            response = None

        logging.debug('complete(%s, %s) => %s', repr(text), state, repr(response))
        return response

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
    letter_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G']

    def __init__(self, group_num, num_players):
        self.group_num = group_num
        self.group_name = "Group {}".format(self.group_num)
        self.num_players = num_players
        self.players = []
        self.sorted_players = []
        self.group_winner = None

    def get_info(self, league_roster, league_roster_dict, backtrack=False):
        if backtrack:
            i = len(self.players) - 1
        else:
            print('_______________________________________________________________________________\n')
            print("Please input the names and ratings for each person in {}.\n".format(self.group_name))
            i = 0
        if not league_roster_dict:
            go_back = False
            while 0 <= i < self.num_players:
                if go_back or backtrack:
                    j = 1
                else:
                    j = 0
                while 0 <= j < 2:
                    if j == 0:
                        player_name = correct_input("Name of person {} in {}: "
                                                    .format(i + 1, self.group_name), str).title()
                        if player_name.lower() == 'back':
                            i -= 1
                            j -= 1
                            if i >= 0:
                                go_back = True
                                if i <= len(self.players) - 1:
                                    print("{} has been removed from {}.\n".format(self.players.pop(), self.group_name))
                        else:
                            j += 1
                    elif j == 1:
                        player_rating = correct_input("Rating of person {} in {}: "
                                                      .format(i + 1, self.group_name), 'rating_input')
                        if player_rating == 'back':
                            j -= 1
                        else:
                            if go_back or backtrack:
                                old_player_info = copy.deepcopy(self.players[i])
                                self.players[i].player_rating = player_rating
                                print("{} has been amended to {}.\n".format(old_player_info, self.players[i]))
                            else:
                                player_info = Player(player_name=player_name, player_rating=player_rating)
                                self.players.append(player_info)
                                print("{} has been added to {}.\n".format(player_info, self.group_name))
                            go_back = False
                            i += 1
                            j += 1
            if i < 0:
                return 'backtrack'
        else:
            players_in_roster = []
            in_roster = False
            go_back = False

            readline.set_completer(Completer(league_roster).complete)
            readline.parse_and_bind('tab: complete')
            print('Press tab to autocomplete names.')

            while 0 <= i < self.num_players:
                if go_back and not in_roster:
                    j = 1
                else:
                    j = 0
                while 0 <= j < 2:
                    if j == 0:
                        player_name = correct_input("Name of person {} in {}: "
                                                    .format(i + 1, self.group_name), str).title()
                        if player_name.lower() == 'back':
                            i -= 1
                            j -= 1
                            if i >= 0:
                                go_back = True
                                if players_in_roster[i]:
                                    in_roster = True
                                else:
                                    in_roster = False
                                if i <= len(self.players) - 1:
                                    print("{} has been removed from {}.\n".format(self.players.pop(), self.group_name))
                        else:
                            j += 1
                    elif j == 1:
                        rating = league_roster_dict.get(player_name)
                        if rating:
                            player_rating = int(rating)
                            j += 1
                            if i < len(players_in_roster):
                                players_in_roster[i] = True
                            else:
                                players_in_roster.append(True)
                            go_back = False
                            player_info = Player(player_name=player_name, player_rating=player_rating)
                            self.players.append(player_info)
                            print("{} has been added to {}.\n".format(player_info, self.group_name))
                            i += 1
                        else:
                            player_rating = correct_input("Rating of person {} in {}: ".format(i + 1, self.group_name),
                                                          'rating_input')
                            if player_rating == 'back':
                                j -= 1
                            else:
                                if i < len(players_in_roster):
                                    players_in_roster[i] = False
                                else:
                                    players_in_roster.append(False)
                                if go_back:
                                    old_player_info = copy.deepcopy(self.players[i])
                                    self.players[i].player_rating = player_rating
                                    print("{} has been amended to {}.\n".format(old_player_info, self.players[i]))
                                else:
                                    player_info = Player(player_name=player_name, player_rating=player_rating)
                                    self.players.append(player_info)
                                    print("{} has been added to {}.\n".format(player_info, self.group_name))
                                go_back = False
                                i += 1
                                j += 1
            if i < 0:
                return 'backtrack'
        self.sort_ratings()
        for index, player in enumerate(self.sorted_players):
            print('{}: {}'.format(Group.letter_list[index], player))
        print('')
        return 'continue'

    def sort_ratings(self):
        self.sorted_players = sorted(self.players, key=lambda player: player.player_rating, reverse=True)

class Groups:
    def __init__(self, num_groups=0, group_list=[]):
        self.num_groups = num_groups
        self.group_list = group_list

    def add_group(self, group):
        self.group_list.append(group)

    def remove_group(self):
        if len(self.group_list) > 0:
            self.group_list.pop()

    def construct_groups(self, backtrack=False):
        if backtrack:
            index = len(self.group_list) - 1
            self.remove_group()
        else:
            num_groups = correct_input('Number of groups: ', int)
            while num_groups == 'back':
                num_groups = correct_input('Try again. Number of groups: ', int)
            self.num_groups = num_groups
            index = 0

        while 0 <= index < self.num_groups:
            group_num = index + 1
            while True:
                num_players = correct_input('How many people were in Group ' + str(group_num) + '? ', int)
                if num_players == 'back':
                    if index == 0:
                        return self.construct_groups()
                    else:
                        self.remove_group()
                        index -= 1
                        group_num = index + 1
                elif num_players < 3:
                    print('There has to be at least three people in a group. Try again.')
                elif num_players > 7:
                    print('There cannot be more than seven people in a group. Try again.')
                else:
                    break
            self.add_group(Group(group_num=group_num, num_players=num_players))
            index += 1

class ResultSheet:
    match_ordering_selection = {3: ['A:C', 'B:C', 'A:B'],
                                4: ['B:D', 'A:C', 'B:C', 'A:D', 'C:D', 'A:B'],
                                5: ['A:D', 'B:C', 'B:E', 'C:D', 'A:E', 'B:D', 'A:C', 'D:E', 'C:E', 'A:B'],
                                6: ['A:D', 'B:C', 'E:F', 'A:E', 'B:D', 'C:F', 'B:F', 'D:E', 'A:C', 'A:F',
                                    'B:E', 'C:D', 'C:E', 'D:F', 'A:B'],
                                7: ['A:F', 'B:E', 'C:D', 'B:G', 'C:F', 'D:E', 'A:E', 'B:D', 'C:G', 'A:C', 'D:F',
                                    'E:G', 'F:G', 'A:D', 'B:C', 'A:B', 'E:F', 'D:G', 'A:G', 'B:F', 'C:E']}
    letter_dict = {'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6}
    last_row_selection = {3: 11, 4: 20, 5: 33, 6: 48, 7: 66}

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
        self.last_row = ResultSheet.last_row_selection[self.num_players]
        self.match_ordering = ResultSheet.match_ordering_selection[self.num_players]
        self.match_winner = None
        self.match_list = []

    def get_match_winner(self, matches):
        #Sort by matches won, ties broken with games won
        match_winner_groups = []
        sort_match_winner = sorted(self.group.sorted_players, key=lambda player: player.games_won, reverse=True)
        sort_match_winner = sorted(sort_match_winner, key=lambda player: player.matches_won, reverse=True)
        #Check for further ties
        start = 0
        end = 1
        while(start != len(sort_match_winner)-1):
            if (sort_match_winner[start].matches_won == sort_match_winner[end].matches_won) and (sort_match_winner[start].games_won == sort_match_winner[end].games_won):
                if end+1 < len(sort_match_winner):
                    end+=1
                else:
                    match_winner_groups.append(sort_match_winner[start:])
                    break
            else:
                #if two-way tie get h2h
                if end-start == 2:
                    player1_index = self.group.sorted_players.index(sort_match_winner[start])
                    player2_index = self.group.sorted_players.index(sort_match_winner[start+1])
                    if player2_index < player1_index:
                        sort_match_winner[start], sort_match_winner[start+1] = sort_match_winner[start+1], sort_match_winner[start]
                        tempIndex = player2_index
                        player2_index = player1_index
                        player1_index = tempIndex
                    for key, val in ResultSheet.letter_dict.items():
                        if val == player1_index:
                            player1_letter = key
                        if val == player2_index:
                            player2_letter = key
                    index = self.match_ordering.index(player1_letter+':'+player2_letter)
                    if int(matches[index][0]) < int(matches[index][2]):
                        sort_match_winner[start], sort_match_winner[start+1] = sort_match_winner[start+1], sort_match_winner[start]
                    match_winner_groups.append([sort_match_winner[start]])
                    match_winner_groups.append([sort_match_winner[start+1]])
                #if more than two-way tie accept tie
                elif end-start > 2:
                    match_winner_groups.append(sort_match_winner[start:end])
                #if no tie
                elif end-start == 1:
                    match_winner_groups.append([sort_match_winner[start]])

                start = end
                if start+1 < len(sort_match_winner):
                    end = start + 1
                else:
                    match_winner_groups.append([sort_match_winner[start]])
                    break
        return match_winner_groups

    def get_group_prize_points(self):
        prize_points_amounts = {1: [10,8,6,4,2,2,2],
                                2: [8,6,4,2,2,2,2],
                                3: [8,6,4,2,2,2,2],
                                4: [7,5,3,1,1,1,1],
                                5: [7,5,3,1,1,1,1],
                                6: [7,5,3,1,1,1,1],
                                7: [7,5,3,1,1,1,1],
                                8: [7,5,3,1,1,1,1]}
        group_prize_points = {}
        rank = 0
        for i in self.match_winner:
            if len(i) == 1:
                group_prize_points[i[0].player_name] = prize_points_amounts[self.group.group_num][rank]
            elif len(i) > 1:
                point_value = math.ceil(sum(prize_points_amounts[self.group.group_num][rank:rank+len(i)])/len(i))
                for j in i:
                    group_prize_points[j.player_name] = point_value
            rank += len(i)
        return group_prize_points

    def higher_rating_is_winner(self, match):
        games_won = match[0]
        games_lost = match[2]
        if games_won == games_lost:
            return 'tied'
        return int(games_won) > int(games_lost)

    def rating_calc(self, higher_rating, lower_rating, higher_rating_is_winner):
        difference = higher_rating - lower_rating
        rating_increment = 25
        min_rating_threshold = 13
        max_rating_threshold = min_rating_threshold + rating_increment * 9 + 1

        if higher_rating_is_winner == 'tied':
            return 0
        elif higher_rating_is_winner:
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
            elif 13 <= difference < 38:
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
        self.sheet.set_column(0, 0, len_longest_substring('Match') + 3)
        self.sheet.set_column(1, 1, self.player_name_col_len + 3)
        self.sheet.set_column(2, 2, len_longest_substring('Score') + 3)
        self.sheet.set_column(3, 3, len_longest_substring('Rating Before') + 3)
        self.sheet.set_column(4, 4, len_longest_substring('Point Change') + 3)

        for player in self.group.sorted_players:
            len_player_name = len(player.player_name)
            if len_player_name + 6 > self.player_name_col_len:
                self.player_name_col_len = len_player_name + 6
                self.sheet.set_column(1, 1, self.player_name_col_len)

        self.sheet.write('A2', 'Match', self.results_header_format)
        self.sheet.write('B2', 'Players', self.results_header_format)
        self.sheet.write('C2', 'Score', self.results_header_format)
        self.sheet.write('D2', 'Rating Before', self.results_header_format)
        self.sheet.write('E2', 'Point Change', self.results_header_format)

    def construct_sheet(self, league_roster_dict, matches):
        self.sheet_merger()
        self.header_writer()

        for index, row_num in enumerate(range(self.first_row, self.last_row, 3)):
            player_one_letter = self.match_ordering[index][0]
            player_two_letter = self.match_ordering[index][2]
            player_one = self.group.sorted_players[ResultSheet.letter_dict[player_one_letter]]
            player_two = self.group.sorted_players[ResultSheet.letter_dict[player_two_letter]]
            match = matches[index]

            if len(match) == 0 or (int(match[0]) == int(match[2]) == 0):
                point_change = 0
                self.sheet.write('C' + str(row_num), 0, self.results_regular_format)
                self.sheet.write('C' + str(row_num + 1), 0, self.results_regular_format)
            else:
                point_change = self.rating_calc(player_one.player_rating, player_two.player_rating,
                                                self.higher_rating_is_winner(match))
                if int(match[0]) > int(match[2]):
                    player_one.matches_won += 1
                elif int(match[0]) < int(match[2]):
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

        for player in self.group.sorted_players:
            league_roster_dict[player.player_name] = player.final_rating

        self.match_winner = self.get_match_winner(matches)

class SummarySheet:
    seed_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G']

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
        self.player_name_col_len = 15

    def set_columns(self):
        self.worksheet.set_column(0, 0, 5)
        self.worksheet.set_column(1, 1, len_longest_substring('Seed') + 3)
        self.worksheet.set_column(2, 2, self.player_name_col_len)
        self.worksheet.set_column(3, 3, len_longest_substring('Rating Before') + 3)
        self.worksheet.set_column(4, 4, len_longest_substring('Matches Won') + 3)
        self.worksheet.set_column(5, 5, len_longest_substring('Games Won') + 3)
        self.worksheet.set_column(6, 6, len_longest_substring('Rating Change') + 3)
        self.worksheet.set_column(7, 7, len_longest_substring('Rating After') + 3)
        self.worksheet.set_column(8, 8, 5)

    def create_title_info(self):
        description = 'Group winners (denoted by **) are promoted to the next higher table during the next week' \
                      ' if they are present.'
        self.worksheet.merge_range(first_row=0, first_col=0, last_row=0, last_col=8,
                                   data='League Summary - {}'.format(self.name),
                                   cell_format=self.summary_main_title_format)
        self.worksheet.merge_range(first_row=2, first_col=2, last_row=3, last_col=6, data=description,
                                   cell_format=self.summary_description_format)

    def make_table(self, title_row_num, header_row_num, group_num):
        self.worksheet.merge_range(first_row=title_row_num, first_col=1, last_row=title_row_num, last_col=7,
                                   data='Group ' + str(group_num), cell_format=self.summary_group_title_format)
        self.worksheet.write(header_row_num, 1, 'Seed', self.summary_header_format)
        self.worksheet.write(header_row_num, 2, 'Player', self.summary_header_format)
        self.worksheet.write(header_row_num, 3, 'Rating Before', self.summary_header_format)
        self.worksheet.write(header_row_num, 4, 'Matches Won', self.summary_header_format)
        self.worksheet.write(header_row_num, 5, 'Games Won', self.summary_header_format)
        self.worksheet.write(header_row_num, 6, 'Rating Change', self.summary_header_format)
        self.worksheet.write(header_row_num, 7, 'Rating After', self.summary_header_format)

    def write_to_table(self, group_size, group, first_data_row_num, match_winner):
        for i in range(0, group_size):
            row_num = i + first_data_row_num
            player_name = group.sorted_players[i].player_name
            if group.sorted_players[i] in match_winner[0]:
                player_name += "**"

            self.worksheet.write(row_num, 1, SummarySheet.seed_letters[i], self.summary_regular_format)
            self.worksheet.write(row_num, 2, player_name, self.summary_bold_format)
            self.worksheet.write(row_num, 3, group.sorted_players[i].player_rating, self.summary_regular_format)
            self.worksheet.write(row_num, 4, group.sorted_players[i].matches_won, self.summary_regular_format)
            self.worksheet.write(row_num, 5, group.sorted_players[i].games_won, self.summary_regular_format)
            self.worksheet.write(row_num, 6, group.sorted_players[i].rating_change, self.summary_regular_format)
            self.worksheet.write(row_num, 7, group.sorted_players[i].final_rating, self.summary_bold_format)

            len_longest_name = len_longest_substring(group.sorted_players[i].player_name)
            len_longest_full_name = len(group.sorted_players[i].player_name)

            if len_longest_full_name + 6 > self.player_name_col_len:
                self.player_name_col_len = len_longest_full_name + 6
                self.worksheet.set_column(2, 2, self.player_name_col_len)

            elif len_longest_name + 6 > self.player_name_col_len:
                self.player_name_col_len = len_longest_name + 6
                self.worksheet.set_column(2, 2, self.player_name_col_len)

def check_quit(input_text):
    if input_text.lower().strip() in ['quit', 'q']:
        try:
            workbook.close()
            os.remove(file_name)
        except:
            pass
        finally:
            sys.exit()

def correct_input(input_text, var_type):
    type_dict = {str: 'string', int: 'integer', 'match_input': 'match input, e.g. 3:2',
                 'date_input': 'date input.', 'rating_input': 'rating input.'}
    pre_input = input(input_text)
    check_quit(pre_input)

    if pre_input.lower().strip() in ['back', 'b'] and var_type != 'date_input':
        return 'back'
    if var_type == 'date_input':
        date_input = pre_input.strip().replace('\'', '')
        while True:
            try:
                date_long, date_short, is_tryouts = tuple(shared_functions.reformat_file_name(date_input, 'try'))
                month, day, year = tuple([int(element) for element in date_short])
                if datetime.datetime(year=year, month=month, day=day):
                    return '{}-{}-{} Tryouts'.format(month, day, year)\
                        if is_tryouts else '{}-{}-{}'.format(month, day, year)
            except:
                print("Please input the correct format for the {}".format(type_dict[var_type]))
                date_input = input('Date: ')
                check_quit(date_input)
    elif var_type == 'match_input':
        match_input = pre_input.strip().replace('.', '')
        while True:
            try:
                if len(match_input) == 2 and match_input[0].isdigit() and match_input[1].isdigit():
                    return match_input[0] + ":" + match_input[1]
                elif len(match_input) == 3 and match_input[0].isdigit() and match_input[2].isdigit():
                    return match_input
                else:
                    raise Exception
            except:
                print("Please input the correct format for the {}".format(type_dict[var_type]))
                match_input = input(input_text)
                check_quit(match_input)
    elif var_type == 'rating_input':
        rating_input = pre_input.replace('-', '').strip()
        while True:
            try:
                if len(rating_input) < 5 and rating_input.isdigit():
                    return abs(int(rating_input))
                else:
                    raise Exception
            except:
                print("\nPlease input the correct format for the {}".format(type_dict[var_type]))
                rating_input = input(input_text)
                check_quit(rating_input)
    else:
        value = pre_input.strip()
        while True:
            try:
                if value == '':
                    return int(value)
                return var_type(value)
            except ValueError:
                print("Please input the correct value of type {}.".format(type_dict[var_type]))
                value = input(input_text).strip()
                check_quit(value)

def len_longest_substring(string):
    return len(max(string.split(' '), key=len))

def set_up_workbook():
    name = correct_input("Please input the date this league took place in 'MM-DD-YY' format.\n"
                         "If you are inputting results for tryouts, input 'MM-DD-YY Tryouts': ", 'date_input')
    print('')
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

    all_info = {'summary_info': (summary_main_title_format, summary_description_format,
                                 summary_group_title_format, summary_header_format, summary_regular_format,
                                 summary_bold_format, name),
                'results_info': (results_regular_format, results_group_title_format, results_merge_format,
                                 results_header_format)}

    return all_info, workbook, file_name

def get_match_inputs(group, backtrack=False):
    print('-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -\n')
    print('Please input the game scores for {}.'.format(group.group_name))
    print("\nFor example, assuming B won 3 - 2 for Match B vs D, input 3:2.")
    print("In the case of B losing to D 2-3, input 2:3.\n")

    match_list = []
    match_ordering = ResultSheet.match_ordering_selection[group.num_players]
    if backtrack:
        match_list.pop()
        index = len(match_list) - 1
    else:
        index = 0
    while 0 <= index < len(match_ordering):
        match = correct_input(match_ordering[index][0] + " versus " + match_ordering[index][2] + ": ", 'match_input')
        if match == 'back':
            index -= 1
            if index > 0:
                match_list.pop()
        else:
            if index < len(match_list):
                match_list[index] = match
            else:
                match_list.append(match)
            index += 1
    if index < 0:
        return 'backtrack'
    return match_list

def set_up_summary_sheet(args):
    worksheet = workbook.add_worksheet('Summary')
    summary_sheet = SummarySheet(worksheet, *args)
    summary_sheet.set_columns()
    return summary_sheet

def get_ratings_sheet_name(file_name):
    semester_month_dict = {(8, 12): 'Fall', (1, 4): 'Spring', (5, 7): 'Summer'}
    date_long, date_short, is_tryouts = shared_functions.reformat_file_name(file_name)
    month, day, year = date_long
    sheet_name = None

    for month_ranges in semester_month_dict.keys():
        if month in range(month_ranges[0], month_ranges[1] + 1):
            sheet_name = '{} {}'.format(semester_month_dict[month_ranges], year)

    return sheet_name

def get_prize_points_sheet_name(file_name):
    if 'Fall' in file_name:
        prize_points_sheet_name = str(file_name[5:])+"-"+str(int(file_name[5:])+1)
    else:
        prize_points_sheet_name = str(int(file_name[7:])-1)+"-"+str(file_name[7:])
    return prize_points_sheet_name

def generate_workbook():
    print('_______________________________________________________________________________')
    print("Basic rules for league at GTTTA:\n")
    print("There can be no more than seven players in any group.")
    print("There can be no less than four people per group.")
    print("Type 'back' or 'b' to go back at any time.")
    print("Type 'quit' or 'q' to exit the program at any time.\n")

    global file_name, workbook
    all_info, workbook, file_name = set_up_workbook()
    summary_sheet = set_up_summary_sheet(all_info['summary_info'])
    title_row_num = 4

    groups = Groups()
    groups.construct_groups()
    group_list = groups.group_list

    print('\nLoading roster, please wait...')
    service = google_sheets_functions.create_service()
    ratings_sheet_name = get_ratings_sheet_name(file_name)
    league_roster_list, league_roster_dict = google_sheets_functions.get_league_roster(service, ratings_sheet_name)
    prize_points_sheet_name = get_prize_points_sheet_name(ratings_sheet_name)
    prize_points_dict, points_used, numLeagues = google_sheets_functions.get_prize_points(service, prize_points_sheet_name)
    ratings_sheet_start_row_index = len(league_roster_dict) + 1

    group_index = 0
    backtrack = False
    group_matches = []
    while 0 <= group_index < len(group_list):
        group = group_list[group_index]
        if league_roster_list == None:
            league_roster = []
        else:
            league_roster = league_roster_list[1]
        if group.get_info(league_roster, league_roster_dict, backtrack=backtrack) == 'backtrack':
            group_index -= 1
            backtrack = False
            if group_index < 0:
                groups.construct_groups(backtrack=True)
                group_list = groups.group_list
                group_index = 0
            else:
                group_matches.pop()
        else:
            readline.set_completer(None)
            match_inputs = get_match_inputs(group)
            if match_inputs == 'backtrack':
                backtrack = True
            else:
                if group_index < len(group_matches):
                    group_matches[group_index] = (group, match_inputs)
                else:
                    group_matches.append((group, match_inputs))
                backtrack = False
                group_index += 1

    summary_sheet.create_title_info()

    prize_points = {}
    for group, matches in group_matches:
        sheet = workbook.add_worksheet(group.group_name)
        result_sheet = ResultSheet(sheet, group, *all_info['results_info'])
        result_sheet.construct_sheet(league_roster_dict, matches)
        group_prize_points = result_sheet.get_group_prize_points()
        for key, value in group_prize_points.items():
            prize_points[key] = value
        header_row_num = title_row_num + 1
        first_data_row_num = header_row_num + 1
        last_row_num = header_row_num + group.num_players
        summary_sheet.make_table(title_row_num=title_row_num, header_row_num=header_row_num,
                                 group_num=group.group_num)
        summary_sheet.write_to_table(group_size=group.num_players, group=group,
                                     first_data_row_num=first_data_row_num,
                                     match_winner=result_sheet.match_winner)
        title_row_num = last_row_num + 2
        group_index += 1

    print('_______________________________________________________________________________')
    print('\nOpening league sheet...')

    workbook.close()
    ratings_sheet_end_row_index = len(league_roster_dict) + 1
    ratings_sheet_data = [[i + 1, element[0], element[1]] for i, element
                          in enumerate(sorted(league_roster_dict.items(),
                                              key=operator.itemgetter(1),
                                              reverse=True))]
    google_sheets_functions.write_to_ratings_sheet(service=service, row_data=ratings_sheet_data,
                                                   start_row_index=ratings_sheet_start_row_index,
                                                   end_row_index=ratings_sheet_end_row_index,
                                                   sheet_name=ratings_sheet_name)

    prize_points_dict[file_name[:-5]] = prize_points
    google_sheets_functions.write_to_prize_points_sheet(service=service, roster=league_roster_dict.keys(),
                                                        prize_points=prize_points_dict,
                                                        points_used=points_used, num_leagues = numLeagues,
                                                        start_row_index=ratings_sheet_start_row_index,
                                                        end_row_index=ratings_sheet_end_row_index,
                                                        sheet_name=prize_points_sheet_name)
    return file_name

if __name__ == "__main__":
    generate_workbook()