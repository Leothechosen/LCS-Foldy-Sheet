import xlsxwriter
import timeit
from tqdm import tqdm
import tqdm.contrib.itertools
import itertools


workbook = xlsxwriter.Workbook('C:/DiscordBots/Expirements/LoL Scenarios/LCS/LCS_Scenarios_SOV.xlsx')
worksheet = workbook.add_worksheet()

two_way_tie_unresolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'red'})
two_way_tie_unresolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'red'})
two_way_tie_unresolved_start_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'red', 'italic': True})
two_way_tie_unresolved_end_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'red', 'italic': True})

two_way_tie_resolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': '#FFCCCB'})
two_way_tie_resolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': '#FFCCCB'})

Multiway_tie_unresolved_begin = workbook.add_format({'bottom': 2, 'top': 2, 'left' : 2, 'bg_color': 'lime'})
Multiway_tie_unresolved_middle = workbook.add_format({'bottom': 2, 'top' : 2, 'bg_color': 'lime'})
Multiway_tie_unresolved_end = workbook.add_format({'bottom' : 2, 'top' : 2, 'right': 2, 'bg_color': 'lime'})

Multiway_tie_unresolved_begin_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left' : 2, 'bg_color': 'lime', 'italic': True})
Multiway_tie_unresolved_middle_tied_SOV = workbook.add_format({'bottom': 2, 'top' : 2, 'bg_color': 'lime', 'italic': True})
Multiway_tie_unresolved_end_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'right': 2, 'bg_color': 'lime', 'italic': True})

Multiway_tie_partially_resolved_begin = workbook.add_format({'bottom' : 2, 'top' : 2, 'left' : 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_middle = workbook.add_format({'bottom' : 2, 'top' : 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_end = workbook.add_format({'bottom' : 2, 'top' : 2, 'right' : 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_begin_locked = workbook.add_format({'bottom' : 2, 'top' : 2, 'left' : 2, 'bg_color': '#00FFFF', 'bold': True})
Multiway_tie_partially_resolved_end_locked = workbook.add_format({'bottom' : 2, 'top' : 2, 'right' : 2, 'bg_color': '#00FFFF', 'bold': True})

Multiway_tie_partially_resolved_begin_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'left' : 2, 'bg_color': '#00FFFF', 'italic': True})
Multiway_tie_partially_resolved_middle_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'bg_color': '#00FFFF', 'italic': True})
Multiway_tie_partially_resolved_end_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'right' : 2, 'bg_color': '#00FFFF', 'italic': True})

Multiway_tie_fully_resolved_begin = workbook.add_format({'bottom': 2, 'top': 2, 'left' : 2, 'bg_color': 'yellow'})
Multiway_tie_fully_resolved_middle = workbook.add_format({'bottom': 2, 'top' : 2, 'bg_color': 'yellow'})
Multiway_tie_fully_resolved_end = workbook.add_format({'bottom' : 2, 'top' : 2, 'right': 2, 'bg_color': 'yellow'})

def Strength_of_victory(tied_teams, teams_h2h, sorted_teams_no_WL):
    #In cases where SoV is needed to determine tiebreaker order, this function will attempt to do so.
    sov_points = {
        1: 5.0,
        2: 4.5,
        3: 4.0,
        4: 3.5,
        5: 3.0,
        6: 2.5,
        7: 2.0,
        8: 1.5,
        9: 1.0,
        10: 0.5
    }
    ordinal = 1
    teams_sov_points = {}
    for teams in sorted_teams_no_WL: # Assigns each team a set SoV points for where they placed in the standings. ex: {'G2': 5.0, 'RGE': 4.5, 'FNC': 4.0, 'MAD': 4.0, 'SK': 4.0, 'S04': 2.5, 'MSF': 2.0, 'XL': 2.0, 'AST': 1.0, 'VIT': 0.5}
        teams = teams.split()
        for team in teams:
            teams_sov_points[team] = sov_points[ordinal] 
        ordinal += len(teams)
    teams_h2h_order = ["100", "C9", "CLG", "DIG", "EG", "FLY", "GG", "IMT", "TL", "TSM"]
    tied_teams_sov = []
    for team in tied_teams: #Calculates each tied team's total SoV points and puts them in a list in the same order as tied_teams
        team_sov = 0
        team_h2h = teams_h2h[team]
        teams_h2h_index = 0
        for h2h in team_h2h: # h2h = One instance of [2, 0], [1, 1], or [0, 2]
            num_wins = h2h[0]
            if num_wins is not None:
                team_sov += (num_wins * teams_sov_points[teams_h2h_order[teams_h2h_index]])
            teams_h2h_index += 1
        tied_teams_sov.append(team_sov)
    return tied_teams_sov

two_way_ties = 0
three_way_ties = 0
four_way_ties = 0
five_way_ties = 0
six_way_ties = 0
seven_way_ties = 0
eight_way_ties = 0
nine_way_ties = 0
ten_way_ties = 0

matches = [
    ["C9", "CLG"],
    ["EG", "TSM"],
    ["100", "IMT"],
    ["TL", "DIG"],
    ["GG", "FLY"],
    ["TSM", "IMT"],
    ["DIG", "C9"],
    ["FLY", "100"],
    ["EG", "TL"],
    ["CLG", "GG"],
    ["FLY", "DIG"],
    ["GG", "TSM"],
    ["100", "TL"],
    ["CLG", "EG"],
    ["IMT", "C9"]
]

# 15 matches - Approximately 7 seconds

teams_chances_no_tie = {
    "C9":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "DIG": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "TSM": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "100": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "TL":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "EG":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "IMT": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "FLY": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "CLG": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "GG":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
}

teams_chances_tie = {
    "C9":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "DIG": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "TSM": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "100": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "TL":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "EG":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "IMT": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "FLY": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "CLG": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "GG":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
}

teams_worst_finish_in_ties = {
    "C9":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "DIG": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "TSM": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "100": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "TL":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "EG":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "IMT": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "FLY": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "CLG": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "GG":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
}
# In cases where there is a multiway tie for a place where not all the TB games need to be placed, and SoV is needed to determine tiebreaker order, if some or all SoVs are equal, it's not known to this script if a team will need to play a tiebreaker game
# As such, teams_chances_unknown lists where a team could potentially be playing for with a tb game, but it's not guaranteed.
teams_chances_unknown = { 
    "C9":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "DIG": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "TSM": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "100": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "TL":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "EG":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "IMT": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "FLY": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "CLG": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "GG":  [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
}

places = ["1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th"]
for place in places:
    globals()[place] = ""
#outcomes = tqdm.contrib.itertools.product(*matches)
start = timeit.default_timer()
outcomes = itertools.product(*matches)
outcomes = zip(range(2**len(matches)), outcomes)
worksheet_data_to_write = {}
for scenario in outcomes:
    row_data = [] # col, data, formatting
    row = scenario[0]
    winners = scenario[1]
    tiebreaker_games = 0
    scenario_num = row+1
    row_data.append([0, row+1, None]) #Writes scenario num to cell A(row)
    col = 1
    for winner in winners:
        row_data.append([col, winner, None])
        col += 1
    sorted_teams = {}
    teams = {
        "C9":  [11, 4],
        "DIG": [10, 5],
        "TSM": [10, 5],
        "100": [9, 6],
        "TL":  [9, 6],
        "EG":  [8, 7],
        "IMT": [7, 8],
        "FLY": [5, 10],
        "CLG": [4, 11],
        "GG":  [2, 13],
    }
    teams_summer = {
        "C9":  [0, 0],
        "DIG": [0, 0],
        "TSM": [0, 0],
        "100": [0, 0],
        "TL":  [0, 0],
        "EG":  [0, 0],
        "IMT": [0, 0],
        "FLY": [0, 0],
        "CLG": [0, 0],
        "GG":  [0, 0],
    }
    teams_h2h = { # 100 |  C9 | CLG | DIG | EG | FLY | GG | IMT | TL | TSM
        "100": [[None, None], [0, 2], [2, 0], [2, 0], [1, 1], [1, 0], [1, 1], [1, 0], [1, 0], [0, 2]],
        "C9":  [[2, 0], [None, None], [1, 0], [1, 0], [1, 1], [2, 0], [2, 0], [1, 0], [0, 2], [1, 1]],
        "CLG": [[0, 2], [0, 1], [None, None], [1, 1], [0, 1], [0, 2], [1, 0], [1, 1], [1, 1], [0, 2]],
        "DIG": [[0, 2], [0, 1], [1, 1], [None, None], [2, 0], [1, 0], [2, 0], [2, 0], [0, 1], [2, 0]],
        "EG":  [[1, 1], [1, 1], [1, 0], [0, 2], [None, None], [2, 0], [2, 0], [0, 2], [1, 0], [0, 1]],
        "FLY": [[0, 1], [0, 2], [2, 0], [0, 1], [0, 2], [None, None], [1, 0], [0, 2], [0, 2], [2, 0]],
        "GG":  [[1, 1], [0, 2], [0, 1], [0, 2], [0, 2], [0, 1], [None, None], [1, 1], [0, 2], [0, 1]],
        "IMT": [[0, 1], [0, 1], [1, 1], [0, 2], [2, 0], [2, 0], [1, 1], [None, None], [1, 1], [0, 1]],
        "TL":  [[0, 1], [2, 0], [1, 1], [1, 0], [0, 1], [2, 0], [2, 0], [1, 1], [None, None], [0, 2]],
        "TSM": [[2, 0], [1, 1], [2, 0], [0, 2], [1, 0], [0, 2], [1, 0], [1, 0], [2, 0], [None, None]]
    }
    teams_h2h_summer = {
        "100": [[None, None], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0]],
        "C9":  [[0, 0], [None, None], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0]],
        "CLG": [[0, 0], [0, 0], [None, None], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0]],
        "DIG": [[0, 0], [0, 0], [0, 0], [None, None], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0]],
        "EG":  [[0, 0], [0, 0], [0, 0], [0, 0], [None, None], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0]],
        "FLY": [[0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [None, None], [0, 0], [0, 0], [0, 0], [0, 0]],
        "GG":  [[0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [None, None], [0, 0], [0, 0], [0, 0]],
        "IMT": [[0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [None, None], [0, 0], [0, 0]],
        "TL":  [[0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [None, None], [0, 0]],
        "TSM": [[0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [0, 0], [None, None]]
    }
    match_num = 0
    for winner in winners:
        teams_summer[winner][0] += 1
        if winner == matches[match_num][0]:
            loser = matches[match_num][1]
        else:
            loser = matches[match_num][0]
        teams_h2h_summer[winner][list(teams_h2h_summer).index(loser)][0] += 1 # Increase winner's wins vs opponent by 1 in teams_h2h
        teams_h2h_summer[loser][list(teams_h2h_summer).index(winner)][1] += 1 # Increase loser's losses vs opponent by 1 in teams_h2h
        teams_summer[loser][1] += 1 # Increase's loser's losses by one in teams
        match_num += 1
    ordinal = 1
    for k in sorted(teams_summer, key=lambda k: (-teams_summer[k][0], teams_summer[k][1]), reverse=False):  # k = team. Sorts the teams dict by Wins descending, then losses ascending
        if sorted_teams.get(str(teams_summer.get(k))) == None:
            sorted_teams.update({str(teams_summer.get(k)): k})
        else:
            sorted_teams.update({str(teams_summer.get(k)): sorted_teams.get(str(teams_summer.get(k))) + " " + k})
    sorted_teams_no_WL = list(sorted_teams.values()) # Assigns just the teams in order (without their W-L) to values
    col += 1
    for teams in sorted_teams_no_WL:
        teams_in_ordinal = teams.split()
        first_team_in_ordinal = teams_in_ordinal[0]
        z = (' '.join(sorted_teams_no_WL[0:(sorted_teams_no_WL.index(teams)+1)]))
        zz = z.split()
        ordinal = zz.index(first_team_in_ordinal)
        globals()[(places[ordinal])] = teams
        if len(teams_in_ordinal) == 1: # If the team isn't tied with anyone
            row_data.append([col, teams_in_ordinal[0], None])
            teams_chances_no_tie[teams_in_ordinal[0]][ordinal] += 1
        elif len(teams_in_ordinal) == 2: # If there is a two way tie, it goes to head-to-head records
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_1_aggregate = teams_h2h_summer[team_1][list(teams_h2h_summer).index(team_2)]
            team_2_aggregate = teams_h2h_summer[team_2][list(teams_h2h_summer).index(team_1)]
            teams_sov = Strength_of_victory([team_1, team_2], teams_h2h_summer, sorted_teams_no_WL)
            team_1_sov = teams_sov[0]
            team_2_sov = teams_sov[1]
            two_way_ties += 1
            if ordinal >= 8: # In the summer split, only the Top 8 teams make playoffs. Two way ties for 9th, therefore, do not have tiebreakers.
                row_data.append([col, team_1, two_way_tie_resolved_start]) 
                col += 1
                row_data.append([col, team_2, two_way_tie_resolved_end])
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal] += 1
                teams_worst_finish_in_ties[team_2][ordinal] += 1
            elif team_1_aggregate == [1, 1]: #If the teams head-to-head are 1-1, it goes to a tiebreaker game.
                tiebreaker_games += 1
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                teams_worst_finish_in_ties[team_2][ordinal+1] += 1
                if team_1_sov > team_2_sov: 
                    row_data.append([col, team_1, two_way_tie_unresolved_start])
                    col += 1
                    row_data.append([col, team_2, two_way_tie_unresolved_end])
                elif team_2_sov > team_1_sov:
                    row_data.append([col, team_2, two_way_tie_unresolved_start])
                    col += 1
                    row_data.append([col, team_1, two_way_tie_unresolved_end])
                elif team_1_sov == team_2_sov:
                    row_data.append([col, team_2, two_way_tie_unresolved_start_tied_SOV])
                    col += 1
                    row_data.append([col, team_1, two_way_tie_unresolved_end_tied_SOV])
            elif team_1_aggregate in [[1, 0], [2, 0]]: #If team 1 has a positive game differential against team 2, team 1 wins the tie
                row_data.append([col, team_1, two_way_tie_resolved_start])
                col += 1
                row_data.append([col, team_2, two_way_tie_resolved_end])
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal+1] += 1
            elif team_2_aggregate in [[1, 0], [2, 0]]: #If team 2 has a positive game differential against team 1, team 2 wins the tie
                row_data.append([col, team_2, two_way_tie_resolved_start]) 
                col += 1
                row_data.append([col, team_1, two_way_tie_resolved_end])
                teams_chances_no_tie[team_2][ordinal] += 1
                teams_chances_no_tie[team_1][ordinal+1] += 1
        elif len(teams_in_ordinal) == 3: #If there is a three way tie, the combined wins of the teams in the tiebreaker are compared.
            three_way_ties += 1
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_3 = teams_in_ordinal[2]
            team_1_combined_wins = teams_h2h[team_1][list(teams_h2h).index(team_2)][0] + teams_h2h[team_1][list(teams_h2h).index(team_3)][0] + teams_h2h_summer[team_1][list(teams_h2h_summer).index(team_2)][0] + teams_h2h_summer[team_1][list(teams_h2h_summer).index(team_3)][0]
            team_2_combined_wins = teams_h2h[team_2][list(teams_h2h).index(team_1)][0] + teams_h2h[team_2][list(teams_h2h).index(team_3)][0] + teams_h2h_summer[team_1][list(teams_h2h_summer).index(team_2)][0] + teams_h2h_summer[team_2][list(teams_h2h_summer).index(team_3)][0]
            team_3_combined_wins = teams_h2h[team_3][list(teams_h2h).index(team_1)][0] + teams_h2h[team_3][list(teams_h2h).index(team_2)][0] + teams_h2h_summer[team_3][list(teams_h2h_summer).index(team_2)][0] + teams_h2h_summer[team_3][list(teams_h2h_summer).index(team_2)][0]
            teams_sov = Strength_of_victory([team_1, team_2, team_3], teams_h2h, sorted_teams_no_WL)
            team_1_sov = teams_sov[0]
            team_2_sov = teams_sov[1]
            team_3_sov = teams_sov[2]
            if team_1_combined_wins == team_2_combined_wins == team_3_combined_wins: #If all 3 times have the same Combined Wins, then it goes to 2 tiebreaker games. Lowest SoVs play each other first.
                tiebreaker_games += 2
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                teams_worst_finish_in_ties[team_3][ordinal+2] += 1 
                if team_1_sov > team_2_sov and team_1_sov > team_3_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    if team_2_sov > team_3_sov:
                        row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_unresolved_end])
                    elif team_3_sov > team_2_sov:
                        row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_unresolved_end])
                    elif team_2_sov == team_3_sov:
                        row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_unresolved_end_tied_SOV])
                elif team_2_sov > team_1_sov and team_2_sov > team_3_sov:
                    row_data.append([col, team_2, Multiway_tie_unresolved_begin])
                    col += 1
                    if team_1_sov > team_3_sov:
                        row_data.append([col, team_1, Multiway_tie_unresolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_unresolved_end])
                    elif team_3_sov > team_1_sov:
                        row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                        col += 1
                        row_data.append([col, team_1, Multiway_tie_unresolved_end])
                    elif team_1_sov == team_3_sov:
                        row_data.append([col, team_1, Multiway_tie_unresolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_unresolved_end_tied_SOV])
                elif team_3_sov > team_1_sov and team_3_sov > team_2_sov:
                    row_data.append([col, team_3, Multiway_tie_unresolved_begin])
                    col += 1
                    if team_1_sov > team_2_sov:
                        row_data.append([col, team_1, Multiway_tie_unresolved_middle])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_unresolved_end])
                    elif team_2_sov > team_1_sov:
                        row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                        col += 1
                        row_data.append([col, team_1, Multiway_tie_unresolved_end])
                    elif team_1_sov == team_2_sov:
                        row_data.append([col, team_1, Multiway_tie_unresolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_unresolved_end_tied_SOV])
                elif team_1_sov == team_2_sov and team_1_sov > team_3_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_end])
                elif team_2_sov == team_3_sov and team_2_sov > team_1_sov:
                    row_data.append([col, team_2, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_1, Multiway_tie_unresolved_end])
                elif team_3_sov == team_1_sov and team_3_sov > team_2_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_end])
                elif team_1_sov == team_2_sov == team_3_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_end_tied_SOV])
            elif team_1_combined_wins > team_2_combined_wins and team_1_combined_wins > team_3_combined_wins: #If T1 has the highest combined wins, they win the highest seed.
                teams_chances_no_tie[team_1][ordinal] += 1
                if team_2_combined_wins > team_3_combined_wins: #If T2 has the 2nd highest combined wins, they win the middle seed
                    teams_chances_no_tie[team_2][ordinal+1] += 1
                    teams_chances_no_tie[team_3][ordinal+2] += 1
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_end])
                elif team_3_combined_wins > team_2_combined_wins: # If T3 has the 2nd highest combined wins, they win the middle seed
                    teams_chances_no_tie[team_2][ordinal+2] += 1
                    teams_chances_no_tie[team_3][ordinal+1] += 1
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_end])
                elif team_2_combined_wins == team_3_combined_wins: #If T2 and T3 have the same number of combined wins, they go to a tiebreaker game, with side selection given to the best h2h. 
                    #TODO: Ask if best h2h means just in summer or in spring and summer. Assume spring and summer
                    tiebreaker_games += 1
                    teams_chances_tie[team_2][ordinal+1] += 1
                    teams_chances_tie[team_3][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_locked])
                    col += 1
                    team_2_aggregate = teams_h2h[team_2][list(teams_h2h).index(team_3)][0] + teams_h2h_summer[team_2][list(teams_h2h_summer).index(team_3)][0] #Outputs team_2's wins vs team_3
                    team_3_aggregate = teams_h2h[team_3][list(teams_h2h).index(team_2)][0] + teams_h2h_summer[team_3][list(teams_h2h_summer).index(team_2)][0] #Outputs team_3's wins vs team_2
                    if team_2_aggregate > team_3_aggregate:
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end])
                    if team_3_aggregate > team_2_aggregate:
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                        col+=1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                    if team_2_aggregate == team_3_aggregate: #TODO: Ask if this would default to SoV. Assume it does
                        if team_2_sov > team_3_sov:
                            row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                            col +=1 
                            row_data.append([col, team_3, Multiway_tie_partially_resolved_end])
                        elif team_3_sov > team_2_sov:
                            row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                            col += 1
                            row_data.append([col, team_2, Multiway_tie_partially_resolved_end])
                        elif team_2_sov == team_3_sov:
                            row_data.append([col, team_2, Multiway_tie_partially_resolved_middle_tied_SOV])
                            col += 1
                            row_data.append([col, team_3, Multiway_tie_partially_resolved_end_tied_SOV])
            elif team_2_combined_wins > team_1_combined_wins and team_2_combined_wins > team_3_combined_wins: #If T2 has the highest combined wins, they win the highest seed.
                teams_chances_no_tie[team_2][ordinal] += 1
                if team_1_combined_wins > team_3_combined_wins: #If T2 has the 2nd highest combined wins, they win the middle seed
                    teams_chances_no_tie[team_1][ordinal+1] += 1
                    teams_chances_no_tie[team_3][ordinal+2] += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_end])
                elif team_3_combined_wins > team_1_combined_wins: # If T3 has the 2nd highest combined wins, they win the middle seed
                    teams_chances_no_tie[team_1][ordinal+2] += 1
                    teams_chances_no_tie[team_3][ordinal+1] += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_end])
                elif team_1_combined_wins == team_3_combined_wins: #If T1 and T3 have the same number of combined wins, they go to a tiebreaker game, with side selection given to the best h2h. 
                    #TODO: Ask if best h2h means just in summer or in spring and summer. Assume spring and summer
                    tiebreaker_games += 1
                    teams_chances_tie[team_1][ordinal+1] += 1
                    teams_chances_tie[team_3][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_begin_locked])
                    col += 1
                    team_1_aggregate = teams_h2h[team_1][list(teams_h2h).index(team_3)][0] + teams_h2h_summer[team_1][list(teams_h2h_summer).index(team_3)][0] #Outputs team_1's wins vs team_3
                    team_3_aggregate = teams_h2h[team_3][list(teams_h2h).index(team_1)][0] + teams_h2h_summer[team_3][list(teams_h2h_summer).index(team_1)][0] #Outputs team_3's wins vs team_1
                    if team_1_aggregate > team_3_aggregate:
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end])
                    if team_3_aggregate > team_1_aggregate:
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                    if team_1_aggregate == team_3_aggregate: #TODO: Ask if this would default to SoV. Assume it does
                        if team_1_sov > team_3_sov:
                            row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                            col += 1 
                            row_data.append([col, team_3, Multiway_tie_partially_resolved_end])
                        elif team_3_sov > team_1_sov:
                            row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                            col += 1
                            row_data.append([col, team_1, Multiway_tie_partially_resolved_end])
                        elif team_1_sov == team_3_sov:
                            row_data.append([col, team_1, Multiway_tie_partially_resolved_middle_tied_SOV])
                            col += 1
                            row_data.append([col, team_3, Multiway_tie_partially_resolved_end_tied_SOV])
            elif team_3_combined_wins > team_2_combined_wins and team_3_combined_wins > team_1_combined_wins: #If T3 has the highest combined wins, they win the highest seed.
                teams_chances_no_tie[team_3][ordinal] += 1
                if team_1_combined_wins > team_2_combined_wins: #If T1 has the 2nd highest combined wins, they win the middle seed
                    teams_chances_no_tie[team_1][ordinal+1] += 1
                    teams_chances_no_tie[team_2][ordinal+2] += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_end])
                elif team_2_combined_wins > team_1_combined_wins: # If T2 has the 2nd highest combined wins, they win the middle seed
                    teams_chances_no_tie[team_1][ordinal+2] += 1
                    teams_chances_no_tie[team_2][ordinal+1] += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_end])
                elif team_1_combined_wins == team_2_combined_wins: #If T1 and T2 have the same number of combined wins, they go to a tiebreaker game, with side selection given to the best h2h. 
                    #TODO: Ask if best h2h means just in summer or in spring and summer. Assume spring and summer
                    tiebreaker_games += 1
                    teams_chances_tie[team_1][ordinal+1] += 1
                    teams_chances_tie[team_2][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    row_data.append([col, team_3, Multiway_tie_partially_resolved_begin_locked])
                    col += 1
                    team_1_aggregate = teams_h2h[team_1][list(teams_h2h).index(team_2)][0] + teams_h2h_summer[team_1][list(teams_h2h_summer).index(team_2)][0] #Outputs team_1's wins vs team_2
                    team_2_aggregate = teams_h2h[team_2][list(teams_h2h).index(team_1)][0] + teams_h2h_summer[team_2][list(teams_h2h_summer).index(team_1_aggregate)][0] #Outputs team_2's wins vs team_1
                    if team_1_aggregate > team_2_aggregate:
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_end])
                    if team_2_aggregate > team_1_aggregate:
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                        col+=1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                    if team_1_aggregate == team_2_aggregate: #TODO: Ask if this would default to SoV. Assume it does
                        if team_1_sov > team_2_sov:
                            row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                            col +=1 
                            row_data.append([col, team_2, Multiway_tie_partially_resolved_end])
                        elif team_2_sov > team_1_sov:
                            row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                            col += 1
                            row_data.append([col, team_1, Multiway_tie_partially_resolved_end])
                        elif team_1_sov == team_2_sov:
                            row_data.append([col, team_1, Multiway_tie_partially_resolved_middle_tied_SOV])
                            col += 1
                            row_data.append([col, team_2, Multiway_tie_partially_resolved_end_tied_SOV])
            elif team_1_combined_wins == team_2_combined_wins and team_1_combined_wins > team_3_combined_wins: #If T1 and T2 have the highest combined wins, they play a tiebreaker game to determine highest and middle seed. T3 gets lowest seed.
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_no_tie[team_3][ordinal+2] += 1
                teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                teams_worst_finish_in_ties[team_2][ordinal+1] += 1
                team_1_aggregate = teams_h2h[team_1][list(teams_h2h).index(team_2)][0] + teams_h2h_summer[team_1][list(teams_h2h_summer).index(team_2)][0] #Outputs team_1's wins vs team_2
                team_2_aggregate = teams_h2h[team_2][list(teams_h2h).index(team_1)][0] + teams_h2h_summer[team_2][list(teams_h2h_summer).index(team_1)][0] #Outputs team_2's wins vs team_1
                if team_1_aggregate > team_2_aggregate:
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                elif team_2_aggregate > team_1_aggregate:
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_begin])
                    col += 1
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                elif team_1_aggregate == team_2_aggregate:
                    if team_1_sov > team_2_sov:
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                    elif team_2_sov > team_1_sov:
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_begin])
                        col += 1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                    elif team_1_sov == team_2_sov:
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_tied_SOV])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_partially_resolved_end_locked])
            elif team_2_combined_wins == team_3_combined_wins and team_2_combined_wins > team_1_combined_wins: #If T2 and T3 have the highest combined wins, they play a tiebreaker game to determine highest and middle seed. T1 gets lowest seed
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_no_tie[team_1][ordinal+2] += 1
                teams_worst_finish_in_ties[team_2][ordinal+1] += 1
                teams_worst_finish_in_ties[team_3][ordinal+1] += 1
                team_2_aggregate = teams_h2h[team_2][list(teams_h2h).index(team_3)][0] + teams_h2h_summer[team_2][list(teams_h2h_summer).index(team_3)][0] #Outputs team_2's wins vs team_3
                team_3_aggregate = teams_h2h[team_3][list(teams_h2h).index(team_2)][0] + teams_h2h_summer[team_3][list(teams_h2h_summer).index(team_2)][0] #Outputs team_3's wins vs team_2
                if team_2_aggregate > team_3_aggregate:
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_begin])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                elif team_3_aggregate > team_2_aggregate:
                    row_data.append([col, team_3, Multiway_tie_partially_resolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                elif team_2_aggregate == team_3_aggregate:
                    if team_2_sov > team_3_sov:
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_begin])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                    elif team_3_sov > team_2_sov:
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_begin])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                    elif team_2_sov == team_3_sov:
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_begin_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_1, Multiway_tie_partially_resolved_end_locked])
            elif team_1_combined_wins == team_3_combined_wins and team_1_combined_wins > team_2_combined_wins: #If T1 and T3 have the highest combined wins, they play a tiebreaker game to determine highest and middle seed. T2 gets lowest seed.
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal+2] += 1
                teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                teams_worst_finish_in_ties[team_3][ordinal+1] += 1
                team_1_aggregate = teams_h2h[team_1][list(teams_h2h).index(team_3)][0] + teams_h2h_summer[team_1][list(teams_h2h_summer).index(team_3)][0] #Outputs team_1's wins vs team_3
                team_3_aggregate = teams_h2h[team_3][list(teams_h2h).index(team_1)][0] + teams_h2h_summer[team_3][list(teams_h2h_summer).index(team_1)][0] #Outputs team_3's wins vs team_1
                if team_1_aggregate > team_3_aggregate:
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_begin])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                elif team_3_aggregate > team_1_aggregate:
                    row_data.append([col, team_3, Multiway_tie_partially_resolved_begin])
                    col += 1
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                elif team_1_aggregate == team_3_aggregate:
                    if team_1_sov > team_3_sov:
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                    elif team_3_sov > team_1_sov:
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_begin])
                        col += 1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                    elif team_1_sov == team_3_sov:
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_partially_resolved_end_locked])
        else: #These assume that in cases where the teams have playin matches, the resulting new ties automatically go to tiebreakers with no head to heads considered.
            if len(teams_in_ordinal) == 4:
                if ordinal == 2: # In the Summer Split, a 4 way tie for 3rd is considered fully resolved. All teams are considered #3 seed, and will be randomly drawn into the LCS Championship.
                    row_data.append([col, teams_in_ordinal[0], Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, teams_in_ordinal[1], Multiway_tie_fully_resolved_middle])
                    col + 1
                    row_data.append([col, teams_in_ordinal[2], Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, teams_in_ordinal[3], Multiway_tie_fully_resolved_end])
                    teams_chances_no_tie[teams_in_ordinal[0]][ordinal] += 1
                    teams_chances_no_tie[teams_in_ordinal[1]][ordinal] += 1
                    teams_chances_no_tie[teams_in_ordinal[2]][ordinal] += 1
                    teams_chances_no_tie[teams_in_ordinal[3]][ordinal] += 1
                    continue
                elif ordinal == 6: # In the Summer Split, a 4 way tie for 7th only needs 3 games. Match X1 is not played.
                    tiebreaker_games += 3
                elif ordinal in [0, 1, 3, 4, 5]:
                    tiebreaker_games += 4
                four_way_ties += 4
            elif len(teams_in_ordinal) == 5:
                if ordinal == 2: #In the Summer Split, a 5 way tie for 3rd only needs the 1 playin game.
                    tiebreaker_games += 1
                else:
                    tiebreaker_games += 5
                five_way_ties += 1
            elif len(teams_in_ordinal) == 6:
                if ordinal == 2: #In the Summer Split, a 6 way tie for 3rd only needs the 3 games. Matches A1-3 and X1 are not played.
                    tiebreaker_games += 3
                elif ordinal == 4: #In the Summer Split, a 6 way tie for 5th only needs 6 games. Match L1 is not played.
                    tiebreaker_games += 6
                else:
                    tiebreaker_games += 7
                six_way_ties += 1
            elif len(teams_in_ordinal) == 7:
                if ordinal == 2: #In the Summer Split, a 7 way tie for 3rd only needs 5 games. Matches A-3 and X1 are not played.
                    tiebreaker_games += 5
                else:
                    tiebreaker_games += 9
            elif len(teams_in_ordinal) == 8: 
                if ordinal == 2: #In the Summer Split, a 8 way tie for 3rd only needs 7 games. The winner's 4-way-tie games are not played.
                    tiebreaker_games += 8
                else:
                    tiebreaker_games += 12
                eight_way_ties += 1
            elif len(teams_in_ordinal) == 9:
                tiebreaker_games += 13
                nine_way_ties += 1
            elif len(teams_in_ordinal) == 10:
                tiebreaker_games += 14
                ten_way_ties += 1
            for team in teams_in_ordinal:
                teams_worst_finish_in_ties[team][ordinal + (len(teams_in_ordinal)-1)] += 1
                if team == teams_in_ordinal[0]: #If it's the first team in the ordinal
                    row_data.append([col, team, Multiway_tie_unresolved_begin])
                    col += 1
                elif team == teams_in_ordinal[len(teams_in_ordinal) - 1]: #If it's the last team in the ordinal
                    row_data.append([col, team, Multiway_tie_unresolved_end])
                else:
                    row_data.append([col, team, Multiway_tie_unresolved_middle])
                    col += 1
                teams_chances_tie[team][ordinal] += 1
        col += 1
    row_data.append([col, tiebreaker_games, None])
    worksheet_data_to_write[row] = row_data
    col += 1
    row += 1



scenarios_stop = timeit.default_timer()
ws_start = timeit.default_timer()
for row in worksheet_data_to_write:
    row_data_to_write = worksheet_data_to_write[row]
    for data in row_data_to_write:
        col = data[0]
        writables = data[1]
        cell_format = data[2]
        worksheet.write(row, col, writables, cell_format)
ws_stop = timeit.default_timer()
ws_close_start = timeit.default_timer()
workbook.close()
ws_close_stop = timeit.default_timer()
teams_chances_start = timeit.default_timer()
print("Chances of endings in Nth place - No Tiebreakers")
for team in teams_chances_no_tie:
    print(f"{team}: {teams_chances_no_tie[team]}")
print("\nChances of playing for Nth place in Tiebreakers")
for team in teams_chances_tie:
    print(f"{team}: {teams_chances_tie[team]}")
print("\nUnknown (Tied SoV in tiebreakers)")
for team in teams_chances_unknown:
    print(f"{team}: {teams_chances_unknown[team]}")
print("\nWorst place a team can finish in ties")
for team in teams_worst_finish_in_ties:
    print(f"{team}: {teams_worst_finish_in_ties[team]}")
teams_chances_stop = timeit.default_timer()
ties_start = timeit.default_timer()
print("2-way ties;", str(two_way_ties))
print("3-way ties;", str(three_way_ties))
print("4-way ties;", str(four_way_ties))
print("5-way ties;", str(five_way_ties))
print("6-way ties;", str(six_way_ties))
print("7-way ties;", str(seven_way_ties))
print("8-way ties:", str(eight_way_ties))
print("9-way ties:", str(nine_way_ties))
print("10-way ties:", str(ten_way_ties))
ties_stop = timeit.default_timer()

stop = timeit.default_timer()
print("\nScenarios generation time: ", scenarios_stop - start)
print("Worksheet write time: ", ws_stop - ws_start)
print("Worksheet close time: ", ws_close_stop - ws_close_start)
print("Teams chances time: ", teams_chances_stop - teams_chances_start)
print("Ties time: ", ties_stop - ties_start)
print("Total time: ", stop-start)


