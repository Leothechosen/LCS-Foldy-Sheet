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
    for teams in sorted_teams_no_WL: # Assigns each team a set SoV points for where they placed in the standings. ex: {'100': 5.0, 'C9': 4.5, 'CLG': 4.0, 'DIG': 4.0, 'EG': 4.0, 'FLY': 2.5, 'GG': 2.0, 'IMT': 2.0, 'TL': 1.0, 'TSM': 0.5}
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
    teams_standings = {
        "C9":  [12, 5],
        "TL":  [11, 6],
        "TSM": [11, 6],
        "100": [11, 6],
        "DIG": [10, 7],
        "EG":  [9, 8],
        "IMT": [7, 10],
        "FLY": [6, 11],
        "CLG": [5, 12],
        "GG":  [3, 14],
    }
    teams_h2h = { # 100 |  C9 | CLG | DIG | EG | FLY | GG | IMT | TL | TSM
        "100": [[None, None], [0, 2], [2, 0], [2, 0], [1, 1], [2, 0], [1, 1], [2, 0], [1, 0], [0, 2]],
        "C9":  [[2, 0], [None, None], [1, 1], [2, 0], [1, 1], [2, 0], [2, 0], [1, 0], [0, 2], [1, 1]],
        "CLG": [[0, 2], [1, 1], [None, None], [1, 1], [0, 1], [0, 2], [1, 1], [1, 1], [1, 1], [0, 2]],
        "DIG": [[0, 2], [0, 2], [1, 1], [None, None], [2, 0], [1, 0], [2, 0], [2, 0], [0, 2], [2, 0]],
        "EG":  [[1, 1], [1, 1], [1, 0], [0, 2], [None, None], [2, 0], [2, 0], [0, 2], [1, 1], [1, 1]],
        "FLY": [[0, 2], [0, 2], [2, 0], [0, 1], [0, 2], [None, None], [1, 0], [0, 2], [0, 2], [2, 0]],
        "GG":  [[1, 1], [0, 2], [1, 1], [0, 2], [0, 2], [0, 1], [None, None], [1, 1], [0, 2], [0, 1]],
        "IMT": [[0, 2], [0, 1], [1, 1], [0, 2], [2, 0], [2, 0], [1, 1], [None, None], [1, 1], [0, 2]],
        "TL":  [[0, 1], [2, 0], [1, 1], [2, 0], [1, 1], [2, 0], [2, 0], [1, 1], [None, None], [0, 2]],
        "TSM": [[2, 0], [1, 1], [2, 0], [0, 2], [1, 1], [0, 2], [1, 0], [2, 0], [2, 0], [None, None]]
    }
    match_num = 0
    for winner in winners:
        teams_standings[winner][0] += 1
        if winner == matches[match_num][0]: # loser == matches[match_num][1]
            loser = matches[match_num][1]
            teams_h2h[winner][list(teams_h2h).index(loser)][0] += 1 # Increase winner's wins vs opponent by 1 in teams_h2h
            teams_h2h[loser][list(teams_h2h).index(winner)][1] += 1 # Increase loser's losses vs opponent by 1 in teams_h2h
            teams_standings[loser][1] += 1 # Increase's loser's losses by one in teams
        else:
            loser = matches[match_num][0]
            teams_h2h[winner][list(teams_h2h).index(loser)][0] += 1 #Increase winner's wins vs opponent by 1 in teams_h2h
            teams_h2h[loser][list(teams_h2h).index(winner)][1] += 1 #Increase loser's losses vs opponent by 1 in teams_h2h
            teams_standings[loser][1] += 1 # Increase's loser's losses by one in teams
        match_num += 1
    ordinal = 1
    for k in sorted(teams_standings, key=lambda k: (-teams_standings[k][0], teams_standings[k][1]), reverse=False):  # k = team. Sorts the teams dict by Wins descending, then losses ascending
        if sorted_teams.get(str(teams_standings.get(k))) == None:
            sorted_teams.update({str(teams_standings.get(k)): k})
        else:
            sorted_teams.update({str(teams_standings.get(k)): sorted_teams.get(str(teams_standings.get(k))) + " " + k})
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
            team_1_aggregate = teams_h2h[team_1][list(teams_h2h).index(team_2)]
            team_2_aggregate = teams_h2h[team_2][list(teams_h2h).index(team_1)]
            teams_sov = Strength_of_victory([team_1, team_2], teams_h2h, sorted_teams_no_WL)
            team_1_sov = teams_sov[0]
            team_2_sov = teams_sov[1]
            two_way_ties += 1
            if ordinal == 2 or ordinal >= 6: # In the spring split, if there is a 2 way tie for 3rd, both teams are considered the #3 seed, and therefore no tiebreaker games are needed. Check for 2-0 H2h anyway
                if team_1_aggregate == [2, 0] or team_1_aggregate == [1, 1]:
                    row_data.append([col, team_1, two_way_tie_resolved_start]) 
                    col += 1
                    row_data.append([col, team_2, two_way_tie_resolved_end])
                elif team_2_aggregate == [2, 0]:
                    row_data.append([col, team_2, two_way_tie_resolved_start]) 
                    col += 1
                    row_data.append([col, team_1, two_way_tie_resolved_end])
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal] += 1
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
            elif team_1_aggregate == [2, 0]: #If team 1 has a positive game differential against team 2, team 1 wins the tie
                row_data.append([col, team_1, two_way_tie_resolved_start])
                col += 1
                row_data.append([col, team_2, two_way_tie_resolved_end])
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal+1] += 1
            elif team_2_aggregate == [2, 0]: #If team 2 has a positive game differential against team 1, team 2 wins the tie
                row_data.append([col, team_2, two_way_tie_resolved_start]) 
                col += 1
                row_data.append([col, team_1, two_way_tie_resolved_end])
                teams_chances_no_tie[team_2][ordinal] += 1
                teams_chances_no_tie[team_1][ordinal+1] += 1
        elif len(teams_in_ordinal) == 3: #If there is a three way tie, the aggregate head-to-heads are compared. There are 5 scenarios. Original length: 543 lines | team_aggs optimization: 220 lines
            three_way_ties += 1
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_3 = teams_in_ordinal[2]
            if ordinal >= 6: #Ties that do not effect post-split seeding are not played
                row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_fully_resolved_end])
                col += 1
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal] += 1
                teams_chances_no_tie[team_3][ordinal] += 1
                continue
            team_1_aggregate = [teams_h2h[team_1][list(teams_h2h).index(team_2)][0] + teams_h2h[team_1][list(teams_h2h).index(team_3)][0], teams_h2h[team_1][list(teams_h2h).index(team_2)][1] + teams_h2h[team_1][list(teams_h2h).index(team_3)][1]]
            team_2_aggregate = [teams_h2h[team_2][list(teams_h2h).index(team_1)][0] + teams_h2h[team_2][list(teams_h2h).index(team_3)][0], teams_h2h[team_2][list(teams_h2h).index(team_1)][1] + teams_h2h[team_2][list(teams_h2h).index(team_3)][1]]
            team_3_aggregate = [teams_h2h[team_3][list(teams_h2h).index(team_1)][0] + teams_h2h[team_3][list(teams_h2h).index(team_2)][0], teams_h2h[team_3][list(teams_h2h).index(team_1)][1] + teams_h2h[team_3][list(teams_h2h).index(team_2)][1]]
            teams_aggs_dict = {team_1: team_1_aggregate, team_2: team_2_aggregate, team_3: team_3_aggregate}
            sorted_teams_aggs_dict = {}
            for team in sorted(teams_aggs_dict, key=teams_aggs_dict.get, reverse=True):
                sorted_teams_aggs_dict[team] = teams_aggs_dict[team]
            teams_in_ordinal = list(sorted_teams_aggs_dict)
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_3 = teams_in_ordinal[2]
            team_1_aggregate = sorted_teams_aggs_dict[team_1]
            team_2_aggregate = sorted_teams_aggs_dict[team_2]
            team_3_aggregate = sorted_teams_aggs_dict[team_3]
            teams_aggs = [team_1_aggregate, team_2_aggregate, team_3_aggregate]
            if teams_aggs == [[2, 2], [2, 2], [2, 2]]: # Scenario 1: If all teams are 2-2, it's an unresolved 3 way tie requiring 2 tiebreaker games.
                teams_sovs = Strength_of_victory([team_1, team_2, team_3], teams_h2h, sorted_teams_no_WL)
                team_1_sov = teams_sovs[0]
                team_2_sov = teams_sovs[1]
                team_3_sov = teams_sovs[2]
                teams_sov_dict = {team_1: team_1_sov, team_2: team_2_sov, team_3: team_3_sov}
                sorted_teams_sov_dict = {}
                for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                    sorted_teams_sov_dict[team] = teams_sov_dict[team]
                teams_in_ordinal = list(sorted_teams_sov_dict)
                team_1 = teams_in_ordinal[0]
                team_2 = teams_in_ordinal[1]
                team_3 = teams_in_ordinal[2]
                team_1_sov = sorted_teams_sov_dict[team_1]
                team_2_sov = sorted_teams_sov_dict[team_2]
                team_3_sov = sorted_teams_sov_dict[team_3]
                if ordinal == 2:  #If playing for 3rd in the Spring Split, only the lowest SOV match happens. The team with the highest SOV is automatically considered tied for 3rd with the winner of the lowest SOV match.
                    tiebreaker_games += 1
                    if team_1_sov > team_2_sov > team_3_sov:
                        teams_chances_no_tie[team_1][ordinal] += 1
                        teams_chances_tie[team_2][ordinal+1] += 1
                        teams_chances_tie[team_3][ordinal+1] += 1
                        teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                        teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_locked])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end])
                    elif team_1_sov > team_2_sov == team_3_sov:
                        teams_chances_no_tie[team_1][ordinal] += 1
                        teams_chances_tie[team_2][ordinal+1] += 1
                        teams_chances_tie[team_3][ordinal+1] += 1
                        teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                        teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_locked])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end_tied_SOV])
                    elif team_1_sov == team_2_sov > team_3_sov:
                        teams_chances_unknown[team_1][ordinal] += 1
                        teams_chances_unknown[team_2][ordinal] += 1
                        teams_chances_tie[team_3][ordinal+1] += 1
                        teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                        teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                        teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_tied_SOV])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end])
                    elif team_1_sov == team_2_sov == team_3_sov:
                        teams_chances_unknown[team_1][ordinal] += 1
                        teams_chances_unknown[team_2][ordinal] += 1
                        teams_chances_unknown[team_3][ordinal] += 1
                        teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                        teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                        teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_tied_SOV])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end_tied_SOV])
                else:
                    teams_chances_tie[team_1][ordinal] += 1
                    teams_chances_tie[team_2][ordinal] += 1
                    teams_chances_tie[team_3][ordinal] += 1
                    teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                    tiebreaker_games += 2
                    if team_1_sov > team_2_sov > team_3_sov:
                        row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_unresolved_end])
                    elif team_1_sov > team_2_sov == team_3_sov:
                        row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_unresolved_end_tied_SOV])
                    elif team_1_sov == team_2_sov > team_3_sov:
                        row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_unresolved_end])
                    elif team_1_sov == team_2_sov == team_3_sov:
                        row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_unresolved_end_tied_SOV])   
            elif teams_aggs == [[3, 1], [2, 2], [1, 3]]: # Scenario 2: If the teams have a 3-1, 2-2, and 1-3 aggregate, it's an unresolved 3 way tie requiring 2 tiebreaker games.
                if ordinal == 2: #If playing for 3rd in the Spring Split, only the match between the 2-2 and 1-3 teams happen. The 3-1 team is considered automatically tied for 3rd
                    teams_chances_no_tie[team_1][ordinal] += 1
                    teams_chances_tie[team_2][ordinal+1] += 1
                    teams_chances_tie[team_3][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    tiebreaker_games += 1
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_locked])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_partially_resolved_end])
                else:
                    teams_chances_tie[team_1][ordinal] += 1
                    teams_chances_tie[team_2][ordinal] += 1
                    teams_chances_tie[team_3][ordinal] += 1
                    teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                    tiebreaker_games += 2
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_end]) 
            elif teams_aggs == [[3, 1], [3, 1], [0, 4]]: # Scenario 3: If 2 teams are 3-1 and the 3rd team is 0-4, it's a partially resolved 3-way tie, with a tiebreaker between the 3-1 teams
                if ordinal == 2: #If playing for 3rd in the Spring Split, the two 3-1 teams are automatically considered 3rd seeds. No tiebreakers are needed.
                    teams_chances_no_tie[team_1][ordinal] += 1
                    teams_chances_no_tie[team_2][ordinal] += 1
                    teams_chances_no_tie[team_3][ordinal+2] += 1
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_end])
                else:
                    teams_chances_tie[team_1][ordinal] += 1
                    teams_chances_tie[team_2][ordinal] += 1
                    teams_chances_no_tie[team_3][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+1] += 1
                    tiebreaker_games += 1
                    teams_sovs = Strength_of_victory([team_1, team_2], teams_h2h, sorted_teams_no_WL)
                    team_1_sov = teams_sovs[0]
                    team_2_sov = teams_sovs[1]
                    if team_1_sov > team_2_sov:
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end_locked])
                    elif team_2_sov > team_1_sov:
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_begin])
                        col += 1
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end_locked])
                    elif team_1_sov == team_2_sov:
                        row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_tied_SOV])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle_tied_SOV])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end_locked])
            elif teams_aggs == [[4, 0], [1, 3], [1, 3]]: # Scenario 4: If 1 team is 4-0 and the other two teams are 1-3, it's a partially resolved 3-way tie, with a tiebreaker between the 1-3 teams.
                teams_chances_no_tie[team_1][ordinal] += 1
                if ordinal == 1: #If playing for 2nd in the Spring Split, the two 1-3 teams are considered 3rd seeds. No tiebreakers are needed.
                    teams_chances_no_tie[team_2][ordinal+1] += 1
                    teams_chances_no_tie[team_3][ordinal+1] += 1
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_end])
                else:
                    teams_chances_tie[team_2][ordinal+1] += 1
                    teams_chances_tie[team_3][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_locked])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                    col += 1
                    row_data.append([col, team_3,Multiway_tie_partially_resolved_end])
            elif teams_aggs == [[4, 0], [2, 2], [0, 4]]: # Scenario 5: If 1 team is 4-0, 1 team is 2-2, and 1 team is 0-4, it's a fully resolved 3 way tie.
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal+1] += 1
                teams_chances_no_tie[team_3][ordinal+2] += 1
                row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_fully_resolved_end])
        elif len(teams_in_ordinal) == 4: #SOV doesn't matter for seeding into a 4way tiebreaker, only side selection
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_3 = teams_in_ordinal[2]
            team_4 = teams_in_ordinal[3]
            if ordinal == 6: #If teams are playing for 7th
                row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_fully_resolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_fully_resolved_end])
                teams_chances_no_tie[teams_in_ordinal[0]][ordinal] += 1
                teams_chances_no_tie[teams_in_ordinal[1]][ordinal] += 1
                teams_chances_no_tie[teams_in_ordinal[2]][ordinal] += 1
                teams_chances_no_tie[teams_in_ordinal[3]][ordinal] += 1
            else:
                four_way_ties += 1
                teams_sov = Strength_of_victory([team_1, team_2, team_3, team_4], teams_h2h, sorted_teams_no_WL)
                teams_sov_dict = {team_1: teams_sov[0], team_2: teams_sov[1], team_3: teams_sov[2], team_4: teams_sov[3]}
                sorted_teams_sov_dict = {}
                for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                    sorted_teams_sov_dict[team] = teams_sov_dict[team]
                teams_in_ordinal = list(sorted_teams_sov_dict)
                team_1 = teams_in_ordinal[0]
                team_2 = teams_in_ordinal[1]
                team_3 = teams_in_ordinal[2]
                team_4 = teams_in_ordinal[3]
                team_1_sov = sorted_teams_sov_dict[team_1]
                team_2_sov = sorted_teams_sov_dict[team_2]
                team_3_sov = sorted_teams_sov_dict[team_3]
                team_4_sov = sorted_teams_sov_dict[team_4]
                if team_1_sov > team_2_sov > team_3_sov > team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end])
                elif team_1_sov == team_2_sov == team_3_sov == team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end_tied_SOV])
                elif team_1_sov > team_2_sov == team_3_sov == team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end_tied_SOV])
                elif team_1_sov > team_2_sov > team_3_sov == team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end_tied_SOV])
                else:
                    print("4 way tie unknown SOV resolution: ", sorted_teams_sov_dict)
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal] += 1
                teams_worst_finish_in_ties[team_2][ordinal] += 1
                teams_worst_finish_in_ties[team_3][ordinal] += 1
                teams_worst_finish_in_ties[team_4][ordinal] += 1
                if ordinal in [5, 4, 2, 0]: #If teams playing for 6th, 5th, 3rd, or 1st, only 3 games are needed for postseason seeding
                    tiebreaker_games += 3
                elif ordinal in [3, 1]: #4th and 2nd
                    tiebreaker_games += 4
        elif len(teams_in_ordinal) == 5: #2 lowest SOVs go to play-in
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_3 = teams_in_ordinal[2]
            team_4 = teams_in_ordinal[3]
            team_5 = teams_in_ordinal[4]
            teams_sov = Strength_of_victory([team_1, team_2, team_3, team_4, team_5], teams_h2h, sorted_teams_no_WL)
            teams_sov_dict = {team_1: teams_sov[0], team_2: teams_sov[1], team_3: teams_sov[2], team_4: teams_sov[3], team_5: teams_sov[4]}
            sorted_teams_sov_dict = {}
            for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                sorted_teams_sov_dict[team] = teams_sov_dict[team]
            teams_in_ordinal = list(sorted_teams_sov_dict)
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_3 = teams_in_ordinal[2]
            team_4 = teams_in_ordinal[3]
            team_5 = teams_in_ordinal[4]
            team_1_sov = sorted_teams_sov_dict[team_1]
            team_2_sov = sorted_teams_sov_dict[team_2]
            team_3_sov = sorted_teams_sov_dict[team_3]
            team_4_sov = sorted_teams_sov_dict[team_4]
            team_5_sov = sorted_teams_sov_dict[team_5]
            if team_3_sov == team_4_sov and team_4_sov > team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+3] += 1
                teams_chances_unknown[team_3][ordinal+4] += 1
                teams_chances_unknown[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1
            elif team_3_sov != team_4_sov and team_4_sov > team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+3] += 1
                teams_worst_finish_in_ties[team_3][ordinal+3] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1
            elif team_3_sov != team_4_sov and team_4_sov == team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+3] += 1
                teams_worst_finish_in_ties[team_3][ordinal+3] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1
            else:
                print("5 way tie unknown SOV resolution: ", sorted_teams_sov_dict)
            if ordinal in [5, 4, 2, 0]: #If playing for 1st, 3rd, 5th or 6th, only 4 games are needed
                tiebreaker_games += 4
            elif ordinal in [3, 1]:
                tiebreaker_games += 5
            five_way_ties += 1
        elif len(teams_in_ordinal) == 6: #4 lowest SOVs randomly drawn into playins
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_3 = teams_in_ordinal[2]
            team_4 = teams_in_ordinal[3]
            team_5 = teams_in_ordinal[4]
            team_6 = teams_in_ordinal[5]
            teams_sov = Strength_of_victory([team_1, team_2, team_3, team_4, team_5, team_6], teams_h2h, sorted_teams_no_WL)
            teams_sov_dict = {team_1: teams_sov[0], team_2: teams_sov[1], team_3: teams_sov[2], team_4: teams_sov[3], team_5: teams_sov[4], team_6: teams_sov[5]}
            sorted_teams_sov_dict = {}
            for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                sorted_teams_sov_dict[team] = teams_sov_dict[team]
            teams_in_ordinal = list(sorted_teams_sov_dict)
            team_1 = teams_in_ordinal[0]
            team_2 = teams_in_ordinal[1]
            team_3 = teams_in_ordinal[2]
            team_4 = teams_in_ordinal[3]
            team_5 = teams_in_ordinal[4]
            team_6 = teams_in_ordinal[5]
            team_1_sov = sorted_teams_sov_dict[team_1]
            team_2_sov = sorted_teams_sov_dict[team_2]
            team_3_sov = sorted_teams_sov_dict[team_3]
            team_4_sov = sorted_teams_sov_dict[team_4]
            team_5_sov = sorted_teams_sov_dict[team_5]
            team_6_sov = sorted_teams_sov_dict[team_6]
            if team_2_sov != team_3_sov: 
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                col += 1
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+3] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            else:
                print("6 way tie unknown SOV resolution: ", sorted_teams_sov_dict)
            if ordinal == 2 or ordinal == 4: #If playing for 3rd or 5th, 5 games
                tiebreaker_games += 5
            elif ordinal == 0 or ordinal == 3: # If playing for 1st or 4th, 6 games
                tiebreaker_games += 6
            elif ordinal == 1: # If playing for 2nd, 7 games
                tiebreaker_games += 7
            six_way_ties += 1
        else:
            if len(teams_in_ordinal) == 7:
                if ordinal == 2 or ordinal == 3:
                    tiebreaker_games += 7
                else:
                    tiebreaker_games += 9
                seven_way_ties += 1
            elif len(teams_in_ordinal) == 8:
                if ordinal == 2:
                    tiebreaker_games += 8
                elif ordinal == 0 or ordinal == 1:
                    tiebreaker_games += 11
                eight_way_ties += 1
            elif len(teams_in_ordinal) == 9:
                tiebreaker_games += 12
                nine_way_ties += 1
            elif len(teams_in_ordinal) == 10:
                tiebreaker_games += 13
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
no_tie_output, tie_output, unknown_output, worst_finish_output = "", "", "", ""
for team in teams_standings:
    no_tie_output += f"{team}: {teams_chances_no_tie[team]}\n"
    tie_output += f"{team}: {teams_chances_tie[team]}\n"
    unknown_output += f"{team}: {teams_chances_unknown[team]}\n"
    worst_finish_output += f"{team}: {teams_worst_finish_in_ties[team]}\n"
print("Chances of endings in Nth place - No Tiebreakers")
print(no_tie_output)
print("\nChances of playing for Nth place in Tiebreakers")
print(tie_output)
print("\nUnknown (Tied SoV in tiebreakers)")
print(unknown_output)
print("\nWorst place a team can finish in ties")
print(worst_finish_output)
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


