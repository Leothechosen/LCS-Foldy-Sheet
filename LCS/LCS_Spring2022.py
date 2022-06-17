#Changes from Summer 2021

#Took advantage of the new match/case functionality in Python 3.10
#Made appending to row_data it's own function, improving code efficiency
#Condensed initialization of Dict of Lists using dict comprehension
#Removed unnecessary variables
#No more variables that are randomly named

import xlsxwriter
import itertools

matches = [
    ["TL", "GG"],
    ["EG", "FLY"],
    ["DIG", "C9"],
    ["IMT", "CLG"],
    ["TSM", "100"],

    ["TL", "DIG"],
    ["GG", "EG"],
    ["C9", "100"],
    ["TSM", "CLG"],
    ["IMT", "FLY"],

    ["DIG", "100"],
    ["EG", "CLG"],
    ["FLY", "C9"],
    ["IMT", "GG"],
    ["TSM", "TL"]
]

workbook = xlsxwriter.Workbook(f"C:/DiscordBots/Expirements/LoL Scenarios/LCS-Foldy-Sheet/LCS/LCS_Scenarios_Spring2022_{len(matches)}_matches.xlsx")
worksheet = workbook.add_worksheet()

two_way_tie_unresolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'red'})
two_way_tie_unresolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'red'})
two_way_tie_unresolved_start_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'red', 'italic': True})
two_way_tie_unresolved_end_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'red', 'italic': True})

two_way_tie_resolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': '#FFCCCB'})
two_way_tie_resolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': '#FFCCCB'})

Multiway_tie_unresolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'lime'})
Multiway_tie_unresolved_middle = workbook.add_format({'bottom': 2, 'top': 2, 'bg_color': 'lime'})
Multiway_tie_unresolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'lime'})

Multiway_tie_unresolved_start_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'lime', 'italic': True})
Multiway_tie_unresolved_middle_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'bg_color': 'lime', 'italic': True})
Multiway_tie_unresolved_middle_new_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left': 1, 'bg_color': 'lime', 'italic': True})
Multiway_tie_unresolved_end_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'lime', 'italic': True})

Multiway_tie_partially_resolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_middle = workbook.add_format({'bottom': 2, 'top': 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_start_locked = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': '#00FFFF', 'bold': True})
Multiway_tie_partially_resolved_middle_locked = workbook.add_format({'bottom': 2, 'top': 2, 'bg_color': '#00FFFF', 'bold': True})
Multiway_tie_partially_resolved_end_locked = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': '#00FFFF', 'bold': True})

Multiway_tie_partially_resolved_start_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': '#00FFFF', 'italic': True})
Multiway_tie_partially_resolved_middle_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'bg_color': '#00FFFF', 'italic': True})
Multiway_tie_partially_resolved_end_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': '#00FFFF', 'italic': True})
Multiway_tie_partially_resolved_middle_new_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left': 1, 'bg_color': '#00FFFF', 'italic': True})

Multiway_tie_resolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'yellow'})
Multiway_tie_resolved_middle = workbook.add_format({'bottom': 2, 'top': 2, 'bg_color': 'yellow'})
Multiway_tie_resolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'yellow'})

def strength_of_victory(tied_teams, teams_h2h, sorted_teams):
    #Calculates SoV when necessary
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
    for teams in sorted_teams:
        teams = teams.split()
        for team in teams:
            teams_sov_points[team] = sov_points[ordinal] # Assigns each team the SoV value if you had beaten them.
        ordinal += len(teams)
    tied_teams_sov = []
    for team in tied_teams:
        team_sov = 0
        team_h2h = teams_h2h[team]
        for opp_team in team_h2h:
            team_sov += (team_h2h[opp_team] * teams_sov_points[opp_team])
        tied_teams_sov.append(team_sov)
    return tied_teams_sov

def append_row_data(row_data, col, teams, ties=None, sov_ties=None):
    # col - int:column num
    # teams - list:teams to write
    # ties - None, str("Resolved" or "Unresolved"), or list["Locked", None, None]
    # sov_ties - None or list, Example [False, True, True]. If more than 1 sov tie, [False, True, True, "New", True]
    # scenario_num - Debugging purposes
    if len(teams) == 1:
        row_data.append([col, teams[0], None])
        col += 1
    elif len(teams) == 2:
        team_1, team_2 = teams
        if ties == "Resolved":
            team_1_fmt = two_way_tie_resolved_start
            team_2_fmt = two_way_tie_resolved_end
        elif sov_ties == None:
            team_1_fmt = two_way_tie_unresolved_start
            team_2_fmt = two_way_tie_unresolved_end
        else:
            team_1_fmt = two_way_tie_unresolved_start_tied_SOV
            team_2_fmt = two_way_tie_unresolved_end_tied_SOV
        row_data.append([col, team_1, team_1_fmt])
        row_data.append([col+1, team_2, team_2_fmt])
        col += 2
    else:
        for team in teams:
            tie = ties[teams.index(team)] if type(ties) == list else ties
            sov_tie = sov_ties[teams.index(team)] if type(sov_ties) == list else None
            match (tie, teams.index(team), sov_tie):
                case ("Resolved", 0, _): # Resolved Tie, First Team
                    fmt = Multiway_tie_resolved_start
                case ("Resolved", x, _) if x == len(teams)-1: # Resolved Tie, Last Team
                    fmt = Multiway_tie_resolved_end
                case ("Resolved", _, _): # Resolved Tie, Middle Team
                    fmt = Multiway_tie_resolved_middle
                case ("Unresolved", 0, True): #Unresolved Tie, First Team, SOV Tie
                    fmt = Multiway_tie_unresolved_start_tied_SOV
                case ("Unresolved", 0, _): # Unresolved Tie, First Team, No SOV Tie
                    fmt = Multiway_tie_unresolved_start
                case ("Unresolved", x, True) if x == len(teams)-1: # Unresolved Tie, Last Team, SOV Tie
                    fmt = Multiway_tie_unresolved_end_tied_SOV
                case ("Unresolved", x, _) if x == len(teams)-1: # Unresolved Tie, Last Team, No SOV Tie
                    fmt = Multiway_tie_unresolved_end
                case ("Unresolved", _, True): # Unresolved Tie, Middle Team, SOV Tie
                    fmt = Multiway_tie_unresolved_middle_tied_SOV
                case ("Unresolved", _, "New"): # Unresolved Tie, Middle Team, New SOV Tie
                    fmt = Multiway_tie_unresolved_middle_new_tied_SOV
                case ("Unresolved", _, _): # Unresolved Tie, Middle Team, No SOV Tie
                    fmt = Multiway_tie_unresolved_middle
                case ("Locked", 0, True): # Partially Resolved Tie, First Team Locked, SOV Tie
                    print("Multiway_tie_partially_resolved_start_locked_tied_SOV")
                    exit()
                    #fmt = Multiway_tie_partially_resolved_start_locked_tied_SOV
                case ("Locked", 0, _): # Partially Resolved Tie, First Team Locked, No SOV Tie
                    fmt = Multiway_tie_partially_resolved_start_locked
                case ("Locked", x, True) if x == len(teams)-1: # Partially Resolved Tie, Last Team Locked, SOV Tie
                    print("Multiway_tie_partially_resolved_end_locked_tied_SOV")
                    exit()
                    #fmt = Multiway_tie_partially_resolved_end_locked_tied_SOV
                case ("Locked", x, _) if x == len(teams)-1: # Partially Resolved Tie, Last Team Locked, No SOV Tie
                    fmt = Multiway_tie_partially_resolved_end_locked
                case ("Locked", _, True): # Partially Resolved Tie, Middle Team Locked, SOV Tie
                    print("Multiway_tie_partially_resolved_middle_locked_tied_SOV")
                    exit()
                    #fmt = Multiway_tie_partially_resolved_middle_locked_tied_SOV
                case ("Locked", _, "New"): # Partially Resolved Tie, Middle Team Locked, New SOV Tie
                    print("Multiway_tie_partially_resolved_middle_locked_new_tied_SOV")
                    exit()
                    #fmt = Multiway_tie_partially_resolved_middle_locked_new_tied_SOV
                case ("Locked", _, _): # Partially Resolved Tie, Middle Team Locked, No SOV Tie
                    fmt = Multiway_tie_partially_resolved_middle_locked
                case (None, 0, True): # Partially Resolved Tie, First Team, SOV Tie
                    fmt = Multiway_tie_partially_resolved_start_tied_SOV
                case (None, 0, _): # Partially Resolved Tie, First Team, No SOV Tie
                    fmt = Multiway_tie_partially_resolved_start
                case (None, x, True) if x == len(teams)-1: # Partially Resolved Tie, Last Team, SOV Tie
                    fmt = Multiway_tie_partially_resolved_end_tied_SOV
                case (None, x, _) if x == len(teams)-1: # Partially Resolved Tie, Last Team, No SOV Tie
                    fmt = Multiway_tie_partially_resolved_end
                case (None, _, True): # Partially Resolved Tie, Middle Team, SOV Tie
                    fmt = Multiway_tie_partially_resolved_middle_tied_SOV
                case (None, _, "New"): # Partially Resolved Tie, Middle Team, New SOV Tie
                    fmt = Multiway_tie_partially_resolved_middle_new_tied_SOV
                case (None, _, _): # Partially Resolved Tie, Middle Team, No SOV Tie
                    fmt = Multiway_tie_partially_resolved_middle
                case (_, _, _):
                    print("Oh no, there's a format not worked out.")
                    print(tie, teams.index(team), sov_tie)
                    exit()
            row_data.append([col, team, fmt])
            col += 1
    return row_data, col
x_way_ties = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

teams = ["C9", "DIG", "TSM", "100", "TL", "EG", "IMT", "FLY", "CLG", "GG"]
teams_chances_no_tie = {team: [0] * 10 for team in teams}
teams_chances_tie = {team: [0] * 10 for team in teams}
teams_worst_finish_in_ties = {team: [0] * 10 for team in teams}
teams_chances_unknown = {team: [0] * 10 for team in teams}

outcomes = itertools.product(*matches)
worksheet_data_to_write = {}

for scenario_num, winners in enumerate(outcomes, start=1):
    tiebreaker_games = 0
    row_data = [] # row_data.append([column, data, format])
    row = scenario_num
    row_data.append([0, row, None])
    col = 1
    standings = {
        "C9":  12,
        "TL":  11,
        "100": 10,
        "FLY": 8,
        "DIG": 7,
        "EG":  7,
        "GG":  7,
        "CLG": 5,
        "IMT": 4,
        "TSM": 4,
        }
    h2h = { #Believe it or not, representing it like this instead of a dictionary of lists cuts interaction time by ~75%.
        '100': {   
            'C9':  0,
            'CLG': 2,
            'DIG': 0,
            'EG':  2,
            'FLY': 1,
            'GG':  2,
            'IMT': 1,
            'TL':  1,
            'TSM': 1},
        'C9': {
            '100': 1,
            'CLG': 1,
            'DIG': 1,
            'EG':  2,
            'FLY': 1,
            'GG':  2,
            'IMT': 2,
            'TL':  1,
            'TSM': 1},
        'CLG': {   
            '100': 0,
            'C9':  1,
            'DIG': 0,
            'EG':  1,
            'FLY': 1,
            'GG':  1,
            'IMT': 0,
            'TL':  0,
            'TSM': 1},
        'DIG': {   
            '100': 1,
            'C9':  0,
            'CLG': 2,
            'EG':  0,
            'FLY': 0,
            'GG':  0,
            'IMT': 2,
            'TL':  0,
            'TSM': 2},
        'EG': {   
            '100': 0,
            'C9':  0,
            'CLG': 0,
            'DIG': 2,
            'FLY': 1,
            'GG':  1,
            'IMT': 1,
            'TL':  0,
            'TSM': 2},
        'FLY': {   
            '100': 1,
            'C9':  0,
            'CLG': 1,
            'DIG': 2,
            'EG':  0,
            'GG':  1,
            'IMT': 0,
            'TL':  1,
            'TSM': 2},
        'GG': {   
            '100': 0,
            'C9':  0,
            'CLG': 1,
            'DIG': 2,
            'EG':  0,
            'FLY': 1,
            'IMT': 1,
            'TL':  1,
            'TSM': 1},
        'IMT': {   
            '100': 1,
            'C9':  0,
            'CLG': 1,
            'DIG': 0,
            'EG':  1,
            'FLY': 1,
            'GG':  0,
            'TL':  0,
            'TSM': 0},
        'TL': {   
            '100': 1,
            'C9':  1,
            'CLG': 2,
            'DIG': 1,
            'EG':  2,
            'FLY': 1,
            'GG':  0,
            'IMT': 2,
            'TSM': 1},
        'TSM': {   
            '100': 0,
            'C9':  1,
            'CLG': 0,
            'DIG': 0,
            'EG':  0,
            'FLY': 0,
            'GG':  1,
            'IMT': 2,
            'TL':  0}}
    match_num = 0
    for winner in winners:
        standings[winner] += 1
        if winner == matches[match_num][0]:
            loser = matches[match_num][1]
        else:
            loser = matches[match_num][0]
        match_num += 1
        h2h[winner][loser] += 1
        row_data.append([col, winner, None])
        col += 1
    sorted_teams = {}
    for team in sorted(standings, key=lambda team: (-standings[team]), reverse=False):
        if sorted_teams.get(standings.get(team)) == None:
            sorted_teams.update({standings.get(team): team})
        else:
            sorted_teams.update({standings.get(team): sorted_teams.get(standings.get(team)) + " " + team})
    sorted_teams = list(sorted_teams.values())
    col += 1
    ordinal = 0
    for teams in sorted_teams:
        teams_in_ordinal = teams.split()
        if len(teams_in_ordinal) != 1:
            x_way_ties[len(teams_in_ordinal) - 1] += 1
        if len(teams_in_ordinal) == 1:
            row_data, col = append_row_data(row_data, col, teams_in_ordinal)
            teams_chances_no_tie[teams_in_ordinal[0]][ordinal] += 1
        elif len(teams_in_ordinal) in [2, 3]:
            teams_aggs = {}
            for team in teams_in_ordinal:
                team_agg = 0
                for other_team in teams_in_ordinal:
                    if team != other_team:
                        team_agg += h2h[team][other_team]
                teams_aggs.update({team: team_agg})
            sorted_teams_aggs = {}
            for team in sorted(teams_aggs, key = teams_aggs.get, reverse=True):
                if sorted_teams_aggs.get(teams_aggs.get(team)) == None:
                    sorted_teams_aggs.update({teams_aggs.get(team): team})
                else:
                    sorted_teams_aggs.update({teams_aggs.get(team): sorted_teams_aggs.get(teams_aggs.get(team)) + " " + team})
            sorted_teams_no_aggs = list(sorted_teams_aggs.values()) # Returns ["TSM TL"] if 2 way tie after H2H. Returns ["TSM", "TL"] if no 2 way tie after H2H.
            match (len(teams_in_ordinal), len(sorted_teams_no_aggs)): #If equal, then there are no H2H ties. If not equal, there are H2H ties.
                case (2, 2): # No ties in 2-way h2h
                    team_1, team_2 = sorted_teams_no_aggs
                    row_data, col = append_row_data(row_data, col, [team_1, team_2], "Resolved")
                    teams_chances_no_tie[team_1][ordinal] += 1
                    teams_chances_no_tie[team_2][ordinal+1] += 1
                case (2, 1): # 2-way-tie after h2h
                    if ordinal >= 6: #If tied for 7th or worse, no tiebreakers are played and the tie is resolved.
                        row_data, col = append_row_data(row_data, col, sorted_teams_no_aggs[0].split(), "Resolved")
                        for team in sorted_teams_no_aggs[0].split():
                            teams_chances_no_tie[team][ordinal] += 1
                    else:
                        tiebreaker_games += 1
                        team_1, team_2 = teams_in_ordinal
                        team_1_sov, team_2_sov = strength_of_victory(teams_in_ordinal, h2h, sorted_teams)
                        if team_1_sov > team_2_sov:
                            row_data, col = append_row_data(row_data, col, [team_1, team_2], "Unresolved")
                        elif team_2_sov > team_1_sov:
                            row_data, col = append_row_data(row_data, col, [team_2, team_1], "Unresolved")
                        else:
                            row_data, col = append_row_data(row_data, col, [team_1, team_2], "Unresolved", [True, True])
                        teams_chances_tie[team_1][ordinal] += 1
                        teams_chances_tie[team_2][ordinal] += 1
                        teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                        teams_worst_finish_in_ties[team_2][ordinal+1] += 1
                case (3, 1): # All 3 teams are tied in h2h. Scenario 1 in LCS Rulebook 8.6
                    if ordinal >= 6:
                        row_data, col = append_row_data(row_data, col, sorted_teams_no_aggs[0].split(), "Resolved")
                        for team in sorted_teams_no_aggs[0].split():
                            teams_chances_no_tie[team][ordinal] += 1
                    else:
                        tiebreaker_games += 2
                        team_1, team_2, team_3 = teams_in_ordinal
                        teams_sovs = strength_of_victory([team_1, team_2, team_3], h2h, sorted_teams)
                        sovs_dict = {team_1: teams_sovs[0], team_2: teams_sovs[1], team_3: teams_sovs[2]}
                        sorted_sovs = {}
                        for team in sorted(sovs_dict, key = sovs_dict.get, reverse=True):
                            sorted_sovs[team] = sovs_dict[team]
                        team_1, team_2, team_3 = list(sorted_sovs)
                        team_1_sov, team_2_sov, team_3_sov = list(sorted_sovs.values())
                        if (team_1_sov == team_2_sov == team_3_sov) or (team_1_sov == team_2_sov > team_3_sov):
                            # In both scenarios, one team's worst finish is ordinal+1, but it's not possible to know without Total Game Victory Time.
                            # TGVT is only known once all games are played. As such, this script cannot determine who plays in the 1st of 2 tiebreaker games.
                            # There for, I'm just making their worst finishes be ordinal+2.
                            if team_2_sov == team_3_sov:
                                row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Unresolved", [True, True, True])
                            else:
                                row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Unresolved", [True, True, False])
                            teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                            teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                            teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                        else:
                            if team_1_sov > team_2_sov > team_3_sov:
                                row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Unresolved")
                            elif team_1_sov > team_2_sov == team_3_sov:
                                row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Unresolved", [False, True, True])
                            else:
                                print("Scenario 1 error 1")
                            teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                            teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                            teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                        teams_chances_tie[team_1][ordinal] += 1
                        teams_chances_tie[team_2][ordinal] += 1
                        teams_chances_tie[team_3][ordinal] += 1
                case (3, 2) if len(sorted_teams_no_aggs[0].split()) == 2: # Top 2 teams have the same aggregate. Scenario 3 in LCS Rulebook 8.6
                    team_1, team_2 = sorted_teams_no_aggs[0].split()
                    team_3 = sorted_teams_no_aggs[1]
                    if ordinal >= 6:
                        row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Resolved")
                        teams_chances_no_tie[team_1][ordinal] += 1
                        teams_chances_no_tie[team_2][ordinal] += 1
                        teams_chances_no_tie[team_3][ordinal+2] += 1
                    else:
                        tiebreaker_games += 1
                        team_1_sov, team_2_sov = strength_of_victory([team_1, team_2], h2h, sorted_teams)
                        if team_1_sov > team_2_sov:
                            row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], [None, None, "Locked"])
                        elif team_2_sov > team_1_sov:
                            row_data, col = append_row_data(row_data, col, [team_2, team_1, team_3], [None, None, "Locked"])
                        else:
                            row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], [None, None, "Locked"], [True, True, False])
                        teams_chances_tie[team_1][ordinal] += 1
                        teams_chances_tie[team_2][ordinal] += 1
                        teams_chances_no_tie[team_3][ordinal+2] += 1
                        teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                        teams_worst_finish_in_ties[team_2][ordinal+1] += 1
                case (3, 2) if len(sorted_teams_no_aggs[0].split()) == 1: # Bottom 2 teams have the same aggregate. Scenario 4 in LCS Rulebook 8.6
                    team_1 = sorted_teams_no_aggs[0]
                    team_2, team_3 = sorted_teams_no_aggs[1].split()
                    if ordinal >= 5:  #If the tie is for 6th or worse, the tie has no impact on seeding. Tie is considered resolved with no tiebreakers.
                        row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Resolved")
                        teams_chances_no_tie[team_1][ordinal] += 1
                        teams_chances_no_tie[team_2][ordinal+1] += 1
                        teams_chances_no_tie[team_3][ordinal+1] += 1
                    else:
                        tiebreaker_games += 1
                        team_2_sov, team_3_sov = strength_of_victory([team_2, team_3], h2h, sorted_teams)
                        if team_2_sov > team_3_sov:
                            row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], ["Locked", None, None])
                        elif team_3_sov > team_2_sov:
                            row_data, col = append_row_data(row_data, col, [team_1, team_3, team_2], ["Locked", None, None])
                        else:
                            row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], ["Locked", None, None], [False, True, True])
                        teams_chances_no_tie[team_1][ordinal] += 1
                        teams_chances_tie[team_2][ordinal+1] += 1
                        teams_chances_tie[team_3][ordinal+1] += 1
                        teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                        teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                case (3, 3) : # No ties in 3-way h2h. Scenario 2 or 5 in 8.6
                    team_1, team_2, team_3 = sorted_teams_no_aggs
                    if list(sorted_teams_aggs) == [3, 2, 1]: #Scenario 2 - For some reason, this is a tie requiring 2 tiebreaker games.
                        if ordinal >= 6:
                            row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Resolved")
                            teams_chances_no_tie[team_1][ordinal] += 1
                            teams_chances_no_tie[team_2][ordinal] += 1
                            teams_chances_no_tie[team_3][ordinal] += 1
                        else:
                            tiebreaker_games += 2
                            row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Unresolved")
                            teams_chances_tie[team_1][ordinal] += 1
                            teams_chances_tie[team_2][ordinal] += 1
                            teams_chances_tie[team_3][ordinal] += 1
                            teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                            teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                            teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                    elif list(sorted_teams_aggs) == [4, 2, 0]:
                        row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Resolved")
                        teams_chances_no_tie[team_1][ordinal] += 1
                        teams_chances_no_tie[team_2][ordinal+1] += 1
                        teams_chances_no_tie[team_3][ordinal+2] += 1
                    else:
                        print("Scenario 2/5 error")
        else: # 4 or more teams in a tie automatically goes to tiebreaker games, unless said tiebreakers don't effect post-split seeding. SOV is considered for side selection.
            if ordinal >= 6: 
                team_order = []
                print(teams)
                for team in teams.split():
                    teams_chances_no_tie[team][ordinal] += 1
                    team_order.append(team)
                row_data, col = append_row_data(row_data, col, team_order, "Resolved")
            else:
                teams_sovs = strength_of_victory(teams_in_ordinal, h2h, sorted_teams)
                sovs_dict = {}
                for team in teams_in_ordinal:
                    sovs_dict[team] = teams_sovs[teams_in_ordinal.index(team)]
                sorted_sovs = {}
                teams_in_tie = sorted(sovs_dict, key=sovs_dict.get, reverse=True) # Returns list of teams in SOV order descending, but with no grouping or SOV number. ex: ['TSM', 'EG', '100', 'C9']
                for team in teams_in_tie:
                    if sorted_sovs.get(sovs_dict.get(team)) == None:
                        sorted_sovs.update({sovs_dict.get(team): team})
                    else:
                        sorted_sovs.update({sovs_dict.get(team): sorted_sovs.get(sovs_dict.get(team)) + " " + team})
                sorted_sov_teams = list(sorted_sovs.values())
                if len(teams_in_tie) == len(sorted_sov_teams): # No SOV Ties
                    sov_ties = None
                else:
                    sov_ties = []
                    new_sov = False
                    for teams in sorted_sov_teams:
                        if len(teams.split()) == 1:
                            sov_ties.append(False)
                        else:
                            if not sov_ties:
                                pass
                            elif sov_ties[-1] is True:
                                new_sov = True
                            for team in teams.split():
                                if new_sov:
                                    sov_ties.append("New")
                                    new_sov = False
                                else:
                                    sov_ties.append(True)
                row_data, col = append_row_data(row_data, col, teams_in_tie, "Unresolved", sov_ties)
                match (len(teams_in_ordinal), ordinal):
                    case (4, 0|1|2|3):   # 4 way tie for 1st, 2nd, 3rd, and 4th have 4 tiebreaker games.
                        tiebreaker_games += 4
                    case (4, 4|5):       # 4 way tie for 5th and 6th have 3 tiebreaker games
                        tiebreaker_games += 3
                    case (5, 0|1|2|3):   # 5 way tie for 1st, 2nd, 3rd, and 4th have 5 tiebreaker games.
                        tiebreaker_games += 5
                    case (5, 4|5):       # 5 way tie for 5th and 6th have 4 tiebreaker games.
                        tiebreaker_games += 4
                    case (6, 0|1|2|3):   # 6 way tie for 1st and 2nd have 6-7 tiebreaker games. 3rd and 4th have 6 tiebreaker games. If H2H doesnt matter, then 1st and 2nd are 7 games. Unlikely to find out since 6-way-ties are near-impossible at EoS.
                        tiebreaker_games += 6
                    case (6, 4):         # 6 way tie for 5th has 5 tiebreaker games.
                        tiebreaker_games += 5
                    case (7, _):         # 7 way tie for 1st and 2nd have 7-9 tiebreaker games. 3rd and 4th have 7 games. If H2H doesnt matter, then 1st and 2nd are 9 games.
                        tiebreaker_games += 7
                    case (8, 0|1):       # 8 way tie for 1st and 2nd have 11 tiebreaker gamesgames. 
                        tiebreaker_games += 11
                    case (8, 2):         # 8 way tie for 3rd has 8 tiebreaker games.
                        tiebreaker_games += 8
                    case (9, _):         # 9 way ties have 12 tiebreaker games.
                        tiebreaker_games += 12
                    case (10, _):        # 10 way ties have 14 tiebreaker games.
                        tiebreaker_games += 14
                for team in teams_in_ordinal:
                    teams_chances_tie[team][ordinal] += 1
                    teams_worst_finish_in_ties[team][ordinal + len(teams_in_ordinal)-1] += 1
        ordinal += len(teams_in_ordinal)
    row_data.append([col, tiebreaker_games, None])
    worksheet_data_to_write[row] = row_data

for row in worksheet_data_to_write:
    row_data_to_write = worksheet_data_to_write[row]
    for data in row_data_to_write:
        col, writables, cell_format = data
        worksheet.write(row, col, writables, cell_format)

workbook.close()

no_tie_output, tie_output, worst_output = "", "", ""
for team in standings:
    no_tie_output += f"{team}: {teams_chances_no_tie[team]}\n"
    tie_output += f"{team}: {teams_chances_tie[team]}\n"
    worst_output += f"{team}: {teams_worst_finish_in_ties[team]}\n"
print(f"No ties\n{no_tie_output}")
print(f"Ties\n{tie_output}")
print(f"Worst Finish\n{worst_output}")

# Spring 2022 Tiebreaker rules
# All ties: If a tiebreaker match would not have an effect on post-split seeding, it is not played.

#region Ties Explanations
#  3 way tie:  0 games - Team aggregates are 4-0, 2-2, 0-4
#             1 game - Team aggregates are 3-1, 3-1, 0-4 or 4-0, 1-3, 1-3
#             2 games - Team aggregates are 2-2, 2-2, 2-2 or 3-1, 2-2, 1-3
#  4 way tie: Teams are drawn into 2 "1st round matches". 
#             Losers play for Bottom 2 Seeds
#             Winners play for Top 2 seeds. 
#             Max: 4 games
#  5 way tie: 1 play-in game between 2 lowest SoV Teams.
#             Loser gets lowest seed
#             Winner + 3 remaining teams go to 4-way-tie procedure for highest seed
#             Max: 5 games
#  6 way tie: 2 randomly drawn play-in games between 4 lowest SoV Teams
#             Losers go to 2-way-tie procedure for 2nd lowest seed (Look at H2H?)
#             Winners + 2 remaining teams go to 4-way-tie procedure for highest seed.
#             Max: 7 games
#  7 way tie: 3 randomly drawn play-in games between 6 lowest SoV Teams
#             Losers go to 3-way-tie procedure for 3rd lowest seed (Look at H2H?)
#             Winners + remaining team go to 4-way-tie procedure for highest seed
#             Max: 9 games
#  8 way tie: 4 randomly drawn play-in games between all teams.
#             Losers go to 4-way-tie proceedure for 4th lowest seed
#             Winners go to 4-way-tie procedure for 4th highest seed.
#             Max: 12 games
#  9 way tie: 1 play-in game between 2 lowest SoV Teams
#             Loser gets lowets seed
#             Winner + 7 remaining teams go to 8-way-tie procedure
#             Max: 13 games
# 10 way tie: 2 play-in games between 4 lowest SoV teamms
#             Losers go to 2-way-tie procedure for 9th (Look at H2H?)
#             Winner + 7 remaining teams to go to 8-way-tie procedure for 1st
#             Max: 15 games
#endregion Ties Explanations

# 4+ way ties
# Ties for 1st: 
# 4 way tie (1st-4th): 2 play-in games. Losers play for 3rd seed. Winners play for 1st seed. 4 games.
# 5 way tie (1st-5th): 1 play-in game. Loser is considered 5th seed. Winner + 3 remaining teams go to 4-way-tie for 1st. 5 games.
# 6 way tie (1st-6th): 2 play-in games. Losers go to 2-way-tie for 5th. Winners + 2 remaining teams go to 4-way-tie for 1st. 6-7 games.
# 7 way tie (1st-7th): 3 play-in games. Losers go to 3-way-tie for 5th-7th. Winners + 1 remaining team go to 4 way tie for 1st. 7-9 games
# 8 way tie (1st-8th): 4 play-in games. Losers go to 4-way-tie for 5th-8th. Winners go to 4-way-tie for 1st. 11 games.
# 9 way tie (1st-9th): 1 play-in game. Loser is considered 9th seed. Winners go to 8-way-tie for 1st. 12 games.
# 10 way tie (1st-10th): 2 play-in games. Losers go to 2-way-tie for 9th, no games. Winners go to 8-way-tie for 1st. 14 games.

# Ties for 2nd:
# 4 way tie (2nd-5th): 2 play-in games. Losers play for 4th seed. Winners play for 2nd seed. 4 games.
# 5 way tie (2nd-6th): 1 play-in game. Loser is considered 6th seed. Winner + 3 remaining teams go to 4-way-tie for 2nd. 5 games.
# 6 way tie (2nd-7th): 2 play-in games. Losers go to 2-way-tie for 6th. Winners + 2 remaining teams go to 4-way-tie for 2nd. 6-7 games.
# 7 way tie (2nd-8th): 3 play-in games. Losers go to 3-way-tie for 6th. Winners + remaining team go to 4-way-tie for 2nd.  7-9 games.
# 8 way tie (2nd-9th): 4 play-in games. Losers go to 4-way-tie for 6th. Winners go to 4-way-tie for 2nd. 11 games.
# 9 way tie (2nd-10th): 1 play-in games. Loser is considered 10th seed. Winner + 7 remaining teams go to 8-way-tie procedure for 2nd. 12 games.

# Ties for 3rd:
# 4 way tie (3rd-6th): 2 play-in games. Losers play for 5th seed. Winners play for 3rd seed. 4 games.
# 5 way tie (3rd-7th): 1 play-in game. Losers is considered 7th seed. Winners + 3 remaining teams go to 4-way-tie for 3rd. 5 games.
# 6 way tie (3rd-8th): 2 play-in games. Losers go to 2-way-tie for 7th, no games. Winners + 2 remaining teams go to 4-way-tie for 3rd. 6 games.
# 7 way tie (3rd-9th): 3 play-in games. Losers go to 3-way-tie procedure for 7th, no games. Winners + remaining team go to 4-way-tie for 3rd. 7 games.
# 8 way tie (3rd-10th): 4 play-in games. Losers go to 4-way-tie procedure for 7th, no games. Winners go to 4-way-tie for 3rd. 8 games.

# Ties for 4th:
# 4 way tie (4th-7th): 2 play-in games. Losers play for 6th seed. Winners play for 4th. 4 games.
# 5 way tie (4th-8th): 1 play-in game. Loser is considered 8th seed. Winner + 3 remaining teams go to 4-way-tie for 4th. 5 games
# 6 way tie (4th-9th): 2 play-in games. Losers are considered tied for 8th, no game. Winners + 2 remaining teams go to 4-way-tie for 4th. 6 games.
# 7 way tie (4th-10th): 3 play-in games. Losers go to 3-way-tie procedure for 8th, no games. Winners + remaining team go to 4-way-tie for 4th. 7 games.

# Ties for 5th:
# 4 way tie (5th-8th): 2 play-in games. Losers are considered tied for 7th, no game. Winners play for 5th and 6th. 3 games.
# 5 way tie (5th-9th): 1 play-in game. Loser is considered 9th seed. Winner + 3 remaining teams go to 4-way-tie for 5th. 4 games.
# 6 way tie (5th-10th): 2 play-in games. Losers are considered tied for 9th, no game. Winners + 2 remaining teams go to 4-way-tie for 5th. 5 games.

# Ties for 6th:
# 4 way tie (6th-9th): 2 play-in games. Losers are considered tied for 8th, no game. Winners play for 6th and 7th. 3 games.
# 5 way tie (6th-10th): 1 play-in game. Loser is considered 10th seed. Winner + 3 remaining teams go to a 4-way-tie for 6th. 4 games.