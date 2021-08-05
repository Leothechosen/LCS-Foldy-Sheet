#c:\Users\Leosm\Downloads\pypy3.7-v7.3.5-win64\pypy3.exe "c:/DiscordBots/Expirements/LoL Scenarios/LCS-Foldy-Sheet/LCS/LCS_Summer2021_Shorter.py"

import xlsxwriter
import itertools
import copy
import timeit

workbook = xlsxwriter.Workbook('C:/DiscordBots/Expirements/LoL Scenarios/LCS-Foldy-Sheet/LCS/LCS_Scenarios_Summer2021_Testing.xlsx')
worksheet = workbook.add_worksheet()

two_way_tie_unresolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'red'})
two_way_tie_unresolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'red'})
two_way_tie_unresolved_start_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': 'red', 'italic': True})
two_way_tie_unresolved_end_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': 'red', 'italic': True})

two_way_tie_resolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left': 2, 'bg_color': '#FFCCCB'})
two_way_tie_resolved_end = workbook.add_format({'bottom': 2, 'top': 2, 'right': 2, 'bg_color': '#FFCCCB'})

Multiway_tie_unresolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left' : 2, 'bg_color': 'lime'})
Multiway_tie_unresolved_middle = workbook.add_format({'bottom': 2, 'top' : 2, 'bg_color': 'lime'})
Multiway_tie_unresolved_end = workbook.add_format({'bottom' : 2, 'top' : 2, 'right': 2, 'bg_color': 'lime'})

Multiway_tie_unresolved_start_tied_SOV = workbook.add_format({'bottom': 2, 'top': 2, 'left' : 2, 'bg_color': 'lime', 'italic': True})
Multiway_tie_unresolved_middle_tied_SOV = workbook.add_format({'bottom': 2, 'top' : 2, 'bg_color': 'lime', 'italic': True})
Multiway_tie_unresolved_middle_new_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'left': 1, 'bg_color': 'lime', 'italic': True})
Multiway_tie_unresolved_end_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'right': 2, 'bg_color': 'lime', 'italic': True})

Multiway_tie_partially_resolved_start = workbook.add_format({'bottom' : 2, 'top' : 2, 'left' : 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_middle = workbook.add_format({'bottom' : 2, 'top' : 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_end = workbook.add_format({'bottom' : 2, 'top' : 2, 'right' : 2, 'bg_color': '#00FFFF'})
Multiway_tie_partially_resolved_start_locked = workbook.add_format({'bottom' : 2, 'top' : 2, 'left' : 2, 'bg_color': '#00FFFF', 'bold': True})
Multiway_tie_partially_resolved_middle_locked = workbook.add_format({'bottom' : 2, 'top' : 2, 'bg_color': '#00FFFF', 'bold': True})
Multiway_tie_partially_resolved_end_locked = workbook.add_format({'bottom' : 2, 'top' : 2, 'right' : 2, 'bg_color': '#00FFFF', 'bold': True})

Multiway_tie_partially_resolved_start_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'left' : 2, 'bg_color': '#00FFFF', 'italic': True})
Multiway_tie_partially_resolved_middle_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'bg_color': '#00FFFF', 'italic': True})
Multiway_tie_partially_resolved_end_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'right' : 2, 'bg_color': '#00FFFF', 'italic': True})

Multiway_tie_resolved_start = workbook.add_format({'bottom': 2, 'top': 2, 'left' : 2, 'bg_color': 'yellow'})
Multiway_tie_resolved_middle = workbook.add_format({'bottom': 2, 'top' : 2, 'bg_color': 'yellow'})
Multiway_tie_resolved_end = workbook.add_format({'bottom' : 2, 'top' : 2, 'right': 2, 'bg_color': 'yellow'})

def Strength_of_victory(tied_teams, teams_wins, sorted_teams_no_WL):
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
    for teams in sorted_teams_no_WL: # Assigns each team a set SoV points for where they placed in the standings. 
    #ex: {'100': 5.0, 'C9': 4.5, 'CLG': 4.0, 'DIG': 3.5, 'EG': 3.0, 'FLY': 2.5, 'GG': 2.0, 'IMT': 2.0, 'TL': 1.0, 'TSM': 0.5}
        teams = teams.split()
        for team in teams:
            teams_sov_points[team] = sov_points[ordinal] 
        ordinal += len(teams)
    teams_h2h_order = ["100", "C9", "CLG", "DIG", "EG", "FLY", "GG", "IMT", "TL", "TSM"]
    tied_teams_sov = []
    for team in tied_teams: #Calculates each tied team's total SoV points and puts them in a list in the same order as tied_teams
        team_sov = 0
        team_h2h = teams_wins[team]
        teams_h2h_index = 0
        for wins in team_h2h: # wins can be a single instance of 0 through 5 
            if wins is not None:
                team_sov += (wins * teams_sov_points[teams_h2h_order[teams_h2h_index]])
            teams_h2h_index += 1
        tied_teams_sov.append(team_sov)
    return tied_teams_sov

def append_row_data(row_data, col, teams, ties=None, sov_ties=None, scenario_num=None):
    # col - int:column num
    # teams - list:teams to write
    # ties - None, str("Resolved" or "Unresolved"), or list["Locked", None, None]
    # sov_ties - None or list, Example [False, True, True]. If more than 1 sov tie, [True, True, "New", True]
    # scenario_num - Debugging purposes
    if len(teams) == 1:
        row_data.append([col, teams[0], None])
        col += 1
    elif len(teams) == 2:
        team_1, team_2 = teams
        if ties == "Resolved":
            row_data.append([col, team_1, two_way_tie_resolved_start])
            col += 1
            row_data.append([col, team_2, two_way_tie_resolved_end])
        elif sov_ties == None:
            row_data.append([col, team_1, two_way_tie_unresolved_start])
            col += 1
            row_data.append([col, team_2, two_way_tie_unresolved_end])
        else:
            row_data.append([col, team_1, two_way_tie_unresolved_start_tied_SOV])
            col += 1
            row_data.append([col, team_2, two_way_tie_unresolved_end_tied_SOV])
        col += 1
    else:
        for team in teams:
            tie = ties[teams.index(team)] if type(ties) == list else ties
            sov_tie = sov_ties[teams.index(team)] if type(sov_ties) == list else None
            if tie == "Resolved":
                if teams.index(team) == 0: #If first team in tie
                    row_data.append([col, team, Multiway_tie_resolved_start])
                elif teams.index(team) == len(teams) - 1: #If last team in the tie
                    row_data.append([col, team, Multiway_tie_resolved_end])
                else:
                    row_data.append([col, team, Multiway_tie_resolved_middle])
            elif tie == "Unresolved":
                if teams.index(team) == 0:
                    if sov_tie is True:
                        row_data.append([col, team, Multiway_tie_unresolved_start_tied_SOV])
                    else:
                        row_data.append([col, team, Multiway_tie_unresolved_start])
                elif teams.index(team) == len(teams) - 1:
                    if sov_tie is True:
                        row_data.append([col, team, Multiway_tie_unresolved_end_tied_SOV])
                    else:
                        row_data.append([col, team, Multiway_tie_unresolved_end])
                else:
                    if sov_tie is True:
                        row_data.append([col, team, Multiway_tie_unresolved_middle_tied_SOV])
                    elif sov_tie == "New":
                        row_data.append([col, team, Multiway_tie_unresolved_middle_new_tied_SOV])
                    else:
                        row_data.append([col, team, Multiway_tie_unresolved_middle])
            else: #When the tie is partially resolved
                if teams.index(team) == 0:
                    if tie == "Locked":
                        row_data.append([col, team, Multiway_tie_partially_resolved_start_locked])
                    elif sov_tie is True:
                        row_data.append([col, team, Multiway_tie_partially_resolved_start_tied_SOV])
                    else:
                        row_data.append([col, team, Multiway_tie_partially_resolved_start])
                elif teams.index(team) == len(teams) - 1:
                    if tie == "Locked":
                        row_data.append([col, team, Multiway_tie_partially_resolved_end_locked])
                    elif sov_tie is True:
                        row_data.append([col, team, Multiway_tie_partially_resolved_end_tied_SOV])
                    else:
                        row_data.append([col, team, Multiway_tie_partially_resolved_end])
                else:
                    if tie == "Locked":
                        row_data.append([col, team, Multiway_tie_partially_resolved_middle_locked])
                    elif sov_tie is True:
                        row_data.append([col, team, Multiway_tie_partially_resolved_middle_tied_SOV])
                    elif sov_tie == "New":
                        row_data.append([col, team, Multiway_tie_partially_resolved_middle_new_tied_SOV])
                    else:
                        row_data.append([col, team, Multiway_tie_partially_resolved_middle])
            col += 1
    return row_data, col

matches = [
    ["TSM", "C9"],
    ["EG", "CLG"],
    ["FLY", "DIG"],
    ["TSM", "100"],
    ["EG", "C9"],
    ["TL", "IMT"],
    ["FLY", "GG"],
    ["DIG", "CLG"],
    ["100", "GG"],
    ["TL", "EG"],
    ["C9", "DIG"],
    ["IMT", "FLY"],
    ["TSM", "CLG"],
    ["DIG", "TSM"],
    ["IMT", "GG"],
    ["TL", "C9"],
    ["100", "EG"],
    ["FLY", "CLG"]
]

ties = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0] #The first 0 will always be 0, since there's no such thing as 1-way-ties
teams_chances_no_tie = {team:[0]*10 for team in ["C9", "DIG", "TSM", "100", "TL", "EG", "IMT", "FLY", "CLG", "GG"]}

teams_chances_tie = {team:[0]*10 for team in ["C9", "DIG", "TSM", "100", "TL", "EG", "IMT", "FLY", "CLG", "GG"]}

teams_worst_finish_in_ties = {team:[0]*10 for team in ["C9", "DIG", "TSM", "100", "TL", "EG", "IMT", "FLY", "CLG", "GG"]}

# In cases where there is a multiway tie for a place where not all the TB games need to be placed, and SoV is needed to determine tiebreaker order, 
# if some or all SoVs are equal, it's not known to this script if a team will need to play a tiebreaker game
# As such, teams_chances_unknown lists where a team could potentially be playing for with a tb game, but it's not guaranteed.
teams_chances_unknown = {team:[0]*10 for team in ["C9", "DIG", "TSM", "100", "TL", "EG", "IMT", "FLY", "CLG", "GG"]}

start = timeit.default_timer()

outcomes = itertools.product(*matches)
outcomes = zip(range(2**len(matches)), outcomes)
worksheet_data_to_write = {}
for scenario in outcomes:
    row_data = [] # column, data, format
    row = scenario[0]
    winners = scenario[1]
    tiebreaker_games = 0
    scenario_num = row+1
    row_data.append([0, row+1, None])
    col = 1
    teams_standings = { #The order of this list doesn't matter. I like ordering it by how the standings are though.
        "TSM": [27, 14],
        "100": [28, 14],
        "EG":  [25, 16],
        "C9":  [26, 15],
        "TL":  [25, 17],
        "DIG": [20, 21],
        "IMT": [20, 22],
        "FLY": [13, 28],
        "GG":  [12, 30],
        "CLG": [11, 30]
    }

    teams_combined_wins = { # 100, C9, CLG, DIG, EG, FLY, GG, IMT, TL, TSM
        "100": [None, 2, 4, 4, 2, 5, 2, 3, 4, 2],
        "C9":  [3, None, 4, 4, 1, 4, 4, 3, 2, 1],
        "CLG": [1, 1, None, 2, 1, 0, 3, 1, 1, 1],
        "DIG": [1, 0, 2, None, 3, 4, 3, 3, 1, 3],
        "EG":  [2, 3, 3, 2, None, 4, 4, 2, 2, 3],
        "FLY": [0, 1, 4, 0, 1, None, 2, 1, 0, 4],
        "GG":  [2, 1, 2, 2, 1, 2, None, 1, 1, 0],
        "IMT": [2, 2, 4, 2, 3, 3, 3, None, 1, 0],
        "TL":  [1, 2, 4, 4, 2, 5, 4, 3, None, 0],
        "TSM": [2, 3, 3, 1, 2, 1, 5, 5, 5, None]  
    }
    match_num = 0
    for winner in winners:
        teams_standings[winner][0] += 1
        if winner == matches[match_num][0]:
            loser = matches[match_num][1]
        else:
            loser = matches[match_num][0]
        match_num += 1
        teams_combined_wins[winner][list(teams_combined_wins).index(loser)] += 1
        teams_standings[loser][1] += 1
        row_data.append([col, winner, None])
        col += 1
    sorted_teams = {}
    for k in sorted(teams_standings, key=lambda k: (-teams_standings[k][0], teams_standings[k][1]), reverse=False):  # k = team. Sorts the teams dict by Wins descending
        if sorted_teams.get(str(teams_standings.get(k))) == None:
            sorted_teams.update({str(teams_standings.get(k)): k})
        else:
            sorted_teams.update({str(teams_standings.get(k)): sorted_teams.get(str(teams_standings.get(k))) + " " + k})
    sorted_teams_no_WL = list(sorted_teams.values()) # Assigns just the teams in order (without their W-L) to values
    col += 1
    for teams in sorted_teams_no_WL: # Sends either 1 team if that team isn't tied in W-L, or a batch of teams tied in W-L. Tiebreaker logic
        teams_in_ordinal = teams.split()
        if len(teams_in_ordinal) != 1:
            ties[len(teams_in_ordinal)-1] += 1
        first_team_in_ordinal = teams_in_ordinal[0]
        teams_already_processed = (' '.join(sorted_teams_no_WL[0:(sorted_teams_no_WL.index(teams)+1)]))
        teams_already_processed_list = teams_already_processed.split()
        ordinal = teams_already_processed_list.index(first_team_in_ordinal)
        if len(teams_in_ordinal) == 1:
            row_data, col = append_row_data(row_data, col, [teams_in_ordinal[0]])
            teams_chances_no_tie[teams_in_ordinal[0]][ordinal] += 1
        elif len(teams_in_ordinal) == 2 or len(teams_in_ordinal) == 3: #Only 2 and 3 way ties are actually attempted to be resolved.
            teams_aggs = {}
            for team in teams_in_ordinal:
                team_agg = 0
                other_teams_in_ordinal = copy.deepcopy(teams_in_ordinal)
                other_teams_in_ordinal.remove(team)
                for other_team in other_teams_in_ordinal:
                    team_agg += teams_combined_wins[team][list(teams_combined_wins).index(other_team)]
                teams_aggs.update({team: team_agg})
            sorted_teams_aggs = {}
            for k in sorted(teams_aggs, key=teams_aggs.get, reverse=True):
                if sorted_teams_aggs.get(str(teams_aggs.get(k))) == None:
                    sorted_teams_aggs.update({str(teams_aggs.get(k)): k})
                else:
                    sorted_teams_aggs.update({str(teams_aggs.get(k)): sorted_teams_aggs.get(str(teams_aggs.get(k))) + " " + k})
            sorted_teams_no_aggs = list(sorted_teams_aggs.values())
            if len(teams_in_ordinal) == 2 and len(sorted_teams_no_aggs) == 2: #No ties in 2-way h2h.
                team_1, team_2 = sorted_teams_no_aggs
                row_data, col = append_row_data(row_data, col, [team_1, team_2], "Resolved")
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal+1] += 1
            elif len(teams_in_ordinal) == 3 and len(sorted_teams_no_aggs) == 3: #No ties in 3-way h2h.
                team_1, team_2, team_3 = sorted_teams_no_aggs
                row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Resolved")
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal+1] += 1
                teams_chances_no_tie[team_3][ordinal+2] += 1
            else: #If ties still exist after H2H
                if len(teams_in_ordinal) == 2: # 2-way-tie in H2H. Not actually possible in Summer 2021.
                    tiebreaker_games += 1
                    team_1, team_2 = teams_in_ordinal
                    teams_sov == Strength_of_victory([team_1, team_2], teams_combined_wins, sorted_teams_no_WL)
                    team_1_sov, team_2_sov = teams_sovs
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
                else: # 3-way-tie not fully resolved by H2H
                    if len(sorted_teams_no_aggs) == 1: #Means that all 3 teams are tied in H2H
                        tiebreaker_games += 2
                        team_1, team_2, team_3 = sorted_teams_no_aggs[0].split()
                        teams_sovs = Strength_of_victory([team_1, team_2, team_3], teams_combined_wins, sorted_teams_no_WL)
                        teams_sov_dict = {team_1: teams_sovs[0], team_2: teams_sovs[1], team_3: teams_sovs[2]}
                        sorted_teams_sov_dict = {}
                        for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                            sorted_teams_sov_dict[team] = teams_sov_dict[team]
                        team_1, team_2, team_3 = list(sorted_teams_sov_dict)
                        team_1_sov, team_2_sov, team_3_sov = sorted_teams_sov_dict.values()
                        if team_1_sov == team_2_sov == team_3_sov or team_1_sov == team_2_sov > team_3_sov: 
                            # In both sceanrios, One team's worst finish is ordinal+1, but it's not possible to determine before all games are played. 
                            # Therefore, make their worst finishes ordinal+2
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
                            teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                            teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                            teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                        teams_chances_tie[team_1][ordinal] += 1
                        teams_chances_tie[team_2][ordinal] += 1
                        teams_chances_tie[team_3][ordinal] += 1
                    else: #Means that 1 team is able to be seed while the other 2 are still tied. The 2 tied go to a tiebreaker game, with side selection given to the favored H2H team.
                        if len(sorted_teams_no_aggs[0].split()) == 1: #Bottom 2 teams have the same aggregate
                            team_1 = sorted_teams_no_aggs[0]
                            team_2, team_3 = sorted_teams_no_aggs[1].split()
                            if ordinal == 7: #If the 3-way-tie is for 8th, the bottom two teams do not have a tiebreaker game.
                                row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], "Resolved")
                                teams_chances_no_tie[team_1][ordinal] += 1
                                teams_chances_no_tie[team_2][ordinal+1] += 1
                                teams_chances_no_tie[team_3][ordinal+1] += 1
                            else:
                                tiebreaker_games += 1
                                team_2_agg = teams_combined_wins[team_2][list(teams_combined_wins).index(team_3)]
                                team_3_agg = teams_combined_wins[team_3][list(teams_combined_wins).index(team_2)]
                                if team_2_agg > team_3_agg:
                                    row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], ["Locked", None, None])
                                elif team_3_agg > team_2_agg:
                                    row_data, col = append_row_data(row_data, col, [team_1, team_3, team_2], ["Locked", None, None])
                                else:
                                    pass
                                teams_chances_no_tie[team_1][ordinal] += 1
                                teams_chances_tie[team_2][ordinal+1] += 1
                                teams_chances_tie[team_3][ordinal+1] += 1
                                teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                                teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                        else: #Top 2 teams have the same aggregate
                            tiebreaker_games += 1
                            team_1, team_2 = sorted_teams_no_aggs[0].split()
                            team_3 = sorted_teams_no_aggs[1]
                            team_1_agg = teams_combined_wins[team_1][list(teams_combined_wins).index(team_2)]
                            team_2_agg = teams_combined_wins[team_2][list(teams_combined_wins).index(team_1)]
                            if team_1_agg > team_2_agg:
                                row_data, col = append_row_data(row_data, col, [team_1, team_2, team_3], [None, None, "Locked"])
                            elif team_2_agg > team_1_agg:
                                row_data, col = append_row_data(row_data, col, [team_2, team_1, team_3], [None, None, "Locked"])
                            else:
                                pass
                            teams_chances_tie[team_1][ordinal] += 1
                            teams_chances_tie[team_2][ordinal] += 1
                            teams_chances_no_tie[team_3][ordinal+2] += 1
                            teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                            teams_worst_finish_in_ties[team_2][ordinal+1] += 1                      
        else: #When teams_in_ordinal is greater than 3, they automatically go to tiebreakers. Sort by SOV.
            if len(teams_in_ordinal) == 4 and ordinal == 2: #If teams are playing for 3rd, there are no tiebreaker game and they're all considered 3rd seed. - Special Case
                row_data, col = append_row_data(row_data, col, teams_in_ordinal, "Resolved")
                for team in teams_in_ordinal:
                    teams_chances_no_tie[team][ordinal] += 1
                continue
            teams_sovs = Strength_of_victory(teams_in_ordinal, teams_combined_wins, sorted_teams_no_WL)
            sovs_dict = {}
            for team in teams_in_ordinal:
                sovs_dict[team] = teams_sovs[teams_in_ordinal.index(team)]
            sorted_sovs = {}
            teams_in_tie = sorted(sovs_dict, key=sovs_dict.get, reverse=True) # Returns list of teams in SOV order descending, but with no grouping or SOV number. ex: ['TSM', 'EG', '100', 'C9']
            for k in sorted(sovs_dict, key=sovs_dict.get, reverse=True): #Returns dict of SOV nums descending with corresponding teams. Teams with same SOV num are grouped. ex: {'83.0': 'TSM EG', '73.5': '100', '73.0': 'C9'}
                if sorted_sovs.get(str(sovs_dict.get(k))) == None:
                    sorted_sovs.update({str(sovs_dict.get(k)): k})
                else:
                    sorted_sovs.update({str(sovs_dict.get(k)): sorted_sovs.get(str(sovs_dict.get(k))) + " " + k})
            sorted_sov_teams = list(sorted_sovs.values()) # Returns list of teams in descending SOV number. No numbers. Teams with same SOV num are grouped. ex: ["TSM EG", "100", "C9"]
            if len(teams_in_tie) == len(sorted_sov_teams): # No SOV Ties
                sov_ties = None
            else: # Creates list:sov_ties that contain whether or not teams are in an SOV tie. Passed to append_row_data
                sov_ties = []
                new_sov = False
                for team in sorted_sov_teams:
                    if len(teams.split()) == 1:
                        sov_ties.append(False)
                    else:
                        if sov_ties == []:
                            pass
                        elif sov_ties[-1] is True:
                            new_sov = True
                        for team in teams.split():
                            if new_sov is True:
                                sov_ties.append("New")
                                new_sov = False
                            else:
                                sov_ties.append(True)
            if len(teams_in_ordinal) == 5 and ordinal == 2: #5-way-tie for 3rd - Special Case
                if sov_ties is None: #No SOV tie, means that there is a definite bottom 2. Top 3 teams are locked in 3rd.
                    print(f"5 way tie for 3rd with no tied SOVs - Scenario {scenario_num} - Teams {teams_in_tie} - Sorted SOV Teams - {sorted_sov_teams}")
                elif (len(sorted_sov_teams[-1].split()) == 2) or (len(sorted_sov_teams[-1].split()) == 1 and len(sorted_sov_teams[-2].split()) == 1): #Means that there is a definite bottom 2. Top 3 teams being locked in for 3rd.
                    print(f"5 way tie for 3rd with tied SOVs - Scenario {scenario_num} - Teams {teams_in_tie} - Sorted SOV Teams - {sorted_sov_teams}")
                else:
                    print(f"5-way-tie for 3rd but something is wrong - Scenario {scenario_num} - Teams {teams_in_tie} - Sorted SOV Teams - {sorted_sov_teams}")
                row_data, col = append_row_data(row_data, col, sorted_sov_teams, ["Locked", "Locked", "Locked", None, None], sov_ties)
                for team in sorted_sov_teams[:3]:
                    teams_chances_no_tie[team][ordinal] += 1
                for team in sorted_sov_teams[3:]:
                    teams_chances_tie[team][ordinal] += 1
                    teams_worst_finish_in_ties[team][ordinal+4] += 1
                continue
            elif len(teams_in_ordinal) == 6 and ordinal == 2: #6-way-tie for 3rd - Special Case
                if sov_ties is None: #No SOV tie, means that there is a definite Top 2. Top 2 are locked in for 3rd
                    print(f"6 way tie for 3rd with no tied SOVs - Scenario {scenario_num} - Teams {teams_in_tie} - Sorted SOV Teams - {sorted_sov_teams}")
                elif (len(sorted_sov_teams[0].split()) == 2) or (len(sorted_sov_teams[0].split()) == 1 and len(sorted_sov_teams[1].split()) == 1):
                    print(f"6 way tie for 3rd with tied SOVs - Scenario {scenario_num} - Teams {teams_in_tie} - Sorted SOV Teams - {sorted_sov_teams}")
                else:
                    print(f"6-way-tie for 3rd but something is wrong - Scenario {scenario_num} - Teams {teams_in_tie} - Sorted SOV Teams - {sorted_sov_teams}")
                row_data, col = append_row_data(row_data, col, sorted_sov_teams, ["Locked", "Locked", None, None, None, None], sov_ties)
                teams_chances_no_tie[sorted_sov_teams[0]][ordinal] += 1
                teams_chances_no_tie[sorted_sov_teams[1]][ordinal] += 1
                for team in sorted_sov_teams[2:]:
                    teams_chances_tie[team][ordinal] += 1
                    teams_worst_finish_in_ties[team][ordinal+5] += 1
                continue
            else: #There are more special cases, however there's never been a 6-way-tie possible in the last 15 matches of the LCS, which is when my foldy sheet is posted, so no real need to write it out.
                row_data, col = append_row_data(row_data, col, teams_in_tie, "Unresolved", sov_ties)
            if len(teams_in_ordinal) == 4:
                if ordinal == 2:
                    pass
                elif ordinal == 6:
                    tiebreaker_games += 3
                else:
                    tiebreaker_games += 4
            elif len(teams_in_ordinal) == 5:
                if ordinal == 2:
                    tiebreaker_games += 1
                else:
                    tiebreaker_games += 5
            elif len(teams_in_ordinal) == 6:
                if ordinal == 2:
                    tiebreaker_games += 2
                else:
                    tiebreaker_games += 6
            elif len(teams_in_ordinal) == 7:
                if ordinal == 2:
                    tiebreaker_games += 3
                else:
                    tiebreaker_games += 7
            elif len(teams_in_ordinal) == 8:
                if ordinal == 2:
                    tiebreaker_games += 3
                else:
                    tiebrekaer_games += 7
            elif len(teams_in_ordinal) == 9:
                tiebreaker_games += 13
            else:
                tiebreaker_games += 14
            for team in teams_in_ordinal:
                teams_chances_tie[team][ordinal] += 1
                teams_worst_finish_in_ties[team][ordinal + len(teams_in_ordinal)-1] += 1
    row_data.append([col, tiebreaker_games, None])
    worksheet_data_to_write[row] = row_data

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
print(ties)
stop = timeit.default_timer()

no_tie_output, tie_output, worst_output = "", "", ""
for team in teams_standings:
    no_tie_output += f"{team}: {teams_chances_no_tie[team]}\n"
    tie_output += f"{team}: {teams_chances_tie[team]}\n"
    worst_output += f"{team}: {teams_worst_finish_in_ties[team]}\n"
print(f"No ties\n{no_tie_output}")
print(f"Ties\n{tie_output}")
print(f"Worst Finish\n{worst_output}")

print(f"\nScenarios generation time: {scenarios_stop - start}")
print(f"Worksheet write time: {ws_stop - ws_start}")
print(f"Worksheet close time: {ws_close_stop - ws_close_start}")
print(f"Total time: {stop - start}")