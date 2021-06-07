import xlsxwriter
import timeit
from tqdm import tqdm
import tqdm.contrib.itertools
import itertools

# --Notes--
# This sheet calculates remaining scenarios in the LCS, specifically Summer 2021. 
# This sheet figures out what ties will exist in each scenario, from 2 way ties to 10 way ties.
# This sheet puts tied teams in Strength of Victory order when it's needed for both side selection, and seeding play-in matches. 7 to 10 way ties do not calculate SoVs at this time, as those are very rare,
# and there's never been an instance where in the final 20 matches, a 7 to 10 way tie is possible.

# With records from Spring being carried over, it likely reduces chances of tiebreakers.
# 2 way ties: Unlike in Spring, unresolved 2 way ties are impossible, since each team will have 5 games against each team. 2 from Spring, 3 from Summer.
# 3 way ties use the aggregate wins against other teams in the tie. If those are tied, go to tiebreaker games.
# 4+ way ties are the same as spring, automatically go to tiebreaker games with no considerations except maybe SoV

# --Summer Tiebreaker Rules--
# All ties: If a 4 way tie for 3rd exists, it's auto-resolved with no tiebreaker games. This extends to any tie that moves 4 teams to the 4-way-tie procedure.

# 2 way tie: Tiebreaker games are impossible. These will always resolve via H2H. 
# 3 way tie: Check Combined Wins of teams. If there are ties there, go to games. 0-2 additional games possible.
#                    0 games: All teams have different Combined Wins
#                    1 game: 2 teams have the same Combined Wins
#                    2 games: All teams have the same Combined Wins                       
# 4 way tie: Teams are drawn into 2 "1st round" matches. Winners play for top 2 seeds, Losers play for bottom 2 seeds. Maximum of 4 games
# 5 way tie: 1 play-in game between 2 lowest SoV Teams. Loser gets lowest seed. Winner + 3 remaining teams go to 4-way-tie procedure. Maximum of 5 games
# 6 way tie: 2 randomly drawn play-in games between 4 lowest SoV Teams. Losers go to 2-way-tie procedure. Winners + 2 remaining teams go to 4-way-tie procedure. Maximum of 6 games
# 7 way tie: 3 randomly drawn play-in games between 6 lowest SoV Teams. Losers go to 3-way-tie procedure. Winners + remaining team go to 4 way-tie procedure. Maximum of 9 games.
# 8 way tie: 4 randomly drawn play-in games. Losers go to 4-way-tie procedure for bottom 4 seeds. Winners go to 4-way-tie procedure for top 4 seeds. Maximum of 12 games.
# 9 way tie: 1 play-in games between 2 lowest SoV Teams. Loser gets lowest seed. Winner + 7 remaining teams go to 8-way-tie procedure. Maximum of 13 games.
# 10 way tie: 2 play-in games between 4 lowest SoV Teams. Losers go to 2-way-tie for bottom 2 seeds. Winners + remaining 6 taems go to 8-way-tie procedure. Maximum of 14 games.

# --Specific Tiebreaker Scenarios--

#Ties for 1st:
# 3 way tie (1st-3rd): Refer to Summer Tiebreaker Rules for 3 way tie for scenarios. 0-2 games
# 4 way tie (1st-4th): 2 playin games. Losers play for 3rd/4th. Winners play for 1st/2nd. 4 games
# 5 way tie (1st-5th): 1 playin games. Loser is 5th seed. Winner + 3 remaining teams go to 4-way-tie for 1st-4th. 5 games
# 6 way tie (1st-6th): 2 playin games. Losers go to 2-way-tie for 5th/6th. Winners + 2 remaining teams go to 4-way-tie for 1st-4th. 6 games 
# 7 way tie (1st-7th): 3 playin games. Losers go to 3-way-tie for 5th-7th. Winners + 1 remaining team go to 4-way-tie for 1st-4th. 7-9 games
# 8 way tie (1st-8th): 4 playin games. Losers go to 4-way-tie for 5th-8th. Winners go to 4-way-tie for 1st-4th. 12 games.
# 9 way tie (1st-9th): 1 playin game. Loser is 9th seed. Winner + remaining 7 teams go to 8-way-tie for 1st-8th. 13 games. 
# 10 way tie (1st-10th): 2 playin games. Losers go to 2-way-tie for 9th/10th. Winners + 6 remaining teams go to 8-way-tie for 1st-8th. 14 games

#Ties for 2nd:
# 3 way tie (2nd-4th): Refer to Summer Tiebreaker Rules for 3 way tie for scenarios. 0-2 games
# 4 way tie (2nd-5th): 2 playin games. Losers play for 4th/5th. Winners play for 2nd/3rd. 4 games
# 5 way tie (2nd-6th): 1 playin games. Loser is 6th seed. Winner + 3 remainin teams go to a 3-way-tie for 2nd-5th. 5 games
# 6 way tie (2nd-7th): 2 playin games. Losers go to 2-way-tie for 6th/7th. Winners + 2 remaining teams go to 4-way-tie for 2nd-5th. 6 games
# 7 way tie (2nd-8th): 3 playin games. Losers go to 3-way-tie for 6th-8th. Winners + 1 remaining team go to 4-way-tie for 2nd-5th. 7-9 games
# 8 way tie (2nd-9th): 4 playin games. Losers go to 4-way-tie for 6th-9th. Winners go to 4-way-tie for 2nd-5th. 12 games.
# 9 way tie (2nd-10th): 1 playin game. Loser is 10th seed. Winner + remaining 7 teams go to 8-way-tie for 2nd-9th. 13 games.

#Ties for 3rd
# 3 way tie (3rd-5th): Refer to Summer Tiebreaker Rules for 3 way tie for scenarios. 0-2 games
# 4 way tie (3rd-6th): No tiebreaker games are played. They are randomly drawn into Quarterfinals, with SoV giving side selection. 0 games
# 5 way tie (3rd-7th): 1 playin game. Loser is 7th seed. Winner + 3 remaining teams go to 4-way-tie for 3rd-6th. 1 game
# 6 way tie (3rd-8th): 2 playin games. Losers go to 2-way-tie for 7th/8th. Winners + 2 remaining teams go to 4-way-tie for 3rd-6th. 2 games
# 7 way tie (3rd-9th): 3 playin games. Losers go to 3-way-tie for 7th-9th. Winners + 1 remaining team go to 4-way-tie for 3rd-6th. 3-5 games
# 8 way tie (3rd-10th): 4 playin games. Losers go to 4-way-tie for 7th-10th. Winners go to 4-way-tie for 3rd-6th. 7 games

#Ties for 4th
# 3 way tie (4th-6th): Refer to Summer Tiebreaker Rules for 3 way tie for scenarios. 0-2 games
# 4 way tie (4th-7th): 2 playin games. Losers play for 6th/7th. Winners play for 4th/5th. 4 games
# 5 way tie (4th-8th): 1 playin game. Loser is 8th seed. Winner + 3 remaining teams go to a 4-way-tie for 4th-7th. 5 games
# 6 way tie (4th-9th): 2 playin games. Losers go to 2-way-tie for 8th/9th. Winners + 2 remaining teams go to 4-way-tie for 4th-7th. 6 games
# 7 way tie (4th-10th): 3 playin games. Losers go to 3-way-tie for 8th-10th. Winners + 1 remaining team go to 4-way-tie for 4th-7th. 7-9 games.

#Ties for 5th
# 3 way tie (5th-7th): Refer to Summer Tiebreaker Rules for 3 way tie for scenarios. 0-2 games
# 4 way tie (5th-8th): 2 playin games. Losers play for 7th/8th. Winners play for 5th/6th. 4 games
# 5 way tie (5th-9th): 1 playin game. Loser is 9th seed. Winner + 3 remaining teams go to 4-way-tie for 5th-8th. 5 games
# 6 way tie (5th-10th): 2 playin games. Losers go to 2-way-tie for 9th/10th. Winner + 2 remaining teams go to 4-way-tie for 5th-8th. 6 games

#Ties for 6th
# 3 way tie (6th-8th): Refer to Summer Tiebreaker Rules for 3 way tie for scenarios. 0-2 games
# 4 way tie (6th-9th): 2 playin games. Losers play for 8th/9th. Winners play for 6th/7th. 4 games
# 5 way tie (6th-10th): 1 playin game. Loser is 10th seed. Winner + 3 remaining taems go to a 4-way-tie for 6th-9th. 5 games

#Ties for 7th
# 3 way tie (7th-9th): Refer to Summer Tiebreaker Rules for 3 way tie for scenarios. 0-2 games
# 4 way tie (7th-10th): 2 playin games. Losers DONT play for 9th/10th. Winners play for 7th/8th. 3 games

#Ties for 8th
# 3 way tie (8th-10th). There will be 0-2 tiebreaker games. Different scenarios apply:
#                              0 games: All teams have different Combined Wins, OR 2 teams have the same lowest Combined Wins.
#                              1 game: 2 teams have the same highest Combined Wins
#                              2 games: All taems have the same Combined Wins

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
Multiway_tie_unresolved_middle_new_tied_SOV = workbook.add_format({'bottom' : 2, 'top' : 2, 'left': 2, 'bg_color': 'lime', 'italic': True})
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
    for teams in sorted_teams_no_WL: # Assigns each team a set SoV points for where they placed in the standings. ex: {'100': 5.0, 'C9': 4.5, 'CLG': 4.0, 'DIG': 4.0, 'EG': 4.0, 'FLY': 2.5, 'GG': 2.0, 'IMT': 2.0, 'TL': 1.0, 'TSM': 0.5}
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
    []
]

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

# In cases where there is a multiway tie for a place where not all the TB games need to be placed, and SoV is needed to determine tiebreaker order, 
# if some or all SoVs are equal, it's not known to this script if a team will need to play a tiebreaker game
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
        "TL":  [12, 6],
        "TSM": [12, 6],
        "100": [11, 7],
        "DIG": [11, 7],
        "EG":  [10, 8],
        "IMT": [7, 10],
        "FLY": [6, 12],
        "CLG": [5, 13],
        "GG":  [3, 15],
    }

    # In a best case scenario, I would combine these 2 tables, but Leaguepedia, which is where I get my h2h info,
    # doesn't have a combined h2h table, it's easier to have 2 seperate tables for each split and combine them in the script.
    teams_wins_spring = { # 100 |  C9 | CLG | DIG | EG | FLY | GG | IMT | TL | TSM
        "100": [None, 0, 2, 2, 1, 1, 1, 2, 1, 0], #11 wins
        "C9":  [2, None, 1, 1, 1, 2, 2, 1, 0, 1], #13 wins
        "CLG": [0, 1, None, 1, 0, 0, 1, 1, 1, 0], #5 wins
        "DIG": [0, 0, 1, None, 2, 1, 2, 2, 0, 2], #11 wins
        "EG":  [1, 1, 2, 0, None, 2, 2, 0, 1, 1], #10 wins
        "FLY": [0, 0, 2, 0, 0, None, 2, 0, 0, 2], #6 wins
        "GG":  [1, 0, 1, 0, 0, 0, None, 1, 0, 0], #3 wins
        "IMT": [0, 0, 1, 0, 2, 2, 1, None, 1, 0], #7 wins
        "TL":  [1, 2, 1, 2, 1, 2, 2, 1, None, 0], #12 wins
        "TSM": [2, 1, 2, 0, 1, 0, 2, 2, 2, None]  #12 wins
    }
    teams_wins_summer = { #100, C9, CLG, DIG, EG, FLY, GG, IMT, TL, TSM
        "100": [None, 1, 0, 1, 0, 0, 0, 0, 0, 0],
        "C9":  [0, None, 0, 0, 0, 0, 0, 0, 1, 0],
        "CLG": [0, 0, None, 0, 0, 0, 0, 0, 0, 0],
        "DIG": [0, 0, 0, None, 1, 1, 0, 0, 0, 0],
        "EG":  [0, 0, 0, 0, None, 0, 0, 0, 0, 0],
        "FLY": [0, 0, 1, 0, 1, None, 0, 0, 0, 0],
        "GG":  [0, 1, 0, 0, 0, 0, None, 0, 0, 0],
        "IMT": [1, 0, 1, 0, 0, 0, 1, None, 0, 0],
        "TL":  [0, 0, 1, 0, 0, 0, 0, 0, None, 0],
        "TSM": [0, 0, 0, 0, 1, 0, 1, 0, 1, None]
    }
    teams_combined_wins = {}
    for team in teams_wins_spring:
        total_wins = []
        spring_record = teams_wins_spring[team]
        summer_record = teams_wins_summer[team]
        for x in range(len(spring_record)):
            if spring_record[x] is None:
                total_wins.append(None)
            else:
                total_wins.append(spring_record[x] + summer_record[x])
        teams_combined_wins[team] = total_wins    
    match_num = 0
    for winner in winners:
        teams_standings[winner][0] += 1
        teams_standings[loser][1] += 1
        if winner == matches[match_num][0]: # loser == matches[match_num][1]
            loser = matches[match_num][1]
        else:
            loser = matches[match_num][0]
        teams_combined_wins[winner][list(teams_combined_wins).index(loser)][0] += 1 #Increase winner's wins vs opponent by 1 in teams_combined_wins
        teams_standings[loser][1] += 1 # Increase's loser's losses by one in teams_standings
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
        elif len(teams_in_ordinal) == 2: # If there is a two way tie, it goes to head-to-head records. Tiebreaker games are impossible in Summer.
            team_1, team_2 = teams_in_ordinal
            team_1_aggregate = teams_combined_wins[team_1][list(teams_combined_wins).index(team_2)]
            team_2_aggregate = teams_combined_wins[team_2][list(teams_combined_wins).index(team_1)]
            if ordinal >= 8: # Ties that don't effect seeding for playoffs are not played out, therefore no tiebreaker games are needed. Check for >50% h2h anyway.
                if team_1_aggregate > team_2_aggregate or team_1_aggregate == team_2_aggregate:
                    row_data.append([col, team_1, two_way_tie_resolved_start]) 
                    col += 1
                    row_data.append([col, team_2, two_way_tie_resolved_end])
                    teams_chances_no_tie[team_1][ordinal] += 1
                    teams_chances_no_tie[team_2][ordinal+1] += 1
                elif team_2_aggregate > team_1_aggregate:
                    row_data.append([col, team_2, two_way_tie_resolved_start]) 
                    col += 1
                    row_data.append([col, team_1, two_way_tie_resolved_end])
                    teams_chances_no_tie[team_1][ordinal+1] += 1
                    teams_chances_no_tie[team_2][ordinal] += 1
            elif team_1_aggregate > team_2_aggregate: #If team 1 has a positive game differential against team 2, team 1 wins the tie
                row_data.append([col, team_1, two_way_tie_resolved_start])
                col += 1
                row_data.append([col, team_2, two_way_tie_resolved_end])
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal+1] += 1
            elif team_2_aggregate > team_1_aggregate: #If team 2 has a positive game differential against team 1, team 2 wins the tie
                row_data.append([col, team_2, two_way_tie_resolved_start]) 
                col += 1
                row_data.append([col, team_1, two_way_tie_resolved_end])
                teams_chances_no_tie[team_2][ordinal] += 1
                teams_chances_no_tie[team_1][ordinal+1] += 1
        elif len(teams_in_ordinal) == 3: #If there is a three way tie, the aggregate head-to-heads are compared.
            three_way_ties += 1
            team_1, team_2, team_3 = teams_in_ordinal
            team_1_aggregate = teams_combined_wins[team_1][list(teams_combined_wins).index(team_2)] + teams_combined_wins[team_1][list(teams_combined_wins).index(team_3)]
            team_2_aggregate = teams_combined_wins[team_2][list(teams_combined_wins).index(team_1)] + teams_combined_wins[team_2][list(teams_combined_wins).index(team_3)]
            team_3_aggregate = teams_combined_wins[team_3][list(teams_combined_wins).index(team_1)] + teams_combined_wins[team_3][list(teams_combined_wins).indeX(team_2)]
            teams_aggs_dict = {team_1: team_1_aggregate, team_2: team_2_aggregate, team_3: team_3_aggregate}
            sorted_teams_aggs_dict = {}
            for team in sorted(teams_aggs_dict, key=teams_aggs_dict.get, reverse=True):
                sorted_teams_aggs_dict[team] = teams_aggs_dict[team]
            team_1, team_2, team_3 = sorted_teams_aggs_dict #Returns teams in order of team aggregate
            team_1_aggregate, team_2_aggregate, team_3_aggregate = sorted_teams_aggs_dict.values() #Returns aggregates in descending order. team_1_aggregate is in fact team_1's aggregate.
            if team_1_aggregate == team_2_aggregate == team_3_aggregate: # Scenario 1: All teams have the same aggregates. 2 tiebreaker games. Not sure if this is actually possible or not with 15 total aggregate games, but here we are.
                tiebreaker_games += 2
                teams_sovs = Strength_of_victory([team_1, team_2, team_3], teams_combined_wins, sorted_teams_no_WL)
                team_1_sov = teams_sovs[0]
                team_2_sov = teams_sovs[1]
                team_3_sov = teams_sovs[2]
                teams_sov_dict = {team_1: team_1_sov, team_2: team_2_sov, team_3: team_3_sov}
                sorted_teams_sov_dict = {}
                for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                    sorted_teams_sov_dict[team] = teams_sov_dict[team]
                team_1, team_2, team_3 = list(sorted_teams_sov_dict)
                team_1_sov, team_2_sov, team_3_sov = sorted_teams_sov_dict.values()
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                tiebreaker_games += 2
                if team_1_sov > team_2_sov > team_3_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_end])
                    teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                elif team_1_sov > team_2_sov == team_3_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_end_tied_SOV])
                    teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                elif team_1_sov == team_2_sov > team_3_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_end])
                    teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                elif team_1_sov == team_2_sov == team_3_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_end_tied_SOV])
                    teams_worst_finish_in_ties[team_1][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
            elif team_1_aggregate == team_2_aggregate > team_3_aggregate: # Scenario 2: If the top 2 teams have the same aggregate, they will play a tiebreaker game, with side selection given to the favored team in their h2h
                tiebreaker_games += 1
                team_1_aggregate = teams_combined_wins[team_1][list(teams_combined_wins).index(team_2)]
                team_2_aggregate = teams_combined_wins[team_2][list(teams_combined_wins).index(team_1)]
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_no_tie[team_3][ordinal+2] += 1
                teams_worst_finish_in_ties[team_1][ordinal+1] += 1
                teams_worst_finish_in_ties[team_2][ordinal+1] += 1
                if team_1_aggregate > team_2_aggregate:
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_partially_resolved_end_locked])
                else:
                    row_data.append([col, team_2, Multiway_tie_partially_resolved_begin])
                    col += 1
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_partially_resolved_end_locked])
            elif team_1_aggregate > team_2_aggregate == team_3_aggregate: # Scenario 3: If the bottom 2 teams have the same aggregate, they will play a tiebreaker game, with side selection given to the favored team in their h2h, unless...
                if ordinal == 7: #If the 3 way tie is for 8th, then the bottom two teams do not have a tiebreaker game, as they would play for 9th/10th.
                    row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_fully_resolved_end])
                else:
                    tiebreaker_games += 1
                    team_2_aggregate = teams_combined_wins[team_2][list(teams_combined_wins).index(team_3)]
                    team_3_aggregate = teams_combined_wins[team_3][list(teams_combined_wins).index(team_2)]
                    teams_chances_no_tie[team_1][ordinal] += 1
                    teams_chances_tie[team_2][ordinal+1] += 1
                    teams_chances_tie[team_3][ordinal+1] += 1
                    teams_worst_finish_in_ties[team_2][ordinal+2] += 1
                    teams_worst_finish_in_ties[team_3][ordinal+2] += 1
                    row_data.append([col, team_1, Multiway_tie_partially_resolved_begin_locked])
                    col += 1
                    if team_2_aggregate > team_3_aggregate:
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_end])
                    else:
                        row_data.append([col, team_3, Multiway_tie_partially_resolved_middle])
                        col += 1
                        row_data.append([col, team_2, Multiway_tie_partially_resolved_end])
            elif team_1_aggregate > team_2_aggregate > team_3_aggregate: # Scenario 4: All teams have different aggregates. 0 tiebreaker games.
                teams_chances_no_tie[team_1][ordinal] += 1
                teams_chances_no_tie[team_2][ordinal+1] += 1
                teams_chances_no_tie[team_3][ordinal+2] += 1
                row_data.append([col, team_1, Multiway_tie_fully_resolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_fully_resolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_fully_resolved_end])
        elif len(teams_in_ordinal) == 4: #All teams randomly seeded. SOVs give side selection.
            team_1, team_2, team_3, team_4 = teams_in_ordinal
            if ordinal == 2: #If teams are playing for 3rd, there are no tiebreaker games, and they're all considered 3rd seed.
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
                teams_sov = Strength_of_victory([team_1, team_2, team_3, team_4], teams_combined_wins, sorted_teams_no_WL)
                teams_sov_dict = {team_1: teams_sov[0], team_2: teams_sov[1], team_3: teams_sov[2], team_4: teams_sov[3]}
                sorted_teams_sov_dict = {}
                for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                    sorted_teams_sov_dict[team] = teams_sov_dict[team]
                team_1, team_2, team_3, team_4 = teams_in_ordinal = list(sorted_teams_sov_dict)
                team_1_sov, team_2_sov, team_3_sov, team_4_sov = sorted_teams_sov_dict.values()
                if team_1_sov > team_2_sov > team_3_sov > team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end])
                elif team_1_sov > team_2_sov > team_3_sov == team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end_tied_SOV])
                elif team_1_sov > team_2_sov == team_3_sov > team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end])
                elif team_1_sov > team_2_sov == team_3_sov == team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end_tied_SOV])
                elif team_1_sov == team_2_sov > team_3_sov > team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                    col +1 
                    row_data.append([col, team_4, Multiway_tie_unresolved_end])
                elif team_1_sov == team_2_sov > team_3_sov == team_4_sov: 
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_new_tied_SOV])
                    col += 1
                    row_data.append([col, team_4, Multiway_tie_unresolved_end_tied_SOV])
                elif team_1_sov == team_2_sov == team_3_sov > team_4_sov:
                    row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                    col += 1
                    row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                    col += 1
                    row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
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
                if ordinal in [5, 4]:
                    tiebreaker_games += 3
                elif ordinal in [0, 1, 2, 3]:
                    tiebreaker_games += 4
                tiebreaker_games += 4
        elif len(teams_in_ordinal) == 5: #2 lowest SOVs go to play-in | Winner of playin + other 3 teams go to 4-way-tie
            team_1, team_2, team_3, team_4, team_5 = teams_in_ordinal
            teams_sov = Strength_of_victory([team_1, team_2, team_3, team_4, team_5], teams_combined_wins, sorted_teams_no_WL)
            teams_sov_dict = {team_1: teams_sov[0], team_2: teams_sov[1], team_3: teams_sov[2], team_4: teams_sov[3], team_5: teams_sov[4]}
            sorted_teams_sov_dict = {}
            for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                sorted_teams_sov_dict[team] = teams_sov_dict[team]
            team_1, team_2, team_3, team_4, team_5 = list(sorted_teams_sov_dict)
            team_1_sov, team_2_sov, team_3_sov, team_4_sov, team_5_sov = sorted_teams_sov_dict.values()
            #The below if/elifs are coded so that no matter what, the teams will be in order of SOV to indicate which 2 teams will be in the play-in, as well as which teams have side selection priority when they get drawn in to the 4-way-tie.
                #If you only want to know what teams are the bottom 2, and you don't care about SOV order, you can optimize this section with the following code block. Including this destroys my line count, but who really cares in the end?
                    # #if team_3_sov > team_4_sov > team_5_sov:
                    #     pass  
                    # elif team_3_sov > team_4_sov == team_5_sov:
                    #     pass
                    # elif team_2_sov > team_3_sov == team_4_sov > team_5_sov:
                    #     pass
                    # elif team_2_sov > team_3_sov == team_4_sov == team_5_sov:
                    #     pass
                    # elif team_1_sov > team_2_sov == team_3_sov == team_4_sov > team_5_sov:
                    #     pass
                    # elif team_1_sov > team_2_sov == team_3_sov == team_4_sov == team_5_sov:
                    #     pass
                    # elif team_1_sov == team_2_sov == team_3_sov == team_4_sov > team_5_sov:
                    #     pass
                    # elif team_1_sov == team_2_sov == team_3_sov == team_4_sov == team_5_sov:
                    #     pass
            if team_1_sov > team_2_sov > team_3_sov > team_4_sov > team_5_sov:
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
            elif team_1_sov > team_2_sov > team_3_sov > team_4_sov == team_5_sov:
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
            elif team_1_sov > team_2_sov > team_3_sov == team_4_sov > team_5_sov:
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
                teams_worst_finish_in_ties[team_3][ordinal+4] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1 
            elif team_1_sov > team_2_sov > team_3_sov == team_4_sov == team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
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
                teams_worst_finish_in_ties[team_3][ordinal+4] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1 
            elif team_1_sov > team_2_sov == team_3_sov > team_4_sov > team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
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
            elif team_1_sov > team_2_sov == team_3_sov > team_4_sov == team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_new_tied_SOV])
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
            elif team_1_sov > team_2_sov == team_3_sov == team_4_sov > team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
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
                teams_worst_finish_in_ties[team_2][ordinal+4] += 1
                teams_worst_finish_in_ties[team_3][ordinal+4] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1 
            elif team_1_sov > team_2_sov == team_3_sov == team_4_sov == team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
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
                teams_worst_finish_in_ties[team_2][ordinal+4] += 1
                teams_worst_finish_in_ties[team_3][ordinal+4] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1 
            elif team_1_sov == team_2_sov > team_3_sov > team_4_sov > team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
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
            elif team_1_sov == team_2_sov > team_3_sov > team_4_sov == team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
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
            elif team_1_sov == team_2_sov > team_3_sov == team_4_sov > team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_new_tied_SOV])
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
                teams_worst_finish_in_ties[team_3][ordinal+4] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1 
            elif team_1_sov == team_2_sov > team_3_sov == team_4_sov == team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_new_tied_SOV])
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
                teams_worst_finish_in_ties[team_3][ordinal+4] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1 
            elif team_1_sov == team_2_sov == team_3_sov > team_4_sov > team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
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
            elif team_1_sov == team_2_sov == team_3_sov > team_4_sov == team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_new_tied_SOV])
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
            elif team_1_sov == team_2_sov == team_3_sov == team_4_sov > team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
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
                teams_worst_finish_in_ties[team_1][ordinal+4] += 1
                teams_worst_finish_in_ties[team_2][ordinal+4] += 1
                teams_worst_finish_in_ties[team_3][ordinal+4] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1 
            elif team_1_sov == team_2_sov == team_3_sov == team_4_sov == team_5_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+4] += 1
                teams_worst_finish_in_ties[team_2][ordinal+4] += 1
                teams_worst_finish_in_ties[team_3][ordinal+4] += 1
                teams_worst_finish_in_ties[team_4][ordinal+4] += 1
                teams_worst_finish_in_ties[team_5][ordinal+4] += 1 
            if ordinal == 2: #If playing for 3rd, only the play-in is needed, since the resulting 4-way-tie would be for 3rd. TODO: When there is a definite bottom 2 teams, make it partially resolved.
                tiebreaker_games += 1
            else:
                tiebreaker_games += 5
            five_way_ties += 1
        elif len(teams_in_ordinal) == 6: #4 lowest SOVs randomly drawn into 2 play-ins | Losers go to 2 way tie, Winners go to 4 way tie
            team_1, team_2, team_3, team_4, team_5, team_6 = teams_in_ordinal
            teams_sov = Strength_of_victory([team_1, team_2, team_3, team_4, team_5, team_6], teams_combined_wins, sorted_teams_no_WL)
            teams_sov_dict = {team_1: teams_sov[0], team_2: teams_sov[1], team_3: teams_sov[2], team_4: teams_sov[3], team_5: teams_sov[4], team_6: teams_sov[5]}
            sorted_teams_sov_dict = {}
            for team in sorted(teams_sov_dict, key=teams_sov_dict.get, reverse=True):
                sorted_teams_sov_dict[team] = teams_sov_dict[team]
            team_1, team_2, team_3, team_4, team_5, team_6 = list(sorted_teams_sov_dict)
            team_1_sov, team_2_sov, team_3_sov, team_4_sov, team_5_sov, team_6_sov = sorted_teams_sov_dict.values()
            # Same scenario as the 5-way-tiebreaker. Can be optimized if you just want to know the bottom 4 teams. Don't want to bother writing it out now, good luck future me. 
            # Also at what point does all the unique scenarios not actually become possible? Cause ain't no way 10-way-tie has 2^9 unique scenarios for SOVs...
            if team_1_sov > team_2_sov > team_3_sov > team_4_sov > team_5_sov > team_6_sov:
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
            elif team_1_sov > team_2_sov > team_3_sov > team_4_sov > team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
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
            elif team_1_sov > team_2_sov > team_3_sov > team_4_sov == team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
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
            elif team_1_sov > team_2_sov > team_3_sov > team_4_sov == team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
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
            elif team_1_sov > team_2_sov > team_3_sov == team_4_sov > team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
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
            elif team_1_sov > team_2_sov > team_3_sov == team_4_sov > team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
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
            elif team_1_sov > team_2_sov > team_3_sov == team_4_sov == team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
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
            elif team_1_sov > team_2_sov > team_3_sov == team_4_sov == team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
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
            elif team_1_sov > team_2_sov == team_3_sov > team_4_sov > team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov > team_2_sov == team_3_sov > team_4_sov > team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov > team_2_sov == team_3_sov > team_4_sov == team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov > team_2_sov == team_3_sov > team_4_sov == team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov > team_2_sov == team_3_sov == team_4_sov > team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov > team_2_sov == team_3_sov == team_4_sov > team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov > team_2_sov == team_3_sov == team_4_sov == team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov > team_2_sov == team_3_sov == team_4_sov == team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+3] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov == team_2_sov > team_3_sov > team_4_sov > team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
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
            elif team_1_sov == team_2_sov > team_3_sov > team_4_sov > team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
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
            elif team_1_sov == team_2_sov > team_3_sov > team_4_sov == team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
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
            elif team_1_sov == team_2_sov > team_3_sov > team_4_sov == team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
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
            elif team_1_sov == team_2_sov > team_3_sov == team_4_sov > team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
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
            elif team_1_sov == team_2_sov > team_3_sov == team_4_sov > team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
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
            elif team_1_sov == team_2_sov > team_3_sov == team_4_sov == team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
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
            elif team_1_sov == team_2_sov > team_3_sov == team_4_sov == team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
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
            elif team_1_sov == team_2_sov == team_3_sov > team_4_sov > team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+5] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov == team_2_sov == team_3_sov > team_4_sov > team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+5] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov == team_2_sov == team_3_sov > team_4_sov == team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+5] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov == team_2_sov == team_3_sov > team_4_sov == team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+5] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov == team_2_sov == team_3_sov == team_4_sov > team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+5] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov == team_2_sov == team_3_sov == team_4_sov > team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_new_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+5] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov == team_2_sov == team_3_sov == team_4_sov == team_5_sov > team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+5] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            elif team_1_sov == team_2_sov == team_3_sov == team_4_sov == team_5_sov == team_6_sov:
                row_data.append([col, team_1, Multiway_tie_unresolved_begin_tied_SOV])
                col += 1
                row_data.append([col, team_2, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_3, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_4, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_5, Multiway_tie_unresolved_middle_tied_SOV])
                col += 1
                row_data.append([col, team_6, Multiway_tie_unresolved_end_tied_SOV])
                teams_chances_tie[team_1][ordinal] += 1
                teams_chances_tie[team_2][ordinal] += 1
                teams_chances_tie[team_3][ordinal] += 1
                teams_chances_tie[team_4][ordinal] += 1
                teams_chances_tie[team_5][ordinal] += 1
                teams_chances_tie[team_6][ordinal] += 1
                teams_worst_finish_in_ties[team_1][ordinal+5] += 1
                teams_worst_finish_in_ties[team_2][ordinal+5] += 1
                teams_worst_finish_in_ties[team_3][ordinal+5] += 1
                teams_worst_finish_in_ties[team_4][ordinal+5] += 1
                teams_worst_finish_in_ties[team_5][ordinal+5] += 1
                teams_worst_finish_in_ties[team_6][ordinal+5] += 1
            if ordinal == 2: #If playing for 3rd, only the 2 playin games are needed, as the losers would go to a resolved 2-way-tie, and winners go to a 4-way-tie for 3rd with no tiebreakers.
                tiebreaker_games += 2
            else:
                tiebreaker_games += 6
            six_way_ties += 1
        else:
            if len(teams_in_ordinal) == 7: 
                if ordinal == 2: #3-5 tiebreaker games. Only add the minimum.
                    tiebreaker_games += 3
                else: #7-9 tiebreaker games
                    tiebreaker_games += 7
                seven_way_ties += 1
            elif len(teams_in_ordinal) == 8:
                if ordinal == 2:
                    tiebreaker_games += 3
                else:
                    tiebreaker_games += 7
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
