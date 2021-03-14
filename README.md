# LCS-Foldy-Sheet
Unofficial LCS Foldy Sheet Generator

Required packages:
  - xlxswriter
  - itertools

Optional packages:
  - tqdm (Progress bar for large scenario processing)
  - timeit (Manual timing of different stages of the script)

LCS Spring 2021:
   1. Completely does 2-6 way ties, including SOVs where needed for side selection order.
   2. Outputs Foldy Sheet to an .xlsx file
   3. Outputs lists/arrays indicating the following:
        - Teams chances of ending in Nth place with no tiebreaker games played. 
        - Teams chances of ending tied for Nth place with tiebreaker games played.
        - Teams chances of ending tied for Nth place with tiebreaker games played, but unknown if they really need to play or not (Mainly a 3rd/4th seed tiebreaker thing.)
        - Teams chances of finishing in the worst place possible in tiebreakers (Used to determine if X team has locked certain spots
        
LCS Summer 2021:
    1. In Progress.
