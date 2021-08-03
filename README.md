# LCS-Foldy-Sheet
Unofficial LCS Foldy Sheet Generator

Required packages:
  - xlxswriter
  - itertools

Optional packages:
  - tqdm (Progress bar for large scenario processing)
  - timeit (Manual timing of different stages of the script)

What the sheets do:
   1. Completely does 2-6 way ties, including SOVs where needed for side selection order. 7-10 way ties are calculated but not in SOV order.
   2. Outputs Foldy Sheet to an .xlsx file
   3. Outputs lists/arrays indicating the following:
        - Teams chances of ending in Nth place with no tiebreaker games played. 
        - Teams chances of ending tied for Nth place with tiebreaker games played.
        - Teams chances of ending tied for Nth place with tiebreaker games played, but unknown if they really need to play or not (Mainly a 3rd/4th seed tiebreaker thing).
        - Teams chances of finishing in the worst place possible in tiebreakers (Used to determine if X team has locked certain spots).
        
Each file is coded to fit within that split's tiebreaker rules, and newer files likely have optimizations not present in previous files.

Update (August 3rd): So fun fact, xlsxwriter is compatible with xlsxwriter. The time to execute the script is cut in half at 18 matches.
