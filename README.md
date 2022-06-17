# LCS-Foldy-Sheet 2022
Unofficial LCS Foldy Sheet Generator

Required packages:
  - xlxswriter
  - itertools

Optional:
  - PyPy, for faster runtime

What the sheets do:
   1. Generates every single possible scenario of which teams win each match
   2. Takes each scenario and calculate any ties, and if those ties need tiebreakers.
   3. Writes data to an .xlsx file, known as a "Foldy Sheet"
   4. Prints lists/arrays indicating the following:
        - Teams chances of ending in Nth place with no tiebreaker games played. 
        - Teams chances of ending tied for Nth place with tiebreaker games played.
        - Teams chances of finishing in the worst place possible in tiebreakers (Used to determine if X team has locked certain spots).
        
Each file is coded to fit within that split's tiebreaker rules, and newer files likely have optimizations not present in previous files.
