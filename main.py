from openpyxl import load_workbook
import pandas as pd
import re

file_path = "source/ALL_ROSTERS.xlsx"

#! This blocks may change according to the Excel file structure
# Check your the excel file to confirm these ranges
team_blocks = [
    ('A', 5, 32),    # 1st team
    ('F', 5, 32),    # 2nd team
    ('A', 34, 61),   # 3rd team
    ('F', 34, 61),   # 4th team
    ('A', 63, 90),   # 5th team
    ('F', 63, 90),   # 6th team
    ('A', 92, 119),  # 7th team
    ('F', 92, 119),  # 8th team
    ('A', 121, 148), # 9th team
    ('F', 121, 148)  # 10th team
]

def get_all_teams(file_path):
    wb = load_workbook(filename=file_path, data_only=True)
    ws = wb.active

    team_names = []
    for col, start_row, end_row in team_blocks:
        # Team name is always in the first row of the block
        team_name = ws[f'{col}{start_row}'].value
        team_names.append(team_name)
    return team_names

def get_team(file_path, team_index):
    """
    team_index: 0-based, 0 = first team, 1 = second team, etc.
    """
    wb = load_workbook(filename=file_path, data_only=True)
    ws = wb.active

    col, start_row, end_row = team_blocks[team_index]

    # Team name
    team_name = ws[f'{col}{start_row}'].value

    # Remaining coins (credits) in the last row of the block
    coin_str = ws[f'{col}{end_row}'].value
    coins_left = int(re.search(r'\d+', coin_str).group())

    # Players: from start_row+2 to end_row-1 (row after the name is the header)
    df_player = pd.DataFrame(columns=['Role', 'Player', 'Team', 'Cost'])
    for row in range(start_row+2, end_row):
        role = ws[f'{col}{row}'].value
        player_name = ws[f'{col}{row}'].offset(column=1).value
        player_team = ws[f'{col}{row}'].offset(column=2).value
        cost = ws[f'{col}{row}'].offset(column=3).value
        df_player.loc[row-(start_row+2)] = [role, player_name, player_team, cost]

    return team_name, df_player, coins_left

# --- Example: get all team names ---
team_names = get_all_teams(file_path)
print("All teams:", team_names)

# --- Example: get the 3rd team ---
team_name, players, coins = get_team(file_path, 4)
print("\nTeam:", team_name)
print("\nPlayers:")
print(players)
print("\nCoins left:", coins)