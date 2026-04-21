import openpyxl
import numpy as np

def get_heights_pairwise(frame):
    male_female_pairs = []

    for row in range(2, frame.max_row):
        pair = []
        for col in frame.iter_cols(1, frame.max_column):
            pair.append(col[row].value)
        male_female_pairs.append(pair)

    return male_female_pairs

def get_heights(dataframe, column):
    heights = []
    for row in range(2, dataframe.max_row):
        heights.append(dataframe[row][column].value)

    return heights



marsh_data = openpyxl.load_workbook("heights-data-marsh-1988-great-britain.xlsx")
marsh_active_sheet = marsh_data.active

male_column = 0
female_column = 1

male_heights = get_heights(marsh_active_sheet, male_column)
female_heights = get_heights(marsh_active_sheet, female_column)

print(f"Mean male height: {np.mean(male_heights):.3f} mm")
print(f"Mean female height: {np.mean(female_heights):.3f} mm")