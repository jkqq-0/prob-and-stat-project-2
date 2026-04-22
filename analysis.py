import openpyxl
import numpy as np
import scipy.stats as stats

def get_heights_pairwise(frame, conversion_function=lambda x: x):
    male_female_pairs = []

    for row in range(2, frame.max_row):
        pair = []
        for col in frame.iter_cols(1, frame.max_column):
            pair.append(conversion_function(col[row].value))
        male_female_pairs.append(pair)

    return male_female_pairs

def get_heights(dataframe, column, conversion_function=lambda x: x):
    heights = []
    for row in range(2, dataframe.max_row):
        heights.append(conversion_function(dataframe[row][column].value))
    return heights

def get_z_score(x, mu, sigma):
    return (x - mu) / sigma

def process_dataset(name, filename, conversion_function=lambda x: x):
    data = openpyxl.load_workbook(filename)
    sheet = data.active
    
    male_col, female_col = 0, 1
    
    males = get_heights(sheet, male_col, conversion_function)
    females = get_heights(sheet, female_col, conversion_function)
    
    male_mean, female_mean = np.mean(males), np.mean(females)
    male_std, female_std = np.std(males), np.std(females)
    
    correlation = np.corrcoef(males, females)[0, 1]
    
    # Theoretical difference
    diff_mean = male_mean - female_mean
    var_of_diff = male_std**2 + female_std**2 - 2 * correlation * male_std * female_std
    std_of_diff = np.sqrt(var_of_diff)
    
    z_score = get_z_score(0, diff_mean, std_of_diff)
    p_value = stats.norm.sf(z_score)
    
    # Sample statistical height differences
    pairs = get_heights_pairwise(sheet, conversion_function)
    diffs = np.array([p[0] - p[1] for p in pairs])
    sample_diff_mean = np.mean(diffs)
    sample_diff_std = np.std(diffs)
    sample_diff_z_score = get_z_score(0, sample_diff_mean, sample_diff_std)
    p_value = stats.norm.sf(sample_diff_z_score)

    # actual count with height difference > 0
    positive_diffs = [k for k in diffs if k > 0]
    positive_diff_proportion = len(positive_diffs) / len(diffs)

    return {
        "name": name,
        "male_mean": male_mean,
        "female_mean": female_mean,
        "male_std": male_std,
        "female_std": female_std,
        "correlation": correlation,
        "diff_mean": diff_mean,
        "var_of_diff": var_of_diff,
        "std_of_diff": std_of_diff,
        "z_score": z_score,
        "p_value": p_value,
        "sample_diff_mean": sample_diff_mean,
        "sample_diff_std": sample_diff_std,
        "sample_diff_z_score": sample_diff_z_score,
        "sample_diff_p_value": p_value,
        "positive_diff_proportion": positive_diff_proportion
    }

datasets = [
    process_dataset("Marsh", "heights-data-marsh-1988-great-britain.xlsx"),
    process_dataset("Galton", "GaltonFamilies-1886.xlsx", lambda x: 25.4 * x)
]

# Output
print("SAMPLE MEANS")
for d in datasets:
    print(f"{d['name']} Mean male height: {d['male_mean']:.3f} mm")
    print(f"{d['name']} Mean female height: {d['female_mean']:.3f} mm")
print()

print("SAMPLE STANDARD DEVIATIONS")
for d in datasets:
    print(f"{d['name']} male height standard deviation: {d['male_std']:.3f} mm")
    print(f"{d['name']} female height standard deviation: {d['female_std']:.3f} mm")
print()

print("SAMPLE MEAN CORRELATION")
for d in datasets:
    print(f"Correlation between male and female heights ({d['name']}) {d['correlation']:.3f}")
print()

print("THEORETICAL DIFFERENCE MEAN AND STANDARD DEVIATIONS")
for d in datasets:
    print(f"{d['name']} E(x1-x2) = {d['diff_mean']:.3f} mm")
    print(f"{d['name']} Var(x1-x2) = {d['var_of_diff']:.3f} mm^2")
    print(f"\tstd of that is {d['std_of_diff']:.3f} mm")
    print(f"({d['name']}) Z score of difference=0: {d['z_score']:.3f}")
    print(f"P(z > {d['z_score']:.3f}): {d['p_value']:.3f}")
print()

print("SAMPLE DIFFERENCE MEAN AND STANDARD DEVIATIONS")
for d in datasets:
    print(f"{d['name']} height difference mean: {d['sample_diff_mean']:.3f} mm")
    print(f"{d['name']} height difference std: {d['sample_diff_std']:.3f} mm")
    print(f"({d['name']}) Z score of difference=0: {d['sample_diff_z_score']:.3f}")
    print(f"P(z > {d['sample_diff_z_score']:.3f}): {d['p_value']:.3f}")
    print(f"({d['name']}) Actual positive difference proportion: {d['positive_diff_proportion']:.3f}")
print()
