import pandas as pd
import numpy as np
import scipy.stats as stats


def get_z_score(x, mu, sigma):
    return (x - mu) / sigma


def process_dataset(name, filename, conversion_function=lambda x: x):
    df = pd.read_excel(filename)
    df = df.map(conversion_function)
    males = df.iloc[:, 0]
    females = df.iloc[:, 1]

    sample_size = len(df)

    male_mean, female_mean = males.mean(), females.mean()
    male_std, female_std = males.std(), females.std()

    correlation = males.corr(females)

    # Theoretical difference (Bivariate Normal model)
    diff_mean = male_mean - female_mean
    var_of_diff = male_std**2 + female_std**2 - 2 * correlation * male_std * female_std
    std_of_diff = np.sqrt(var_of_diff)

    # P(X1 - X2 > 0) using theoretical model
    z_score = get_z_score(0, diff_mean, std_of_diff)
    p_value = stats.norm.sf(z_score)

    # Sample statistical height differences (Y)
    diffs = males - females
    sample_diff_mean = diffs.mean()
    sample_diff_std = diffs.std()
    sample_diff_z_score = get_z_score(0, sample_diff_mean, sample_diff_std)
    sample_diff_p_value = stats.norm.sf(sample_diff_z_score)

    positive_diff_proportion = (diffs > 0).mean()

    # Hypothesis testing (Question 11)
    test_statistic = (sample_diff_p_value - p_value) / np.sqrt(
        (p_value * (1 - p_value)) / sample_size
    )

    # Two-tailed check (95% confidence)
    ci_lower_bound, ci_upper_bound = stats.norm.interval(0.95)
    interval_check = ci_lower_bound < test_statistic < ci_upper_bound

    # One-tailed check
    one_tailed_critical = stats.norm.ppf(0.95)
    one_tailed_check = test_statistic > one_tailed_critical

    return {
        "name": name,
        "sample_size": sample_size,
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
        "sample_diff_p_value": sample_diff_p_value,
        "positive_diff_proportion": positive_diff_proportion,
        "test_statistic": test_statistic,
        "interval_check": interval_check,
        "one_tailed_check": one_tailed_check,
    }


def joint_hypothesis_test(results1, results2):
    n_1 = results1["sample_size"]
    n_2 = results2["sample_size"]

    p1_hat = results1["sample_diff_p_value"]
    p2_hat = results2["sample_diff_p_value"]

    # Calculate x_i using the hint from Q16
    x1 = p1_hat * n_1
    x2 = p2_hat * n_2

    p_bar = (x1 + x2) / (n_1 + n_2)
    q_bar = 1 - p_bar

    # Null hypothesis: p1 = p2
    return (p1_hat - p2_hat) / np.sqrt(
        ((p_bar * q_bar) / n_1) + ((p_bar * q_bar) / n_2)
    )


def t_test(results1, results2, sex: str):
    assert sex == "male" or sex == "female"

    n_1 = results1["sample_size"]
    n_2 = results2["sample_size"]

    df = min([n_1 - 1, n_2 - 1])

    xbar_1 = results1[f"{sex}_mean"]
    xbar_2 = results2[f"{sex}_mean"]
    s_1 = results1[f"{sex}_std"]
    s_2 = results2[f"{sex}_std"]

    t = (xbar_1 - xbar_2) / np.sqrt((s_1**2 / n_1) + (s_2**2 / n_2))
    p_value = 2 * stats.t.sf(np.abs(t), df=df)

    return t, p_value


def get_degrees_of_freedom(results1, results2):
    n_1 = results1["sample_size"]
    n_2 = results2["sample_size"]
    return min([n_1 - 1, n_2 - 1])


# Dataset processing
datasets = [
    process_dataset("Marsh", "heights-data-marsh-1988-great-britain.xlsx"),
    process_dataset("Galton", "GaltonFamilies-1886.xlsx", lambda x: 25.4 * x),
]

# Formatting and Output
print("SAMPLE SIZES")
for d in datasets:
    print(f"{d['name']}: {d['sample_size']}")
print()

print("SAMPLE MEANS")
for d in datasets:
    print(f"{d['name']} Mean male height: {d['male_mean']:.4f} mm")
    print(f"{d['name']} Mean female height: {d['female_mean']:.4f} mm")
print()

print("SAMPLE STANDARD DEVIATIONS (StDev.S)")
for d in datasets:
    print(f"{d['name']} male height standard deviation: {d['male_std']:.4f} mm")
    print(f"{d['name']} female height standard deviation: {d['female_std']:.4f} mm")
print()

print("SAMPLE MEAN CORRELATION")
for d in datasets:
    print(
        f"Correlation between male and female heights ({d['name']}): {d['correlation']:.4f}"
    )
print()

print("THEORETICAL DIFFERENCE (Bivariate Normal Model)")
for d in datasets:
    print(f"{d['name']} E(x1-x2) = {d['diff_mean']:.4f} mm")
    print(f"{d['name']} Var(x1-x2) = {d['var_of_diff']:.4f} mm^2")
    print(f"\tstd of that is {d['std_of_diff']:.4f} mm")
    print(f"({d['name']}) Z score of difference=0: {d['z_score']:.4f}")
    print(f"P(z > {d['z_score']:.4f}): {d['p_value']:.4f}")
print()

print("SAMPLE DIFFERENCE (Paired Y values)")
for d in datasets:
    print(f"{d['name']} height difference mean: {d['sample_diff_mean']:.4f} mm")
    print(f"{d['name']} height difference std: {d['sample_diff_std']:.4f} mm")
    print(f"({d['name']}) Z score of difference=0: {d['sample_diff_z_score']:.4f}")
    print(f"P(z > {d['sample_diff_z_score']:.4f}): {d['sample_diff_p_value']:.4f}")
    print(
        f"({d['name']}) Actual positive difference proportion: {d['positive_diff_proportion']:.4f}"
    )
print()

print("HYPOTHESIS TESTING (Single Dataset)")
for d in datasets:
    print(f"{d['name']}, Test Statistic (z): {d['test_statistic']:.4f}")
    if d["interval_check"]:
        print(
            f"\tTest statistic is within 95% confidence interval (Fail to reject H0)."
        )
    else:
        print(f"\tTest statistic is outside 95% confidence interval (Reject H0).")

    if d["one_tailed_check"]:
        print(
            f"\tTest statistic suggests proportion is greater than model (Reject H0)."
        )
    else:
        print(f"\tTest statistic does not suggest proportion is greater than model.")
print()

# Joint Hypothesis Test (Part 2, Q16)
joint_test_z = joint_hypothesis_test(datasets[0], datasets[1])
conf_lower, conf_upper = stats.norm.interval(0.95)

print("JOINT HYPOTHESIS TESTING (Marsh vs Galton)")
print(f"Joint test statistic (z): {joint_test_z:.4f}")
if conf_lower < joint_test_z < conf_upper:
    print(
        f"\tStatistic {joint_test_z:.4f} is within ({conf_lower:.4f}, {conf_upper:.4f})."
    )
    print("\tNo significant cultural difference detected.")
else:
    print(
        f"\tStatistic {joint_test_z:.4f} is outside ({conf_lower:.4f}, {conf_upper:.4f})."
    )
    print("\tSignificant cultural difference detected.")
print()

# T-Testing (Part 2, Q18)
t_m, pt_m = t_test(datasets[0], datasets[1], "male")
t_f, pt_f = t_test(datasets[0], datasets[1], "female")
t_ci = stats.t.interval(0.95, df=get_degrees_of_freedom(datasets[0], datasets[1]))

print("T TESTING (By sex across datasets)")
print(f"Male test statistic (t): {t_m:.4f}. P-value: {pt_m:.4f}")
if t_ci[0] < t_m < t_ci[1]:
    print("\tNo significant difference detected between datasets.")
else:
    print("\tSignificant difference detected between datasets.")
print(f"Female test statistic (t): {t_f:.4f}. P-value: {pt_f:.4f}")
if t_ci[0] < t_f < t_ci[1]:
    print("\tNo significant difference detected between datasets.")
else:
    print("\tSignificant difference detected between datasets.")
