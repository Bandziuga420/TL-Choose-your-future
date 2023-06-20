import pandas as pd
import math
import scipy.stats as stats

# Path to the Excel file
excel_file_path = r'C:\Users\Aleks\Documents\University\LUISS - Nova SBE Master\Nova SBE\TechLabs (Data Science)\Project\Data base\Py data.xlsx'

# List of sheets
sheet_names = ["Business and Economics",
               "Technology and Science",
               "Literature and Humanities",
               "Arts and Design",
               "Social Sciences",
               "Languages and Linguistics",
               "History and Geography"]

# Explain the methodology
print("Confidence Intervals for the Mean in statistics is a method to assess the extent of values that are likely to contain the true populace mean."
      "This can be accomplished by calculating a margin of error, which is added and subtracted from the sample mean to form the confidence interval."
      "The level of certainty defines the likelihood that the interval contains the true population mean. We will use the confidence interval 90%,"
      "to give a more comprehensive estimate of the parameters. (hence we are 90% sure that our interval represents the statistical population)"

      "And we will consider as relevant only the parameters whose lower limit is above 0.4"
      "(hence that at least 40% of the population agreed with the statement)")

print("") # Add a line break


# Iterate over each sheet
for sheet_name in sheet_names:
    print(f"The following statements are representative for the course {sheet_name}")

    # Read the Excel file
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name, header=1, usecols="B:AM", nrows=34)

    # Exclude the first column (Observation number)
    df = df.iloc[:, 1:]

    # Calculate the 1 - mean (to calculate bernoulli variance), variance, and standard deviation for each column (statement)
    result_data = []
    for col in df.columns:
        column_title = col
        column_data = df[col]

        mean_value = column_data.mean()
        one_minus_mean = 1 - mean_value
        variance = mean_value * one_minus_mean
        standard_deviation = math.sqrt(variance)

        result_data.append({
            "Statement": column_title,
            "Mean": mean_value,
            "1 - Mean": one_minus_mean,
            "Variance": variance,
            "Standard Deviation": standard_deviation
        })

    # Define the confidence interval = 90%
    confidence_interval = float(0.9)

    # Calculate alpha
    alpha = 1 - confidence_interval

    # Calculate the z-score using the standard normal cumulative distribution
    z_score = stats.norm.ppf(1 - alpha)

    # Calculate the margin of error, upper limit, and lower limit for each column (statement)
    for result in result_data:
        mean_value = result["Mean"]
        standard_deviation = result["Standard Deviation"]
        num_observations = df.shape[0]  # Number of rows in the DataFrame

        margin_of_error = (z_score * standard_deviation) / math.sqrt(num_observations)
        upper_limit = mean_value + margin_of_error
        lower_limit = mean_value - margin_of_error

        result["Margin of Error"] = margin_of_error
        result["Upper Limit"] = upper_limit
        result["Lower Limit"] = lower_limit

    # Filter statements with lower limit above 0.4
    filtered_statements = [result["Statement"] for result in result_data if result["Lower Limit"] > 0.4]


    # Display the filtered statements
    if len(filtered_statements) > 0:
        for statement in filtered_statements:
            print("Statement number:", statement)
    else:
        print("No statements have a Lower Limit above 0.4")

    print("")  # Add a line break between sheets
