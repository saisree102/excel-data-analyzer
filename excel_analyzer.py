import pandas as pd

print("====== Excel Data Analyzer ======\n")

# Ask user for Excel file
file_name = input("Enter Excel file name (example: data.xlsx): ")

try:
    # Read Excel file
    df = pd.read_excel(file_name)

    print("\nData loaded successfully!\n")
    print("First 5 rows of data:\n")
    print(df.head())

    # Find numeric columns
    numeric_columns = df.select_dtypes(include="number").columns

    if len(numeric_columns) == 0:
        print("\nNo numeric columns found for analysis.")
        exit()

    summary_data = []

    print("\n====== ANALYSIS RESULT ======\n")

    for column in numeric_columns:

        total = df[column].sum()
        average = df[column].mean()
        maximum = df[column].max()
        minimum = df[column].min()

        print(f"Column: {column}")
        print("Total:", total)
        print("Average:", round(average, 2))
        print("Max:", maximum)
        print("Min:", minimum)
        print()

        summary_data.append({
            "Column": column,
            "Total": total,
            "Average": average,
            "Max": maximum,
            "Min": minimum
        })

    # Create summary dataframe
    summary_df = pd.DataFrame(summary_data)

    # Save summary file
    output_file = "analysis_summary.xlsx"
    summary_df.to_excel(output_file, index=False)

    print("Summary file saved as:", output_file)

except FileNotFoundError:
    print("Error: Excel file not found.")
