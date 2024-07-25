import pandas as pd

def analyze_csab_data(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Rename columns for easier handling
    df.columns = ['Institute', 'Academic_Program', 'Quota', 'Seat_Type', 'Gender', 'Opening_Rank', 'Closing_Rank']

    # List of categories to analyze
    categories = ['Gender-Neutral', 'OPEN', 'OS', 'AI']

    # Filter data for the specified categories and only Gender-Neutral
    df_filtered = df[(df['Seat_Type'].isin(categories)) & (df['Gender'] == 'Gender-Neutral')]

    # Group by Institute and Academic Program, get the minimum Closing Rank
    grouped = df_filtered.groupby(['Institute', 'Academic_Program', 'Seat_Type', 'Opening_Rank', 'Gender', 'Quota'])['Closing_Rank'].min().reset_index()

    # Sort by Closing Rank in ascending order
    grouped_sorted = grouped.sort_values('Closing_Rank')

    return grouped_sorted

def write_results_to_file(results, output_file):
    with open(output_file, 'w') as f:
        for _, row in results.iterrows():
            f.write(f"Institute: {row['Institute']}\n")
            f.write(f"Academic Program: {row['Academic_Program']}\n")
            f.write(f"Seat Type: {row['Seat_Type']}\n")
            f.write(f"Opening Rank: {row['Opening_Rank']}\n")
            f.write(f"Closing Rank: {row['Closing_Rank']}\n")
            f.write(f"Gender: {row['Gender']}\n")
            f.write(f"Quota: {row['Quota']}\n")
            f.write("-" * 50 + "\n")

if __name__ == "__main__":
    input_file = "Book2.xlsx"  # Replace with your file path
    output_file = "book22.txt"
    
    results = analyze_csab_data(input_file)
    write_results_to_file(results, output_file)
    
    print(f"Analysis complete. Results written to {output_file}")
