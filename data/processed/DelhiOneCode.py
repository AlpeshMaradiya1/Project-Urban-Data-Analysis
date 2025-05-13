import pandas as pd
import os
import glob
import re

# Path to your Excel files
folder_path = r"C:\Users\alpes\OneDrive\Desktop\Delhi's Climate Data Analysis\data\raw"
all_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

monthly_data = []

for file in all_files:
    if os.path.basename(file).startswith('~$'):
        continue

    match = re.search(r'20\d{2}', os.path.basename(file))
    if match:
        year = int(match.group())
    else:
        print(f"❌ Could not extract year from {file}")
        continue

    print(f"📄 Processing file for year {year}...")

    try:
        df = pd.read_excel(file)

        # Check for 'Day' or 'Date' columns
        lower_cols = [str(col).lower() for col in df.columns]
        if 'day' in lower_cols:
            day_column = df.columns[lower_cols.index('day')]
        elif 'date' in lower_cols:
            day_column = df.columns[lower_cols.index('date')]
        else:
            print(f"❌ Skipping {file} — no 'Day' or 'Date' column found.")
            continue

        # Melt the months columns to long format
        df_long = df.melt(id_vars=[day_column], var_name='Month', value_name='AQI')
        df_long['Year'] = year

        # Drop missing AQI values
        df_long.dropna(subset=['AQI'], inplace=True)

        # Group by Year and Month for average
        monthly_avg = df_long.groupby(['Year', 'Month'])['AQI'].mean().reset_index()
        monthly_data.append(monthly_avg)

    except Exception as e:
        print(f"❌ Error processing {file}: {e}")

# Combine and save
if monthly_data:
    final_df = pd.concat(monthly_data, ignore_index=True)

    month_order = ['January', 'February', 'March', 'April', 'May', 'June',
                   'July', 'August', 'September', 'October', 'November', 'December']
    final_df['Month'] = pd.Categorical(final_df['Month'], categories=month_order, ordered=True)
    final_df = final_df.sort_values(['Year', 'Month'])

    output_path = os.path.join(folder_path, '..', 'processed', 'monthly_avg_aqi.csv')
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    final_df.to_csv(output_path, index=False)

    print(f"\n✅ Monthly AQI averages saved to:\n{output_path}")
else:
    print("❌ No data was successfully processed.")
