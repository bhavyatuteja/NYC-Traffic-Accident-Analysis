import pandas as pd

# Define file paths
raw_excel_file = "C:\Projects\New York Accident\Motor_Vehicle_Collisions_-_Crashes.csv"
cleaned_csv_file = "C:\Projects\New York Accident\Cleaned_NYC_Accidents.csv"

# Load dataset
df = pd.read_csv(raw_excel_file)

# Standardize column names
df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

# Convert crash_date to datetime
df['crash_date'] = pd.to_datetime(df['crash_date'], errors='coerce')

# Extract year, month, day, and hour
df['crash_year'] = df['crash_date'].dt.year
df['crash_month'] = df['crash_date'].dt.month
df['crash_day'] = df['crash_date'].dt.day
df['crash_hour'] = pd.to_datetime(df['crash_time'], format='%H:%M', errors='coerce').dt.hour

# Drop rows with missing latitude or longitude
df.dropna(subset=['latitude', 'longitude'], inplace=True)

# Fill missing boroughs
df['borough'] = df['borough'].fillna('UNKNOWN')

# Convert ZIP Code to string
df['zip_code'] = df['zip_code'].fillna(0).astype(int).astype(str)
df['zip_code'] = df['zip_code'].replace('0', 'UNKNOWN')

# Standardize text fields
df[['borough', 'on_street_name', 'cross_street_name']] = df[['borough', 'on_street_name', 'cross_street_name']].apply(lambda x: x.str.upper().str.strip())

# Drop duplicate entries
df.drop_duplicates(subset=['crash_date', 'latitude', 'longitude', 'contributing_factor_vehicle_1'], keep='first', inplace=True)

# Save cleaned data
df.to_csv(cleaned_csv_file, index=False)

print("Data cleaning complete. Cleaned file saved to", cleaned_csv_file)
