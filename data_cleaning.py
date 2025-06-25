import pandas as pd

# Load SIPRI Excel
data = pd.ExcelFile('SIPRI-Milex-data-1949-2024_2.xlsx')

# Load relevant sheets (adjust skiprows if needed)
constant_df = data.parse('Constant (2023) US$', skiprows=5)
gdp_df = data.parse('Share of GDP', skiprows=5)
percapita_df = data.parse('Per capita', skiprows=5)

# Rename first column
constant_df.rename(columns={constant_df.columns[0]: "Country"}, inplace=True)
gdp_df.rename(columns={gdp_df.columns[0]: "Country"}, inplace=True)
percapita_df.rename(columns={percapita_df.columns[0]: "Country"}, inplace=True)

# Drop Notes column if present
for df in [constant_df, gdp_df, percapita_df]:
    if "Notes" in df.columns:
        df.drop(columns=["Notes"], inplace=True)

# Melt to long format
constant_long = constant_df.melt(id_vars=["Country"], var_name="Year", value_name="Expenditure_2023_USD")
gdp_long = gdp_df.melt(id_vars=["Country"], var_name="Year", value_name="Expenditure_percent_GDP")
percapita_long = percapita_df.melt(id_vars=["Country"], var_name="Year", value_name="Expenditure_per_capita_USD")

# Merge datasets
merged = constant_long.merge(gdp_long, on=["Country", "Year"], how="left")
merged = merged.merge(percapita_long, on=["Country", "Year"], how="left")

# Clean Year column
merged['Year'] = pd.to_numeric(merged['Year'], errors='coerce')

# Fix all numeric fields by converting invalid entries to NaN
merged['Expenditure_2023_USD'] = pd.to_numeric(merged['Expenditure_2023_USD'], errors='coerce')
merged['Expenditure_percent_GDP'] = pd.to_numeric(merged['Expenditure_percent_GDP'], errors='coerce')
merged['Expenditure_per_capita_USD'] = pd.to_numeric(merged['Expenditure_per_capita_USD'], errors='coerce')

# Drop rows missing Year or Country
merged.dropna(subset=['Year', 'Country'], inplace=True)

# Save cleaned file ready for Power BI
output_path = r'C:\Users\Rudransh\Downloads\defense_exp_project\Cleaned_SIPRI_Data_Numeric.xlsx'
merged.to_excel(output_path, index=False)

print(f"Cleaned numeric dataset ready for Power BI saved to: {output_path}")