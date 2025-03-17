import pandas as pd
import matplotlib.pyplot as plt

# Load the dataset
file_path = "/mnt/data/PDGES-GHGRP-GHGEmissionsGES-2004-Present.xlsx"
xls = pd.ExcelFile(file_path)
ghg_df = pd.read_excel(xls, sheet_name="GHG Emissions GES 2004-2022")

# Select relevant columns
ghg_cleaned = ghg_df[[
    "Reference Year / Année de référence", 
    "Facility Name / Nom de l'installation", 
    "Facility Province or Territory / Province ou territoire de l'installation", 
    "CO2 (tonnes)", "CH4 (tonnes)", "N2O (tonnes)", 
    "Total Emissions (tonnes CO2e) / Émissions totales (tonnes éq. CO2)"
]]

# Rename columns
ghg_cleaned.columns = ["Year", "Facility_Name", "Province", "CO2", "CH4", "N2O", "Total_Emissions"]

# Convert numeric columns
ghg_cleaned[["CO2", "CH4", "N2O", "Total_Emissions"]] = ghg_cleaned[["CO2", "CH4", "N2O", "Total_Emissions"]].apply(pd.to_numeric, errors="coerce")

# Remove missing emission values
ghg_cleaned = ghg_cleaned.dropna(subset=["Total_Emissions"])

# Trend of total emissions over time
yearly_emissions = ghg_cleaned.groupby("Year")["Total_Emissions"].sum()
plt.figure(figsize=(12, 6))
plt.plot(yearly_emissions.index, yearly_emissions.values, marker="o", linestyle="-", color="blue")
plt.xlabel("Year")
plt.ylabel("Total GHG Emissions (Tonnes CO₂-eq)")
plt.title("Trend of Total GHG Emissions (2004-2022)")
plt.grid(True)
plt.show()

# Top emitting provinces
province_emissions = ghg_cleaned.groupby("Province")["Total_Emissions"].sum().sort_values(ascending=False).head(10)
plt.figure(figsize=(12, 6))
plt.barh(province_emissions.index, province_emissions.values, color="darkred")
plt.xlabel("Total GHG Emissions (Tonnes CO₂-eq)")
plt.ylabel("Province")
plt.title("Top 10 Highest-Emitting Provinces in Canada (2004-2022)")
plt.gca().invert_yaxis()
plt.grid(axis="x", linestyle="--", alpha=0.7)
plt.show()

# Top emitting facilities
top_facilities = ghg_cleaned.groupby("Facility_Name")["Total_Emissions"].sum().sort_values(ascending=False).head(10)
plt.figure(figsize=(12, 6))
plt.barh(top_facilities.index, top_facilities.values, color="purple")
plt.xlabel("Total GHG Emissions (Tonnes CO₂-eq)")
plt.ylabel("Facility Name")
plt.title("Top 10 Highest-Emitting Facilities in Canada (2004-2022)")
plt.gca().invert_yaxis()
plt.grid(axis="x", linestyle="--", alpha=0.7)
plt.show()

# GHG emissions breakdown by gas
ghg_breakdown = ghg_cleaned[["CO2", "CH4", "N2O"]].sum()
plt.figure(figsize=(8, 8))
plt.pie(ghg_breakdown, labels=["CO₂", "CH₄", "N₂O"], autopct="%1.1f%%", colors=["blue", "green", "red"])
plt.title("GHG Emissions Breakdown by Gas Type (2004-2022)")
plt.show()

# Save cleaned dataset for Power BI
power_bi_ghg_df = ghg_cleaned[["Year", "Facility_Name", "Province", "CO2", "CH4", "N2O", "Total_Emissions"]]
power_bi_ghg_path = "/mnt/data/ghg_emissions_cleaned.csv"
power_bi_ghg_df.to_csv(power_bi_ghg_path, index=False)
