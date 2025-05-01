import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

##Load the Excel file
file_path = "C:/Users/shanm/OneDrive/Desktop/Dealer Wise Particular Part Stock.xlsx"  # Update with your file path
df = pd.read_excel(file_path, sheet_name="Export")
## Set style
sns.set(style="whitegrid")

# Objective 1: Total stock per dealer
dealer_stock = df.groupby("Dealer Name")["Stock QTY"].sum().sort_values(ascending=False)

# Plot 1: Total Stock per Dealer (Top 10)
plt.figure(figsize=(10, 6))
dealer_stock.head(10).plot(kind="barh", color='skyblue')
plt.title("Total Stock per Dealer (Top 10)")
plt.xlabel("Stock Quantity")
plt.gca().invert_yaxis()
plt.tight_layout()
plt.show()

# Objective 2: Zonal stock comparison
zonal_stock = df.groupby("Zonal Office")["Stock QTY"].sum().sort_values(ascending=False)

# Plot 2: Stock by Zonal Office
plt.figure(figsize=(10, 6))
zonal_stock.plot(kind="bar", color='salmon')
plt.title("Stock by Zonal Office")
plt.ylabel("Stock Quantity")
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# Objective 3: Top 10 most stocked parts
top_parts = df.groupby("Part Description")["Stock QTY"].sum().sort_values(ascending=False).head(10)

# Plot 3: Top 10 Most Stocked Parts
plt.figure(figsize=(10, 6))
top_parts.plot(kind="barh", color='lightgreen')
plt.title("Top 10 Most Stocked Parts")
plt.xlabel("Stock Quantity")
plt.gca().invert_yaxis()
plt.tight_layout()
plt.show()

# Objective 4: Scatter plot of Stock vs DLC
stock_vs_dlc = df[["Stock QTY", "DLC"]].dropna()

# Plot 4: Scatter Plot
plt.figure(figsize=(10, 6))
sns.scatterplot(data=stock_vs_dlc, x="DLC", y="Stock QTY", color='purple')
plt.title("Stock QTY vs DLC (Cost)")
plt.xlabel("DLC")
plt.ylabel("Stock QTY")
plt.tight_layout()
plt.show()

# Objective 5: Correlation heatmap
correlation = stock_vs_dlc.corr()

# Plot 5: Correlation Heatmap
plt.figure(figsize=(6, 5))
sns.heatmap(correlation, annot=True, cmap="coolwarm")
plt.title("Correlation Matrix")
plt.tight_layout()
plt.show()

# Objective 6: Boxplot of Stock QTY
plt.figure(figsize=(10, 4))
sns.boxplot(x=df["Stock QTY"], color='orange')
plt.title("Boxplot of Stock QTY")
plt.tight_layout()
plt.show()
