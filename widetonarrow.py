import pandas as pd
from datetime import datetime

# Generate timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
df = pd.read_csv(rf"C:\Users\91931\OneDrive - Refract Consulting\Inflation_toPivot.csv")
df_long = pd.melt(
    df,
    id_vars=['code']
)

df_long.to_excel(rf"C:\Users\91931\OneDrive - Refract Consulting\inflation_long_format{timestamp}.xlsx", index=False)