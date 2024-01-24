import pandas as pd
import os
BUTTON_CONFIG_XLSX_PATH_PERSONAL = os.path.join(os.getcwd(), "share/Button_Config.xlsx")

df = pd.read_excel(BUTTON_CONFIG_XLSX_PATH_PERSONAL)
print(df.columns)