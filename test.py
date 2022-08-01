import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

df = pd.read_excel("./Files/Output/Result.xlsx", sheet_name="Reports")

project_ids = list(df["Project ID"].unique())

df = df[df["Project ID"] == project_ids[0]]
df = df.drop("Project ID", axis=1)
df = df.set_index("Date")

plt.plot(df["ACWP"], label="ACWP")
plt.plot(df["BCWP"], label="BCWP")
plt.plot(df["BCWS"], label="ACWS")
plt.legend()
plt.show()