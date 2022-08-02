import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

df_reports = pd.read_excel("./Files/Output/Result.xlsx", sheet_name="Reports")

project_ids = list(df_reports["Project ID"].unique())

for i in range(len(project_ids)):
    df = df_reports[df_reports["Project ID"] == project_ids[i]]
    df = df.drop("Project ID", axis=1)
    df = df.set_index("Date")

    plt.plot(df["ACWP"], label="ACWP", color="red")
    plt.plot(df["BCWP"], label="BCWP", color="green")
    plt.plot(df["BCWS"], label="BCWS", color="purple")
    plt.legend()
    plt.show()
    plt.clf()