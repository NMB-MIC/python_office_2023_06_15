# %% [markdown]
# ## Excel to report

# %%
# read file path
import os
import datetime
import pandas as pd

print("start")
today = datetime.date.today()
year = today.year-1

path = r"D:\My Documents\Desktop\python_office\src\data\sales_data"
#path = os.getcwd() + "\data\sales_data" # change to my floder name

xlxs_file_lists = []

for root,dirs,files in os.walk(path):
      for name in files:
        file_path = os.path.join(root,name)
        #print(file_path.split("\\")[-2])
        if file_path.split("\\")[-2] == str(year): #change filter
            xlxs_file_lists.append(file_path)
#xlxs_file_lists
for i in range(len(xlxs_file_lists)):
    print(xlxs_file_lists[i].split("\\")[-1])

# %%
#export to dataframe
dfs = []
for f in xlxs_file_lists:
  #print(f.split("\\")[-1])
  df = pd.read_excel(f)
  df["file_name"] = f.split("\\")[-1]
  df["year"] = file_path.split("\\")[-2]
  df["amount_10x"] = df["amount"]*10
  dfs.append(df)
#dfs

# %%
dfs_summary = pd.concat(dfs)

# %%
#dfs_summary

# %%
#dfs_summary.info()

# %%
#dfs_summary.nunique()

# %%
#dfs_summary["plan"].unique()

# %%
#dfs_summary.describe()

# %%
#dfs_summary.head(5)

# %%
#dfs_summary.tail(5)

# %%
#dfs_summary

# %%
pivot = pd.pivot_table(dfs_summary,index="transaction_date",columns="store",values="amount",aggfunc="sum")
#pivot

# %%
summary_monthly = pivot.resample("M").sum()
#summary_monthly

# %%
#!pip install matplotlib
import matplotlib

# %%
fig = summary_monthly.plot(kind="bar",figsize=(20,12),fontsize=26,title="daily summary amount").get_figure()

# %% [markdown]
# xlwings for excel

# %%
#!pip install xlwings

# %%
import xlwings as xw

# %%
import xlwings as xw

import datetime
now = datetime.datetime.now()
date_file_name = f'{str(now.date())}_{str(now.time()).split(".")[0].replace(":","_")}'


template = xw.Book(r"D:\My Documents\Desktop\python_office\src\data\sale_template.xlsx")

app = xw.apps.active
sheet = template.sheets["summary"]
sheet["A1"].value = summary_monthly

pivote = template.sheets["pivot"]
pivote["A1"].value = pivot

#add picture
sheet_report = template.sheets["report"]
sheet_report["A1"].value = "Summary by month"
sheet_report['A1'].font.size = 24
sheet_report["A1"].api.Font.Bold = True
plot= sheet_report.pictures.add(fig,top=sheet["A3"].top,left=sheet["A3"].left)
plot.width = plot.width*0.8
plot.height = plot.height*0.8


template.save(f"""D:\My Documents\Desktop\python_office\src\export\summary_sale_report_{date_file_name}.xlsx""")
#template.save(f"export\summary_sale_report_{date_file_name}.xlsx")
template.close()
app.kill()
print("making report success")