import pandas as pd

df = pd.read_excel('112-2開課查詢_2024-03-08.xlsx')

# print(df)
# df2 = df.query("可跨班=='限本班'")
df2 = df[["課程代碼", "課程名稱(中)", "課程性質", "可跨班"]]

crsType = list()
for i in range(len(df2.index)):
    # crsType.append(df2["課程性質"].values[i])
    crsType.append(df2["可跨班"].values[i])

crsType = list(set(crsType))

for item in crsType:
    print(item)
