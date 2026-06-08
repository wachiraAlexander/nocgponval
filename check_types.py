import pandas as pd
df = pd.read_excel("OLT_DATA1.xlsx")
with open("types_output.txt", "w") as f:
    f.write("Types found:\n")
    for t in sorted(df['Type'].dropna().unique()):
        count = (df['Type'] == t).sum()
        f.write(f"  {t}: {count}\n")
print("Done - check types_output.txt")

