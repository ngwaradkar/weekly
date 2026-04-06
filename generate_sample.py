import pandas as pd

data = {
    "Sr. No.": [1, 2, 3, 4, 5],
    "Line Name": ["Arjun-1", "Arjun-1", "AutoLine", "Arjun-2", "Arjun-2"],
    "Part Number": ["P-101", "P-102", "AL-500", "P-201", "P-202"],
    "Part Description": ["Gear A", "Gear B", "Main Shaft", "Housing", "Cover"],
    "Total Plan Qty": [5000, 3000, 20000, 8000, 4000],
    "Major Setup": [1, 0, 1, 1, 0],
    "Minor Setup": [0, 1, 0, 0, 1]
}

df = pd.DataFrame(data)
df.to_excel("sample_production_plan.xlsx", index=False)
print("Sample Excel created.")
