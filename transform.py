import pandas as pd
import json

def transform_excel_to_web_data():
    df = pd.read_excel("C:\\Users\\admin\\OneDrive - Gymnázium Zlín - Lesní čtvrť\\90mistza90dni\\data.xlsx").fillna(0)
    
    # Odstraníme řádky a sloupce, které obsahují "součet" (case-insensitive)
    cols_to_keep = [c for c in df.columns if "celkový součet" not in str(c).lower()]
    df = df[cols_to_keep]
    
    # Odstranění řádků, které mají v prvním sloupci "Součet"
    df = df[~df.iloc[:, 0].astype(str).str.contains("celkový součet", case=False, na=False)]
    
    headers = df.columns.tolist()
    first_col = headers[0] # Název prvního sloupce (např. "Místo")
    class_names = headers[1:] # Vše kromě prvního sloupce
    
    result = []
    
    for class_name in class_names:
        # Převedeme sloupec na řetězec a hledáme "1" nebo číslo 1
        visited_spots = []
        
        for _, row in df.iterrows():
            val = row[class_name]
            # Kontrola, zda je hodnota 1 (číslo) nebo "1" (text)
            if str(val).strip() == "1" or val == 1:
                visited_spots.append(str(row[first_col]))
        
        result.append({
            "trida": str(class_name),
            "points": len(visited_spots),
            "spots": visited_spots
        })
    
    # Seřadíme podle počtu bodů sestupně
    result.sort(key=lambda x: x["points"], reverse=True)

    with open("data.js", "w", encoding="utf-8") as f:
        f.write("var dataJson = ")
        json.dump(result, f, indent=4, ensure_ascii=False)
        f.write(";")
    
transform_excel_to_web_data()