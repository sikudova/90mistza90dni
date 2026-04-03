import pandas as pd
import json

def transform_excel_to_web_data():
    
    df = pd.read_excel("C:\\Users\\admin\\OneDrive - Gymnázium Zlín - Lesní čtvrť\\90mistza90dni\\data.xlsx", skiprows=1, engine='openpyxl').fillna(0)
        
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
            # Kontrola hodnoty 1 v buňkách (včetně ošetření mezer)
            visited_mask = (df[class_name].astype(str).str.strip() == "1") | (df[class_name] == 1)
            visited_spots = df.loc[visited_mask, first_col].tolist()
            
            # Odstranění případných nul z názvů míst a ořezání
            visited_spots = [str(s).strip() for s in visited_spots if str(s).strip() != "0"]
            
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