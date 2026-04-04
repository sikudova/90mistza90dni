import pandas as pd
import json

def transform_excel_to_web_data():
    
    df = pd.read_excel("C:\\Users\\admin\\OneDrive - Gymnázium Zlín - Lesní čtvrť\\90mistza90dni\\data.xlsx", skiprows=1, engine='openpyxl').fillna(0)
        
    cols_to_keep = [c for c in df.columns if "celkový součet" not in str(c).lower()]
    df = df[cols_to_keep]
    
    df = df[~df.iloc[:, 0].astype(str).str.contains("celkový součet", case=False, na=False)]
    
    headers = df.columns.tolist()
    first_col = headers[0]
    class_names = headers[1:]
    
    result = []
    
    for class_name in class_names:
            visited_mask = (df[class_name].astype(str).str.strip() == "1") | (df[class_name] == 1)
            visited_spots = df.loc[visited_mask, first_col].tolist()
            
            visited_spots = [str(s).strip() for s in visited_spots if str(s).strip() != "0"]
            
            result.append({
                "class": str(class_name),
                "points": len(visited_spots),
                "spots": visited_spots
            })
        
    result.sort(key=lambda x: x["points"], reverse=True)

    with open("data.js", "w", encoding="utf-8") as f:
        f.write("var dataJson = ")
        json.dump(result, f, indent=4, ensure_ascii=False)
        f.write(";")
    
transform_excel_to_web_data()