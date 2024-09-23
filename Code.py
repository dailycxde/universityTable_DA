import pandas as pd

df = pd.read_excel("tt.xlsx")
df = df.fillna("missing")

def handleTable(course: str) -> "excel_sheet":
    
    day = []
    time = []
    location = []
    
    for i in range(3, df.shape[0]):
        for j in range(4, df.shape[1]):
            if course in df.iloc[i, j]:
                day.append(df.iloc[0,j])
                time.append(df.iloc[1,j])
                location.append(df.iloc[i,0])
                
    if len(location) == 0:
        return "course does not exist"
    
    day = [i.split()[1][1:-1] for i in day]
    location = [i[-9:-1] for i in location]
    
    with open(f"{course}.txt", "w") as file:
        file.write(f"{course} \n")
        for i in range(len(time)):
            file.write(f"Day: {day[i]} Time: {time[i]} - Location: {location[i]}\n")
            
        print("\nSUCCESSFUL\n")

        
print(">>> WELCOME :) <<<\n")
while True:
    i = input("SELECT A COURSE --> ex: MTH101\n")
    
    if i.lower() == 'q':
        break
    
    if len(i) >= 6:
        handleTable(i)
        print("INPUT Q to EXIT\n")