from pathlib import Path  

home = Path.home()  / "Documents/Matrix One/Influence Line"

with open(home / "Config.txt", "r+") as file:
    a = file.readlines()
for i in a: 
    print(i)