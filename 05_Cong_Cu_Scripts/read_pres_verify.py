with open('index.html', 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i in range(2135, 2170):
    if "isVerified" in lines[i]:
        print(f"{i+1}: {lines[i].rstrip()[:120]}")
