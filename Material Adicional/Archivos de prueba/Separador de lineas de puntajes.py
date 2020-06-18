psu=open("puntajes.csv","r").read()
ex=open("12000.csv", "w")
psu=psu.split("\n")
i=0
for linea in psu:
    ex.write(linea+"\n")
    i=i+1
    if(i==12000):
        break
ex.close()
