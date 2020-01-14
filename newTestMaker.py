import zipfile
import os
from docx import Document


d = {}
with open("dati.txt") as f:
    for line in f:
        (key, *val) = line.split(";")
        for e in range(0, len(val)):
            if val[e] == "null":
                val[e] = " "
        d[key] = val


lista = d['NeoremID']
rad = d['NeoremID'][0]
lun = int(d['NeoremID'][1])
lista[0] = rad+"0"
lista[1] = rad+str(1)

for i in range(2, lun+1):
    lista.append(rad + str(i))

print(lista)

for key, value in d.items():
    if(len(value) != lun):
        val = value[len(value)-1]
        for i in range(len(value), lun):
            value.append(val)


def docx_replace(old_file,new_file,rep,index):
    zin = zipfile.ZipFile (old_file, 'r')
    zout = zipfile.ZipFile (new_file, 'w')
    for item in zin.infolist():
        buffer = zin.read(item.filename)
        if (item.filename == 'word/document.xml'):
            res = buffer.decode("utf-8")
            for r in rep:
                res = res.replace(r,rep[r][index])
            buffer = res.encode("utf-8")
        zout.writestr(item, buffer)
    zout.close()
    zin.close()

names = []
for index in range(lun):
    name = "temp"+str(index)+".docx";
    names.append(name)
    docx_replace("demo.docx",name,d,index)

result = Document("./Template.docx")
for name in names:
    doc = Document(name)
    for element in doc.element.body:
        result.element.body.append(element)
result.save('new.docx')

for name in names:
    os.remove(name)