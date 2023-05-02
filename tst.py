import docx

doc = docx.Document('1-3.docx')

for p in doc.paragraphs:
    print(p.text)
    print("=======\n")

tb=doc.tables[0]
print("rows=" + str(len(tb.rows)))
print("columns=" + str(len(tb.columns)))

cnt=1
for c in tb.columns:
    print(str(cnt)*10)
    cnt+=1
    for ce in c.cells:
        print(ce.text)
        print('==========')
