from docx import Document
import pandas as pd

doc = Document('Test.docx')

# Create dictionary
data = {
    'fb': [],
    'score': [],
    'anz_ma': []
}

# Append fb
r1 = 1
c1 = 2
data['fb'] = [doc.tables[0].cell(r1, c1).text]*11 + [doc.tables[0].cell(r1, c1+1).text]*11 + [doc.tables[0].cell(r1, c1+2).text]*11 + [doc.tables[0].cell(r1, c1+3).text]*11 + [doc.tables[0].cell(r1, c1+4).text]*11

# Append score
data['score'] = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]*5

# Append anz_ma
data['anz_ma'] = [
    doc.tables[0].cell(r1 + 1, c1 + 0).text,
    doc.tables[0].cell(r1 + 3, c1 + 0).text,
    doc.tables[0].cell(r1 + 5, c1 + 0).text,
    doc.tables[0].cell(r1 + 7, c1 + 0).text,
    doc.tables[0].cell(r1 + 9, c1 + 0).text,
    doc.tables[0].cell(r1 + 11, c1 + 0).text,
    doc.tables[0].cell(r1 + 13, c1 + 0).text,
    doc.tables[0].cell(r1 + 15, c1 + 0).text,
    doc.tables[0].cell(r1 + 17, c1 + 0).text,
    doc.tables[0].cell(r1 + 19, c1 + 0).text,
    doc.tables[0].cell(r1 + 21, c1 + 0).text,

    doc.tables[0].cell(r1 + 1, c1 + 1).text,
    doc.tables[0].cell(r1 + 3, c1 + 1).text,
    doc.tables[0].cell(r1 + 5, c1 + 1).text,
    doc.tables[0].cell(r1 + 7, c1 + 1).text,
    doc.tables[0].cell(r1 + 9, c1 + 1).text,
    doc.tables[0].cell(r1 + 11, c1 + 1).text,
    doc.tables[0].cell(r1 + 13, c1 + 1).text,
    doc.tables[0].cell(r1 + 15, c1 + 1).text,
    doc.tables[0].cell(r1 + 17, c1 + 1).text,
    doc.tables[0].cell(r1 + 19, c1 + 1).text,
    doc.tables[0].cell(r1 + 21, c1 + 1).text,

    doc.tables[0].cell(r1 + 1, c1 + 2).text,
    doc.tables[0].cell(r1 + 3, c1 + 2).text,
    doc.tables[0].cell(r1 + 5, c1 + 2).text,
    doc.tables[0].cell(r1 + 7, c1 + 2).text,
    doc.tables[0].cell(r1 + 9, c1 + 2).text,
    doc.tables[0].cell(r1 + 11, c1 + 2).text,
    doc.tables[0].cell(r1 + 13, c1 + 2).text,
    doc.tables[0].cell(r1 + 15, c1 + 2).text,
    doc.tables[0].cell(r1 + 17, c1 + 2).text,
    doc.tables[0].cell(r1 + 19, c1 + 2).text,
    doc.tables[0].cell(r1 + 21, c1 + 2).text,

    doc.tables[0].cell(r1 + 1, c1 + 3).text,
    doc.tables[0].cell(r1 + 3, c1 + 3).text,
    doc.tables[0].cell(r1 + 5, c1 + 3).text,
    doc.tables[0].cell(r1 + 7, c1 + 3).text,
    doc.tables[0].cell(r1 + 9, c1 + 3).text,
    doc.tables[0].cell(r1 + 11, c1 + 3).text,
    doc.tables[0].cell(r1 + 13, c1 + 3).text,
    doc.tables[0].cell(r1 + 15, c1 + 3).text,
    doc.tables[0].cell(r1 + 17, c1 + 3).text,
    doc.tables[0].cell(r1 + 19, c1 + 3).text,
    doc.tables[0].cell(r1 + 21, c1 + 3).text,

    doc.tables[0].cell(r1 + 1, c1 + 4).text,
    doc.tables[0].cell(r1 + 3, c1 + 4).text,
    doc.tables[0].cell(r1 + 5, c1 + 4).text,
    doc.tables[0].cell(r1 + 7, c1 + 4).text,
    doc.tables[0].cell(r1 + 9, c1 + 4).text,
    doc.tables[0].cell(r1 + 11, c1 + 4).text,
    doc.tables[0].cell(r1 + 13, c1 + 4).text,
    doc.tables[0].cell(r1 + 15, c1 + 4).text,
    doc.tables[0].cell(r1 + 17, c1 + 4).text,
    doc.tables[0].cell(r1 + 19, c1 + 4).text,
    doc.tables[0].cell(r1 + 21, c1 + 4).text]

df = pd.DataFrame.from_dict(data)

print(df)
