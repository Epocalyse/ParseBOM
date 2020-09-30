import re
import docx
import pandas as pd

doc = docx.Document('SOPWordCopies/BPBT3107_003BOM.docx')
cells = ['UK SKU', 'SG SKU', 'SKU', 'Supplier']
equipmentData = []
bomData = []
recordEquipment = False
recordMaterial = False
recordReagent = False

for para in doc.paragraphs:
    if para.text == 'PROCEDURE':
        break
    elif para.text == 'Critical reagents' or recordReagent:
        recordReagent = True

        text = re.split(', |\(', para.text)
        cleanText = [text[0]]
        cleanText.extend(word for word in text if any(map(word.__contains__, cells)))

        cleanEntry = [text[0]]
        for tex in text:
            for cell in cells:
                if any(map(tex.__contains__, cells)):
                    if tex.__contains__("UK SKU"):
                        cleanEntry.append(re.split('UK SKU |\)', tex)[1])
                    if tex.__contains__("SG SKU"):
                        # print(re.split('SG SKU:|\s', text)[1])
                        cleanEntry.append(re.split('SG SKU:|\s', tex)[1])
                    if tex.__contains__("SKU"):
                        pass
                        # cleanEntry.extend(text)
                    if tex.__contains__("Supplier"):
                        pass
                        # cleanEntry.extend(text)
                else:
                    # Currently having issues
                    cleanEntry.append("")
        # print("AAAAAAAAAAAAAAAAAAA_________________")
        # print(cleanEntry)
        # print(cleanText)

        row_data = tuple(cleanText)
        bomData.append(row_data)
    elif para.text == 'Equipment':
        recordEquipment = True
    elif recordEquipment:
        equipmentData.append((para.text, 'Equipment'))
    elif para.text == 'Materials' or recordMaterial:
        recordMaterial = True
        equipmentData.append((para.text, 'Material'))

ef = pd.DataFrame(equipmentData)
df = pd.DataFrame(bomData)

# df = df.iloc[1:]
# df.columns = ['MaterialType', 'UKSKU', 'SGSKU']

ef.to_csv('BOMNewCopies/BPBT3107_Equipment.csv', mode='w', header=False)
df.to_csv('BOMNewCopies/BPBT3107_BOM.csv', mode='w', header=False)
