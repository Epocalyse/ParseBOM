import glob
import os
import pandas as pd
from Convert import Convert
from ExtractBOM import ExtractBOM

# Main file to run everything

print("SOP or BOM? (FullList?)")
fileType = input()
loc = ''

extract_BOM = ExtractBOM()
if fileType == "SOP":
    loc = "SOPWordCopies"
    for file in os.listdir("SOPWordCopies"):
        file = extract_BOM.convertDocx(file, loc)
        Convert.sopDocToCSV(file, loc)
elif fileType.__contains__("BOM"):
    loc = "BOMWordCopies"
    for file in os.listdir("BOMWordCopies"):
        extract_BOM.convertDocx(file, loc)
    for file in os.listdir("BOMWordCopies"):
        Convert.bomDocToCSV(file, loc)
    if fileType.__contains__("FullList"):
        os.chdir("BOMNewCopies")
        extension = 'csv'
        all_filenames = [i for i in glob.glob('*.{}'.format(extension))]
        # combine all files in the list
        combined_csv = pd.concat([pd.read_csv(f) for f in all_filenames])
        # export to csv
        combined_csv.to_csv("combined_BOMs.csv", index=False, encoding='utf-8-sig')
else:
    print("Not a valid file type.")
    exit()
