import csv

with open('../transliteration mappings - ortho to APA.csv', encoding="utf-8") as f:
    data = [tuple(line) for line in csv.reader(f)]

with open('to_APA.bas', "w") as outF:
    outF.write("Sub to_APA()"+"\n")
    outF.write("\'"+"\n")
    outF.write("\'to_APA Macro"+"\n")
    outF.write("\'"+"\n")
    outF.write("\'Last revised 1-July-2021 by Helen Zhang\n")
    outF.write("\'"+"\n")
    outF.write("\t"+"Selection.Find.ClearFormatting"+"\n")
    outF.write("\t"+"Selection.Find.Replacement.ClearFormatting"+"\n")
    #outF.write("\t"+"Selection.Find.Font.Name = \"BC Sans\""+"\n")
    #outF.write("\t"+"Selection.Find.Replacement.Font.Name = \"BC Sans\""+"\n")
    outF.write("\t" + "Selection.Find.Format = True" + "\n")
    for row in data[1:87]:
        outF.write("\tWith Selection.Find"+"\n")
        #outF.write("\t\t'" + row[4]+"\n")
        outF.write("\t\t" + ".Text = "+row[1] + "\n")
        outF.write("\t\t" + ".Replacement.Text = "+row[3] + "\n")
        outF.write("\t"+"End With" + "\n")
        outF.write("\t"+"Selection.Find.Execute Replace:=wdReplaceAll" + "\n")

    outF.write("End Sub")

    outF.close()

