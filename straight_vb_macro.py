import csv

with open('transliteration mappings - straight script.csv', encoding="utf-8") as f:
    data = [tuple(line) for line in csv.reader(f)]

with open('StraightDonna_Unicode.bas', "w") as outF:
    outF.write("Sub StraightDonna_Unicode()"+"\n")
    outF.write("\'"+"\n")
    outF.write("\'StraightDonna_Unicode Macro"+"\n")
    outF.write("\'"+"\n")
    outF.write("\'"+"\n")
    outF.write("\t"+"Selection.Find.ClearFormatting"+"\n")
    outF.write("\t"+"Selection.Find.Replacement.ClearFormatting"+"\n")
    outF.write("\t"+"Selection.Find.Font.Name = \"Straight\""+"\n")
    outF.write("\t"+"Selection.Find.Replacement.Font.Name = \"BC Sans\""+"\n")
    outF.write("\t" + "Selection.Find.Format = True" + "\n")
    outF.write("\tCall Straight_Char()"+"\n")
    outF.write("\tCall Imposter_Char()"+"\n")
    outF.write("End Sub\n")

    outF.write("Sub Straight_Char()\n")
    for row in data[1:87]:
        outF.write("\tWith Selection.Find"+"\n")
        outF.write("\t\t'" + row[4]+"\n")
        outF.write("\t\t" + ".Text = "+row[3] + "\n")
        outF.write("\t\t" + ".Replacement.Text = "+row[1] + "\n")
        #outF.write("\t\t\'set Format=True so it only works on Straight font+\n")
        #outF.write("\t\t" + ".Format = True" + "\n")
        outF.write("\t"+"End With" + "\n")
        outF.write("\t"+"Selection.Find.Execute Replace:=wdReplaceAll" + "\n")
    outF.write("End Sub\n")

    outF.write("Sub Imposter_Char()\n")
    outF.write("\t\'clear the formatting because these char don't exist in Straight and reapply replacement font and Format=True\n")
    outF.write("\t" + "Selection.Find.ClearFormatting" + "\n")
    outF.write("\t" + "Selection.Find.Replacement.Font.Name = \"BC Sans\""+"\n")
    outF.write("\t" + "Selection.Find.Format = True" + "\n")
    for row in data[88:]:
        outF.write("\tWith Selection.Find"+"\n")
        outF.write("\t\t'" + row[4]+"\n")
        outF.write("\t\t" + ".Text = "+row[3] + "\n")
        outF.write("\t\t" + ".Replacement.Text = "+row[1] + "\n")
        outF.write("\t"+"End With" + "\n")
        outF.write("\t"+"Selection.Find.Execute Replace:=wdReplaceAll" + "\n")

    outF.write("\t"+"'change all remaining Straight font characters into BC Sans"+"\n")
    outF.write("\t"+"Selection.Find.Font.Name = \"Straight\""+"\n")
    outF.write("\t"+"Selection.Find.Replacement.Font.Name = \"BC Sans\""+"\n")
    outF.write("\t"+"With Selection.Find"+"\n")
    outF.write("\t\t"+".Text = \"\""+"\n")
    outF.write("\t\t"+".Replacement.Text = \"\""+"\n")
    outF.write("\t\t"+".Format = True"+"\n")
    outF.write("\t"+"End With"+"\n")
    outF.write("\t"+"Selection.Find.Execute Replace:=wdReplaceAll"+"\n")
    outF.write("\t\'clear formatting dialog for the user")
    outF.write("\t"+"Selection.Find.ClearFormatting"+"\n")
    outF.write("\t"+"Selection.Find.Replacement.ClearFormatting"+"\n")
    outF.write("\t"+"Selection.Find.Format = False"+"\n")
    outF.write("End Sub")

    outF.close()

