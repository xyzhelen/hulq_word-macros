Sub BCsans_to_APA()
'
'to_APA Macro
'
'Last revised 30-Mar-2021 by Helen Zhang
'
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Font.Name = "BC Sans"
	Selection.Find.Replacement.Font.Name = "BC Sans"
	Selection.Find.Format = True
	With Selection.Find
		.Text = "tth" + ChrW(8217)
		.Replacement.Text = ChrW(116) + ChrW(787) + ChrW(7615)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "tth"
		.Replacement.Text = ChrW(116) + ChrW(7615)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "th"
		.Replacement.Text = ChrW(952)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "xw"
		.Replacement.Text = "x" + ChrW(780) + ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "x"
		.Replacement.Text = "x" + ChrW(780)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "s-hw"
		.Replacement.Text = "sx" + ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "t-hw"
		.Replacement.Text = "tx" + ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "s-h"
		.Replacement.Text = "sh"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "t-h"
		.Replacement.Text = "tx"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ch" + ChrW(8217)
		.Replacement.Text = ChrW(269) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "lh"
		.Replacement.Text = ChrW(322)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "sh"
		.Replacement.Text = ChrW(353)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ch"
		.Replacement.Text = ChrW(269)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "hw"
		.Replacement.Text = "x" + ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "oo"
		.Replacement.Text = "$:"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ou"
		.Replacement.Text = "$"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "u"
		.Replacement.Text = ChrW(601)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "t-l" + ChrW(8217)
		.Replacement.Text = "t"+ChrW(108) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "tl" + ChrW(8217)
		.Replacement.Text = ChrW(411) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ts" + ChrW(8217)
		.Replacement.Text = ChrW(99) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ts"
		.Replacement.Text = "c"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "kw" + ChrW(8217)
		.Replacement.Text = "k" + ChrW(787) + ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "qw" + ChrW(8217)
		.Replacement.Text = "q" + ChrW(787) + ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "qw"
		.Replacement.Text = "q" + ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "t-s"
		.Replacement.Text = "ts"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "q" + ChrW(8217)
		.Replacement.Text = "q" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "l" + ChrW(8217)
		.Replacement.Text = "l" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "m" + ChrW(8217)
		.Replacement.Text = "m" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(8217) + "m"
		.Replacement.Text = "m" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "w" + ChrW(8217)
		.Replacement.Text = "w" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(8217) + "w"
		.Replacement.Text = "w" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "aa"
		.Replacement.Text = "a:"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ee"
		.Replacement.Text = "e:"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ii"
		.Replacement.Text = "i:"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "kw"
		.Replacement.Text = "k" + ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "p" + ChrW(8217)
		.Replacement.Text = "p" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "t" + ChrW(8217)
		.Replacement.Text = "t" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(8217) + "n"
		.Replacement.Text = "n" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "n" + ChrW(8217)
		.Replacement.Text = "n" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "l" + ChrW(8217)
		.Replacement.Text = "l" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(8217) + "l"
		.Replacement.Text = "l" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "y" + ChrW(8217)
		.Replacement.Text = "y" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(8217) + "y"
		.Replacement.Text = "y" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(8217)
		.Replacement.Text = ChrW(660)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ":m" + ChrW(787)
		.Replacement.Text = ":m" + ChrW(660)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ":n" + ChrW(787)
		.Replacement.Text = ":n" + ChrW(660)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ":l" + ChrW(787)
		.Replacement.Text = ":l" + ChrW(660)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ":w" + ChrW(787)
		.Replacement.Text = ":w" + ChrW(660)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ":y" + ChrW(787)
		.Replacement.Text = ":y" + ChrW(660)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "$"
		.Replacement.Text = "u"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
End Sub