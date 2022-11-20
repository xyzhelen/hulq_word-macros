Sub StraightBarbara_Unicode()
'
'Straight Barbara Macro
'
'Last updated: 29-Mar-2021 by Helen Zhang
'
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Replacement.Font.Name = "Times New Roman"
	Selection.Find.Format = True
	Selection.Find.MatchCase = True
	With Selection.Find
		.Text = ChrW(167)
		.Replacement.Text = ChrW(353)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(198)
		.Replacement.Text = ChrW(99) + ChrW(780)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(251)
		.Replacement.Text = "k" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(207)
		.Replacement.Text = ChrW(113) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(195)
		.Replacement.Text = ChrW(411) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(197)
		.Replacement.Text = ChrW(120) + ChrW(780)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(141)
		.Replacement.Text = "c" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(181)
		.Replacement.Text = "m" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(186)
		.Replacement.Text = "n" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(194)
		.Replacement.Text = "l" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(180)
		.Replacement.Text = "y" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(183)
		.Replacement.Text = "w" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(214)
		.Replacement.Text = ChrW(660)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(168)
		.Replacement.Text = ChrW(322)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(59)
		.Replacement.Text = ChrW(601)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(185)
		.Replacement.Text = "p" + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(196)
		.Replacement.Text = ChrW(952)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(165)
		.Replacement.Text = ChrW(183)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(191)
		.Replacement.Text = ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
End Sub