Sub to_Orthography()
'
'to_Orthography Macro
'
'Last revised 12-Dec-2021 by Helen Zhang
'
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Font.Name = "BC Sans"
	'process 1st set of apa -> htg rules
	With Selection.Find
		.Text = ChrW(225)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(233)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(237)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(243)
		.Replacement.Text = "o"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(250)
		.Replacement.Text = "u"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(226) + ChrW(803)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(65) + ChrW(776)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(226)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(224)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(224) + ChrW(803)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(225) + ChrW(803)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(224)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(228)
		.Replacement.Text = "a"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(235)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(233) + ChrW(803)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(234) + ChrW(803)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(232) 
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(233) + ChrW(803)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(224) + ChrW(803)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(203)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601)
		.Replacement.Text = ChrW(601)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601) + ChrW(803) + ChrW(769)
		.Replacement.Text = ChrW(601)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601) + ChrW(770)
		.Replacement.Text = ChrW(601)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601) + ChrW(769)
		.Replacement.Text = ChrW(601)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601) + ChrW(768)
		.Replacement.Text = ChrW(601)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601) + +ChrW(770) + ChrW(803)
		.Replacement.Text = ChrW(601)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601) + ChrW(768) + ChrW(803)
		.Replacement.Text = ChrW(601)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(603)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(603) + ChrW(769)
		.Replacement.Text = "e"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(236)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(239)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(237) + ChrW(803)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(238) + ChrW(803)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(236) + ChrW(803)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "I" + ChrW(776)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(616)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(616) + ChrW(803) + ChrW(768)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(616) + ChrW(803) + ChrW(769)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(616) + ChrW(769)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(616) + ChrW(768)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(616) + ChrW(770)
		.Replacement.Text = "i"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(242) + ChrW(803)
		.Replacement.Text = "o"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(243) + ChrW(803)
		.Replacement.Text = "o"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(244) + ChrW(803)
		.Replacement.Text = "o"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "O" + ChrW(776)
		.Replacement.Text = "o"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(246)
		.Replacement.Text = "o"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(242)
		.Replacement.Text = "o"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(249)
		.Replacement.Text = "u"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(252)
		.Replacement.Text = "u"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(249) + ChrW(803)
		.Replacement.Text = "u"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(251) + ChrW(803)
		.Replacement.Text = "u"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "U" + ChrW(776)
		.Replacement.Text = "u"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(250) + ChrW(803)
		.Replacement.Text = "u"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(121) + ChrW(768)
		.Replacement.Text = "y" + ChrW(787)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "Y" + ChrW(768)
		.Replacement.Text = "y" + ChrW(787)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(116) + ChrW(787) + ChrW(7615)
		.Replacement.Text = "tth" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(116) + ChrW(7615)
		.Replacement.Text = "tth"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(952)
		.Replacement.Text = "th"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "x" + ChrW(780) + ChrW(695)
		.Replacement.Text = "xw"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "x" + ChrW(780)
		.Replacement.Text = "x"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "sx" + ChrW(695)
		.Replacement.Text = "s-hw"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "tx" + ChrW(695)
		.Replacement.Text = "t-hw"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "sh"
		.Replacement.Text = "s-h"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "tx"
		.Replacement.Text = "t-h"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(269) + ChrW(787)
		.Replacement.Text = "ch" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "c" + ChrW(780) + ChrW(787)
		.Replacement.Text = "ch" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(322)
		.Replacement.Text = "lh"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(353)
		.Replacement.Text = "sh"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(269)
		.Replacement.Text = "ch"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "c" + ChrW(780)
		.Replacement.Text = "ch"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "x" + ChrW(695)
		.Replacement.Text = "hw"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(601)
		.Replacement.Text = ChrW(568)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "t"+ChrW(108) + ChrW(787)
		.Replacement.Text = "t-l" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(411) + ChrW(787)
		.Replacement.Text = "tl" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ts"
		.Replacement.Text = "t-s"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(99) + ChrW(787)
		.Replacement.Text = "ts" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "k" + ChrW(787) + ChrW(695)
		.Replacement.Text = "kw" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "q" + ChrW(787) + ChrW(695)
		.Replacement.Text = "qw" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "u:"
		.Replacement.Text = "oo"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "u"
		.Replacement.Text = "ou"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "q" + ChrW(695)
		.Replacement.Text = "qw"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "c"
		.Replacement.Text = "ts"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "q" + ChrW(787)
		.Replacement.Text = "q" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "l" + ChrW(787)
		.Replacement.Text = "l" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "m" + ChrW(787)
		.Replacement.Text = "m" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "w" + ChrW(787)
		.Replacement.Text = "w" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "a:"
		.Replacement.Text = "aa"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "e:"
		.Replacement.Text = "ee"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "i:"
		.Replacement.Text = "ii"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "k" + ChrW(695)
		.Replacement.Text = "kw"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "k" + ChrW(787)
		.Replacement.Text = "k" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "p" + ChrW(787)
		.Replacement.Text = "p" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "t" + ChrW(787)
		.Replacement.Text = "t" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "m" + ChrW(787)
		.Replacement.Text = "m" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "n" + ChrW(787)
		.Replacement.Text = "n" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "l" + ChrW(787)
		.Replacement.Text = "l" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "l" + ChrW(787)
		.Replacement.Text = ChrW(700) + "l"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "w" + ChrW(787)
		.Replacement.Text = "w" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "y" + ChrW(787)
		.Replacement.Text = "y" + ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(660)
		.Replacement.Text = ChrW(700)
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(568)
		.Replacement.Text = "u"
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
'fix glottalized R rules
	Selection.Find.ClearFormatting
	Selection.Find.Format = True
	With Selection.Find
		.Text = ChrW(8217)
		.Replacement.Text = ChrW(700)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "aa"
		.Replacement.Text = ChrW(569)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ee"
		.Replacement.Text = ChrW(1183)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ii"
		.Replacement.Text = ChrW(485)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "im" + ChrW(700) + "i"
		.Replacement.Text = "i" + ChrW(700) + "mi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "in" + ChrW(700) + "i"
		.Replacement.Text = "i" + ChrW(700) + "ni"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "il" + ChrW(700) + "i"
		.Replacement.Text = "i" + ChrW(700) + "li"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "iw" + ChrW(700) + "i"
		.Replacement.Text = "i" + ChrW(700) + "wi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "iy" + ChrW(700) + "i"
		.Replacement.Text = "i" + ChrW(700) + "yi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "em" + ChrW(700) + "i"
		.Replacement.Text = "e" + ChrW(700) + "mi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "em" + ChrW(700) + "a"
		.Replacement.Text = "e" + ChrW(700) + "ma"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "am" + ChrW(700) + "i"
		.Replacement.Text = "a" + ChrW(700) + "mi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "i" + ChrW(700) + "ma"
		.Replacement.Text = "im" + ChrW(700) + "a"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "em" + ChrW(700) + "u"
		.Replacement.Text = "e" + ChrW(700) + "mu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "am" + ChrW(700) + "u"
		.Replacement.Text = "a" + ChrW(700) + "mu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "im" + ChrW(700) + "u"
		.Replacement.Text = "i" + ChrW(700) + "mu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "en" + ChrW(700) + "i"
		.Replacement.Text = "e" + ChrW(700) + "ni"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "en" + ChrW(700) + "a"
		.Replacement.Text = "e" + ChrW(700) + "na"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "an" + ChrW(700) + "i"
		.Replacement.Text = "a" + ChrW(700) + "ni"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "i" + ChrW(700) + "na"
		.Replacement.Text = "in" + ChrW(700) + "a"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "en" + ChrW(700) + "u"
		.Replacement.Text = "e" + ChrW(700) + "nu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "an" + ChrW(700) + "u"
		.Replacement.Text = "a" + ChrW(700) + "nu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "in" + ChrW(700) + "u"
		.Replacement.Text = "i" + ChrW(700) + "nu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "el" + ChrW(700) + "i"
		.Replacement.Text = "e" + ChrW(700) + "li"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "el" + ChrW(700) + "a"
		.Replacement.Text = "e" + ChrW(700) + "la"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "al" + ChrW(700) + "i"
		.Replacement.Text = "a" + ChrW(700) + "li"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "i" + ChrW(700) + "la"
		.Replacement.Text = "il" + ChrW(700) + "a"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "el" + ChrW(700) + "u"
		.Replacement.Text = "e" + ChrW(700) + "lu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "al" + ChrW(700) + "u"
		.Replacement.Text = "a" + ChrW(700) + "lu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "il" + ChrW(700) + "u"
		.Replacement.Text = "i" + ChrW(700) + "lu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ey" + ChrW(700) + "i"
		.Replacement.Text = "e" + ChrW(700) + "yi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ey" + ChrW(700) + "a"
		.Replacement.Text = "e" + ChrW(700) + "ya"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ay" + ChrW(700) + "i"
		.Replacement.Text = "a" + ChrW(700) + "yi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "i" + ChrW(700) + "ya"
		.Replacement.Text = "iy" + ChrW(700) + "a"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ey" + ChrW(700) + "u"
		.Replacement.Text = "e" + ChrW(700) + "yu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ay" + ChrW(700) + "u"
		.Replacement.Text = "a" + ChrW(700) + "yu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "iy" + ChrW(700) + "u"
		.Replacement.Text = "i" + ChrW(700) + "yu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ew" + ChrW(700) + "i"
		.Replacement.Text = "e" + ChrW(700) + "wi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ew" + ChrW(700) + "a"
		.Replacement.Text = "e" + ChrW(700) + "wa"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "aw" + ChrW(700) + "i"
		.Replacement.Text = "a" + ChrW(700) + "wi"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "i" + ChrW(700) + "wa"
		.Replacement.Text = "iw" + ChrW(700) + "a"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "ew" + ChrW(700) + "u"
		.Replacement.Text = "e" + ChrW(700) + "wu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "aw" + ChrW(700) + "u"
		.Replacement.Text = "a" + ChrW(700) + "wu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = "iw" + ChrW(700) + "u"
		.Replacement.Text = "i" + ChrW(700) + "wu"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(700)
		.Replacement.Text = ChrW(8217)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(569)
		.Replacement.Text = "aa"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(1183)
		.Replacement.Text = "ee"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		.Text = ChrW(485)
		.Replacement.Text = "ii"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Text = ""
	Selection.Find.Replacement.Text = ""
End Sub