Sub StraightDonna_Unicode()
'
'StraightDonna_Unicode Macro
'
'Last updated: 7-May-2021 by Helen Zhang
'
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Font.Name = "Straight"
	Selection.Find.Replacement.Font.Name = "BC Sans"
	Selection.Find.Format = True
	With Selection.Find
		'middle dot to ae ligature with acute (not to be confused with bullet)
		.Text = ChrW(183)
		.Replacement.Text = ChrW(509)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'a-ring to a
		.Text = ChrW(229)
		.Replacement.Text = "a"
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'A-CIRCUMFLEX  to a-circumflex with underdot
		.Text = ChrW(194)
		.Replacement.Text = ChrW(226) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'A-GRAVE to A-UMLAUT
		.Text = ChrW(192)
		.Replacement.Text = ChrW(65) + ChrW(776)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'trademark to a-circumflex
		.Text = ChrW(8482)
		.Replacement.Text = ChrW(226)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'sterling to a-grave
		.Text = ChrW(163)
		.Replacement.Text = ChrW(224)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'A-UMLAUT to a-grave with underdot
		.Text = ChrW(196)
		.Replacement.Text = ChrW(224) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'A-ACUTE to a-acute with underdot
		.Text = "A" + ChrW(769)
		.Replacement.Text = ChrW(225) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'a-umlaut to a-grave
		.Text = ChrW(228)
		.Replacement.Text = ChrW(224)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'a-grave to a-umlaut
		.Text = ChrW(224)
		.Replacement.Text = ChrW(228)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'c-cedilla to glottalized c
		.Text = ChrW(231)
		.Replacement.Text = ChrW(99) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'partial differential to glottalized c-hacek
		.Text = ChrW(8706)
		.Replacement.Text = ChrW(269) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'increment (not to be confused with Greek Delta)
		.Text = ChrW(8710)
		.Replacement.Text = ChrW(269)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'e-grave to e-umlaut
		.Text = ChrW(232)
		.Replacement.Text = ChrW(235)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'E-ACUTE to e-acute with underdot
		.Text = "E" + ChrW(769)
		.Replacement.Text = ChrW(233) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'E-CIRCUMFLEX to e-circumflex with underdot
		.Text = ChrW(202)
		.Replacement.Text = ChrW(234) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'e-umlaut to e-grave
		.Text = ChrW(235)
		.Replacement.Text = ChrW(232) 
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'E-ACUTE to e-acute with underdot
		.Text = "E" + ChrW(769)
		.Replacement.Text = ChrW(233) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'E-UMLAUT to e-grave with underdot
		.Text = ChrW(203)
		.Replacement.Text = ChrW(224) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'E-GRAVE to E-UMLAUT
		.Text = ChrW(200)
		.Replacement.Text = ChrW(203)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'semicolon to schwa
		.Text = ChrW(59)
		.Replacement.Text = ChrW(601)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'double dagger to schwa-acute with underdot (not to be confused with dagger)
		.Text = ChrW(8225)
		.Replacement.Text = ChrW(601) + ChrW(803) + ChrW(769)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'single left angle bracket to schwa with circumflex
		.Text = ChrW(8249)
		.Replacement.Text = ChrW(601) + ChrW(770)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'currancy to schwa-grave
		.Text = ChrW(164)
		.Replacement.Text = ChrW(601) + ChrW(768)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'ordinal fem to schwa-hacek with underdot
		.Text = ChrW(170)
		.Replacement.Text = ChrW(601) + +ChrW(770) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'ordinal masc to schwa-grave with underdot
		.Text = ChrW(186)
		.Replacement.Text = ChrW(601) + ChrW(768) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'bar to open e
		.Text = ChrW(124)
		.Replacement.Text = ChrW(603)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'degree to open e (not to be confused with ring diacritic)
		.Text = ChrW(176)
		.Replacement.Text = ChrW(603) + ChrW(769)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'i-umlaut to i-grave
		.Text = ChrW(239)
		.Replacement.Text = ChrW(236)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'i-grave to i-umlaut
		.Text = ChrW(236)
		.Replacement.Text = ChrW(239)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'I-ACUTE to i-acute with underdot
		.Text = "I" + ChrW(769)
		.Replacement.Text = ChrW(237) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'I-CIRCUMFLEX to i-circumflex with underdot
		.Text = ChrW(206)
		.Replacement.Text = ChrW(238) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'I-UMLAUT to i-grave with underdot
		.Text = ChrW(207)
		.Replacement.Text = ChrW(236) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'I-GRAVE to I-UMLAUT
		.Text = "I" + ChrW(768)
		.Replacement.Text = "I" + ChrW(776)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'plus minus to barred i-grave with underdot
		.Text = ChrW(177)
		.Replacement.Text = ChrW(616) + ChrW(803) + ChrW(768)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'single low quote to barred i acute with underdot
		.Text = ChrW(8218)
		.Replacement.Text = ChrW(616) + ChrW(803) + ChrW(769)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'right single angle bracket to barred i with acute
		.Text = ChrW(8250)
		.Replacement.Text = ChrW(616) + ChrW(769)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'latin fi ligature to barred i with grave
		.Text = ChrW(64257)
		.Replacement.Text = ChrW(616) + ChrW(768)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'latin fl ligature to barred i with circumflex
		.Text = ChrW(64258)
		.Replacement.Text = ChrW(616) + ChrW(770)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'ring above (not to be confused with degree)
		.Text = ChrW(730)
		.Replacement.Text = ChrW(107) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'logical not to glottalized l
		.Text = ChrW(172)
		.Replacement.Text = ChrW(108) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'registered to l with cross
		.Text = ChrW(174)
		.Replacement.Text = ChrW(322)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'greek mu to glottalized m (not ot be confused with micro sign)
		.Text = ChrW(956)
		.Replacement.Text = ChrW(109) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'integral to glottalized n
		.Text = ChrW(8747)
		.Replacement.Text = ChrW(110) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'O-UMLAUT to o-grave with underdot
		.Text = ChrW(214)
		.Replacement.Text = ChrW(242) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'O-ACUTE to o-acute with underdot
		.Text = "O" + ChrW(769)
		.Replacement.Text = ChrW(243) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'O-CIRCUMFLEX to o-circumflex with underdot
		.Text = ChrW(212)
		.Replacement.Text = ChrW(244) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'O-GRAVE to O-UMLAUT
		.Text = ChrW(210)
		.Replacement.Text = "O" + ChrW(776)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'o-grave to o-umlaut
		.Text = ChrW(242)
		.Replacement.Text = ChrW(246)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'o-umlaut to o-grave
		.Text = ChrW(246)
		.Replacement.Text = ChrW(242)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'greek lowercase pi to glottalized p
		.Text = ChrW(960)
		.Replacement.Text = ChrW(112) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'small ligature oe to glottalized q
		.Text = ChrW(339)
		.Replacement.Text = ChrW(113) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'sharp b to s-hacek
		.Text = ChrW(223)
		.Replacement.Text = ChrW(353)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'dagger to glottalized t (not to be confused with double dagger)
		.Text = ChrW(8224)
		.Replacement.Text = ChrW(116) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'copyright to t with raised theta
		.Text = ChrW(169)
		.Replacement.Text = ChrW(116) + ChrW(7615)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'cedilla to glottalized t with raised theta
		.Text = ChrW(184)
		.Replacement.Text = ChrW(116) + ChrW(787) + ChrW(7615)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'dot accent to glottalized t with raised theta
		.Text = ChrW(729)
		.Replacement.Text = ChrW(116) + ChrW(787) + ChrW(7615)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'u-umlaut to u-grave
		.Text = ChrW(252)
		.Replacement.Text = ChrW(249)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'u-grave to u-umlaut
		.Text = ChrW(249)
		.Replacement.Text = ChrW(252)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'U-UMLAUT to u-grave with underdot
		.Text = ChrW(220)
		.Replacement.Text = ChrW(249) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'U-CIRCUMFLEX to u-circumflex with underdot
		.Text = ChrW(219)
		.Replacement.Text = ChrW(251) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'U-GRAVE to U-UMLAUT
		.Text = ChrW(217)
		.Replacement.Text = "U" + ChrW(776)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'upside down exclamation to u-acute with underdot
		.Text = ChrW(161)
		.Replacement.Text = ChrW(250) + ChrW(803)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'GREEK SIGMA (not to be confused with summation)
		.Text = ChrW(931)
		.Replacement.Text = ChrW(119) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'U-ACUTE to raised w
		.Text = ChrW(85) + ChrW(769)
		.Replacement.Text = ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'o-slash to raised w
		.Text = ChrW(248)
		.Replacement.Text = ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'apple symbol to raised w (apple encoding F8FF)
		.Text = ChrW(63743)
		.Replacement.Text = ChrW(695)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'almost equals to x-hacek
		.Text = ChrW(8776)
		.Replacement.Text = ChrW(120) + ChrW(780)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'ogonek to X-HACEK
		.Text = ChrW(731)
		.Replacement.Text = ChrW(88) + ChrW(780)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'yen to glottalized y
		.Text = ChrW(165)
		.Replacement.Text = ChrW(121) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'y-umlaut to y-grave
		.Text = ChrW(255)
		.Replacement.Text = ChrW(121) + ChrW(768)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'Y-UMLAUT to Y-GRAVE
		.Text = ChrW(376)
		.Replacement.Text = "Y" + ChrW(768)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'GREEK OMEGA (not to be confused with ohm sign)
		.Text = ChrW(937)
		.Replacement.Text = ChrW(122) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'florin to theta
		.Text = ChrW(402)
		.Replacement.Text = ChrW(952)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'lozenge to lambda with cross
		.Text = ChrW(9674)
		.Replacement.Text = ChrW(411)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'sqrt radical to glottalized lambda with cross
		.Text = ChrW(8730)
		.Replacement.Text = ChrW(411) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'division to glottal stop
		.Text = ChrW(247)
		.Replacement.Text = ChrW(660)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'less equal to pharyngeal voice fricative
		.Text = ChrW(8804)
		.Replacement.Text = ChrW(661)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'per thousand sign to unuhw (not to be confused with percent sign)
		.Text = ChrW(8240)
		.Replacement.Text = ChrW(8217)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'grave to umlaut
		.Text = ChrW(96)
		.Replacement.Text = ChrW(168)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'double angle quote right to umlaut
		.Text = ChrW(187)
		.Replacement.Text = ChrW(168)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'double base quotes to superscript glottal stop
		.Text = ChrW(8222)
		.Replacement.Text = ChrW(704)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'umlaut to grave
		.Text = ChrW(168)
		.Replacement.Text = ChrW(96)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'greater equal to middle dot
		.Text = ChrW(8805)
		.Replacement.Text = ChrW(183)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'ellipsis to semicolon
		.Text = ChrW(8230)
		.Replacement.Text = ChrW(59)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'not equal to right arrow
		.Text = ChrW(8800)
		.Replacement.Text = ChrW(8594)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	'clear the formatting because these char don't exist in Straight and reapply replacement font and Format=True
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.Font.Name = "BC Sans"
	Selection.Find.Format = True
	With Selection.Find
		'GREEK DELTA (not to be confused with increment)
		.Text = ChrW(916)
		.Replacement.Text = ChrW(269)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'euro to barred i
		.Text = ChrW(8364)
		.Replacement.Text = ChrW(616)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'fraction slash to shwa-acute (not to be confused with a normal slash)
		.Text = ChrW(8260)
		.Replacement.Text = ChrW(601) + ChrW(769)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'division slash to schwa-acute (not to be confused with a normal slash)
		.Text = ChrW(8725)
		.Replacement.Text = ChrW(601) + ChrW(769)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'micro symbol to glottalized m (not to be confused with Greek mu)
		.Text = ChrW(181)
		.Replacement.Text = ChrW(109) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'ohm sign (not to be confused with Greek Omega)
		.Text = ChrW(8486)
		.Replacement.Text = ChrW(122) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	With Selection.Find
		'summation (not to be confused with greek sigma)
		.Text = ChrW(8721)
		.Replacement.Text = ChrW(119) + ChrW(787)
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	'change all remaining Straight font characters into BC Sans
	Selection.Find.Font.Name = "Straight"
	Selection.Find.Replacement.Font.Name = "BC Sans"
	With Selection.Find
		.Text = ""
		.Replacement.Text = ""
		.Format = True
	End With
	Selection.Find.Execute Replace:=wdReplaceAll
	'clear formatting dialog for the user
	Selection.Find.ClearFormatting
	Selection.Find.Replacement.ClearFormatting
	Selection.Find.Format = False
End Sub