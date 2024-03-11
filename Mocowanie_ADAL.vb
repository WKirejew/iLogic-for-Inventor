'Deklaracja bibliotek i zapewnienie braku aktualizacji w momencie późniejszej zmiany parametrów
Components.ContentCenterLanguage = "pl-PL"
Parameter.UpdateAfterChange = False

'Pobranie wymaganych objektów z Inventor API
Dim oParams As Parameters
Dim oApp As Inventor.Application = ThisApplication
'Checking if it's a part
Try
	Dim oPartDoc As PartDocument = ThisDoc.Document
	Dim oPartCompDef As PartComponentDefinition = oPartDoc.ComponentDefinition
	oParams = oPartCompDef.Parameters
Catch
	Exit Try
End Try
'Or an assembly:
Try
	Dim oAssyDoc As AssemblyDocument = ThisDoc.Document
	Dim oAssyCompDef As AssemblyComponentDefinition = oAssyDoc.ComponentDefinition
	oParams = oAssyCompDef.Parameters	
Catch
	Exit Try
End Try

Dim oUserParams As UserParameters = oParams.UserParameters

'Deklaracja parametrów
Try
	p = Parameter("Index")
Catch
	oUserParams.AddByExpression("Index", "0", UnitsTypeEnum.kUnitlessUnits)
End Try

Try
  p = Parameter("i0")
Catch
  oUserParams.AddByValue("i0","0", UnitsTypeEnum.kTextUnits)
End Try

Try
  p = Parameter("i1")
Catch
  oUserParams.AddByValue("i1","0", UnitsTypeEnum.kTextUnits)
End Try

Try
  p = Parameter("i2")
Catch
  oUserParams.AddByValue("i2","0", UnitsTypeEnum.kTextUnits)
End Try

Try
  p = Parameter("i3")
Catch
  oUserParams.AddByValue("i3","0", UnitsTypeEnum.kTextUnits)
End Try

Try
  p = Parameter("i4")
Catch
  oUserParams.AddByValue("i4","0", UnitsTypeEnum.kTextUnits)
End Try

Try
  p = Parameter("i5")
Catch
  oUserParams.AddByValue("i5","0", UnitsTypeEnum.kTextUnits)
End Try

'Dodanie jeden do wartości parametru index
ThisDoc.Document.ComponentDefinition.Parameters.UserParameters.Item("Index").Value = ThisDoc.Document.ComponentDefinition.Parameters.UserParameters.Item("Index").Value + 1
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------INTERFACE PART---------------------------------------------------------------------------------------------------------------------------
'Pobieranie informacji od użytkownika
MultiValue.SetList("i0", 3, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 22)
xx = InputListBox("Wybierz i kliknij OK", MultiValue.List("i0"), i0, Title := "Rozmiar mocowania", ListName := "Dostępne rozmiary")
Dim Mx As String = "M" & xx
Dim Lenght As Double = 0
Dim l0 As Double = 0
Dim l1 As Double = 0

MultiValue.SetList("i1", "(PN)", "(PSN)", "(ŚP-PN)", "(ŚP-PSN)", "(ŚP-PN-PSN)")
Typek = InputListBox("Ś-Śruba, P-Podkładka, N-Nakrętka, S-Podkładka Sprężynująca", MultiValue.List("i1"), i1, Title := "Rodzaj połączenia", ListName := "List")

MultiValue.SetList("i2", "A2", "A4", "pok. Zn", "pok. Ox")
Mtrl= InputListBox("Wybierz materiał z listy", MultiValue.List("i2"), i2, Title := "Rodzaj materiału", ListName := "List")

'Okna dialogowe w wypadku występowania śruby
If Typek <> "(PN)" Then
	If Typek <> "(PSN)" Then
    MultiValue.SetList("i3", "PN-EN ISO 4014", "PN-EN ISO 4017")
    ScrewType = InputListBox("ISO 4014 - z gwintem skróconym, ISO 4017 - z gwintem pełnym.", MultiValue.List("i3"), i3, Title := "Typ elementu złącznego", ListName := "Dostępne typy")
    Lenght = InputBox("L:", "Wybór długości", "50")
  	l0 = InputBox("l0:", "Wybór odległości między podkładkami (skrajnymi):", "20")
	End If

	If Typek = "(ŚP-PN-PSN)" Then
	l1 = InputBox("l1:", "Wybór odległości podkładki środkowej od czołowej:", "10")
	End If
End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'----------------------------------AUTOMATED PART---------------------------------------------------------------------------------------------------------------------------
'Wybór klasy dokładności:
Dim Klasa As String
Dim Klasa1 As String
Select Case Mtrl
Case "A2"
	Klasa = "80"
	Klasa1 = "70"
Case "A4"
	Klasa = "80"
	Klasa1 = "70"
Case "pok. Zn"
	Klasa = "8.8"
	Klasa1 = "8"
Case "pok. Ox"
	Klasa = "8.8"
	Klasa1 = "8"
End Select
'Tworzenie stringów aby wstawić elementy z CC:
Dim finalScrew As String = ScrewType & " - "  & Mx & " x " & Lenght & " - " & Klasa & " - " & Mtrl
Dim finalS As String = xx & " - " & Mtrl
Dim finalP As String = xx & " - " & Mtrl &  " - 200HV"
Dim finalN As String = Mx & " - A/B - " & Mtrl & " - " & Klasa1
Dim nameScrew As String = "Śruba" & Parameter("Index").Value.ToString() & ":1"
Dim nameN As String = "Nakrętka" & Parameter("Index").Value.ToString() & ":1"
Dim nameN1 As String = "Nakrętka" & Parameter("Index").Value.ToString() & ":2"
Dim nameP1 As String = "Podkładka" & Parameter("Index").Value.ToString() & ":1"
Dim nameP2 As String = "Podkładka" & Parameter("Index").Value.ToString() & ":2"
Dim nameP3 As String = "Podkładka" & Parameter("Index").Value.ToString() & ":3"
Dim nameS As String = "PodkładkaS" & Parameter("Index").Value.ToString() & ":1"

'Wstawianie elementów z Content Center	
Select Case Typek
Case "(PN)"
	Dim Nakrętka = Components.AddContentCenterPart(nameN, "Mocowania_ADAL:Nakrętki", "ISO 4032: M3-M30",
	                                                 finalN, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
													 
	Dim PodkładkaB = Components.AddContentCenterPart(nameP1, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)

Case "(PSN)"					
	Dim PodkładkaA = Components.AddContentCenterPart(nameP1, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	
	Dim PodkładkaS = Components.AddContentCenterPart(nameS, "Mocowania_ADAL:Podkładki:Sprężyna", "DIN 7980",
	                                                 finalS, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	
	Dim Nakrętka = Components.AddContentCenterPart(nameN, "Mocowania_ADAL:Nakrętki", "ISO 4032: M3-M30",
	                                                 finalN, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)

Case "(ŚP-PN)"	
	If xx > 12 Then
		If ScrewType = "PN-EN ISO 4014" Then
		finalScrew = "M " & xx & " x " & Lenght
		Else
			finalScrew = ScrewType & " - M" & xx & " x " & Lenght & " - 8.8 - " & Mtrl
		End If
    ScrewType = ScrewType & ": M14 - M30"
		Dim Śruba = Components.AddContentCenterPart(nameScrew, "Mocowania_ADAL:Śruby:Śruby z łbem sześciokątnym", ScrewType,
	                                                 finalScrew, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	Else 
    ScrewType = ScrewType & ": M3 - M12"
		Dim Śruba = Components.AddContentCenterPart(nameScrew, "Mocowania_ADAL:Śruby:Śruby z łbem sześciokątnym", ScrewType,
	                                                 finalScrew, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	End If	
											
	Dim PodkładkaA = Components.AddContentCenterPart(nameP1, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	
	Dim Nakrętka = Components.AddContentCenterPart(nameN, "Mocowania_ADAL:Nakrętki", "ISO 4032: M3-M30",
	                                                 finalN, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
													 
	Dim PodkładkaB = Components.AddContentCenterPart(nameP2, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)

Case "(ŚP-PSN)"	
	If xx > 12 Then
		If ScrewType = "PN-EN ISO 4014" Then
		finalScrew = "M " & xx & " x " & Lenght
		Else
			finalScrew = ScrewType & " - M" & xx & " x " & Lenght & " - 8.8 - " & Mtrl
		End If
		ScrewType = ScrewType & ": M14 - M30"
		Dim Śruba = Components.AddContentCenterPart(nameScrew, "Mocowania_ADAL:Śruby:Śruby z łbem sześciokątnym", ScrewType,
	                                                 finalScrew, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	Else 
    ScrewType = ScrewType & ": M3 - M12"
		Dim Śruba = Components.AddContentCenterPart(nameScrew, "Mocowania_ADAL:Śruby:Śruby z łbem sześciokątnym", ScrewType,
	                                                 finalScrew, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	End If	
											
	Dim PodkładkaA = Components.AddContentCenterPart(nameP1, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	
	Dim PodkładkaS = Components.AddContentCenterPart(nameS, "Mocowania_ADAL:Podkładki:Sprężyna", "DIN 7980",
	                                                 finalS, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	
	Dim Nakrętka = Components.AddContentCenterPart(nameN, "Mocowania_ADAL:Nakrętki", "ISO 4032: M3-M30",
	                                                 finalN, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
													 
	Dim PodkładkaB = Components.AddContentCenterPart(nameP2, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
Case "(ŚP-PN-PSN)"
	If xx > 12 Then
		If ScrewType = "PN-EN ISO 4014" Then
		finalScrew = "M " & xx & " x " & Lenght
		Else
			finalScrew = ScrewType & " - M" & xx & " x " & Lenght & " - 8.8 - " & Mtrl
		End If
    ScrewType = ScrewType & ": M14 - M30"
		Dim Śruba = Components.AddContentCenterPart(nameScrew, "Mocowania_ADAL:Śruby:Śruby z łbem sześciokątnym", ScrewType,
	                                                 finalScrew, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	Else 
    ScrewType = ScrewType & ": M3 - M12"
		Dim Śruba = Components.AddContentCenterPart(nameScrew, "Mocowania_ADAL:Śruby:Śruby z łbem sześciokątnym", ScrewType,
	                                                 finalScrew, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	End If	
											
	Dim PodkładkaA = Components.AddContentCenterPart(nameP1, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	
	Dim PodkładkaS = Components.AddContentCenterPart(nameS, "Mocowania_ADAL:Podkładki:Sprężyna", "DIN 7980",
	                                                 finalS, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
	
	Dim NakrętkaA = Components.AddContentCenterPart(nameN1, "Mocowania_ADAL:Nakrętki", "ISO 4032: M3-M30",
	                                                 finalN, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
													 
	Dim PodkładkaB = Components.AddContentCenterPart(nameP2, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
													 	
	Dim NakrętkaB = Components.AddContentCenterPart(nameN, "Mocowania_ADAL:Nakrętki", "ISO 4032: M3-M30",
	                                                 finalN, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)
													 
	Dim PodkładkaC = Components.AddContentCenterPart(nameP3, "Mocowania_ADAL:Podkładki:Płaszczyzna", "PN-EN ISO 7089:2004",
	                                                 finalP, position := Nothing, grounded := False, 
	                                                 visible := True, appearance := Nothing)	
End Select

'Dodawanie wiązań między elementami		
'\\ MUSZĘ MIEĆ NAZWANE KRAWĘDZIE ABY BYĆ W STANIE UŻYĆ WIĄZANIA WSTAWIAJĄCEGO \\

' Albo zparametryzować grubość podkładki zgodnie z normą ISO 7089:
Dim s As Double = 0
Select Case xx
Case "3" 
	s = 0.5
Case "4"
	s = 0.8
Case "5"
	s = 1
Case "10"
	s = 2
Case "6"
	s = 1.6
Case "8"
	s = 1.6
Case "12"
	s = 2.5
Case "14"
	s = 2.5
Case "16" 
	s = 3
Case "18" 
	s = 3
Case "20" 
	s = 3
Case "22" 
	s = 3
End Select
'Niestety potrzebujemy też sparametryzowanej grubości podkładki sprężystej :<
'Zgodnie z normą DIN 7980:
Dim ss As Double = 0
Select Case xx
Case "3" 
	ss = 1
Case "4"
	ss = 1.2
Case "5"
	ss = 1.6
Case "6"
	ss = 1.6
Case "8"
	ss = 2
Case "10"
	ss = 2.5
Case "12"
	ss = 2.5
Case "14" 
	ss = 3
Case "16" 
	ss = 3.5
Case "20"
	ss = 3.5
Case "18"
	ss = 4.5
Case "22"
	ss = 4.5
End Select

'Aby móc wykorzystać więcej mocowań należy sparametryzować ich nazwy, używając paramnetru index
Dim np_SrP As String = "pł-ŚP" & Parameter("Index").Value.ToString() & ":1"
Dim no_SrP As String = "oś-ŚP" & Parameter("Index").Value.ToString() & ":1"
Dim np_PP As String = "pł-PP" & Parameter("Index").Value.ToString() & ":1"
Dim no_PP As String = "oś-PP" & Parameter("Index").Value.ToString() & ":1"
Dim np_PP1 As String = "pł-PP" & Parameter("Index").Value.ToString() & ":2"
Dim no_PP1 As String = "oś-PP" & Parameter("Index").Value.ToString() & ":2"
Dim np_PS As String = "pł-PS" & Parameter("Index").Value.ToString() & ":1"
Dim no_PS As String = "oś-PS" & Parameter("Index").Value.ToString() & ":1"
Dim np_SN As String = "pł-SN" & Parameter("Index").Value.ToString() & ":1"
Dim np_PN As String = "pł-PN" & Parameter("Index").Value.ToString() & ":1"
Dim no_PN As String = "oś-PN" & Parameter("Index").Value.ToString() & ":1"
Dim no_PN1 As String = "oś-PN" & Parameter("Index").Value.ToString() & ":2"
'Ostatnia rodzina zmiennych to zadeklarowane odległości między podkładkami:
Dim lPP As Double = - l0
Dim lPP1 As Double = - l1

'Dodajmy więc wiązania, wykorzystując osie i płaszczyzny elementów
Select Case Typek
Case "(PN)"
	Constraints.AddMate(np_PN, nameN, "YZ Plane",
	                      nameP1, "YZ PLANE")
	Constraints.AddMate(no_PN, nameN, "X Axis",
	                      nameP1, "X Axis")

Case "(PSN)"	
	Constraints.AddMate(np_PS, nameP1, "YZ Plane",
	                      nameS, "YZ PLANE", (s+ss))
	Constraints.AddMate(no_PS, nameP1, "X Axis",
	                      nameS, "X Axis")
	Constraints.AddMate(np_SN, nameS, "YZ Plane",
	                      nameN, "YZ PLANE")
	Constraints.AddMate(no_PN, nameP1, "X Axis",
	                      nameN, "X Axis")

Case "(ŚP-PN)"	
	Constraints.AddMate(np_SrP, nameScrew, "YZ Plane",
	                      nameP1, "YZ PLANE", s)
	Constraints.AddMate(no_SrP, nameScrew, "X Axis",
	                      nameP1, "X Axis")

	Constraints.AddMate(np_PP, nameP1, "YZ Plane",
	                      nameP2, "YZ PLANE", lPP)
	Constraints.AddMate(no_PP, nameP1, "X Axis",
	                      nameP2, "X Axis")
	Constraints.AddFlush(np_PN, nameP2, "YZ Plane",
	                      nameN, "YZ PLANE", s)
	Constraints.AddMate(no_PN, nameP1, "X Axis",
	                      nameN, "X Axis")

Case "(ŚP-PSN)"	
	Constraints.AddMate(np_SrP, nameScrew, "YZ Plane",
	                      nameP1, "YZ PLANE", s)
	Constraints.AddMate(no_SrP, nameScrew, "X Axis",
	                      nameP1, "X Axis")

	Constraints.AddMate(np_PP, nameP1, "YZ Plane",
	                      nameP2, "YZ PLANE", lPP)
	Constraints.AddMate(no_PP, nameP1, "X Axis",
	                      nameP2, "X Axis")
	Constraints.AddMate(np_PS, nameP2, "YZ Plane",
	                      nameS, "YZ PLANE", (s+ss))
	Constraints.AddMate(no_PS, nameP1, "X Axis",
	                      nameS, "X Axis")
	Constraints.AddMate(np_SN, nameS, "YZ Plane",
	                      nameN, "YZ PLANE")
	Constraints.AddMate(no_PN, nameP1, "X Axis",
	                      nameN, "X Axis")
Case "(ŚP-PN-PSN)"
	Constraints.AddMate(np_SrP, nameScrew, "YZ Plane",
	                      nameP1, "YZ PLANE", s)
	Constraints.AddMate(no_SrP, nameScrew, "X Axis",
	                      nameP1, "X Axis")

	Constraints.AddFlush(np_PN, nameN1, "YZ Plane",
	                      nameP3, "YZ PLANE", -s)
	Constraints.AddMate(no_PN1, nameN1, "X Axis",
	                      nameP3, "X Axis")
	Constraints.AddMate(np_PP1, nameP1, "YZ Plane",
	                      nameP3, "YZ PLANE", (lPP1))                      
	Constraints.AddMate(no_PP1, nameP1, "X Axis",
	                      nameP3, "X Axis")

	Constraints.AddMate(np_PP, nameP1, "YZ Plane",
	                      nameP2, "YZ PLANE", lPP)
	Constraints.AddMate(no_PP, nameP1, "X Axis",
	                      nameP2, "X Axis")
	Constraints.AddMate(np_PS, nameP2, "YZ Plane",
	                      nameS, "YZ PLANE", (s+ss))
	Constraints.AddMate(no_PS, nameP2, "X Axis",
	                      nameS, "X Axis")
	Constraints.AddMate(np_SN, nameS, "YZ Plane",
	                      nameN, "YZ PLANE", offset = ss)
	Constraints.AddMate(no_PN, nameP2, "X Axis",
	                      nameN, "X Axis")
				End Select
				


