Attribute VB_Name = "main"    

' ChrW(246) → ö, ChrW(223) → ß, ChrW(228) → ä, ChrW(252) → ü, ChrW(174) → ®, ChrW(8482) → ™, 

' Modell
Private Function ModelDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "E", "Elima-Matic"
        obj.Add "U", "Ultra-Matic"
        obj.Add "V", "V Serie"
        obj.Add "RE", "Air Vantage"
    End If

    Set ModelDictionary = obj
End Function

' Anschlussgröße in Zoll
Private Function ConnectionSizeInchDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "1", "1"
        obj.Add "2", "2"
        obj.Add "3", "3"
        obj.Add "4", "1 1/2"
        obj.Add "4D", "1 1/2"
        obj.Add "40", "1 1/2"
        obj.Add "5", "1/2"
        obj.Add "6", "1/4"
        obj.Add "7", "3/4"
        obj.Add "8", "3/8"
    End If

    Set ConnectionSizeInchDictionary = obj
End Function

' Gehäusematerial (benetzt)
Private Function HousingMaterialWetDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "A", Array("Aluminium", "Aluminum")
        obj.Add "B", Array("Aluminium B", "Aluminum B")
        obj.Add "C", Array("Gusseisen", "Cast Iron")
        obj.Add "G", Array("Leitf" & ChrW(228) & "higes Polypropylen (Acetal)", "Conductive Polypropylene (Acetal)")
        obj.Add "H", Array("Hastelloy C", "Alloy C")
        obj.Add "J", Array("Vernickeltes Aluminium", "Nickel-plated aluminum")
        obj.Add "K", Array("PVDF", "PVDF")
        obj.Add "P", Array("Polypropylen", "Polypropylene")
        obj.Add "Q", Array("Epoxidbeschichtetes Aluminium", "Epoxy-coated aluminum")
        obj.Add "S", Array("Edelstahl", "Stainless Steel")
        obj.Add "Z", Array("PTFE-beschichtetes Aluminium", "PTFE-coated aluminum")
    End If

    Set HousingMaterialWetDictionary = obj
End Function

' Gehäusematerial (nicht benetzt)
Private Function HousingMaterialNotwetDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "A", Array("Aluminium", "Aluminum")
        obj.Add "B", Array("Aluminium B", "Aluminum B")
        obj.Add "C", Array("Gusseisen", "Cast Iron")
        obj.Add "G", Array("Leitf" & ChrW(228) & "higes Polypropylen (Acetal)", "Conductive Polypropylene (Acetal)")
        obj.Add "H", Array("Hastelloy C", "Alloy C")
        obj.Add "J", Array("Vernickeltes Aluminium", "Nickel-plated aluminum")
        obj.Add "K", Array("PVDF", "PVDF")
        obj.Add "P", Array("Polypropylen", "Polypropylene")
        obj.Add "Q", Array("Epoxidbeschichtetes Aluminium", "Epoxy-coated aluminum")
        obj.Add "S", Array("Edelstahl", "Stainless Steel")
        obj.Add "Z", Array("PTFE-beschichtetes Aluminium", "PTFE-coated aluminum")
    End If

    Set HousingMaterialNotwetDictionary = obj
End Function

' Material der Membrane		
Private Function MembraneMaterialDictionary() As Dictionary
    Static obj As Dictionary

    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "1", "Neoprene" & ChrW(174) & " - CR"
        obj.Add "2", "BunaN" & ChrW(174) & " - NBR - Nitrile"
        obj.Add "3", "FKM - Viton" & ChrW(174)
        obj.Add "4", "EPDM - Nordel" & ChrW(8482)
        obj.Add "5", "Teflon" & ChrW(174) & " - PTFE"
        obj.Add "6", "Santoprene" & ChrW(174) & " -  TPE"
        obj.Add "7", "Hytrel" & ChrW(174) & " - TPC"
        obj.Add "9", "Geolast" & ChrW(174)
        obj.Add "Y", "Santoprene" & ChrW(174) & " -  FDA"
    End If

    Set MembraneMaterialDictionary = obj
End Function

' Membranausführung		
Private Function MembraneDesignDictionary() As Dictionary
    Static obj As Dictionary

    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "R", "Versa-Rugged" & ChrW(8482)
        obj.Add "D", "Versa-Dome" & ChrW(8482)
        obj.Add "X", "Thermo-Matic" & ChrW(8482)
        obj.Add "T", "2-piece"
        obj.Add "B", "Versa-Tuff" & ChrW(8482)
        obj.Add "F", "FUSION" & ChrW(8482)
    End If

    Set MembraneDesignDictionary = obj
End Function

' Material Rückschlagventil		
Private Function CheckValveMaterialDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "1", "Neoprene" & ChrW(174) & " - CR"
        obj.Add "2", "BunaN" & ChrW(174) & " - NBR - Nitrile"
        obj.Add "3", "FKM - Viton" & ChrW(174)
        obj.Add "4", "EPDM - Nordel" & ChrW(8482)
        obj.Add "5", "Teflon" & ChrW(174) & " - PTFE"
        obj.Add "6", "Santoprene" & ChrW(174) & " -  TPE"
        obj.Add "7", "Hytrel" & ChrW(174) & " - TPC"
        obj.Add "8", "Polyurethan"
        obj.Add "9", "Geolast" & ChrW(174)
        obj.Add "A", "Acetal"
        obj.Add "S", "Edelstahl"
        obj.Add "Y", "Santoprene" & ChrW(174) & " -  FDA"
        obj.Add "K", "PVDF"
        obj.Add "P", "Polypropylen"
    End If

    Set CheckValveMaterialDictionary = obj
End Function

' Material Ventilsitz		
Private Function ValveSeatMaterialDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "1", "Neoprene" & ChrW(174) & " - CR"
        obj.Add "2", "BunaN" & ChrW(174) & " - NBR - Nitrile"
        obj.Add "3", "FKM - Viton" & ChrW(174)
        obj.Add "4", "EPDM - Nordel" & ChrW(8482)
        obj.Add "5", "Teflon" & ChrW(174) & " - PTFE"
        obj.Add "6", "Santoprene" & ChrW(174) & " -  TPE"
        obj.Add "7", "Hytrel" & ChrW(174) & " - TPC"
        obj.Add "8", "Polyurethan"
        obj.Add "9", "Geolast" & ChrW(174)
        obj.Add "A", "Aluminium"
        obj.Add "S", "Edelstahl"
        obj.Add "C", "Stahl"
        obj.Add "H", "Hastelloy C"
        obj.Add "T", "PTFE-ummanteltes Silikon"
        obj.Add "Y", "Santoprene" & ChrW(174) & " -  FDA"
        obj.Add "P", "Polypropylen"
    End If

    Set ValveSeatMaterialDictionary = obj
End Function

' Gehäuseausführung		
Private Function HousingDesignDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "9", "Geschraubt"
        obj.Add "0", "Geklemmt"
    End If

    Set HousingDesignDictionary = obj
End Function

' Revisionslevel	
Private Function RevisionLevelDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "A", "A"
        obj.Add "B", "B"
        obj.Add "C", "C"
        obj.Add "D", "D"
    End If

    Set RevisionLevelDictionary = obj
End Function
    
' Optionsen		
Private Function OptionsDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "-ATEX", "ATEX Compliant"
        obj.Add "-B", "BSP threaded"
        obj.Add "-CP", "Center Port"
        obj.Add "-DV", "Drop in Bolted Units to Replace VM Clamped (2”+3” Al & SS Only)"
        obj.Add "-FP", "Food Processing"
        obj.Add "-HD", "Horizontal Discharge   "
        obj.Add "-SP", "Sanitary Processing "
        obj.Add "-3A", "3A Sanitary"
        obj.Add "-HP", "High Pressure"
        obj.Add "-DW", "Drop in Bolted Units to Replace Wilden Clamped (2”+3” Al & SS Only)"
        obj.Add "-F", "Flap Valve 2” Al Only)"
        obj.Add "-OB", "Oil Bottle (V-Series Only"
        obj.Add "-SM", "Split Mainfold"
        obj.Add "-UL", "UL Listed"
        obj.Add "-E4", "120VAC Coil"
        obj.Add "-E0", "24VDC Coil "
        obj.Add "-U", "Universal ANSI/DIN Flange"
    End If

    Set OptionsDictionary = obj
End Function

' Min. Temperature of membrane material
Private Function MembraneMaterialMinTempDictionary() As Dictionary
    Static obj As Dictionary

    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "1", "-23"
        obj.Add "2", "-23"
        obj.Add "3", "-40"
        obj.Add "4", "-40"
        obj.Add "5", "-37"
        obj.Add "6", "-40"
        obj.Add "7", "-29"
        obj.Add "9", "-40"
        obj.Add "Y", "-40"
    End If

    Set MembraneMaterialMinTempDictionary = obj
End Function

' Max. Temperature of membrane material
Private Function MembraneMaterialMaxTempDictionary() As Dictionary
    Static obj As Dictionary

    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "1", "93"
        obj.Add "2", "88"
        obj.Add "3", "177"
        obj.Add "4", "138"
        obj.Add "5", "104"
        obj.Add "6", "135"
        obj.Add "7", "104"
        obj.Add "9", "82"
        obj.Add "Y", "135"
    End If

    Set MembraneMaterialMaxTempDictionary = obj
End Function

' Get the model from article number
Private Function GetModelFromArticleNumber(ArticleNumber As String) As Collection
    Dim modelChar As String
    Dim model As String
    Dim remaindArticleNumber As String
    Dim returnCollection As New Collection

    modelChar = Mid(ArticleNumber, 1, 1)
    If ModelDictionary.Exists(modelChar) Then
        model = ModelDictionary.Item(modelChar)
    Else
        model = "Unknown"
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add modelChar
    returnCollection.Add model
    returnCollection.Add remaindArticleNumber
    Set GetModelFromArticleNumber = returnCollection
End Function

' Get the connection size from article number
Private Function GetConnectionSizeFromArticleNumber(ArticleNumber As String) As Collection
    Dim connectionSizeChar As String
    Dim connectionSize As String
    Dim remaindArticleNumber As String
    Dim returnCollection As New Collection

    connectionSizeChar = Mid(ArticleNumber, 1, 1)
    If connectionSizeChar = "4" Then
        connectionSizeChar = Mid(ArticleNumber, 1, 2)
        If ConnectionSizeInchDictionary.Exists(connectionSizeChar) Then
            connectionSize = ConnectionSizeInchDictionary.Item(connectionSizeChar)
            remaindArticleNumber = Mid(ArticleNumber, 3)
            returnCollection.Add connectionSizeChar
            returnCollection.Add connectionSize
            returnCollection.Add remaindArticleNumber
            Set GetConnectionSizeFromArticleNumber = returnCollection
        Else
            connectionSizeChar = Mid(ArticleNumber, 1, 1)
            If ConnectionSizeInchDictionary.Exists(connectionSizeChar) Then
                connectionSize = ConnectionSizeInchDictionary.Item(connectionSizeChar)
            Else
                connectionSize = "Unknown"
            End If
            remaindArticleNumber = Mid(ArticleNumber, 2)
            returnCollection.Add connectionSizeChar
            returnCollection.Add connectionSize
            returnCollection.Add remaindArticleNumber
            Set GetConnectionSizeFromArticleNumber = returnCollection
        End If
    Else
        If ConnectionSizeInchDictionary.Exists(connectionSizeChar) Then
            connectionSize = ConnectionSizeInchDictionary.Item(connectionSizeChar)
        Else
            connectionSize = "Unknown"
        End If
        remaindArticleNumber = Mid(ArticleNumber, 2)
        returnCollection.Add connectionSizeChar
        returnCollection.Add connectionSize
        returnCollection.Add remaindArticleNumber
        Set GetConnectionSizeFromArticleNumber = returnCollection
    End If
End Function

' Get the housing material (wet) from article number
Private Function GetHousingMaterialWetFromArticleNumber(ArticleNumber As String) As Collection
    Dim housingMaterialWetChar As String
    Dim housingMaterialWet As Variant
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    housingMaterialWetChar = Mid(ArticleNumber, 1, 1)
    If HousingMaterialWetDictionary.Exists(housingMaterialWetChar) Then
        housingMaterialWet = HousingMaterialWetDictionary.Item(housingMaterialWetChar)
    Else
        housingMaterialWet = Array("Unbekannt", "Unknown")
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add housingMaterialWetChar
    returnCollection.Add housingMaterialWet
    returnCollection.Add remaindArticleNumber
    Set GetHousingMaterialWetFromArticleNumber = returnCollection
End Function

' Get the housing material (dry) from article number
Private Function GetHousingMaterialNotwetFromArticleNumber(ArticleNumber As String) As Collection
    Dim housingMaterialNotwetChar As String
    Dim housingMaterialNotwet As Variant
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    housingMaterialNotwetChar = Mid(ArticleNumber, 1, 1)
    If HousingMaterialNotwetDictionary.Exists(housingMaterialNotwetChar) Then
        housingMaterialNotwet = HousingMaterialNotwetDictionary.Item(housingMaterialNotwetChar)
    Else
        housingMaterialNotwet =  Array("Unbekannt", "Unknown")
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add housingMaterialNotwetChar
    returnCollection.Add housingMaterialNotwet
    returnCollection.Add remaindArticleNumber
    Set GetHousingMaterialNotwetFromArticleNumber = returnCollection
End Function

' Get the membrane material from article number
Private Function GetMembraneMaterialFromArticleNumber(ArticleNumber As String) As Collection
    Dim membraneMaterialChar As String
    Dim membraneMaterial As String
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    membraneMaterialChar = Mid(ArticleNumber, 1, 1)
    If MembraneMaterialDictionary.Exists(membraneMaterialChar) Then
        membraneMaterial = MembraneMaterialDictionary.Item(membraneMaterialChar)
    Else
        membraneMaterial = "Unknown"
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add membraneMaterialChar
    returnCollection.Add membraneMaterial
    returnCollection.Add remaindArticleNumber
    Set GetMembraneMaterialFromArticleNumber = returnCollection
End Function

' Get the membrane design from article number
Private Function GetMembraneDesignFromArticleNumber(ArticleNumber As String) As Collection
    Dim membraneDesignChar As String
    Dim membraneDesign As String
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    membraneDesignChar = Mid(ArticleNumber, 1, 1)
    If MembraneDesignDictionary.Exists(membraneDesignChar) Then
        membraneDesign = MembraneDesignDictionary.Item(membraneDesignChar)
    Else
        membraneDesign = "Unknown"
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add membraneDesignChar
    returnCollection.Add membraneDesign
    returnCollection.Add remaindArticleNumber
    Set GetMembraneDesignFromArticleNumber = returnCollection
End Function

' Get the check valve material from article number
Private Function GetCheckValveMaterialFromArticleNumber(ArticleNumber As String) As Collection
    Dim checkValveMaterialChar As String
    Dim checkValveMaterial As String
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    checkValveMaterialChar = Mid(ArticleNumber, 1, 1)
    If CheckValveMaterialDictionary.Exists(checkValveMaterialChar) Then
        checkValveMaterial = CheckValveMaterialDictionary.Item(checkValveMaterialChar)
    Else
        checkValveMaterial = "Unknown"
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add checkValveMaterialChar
    returnCollection.Add checkValveMaterial
    returnCollection.Add remaindArticleNumber
    Set GetCheckValveMaterialFromArticleNumber = returnCollection
End Function

' Get the valve seat material from article number
Private Function GetValveSeatMaterialFromArticleNumber(ArticleNumber As String) As Collection
    Dim valveSeatMaterialChar As String
    Dim valveSeatMaterial As String
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    valveSeatMaterialChar = Mid(ArticleNumber, 1, 1)
    If ValveSeatMaterialDictionary.Exists(valveSeatMaterialChar) Then
        valveSeatMaterial = ValveSeatMaterialDictionary.Item(valveSeatMaterialChar)
    Else
        valveSeatMaterial = "Unknown"
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add valveSeatMaterialChar
    returnCollection.Add valveSeatMaterial
    returnCollection.Add remaindArticleNumber
    Set GetValveSeatMaterialFromArticleNumber = returnCollection
End Function

' Get the housing design from article number
Private Function GetHousingDesignFromArticleNumber(ArticleNumber As String) As Collection
    Dim housingDesignChar As String
    Dim housingDesign As String
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    housingDesignChar = Mid(ArticleNumber, 1, 1)
    If HousingDesignDictionary.Exists(housingDesignChar) Then
        housingDesign = HousingDesignDictionary.Item(housingDesignChar)
    Else
        housingDesign = "Unknown"
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add housingDesignChar
    returnCollection.Add housingDesign
    returnCollection.Add remaindArticleNumber
    Set GetHousingDesignFromArticleNumber = returnCollection
End Function

' Get the revision level from article number
Private Function GetRevisionLevelFromArticleNumber(ArticleNumber As String) As Collection
    Dim revisionLevelChar As String
    Dim revisionLevel As String
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    If ArticleNumber = "" Then
        revisionLevelChar = ""
        revisionLevel = "Unknown"
        remainedArticleNumber = ""
        returnCollection.Add revisionLevelChar
        returnCollection.Add revisionLevel
        returnCollection.Add remainedArticleNumber
        Set GetRevisionLevelFromArticleNumber = returnCollection
        Exit Function
    else
        revisionLevelChar = Mid(ArticleNumber, 1, 1)
        if revisionLevelChar = "-" Then
            revisionLevelChar = ""
            revisionLevel = "Unknown"
            remainedArticleNumber = ArticleNumber
        Else
            If RevisionLevelDictionary.Exists(revisionLevelChar) Then
                revisionLevel = RevisionLevelDictionary.Item(revisionLevelChar)
            Else
                revisionLevel = "Unknown"
            End If
            remainedArticleNumber = Mid(ArticleNumber, 2)
        End If
        returnCollection.Add revisionLevelChar
        returnCollection.Add revisionLevel
        returnCollection.Add remainedArticleNumber
        Set GetRevisionLevelFromArticleNumber = returnCollection
    End If
    
End Function

' Get the options from article number
Private Function GetOptionsFromArticleNumber(ArticleNumber As String) As Collection
    Dim options As String
    Dim firstChar As String
    Dim returnCollection As New Collection

    If ArticleNumber = "" Then
        optionsChar = ""
        options = ""
        returnCollection.Add optionsChar
        returnCollection.Add options
        Set GetOptionsFromArticleNumber = returnCollection
        Exit Function
    Else
        firstChar = Mid(ArticleNumber, 1, 1)
        if firstChar = "-" Then
            If InStr(2, ArticleNumber, "-") > 0 Then
                optionsChar = Mid(ArticleNumber, 1, InStr(2, ArticleNumber, "-") - 1)
            Else
                optionsChar = ArticleNumber
            End If
            If OptionsDictionary.Exists(optionsChar) Then
                options = OptionsDictionary.Item(optionsChar)
            Else
                options = ""
            End If
        Else
            optionsChar = ""
            options = ""
        End If

        returnCollection.Add optionsChar
        returnCollection.Add options
        Set GetOptionsFromArticleNumber = returnCollection
    End If
    
End Function

' Get FDA compliance from model, material(wet), membrane material, check valve material, valve seat material, and options
Private Function GetFDACompliance(model As String, meterialWet As String, membraneMaterial As String, checkValveMeterial As String, valveSeatMaterial As String, options As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim returnString As String
    
    ' Set the worksheet where your data is stored
    ' Set ws = ThisWorkbook.Sheets("FDA-Konformität") ' FDA-Konformität
    Set ws = Tabelle10 ' FDA-Konformität
    
    ' Find the last row in the sheet
    ' lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastRow = 102
    
    ' Loop through the rows to find a match
    For i = 3 To lastRow ' Assuming the first row is headers
        ' Debug.Print "Row " & i & ": " & ws.Cells(i, 1).Value & ", " & ws.Cells(i, 3).Value & ", " & ws.Cells(i, 5).Value & ", " & ws.Cells(i, 7).Value & ", " & ws.Cells(i, 8).Value & ", " & ws.Cells(i, 11).Value
        ' Debug.Print "Input: " & model & ", " & meterialWet & ", " & membraneMaterial & ", " & checkValveMeterial & ", " & valveSeatMaterial & ", " & options

        ' Check if the values match
        If ws.Cells(i, 1).Value = model And _
            ws.Cells(i, 3).Value = meterialWet And _
            ws.Cells(i, 5).Value = membraneMaterial And _
            ws.Cells(i, 7).Value = checkValveMeterial And _
            ws.Cells(i, 8).Value = valveSeatMaterial And _
            ws.Cells(i, 11).Value = options Then
            ' Return the corresponding clothes name
            ' Debug.Print "Match found: " & ws.Cells(i, 12).Value
            GetFDACompliance = ws.Cells(i, 12).Value
            Exit Function
        End If
    Next i
    
    ' If no match is found, return the default value "Ohne FDA-Konformität"
    GetFDACompliance = "Ohne FDA-Konformit" & ChrW(228) & "t"
End Function
    
' Get explosion protection from article number
Private Function GetExplosionProtectionFromArticleNumber(ArticleNumber As String) As String
    Dim options As String
    Dim lastChar As String

    If ArticleNumber = "" Then
        GetExplosionProtectionFromArticleNumber = "Ohne Explosionsschutz (ATEX)"
        Exit Function
    Else
        lastChar = Right(ArticleNumber, 5)
        If lastChar = "-ATEX" Then
            GetExplosionProtectionFromArticleNumber = "Inkl. Explosionsschutz (ATEX)"
        Else
            GetExplosionProtectionFromArticleNumber = "Ohne Explosionsschutz (ATEX)"
        End If
    End If
End Function
    
' Get the maximum solid size from connection size, housing material wet and housing design
Private Function GetMaxSolidSize(connectionSize As String, housingMaterialWet As String, housingDesign As String) As String
    Select Case connectionSize
    Case "1"
        GetMaxSolidSize = "3,2"
    Case "2"
        Select Case housingMaterialWet
        Case "S"
            GetMaxSolidSize = "6,0"
        Case "P"
            GetMaxSolidSize = "6,3"
        Case "K"
            GetMaxSolidSize = "6,3"
        Case "A"
            GetMaxSolidSize = "11,1"
        Case Else
            If housingDesign = "0" Then
                GetMaxSolidSize = "6,4"
            Else
                GetMaxSolidSize = "46,0"
            End If
        End Select
    Case "3"
        GetMaxSolidSize = "9,5"
    Case "4"
        GetMaxSolidSize = "4,8"
    Case "40"
        GetMaxSolidSize = "6,0"
    Case "5"
        GetMaxSolidSize = "1,6"
    Case "6"
        GetMaxSolidSize = "1,0"
    Case "7"
        GetMaxSolidSize = "1,6"
    Case "8"
        GetMaxSolidSize = "2,5"
    Case Else
        GetMaxSolidSize = "Unknown"
    End Select

End Function
    
' Get flow rate per stroke
Private Function GetFlowRatePerStroke(connectionSize As String, housingMaterialWet As String, housingDesign As String, options As String) As String
    Select Case connectionSize
    Case "1"
        Select Case housingMaterialWet
        Case "P"
            GetFlowRatePerStroke = "0,42"
        Case "K"
            GetFlowRatePerStroke = "0,42"
        Case Else
            If options = "-ATEX" Then
                GetFlowRatePerStroke = "0,38"
            Else
                GetFlowRatePerStroke = "0,34"
            End If
        End Select
    Case "2"
        Select Case housingMaterialWet
        Case "A"
            GetFlowRatePerStroke = "2,27"
        Case "S"
            GetFlowRatePerStroke = "2,27"
        Case "P"
            GetFlowRatePerStroke = "2,00"
        Case "K"
            GetFlowRatePerStroke = "2,00"
        Case Else
            If housingDesign = "0" Then
                GetFlowRatePerStroke = "2,00"
            Else
                GetFlowRatePerStroke = "1,85"
            End If
        End Select
    Case "3"
        Select Case housingMaterialWet
        Case "K"
            GetFlowRatePerStroke = "3,40"
        Case Else
            If housingDesign = "9" Then
                GetFlowRatePerStroke = "5,50"
            Else
                GetFlowRatePerStroke = "5,10"
            End If
        End Select
    Case "4"
        Select Case options
        Case "-FP"
            GetFlowRatePerStroke = "0,5"
        Case "-3A"
        GetFlowRatePerStroke = "1,1"
        Case Else
            GetFlowRatePerStroke = "1,17"
        End Select
    Case "40"
        GetFlowRatePerStroke = "1,67"
    Case "5"
        GetFlowRatePerStroke = "0,08"
    Case "6"
        GetFlowRatePerStroke = "0,04"
    Case "7"
        GetFlowRatePerStroke = "0,08"
    Case "8"
        GetFlowRatePerStroke = "0,017"
    Case Else
        GetFlowRatePerStroke = "Unknown"
    End Select

End Function

' Get maximum discharge pressure
Private Function GetMaxDischargePressure(connectionSize As String, housingMaterialWet As String, housingMaterialNotwet As String, options As String) As String
    If connectionSize = "6" Or housingMaterialNotwet = "P" Or housingMaterialWet = "P" Or housingMaterialWet = "K" Or options = "-FP" Then
        GetMaxDischargePressure = "6,8"
    Else
        GetMaxDischargePressure = "8,6"
    End If

End Function

' Get conveying capacity
Private Function GetConveyingCapacity(connectionSize As String, housingMaterialWet As String, housingMaterialNotwet As String, membraneMaterial As String, membraneDesign As String, housingDesign As String, options As String) As String
    Select Case connectionSize
    Case "1"
        Select Case housingMaterialWet
        Case "K"
            GetConveyingCapacity = "163"
        Case "P"
            GetConveyingCapacity = "163"
        Case Else
            Select Case housingMaterialNotwet
            Case "A"
                If membraneMaterial = "5" Then
                    GetConveyingCapacity = "144"
                Else
                    GetConveyingCapacity = "185"
                End If
            Case "P"
                If membraneMaterial = "5" Then
                    GetConveyingCapacity = "136"
                Else
                    GetConveyingCapacity = "174"
                End If
            End Select
        End Select
    Case "2"
        Select Case housingMaterialWet
        Case "A"
            If membraneMaterial = "5" Then
                If housingDesign = "9" Then
                    GetConveyingCapacity = "595"
                Else
                    GetConveyingCapacity = "617"
                End If
            Else
                If housingDesign = "9" Then
                    GetConveyingCapacity = "606"
                Else
                    If membraneDesign = "D" Then
                        GetConveyingCapacity = "632"
                    Else
                        GetConveyingCapacity = "700"
                    End If
                End If
            End If
        Case "C"
            Select Case membraneDesign
            Case "F"
                GetConveyingCapacity = "617"
            Case "T"
                GetConveyingCapacity = "617"
            Case "R"
                GetConveyingCapacity = "700"
            Case "X"
                GetConveyingCapacity = "700"
            Case "D"
                GetConveyingCapacity = "632"
            End Select
        Case "K"
            Select Case membraneDesign
            Case "F"
                GetConveyingCapacity = "549"
            Case "T"
                GetConveyingCapacity = "549"
            Case "D"
                GetConveyingCapacity = "670"
            End Select
        Case "P"
            Select Case membraneDesign
            Case "F"
                GetConveyingCapacity = "549"
            Case "T"
                GetConveyingCapacity = "549"
            Case "D"
                GetConveyingCapacity = "670"
            Case "R"
                GetConveyingCapacity = "670"
            End Select
        Case "S"
            Select Case housingMaterialNotwet
            Case "A"
                If membraneMaterial = "5" Then
                    GetConveyingCapacity = "617"
                Else
                    If housingDesign = "9" Then
                        GetConveyingCapacity = "606"
                    Else
                        Select Case membraneDesign
                        Case "T"
                            GetConveyingCapacity = "606"
                        Case "D"
                            GetConveyingCapacity = "632"
                        Case "R"
                            GetConveyingCapacity = "700"
                        Case "X"
                            GetConveyingCapacity = "700"
                        End Select
                    End If
                End If
            Case "J"
                If membraneMaterial = "5" Then
                    If housingDesign = "0" Then
                        GetConveyingCapacity = "617"
                    Else
                        GetConveyingCapacity = "549"
                    End If
                Else
                    Select Case membraneDesign
                    Case "D"
                        GetConveyingCapacity = "632"
                    Case "R"
                        GetConveyingCapacity = "700"
                    Case "X"
                        Select Case options
                        Case "-SP"
                            GetConveyingCapacity = "670"
                        Case "-FP"
                            GetConveyingCapacity = "886"
                        Case "-ATEX"
                            GetConveyingCapacity = "700"
                        End Select
                    End Select
                End If
            Case "P"
                GetConveyingCapacity = "644"
            Case "S"
                If membraneMaterial = "5" Then
                    If housingDesign = "0" Then
                        GetConveyingCapacity = "617"
                    Else
                        GetConveyingCapacity = "549"
                    End If
                Else
                    If housingDesign = "9" Then
                        GetConveyingCapacity = "606"
                    Else
                        If membraneDesign = "D" Then
                            GetConveyingCapacity = "632"
                        Else
                            GetConveyingCapacity = "700"
                        End If
                    End If
                End If
            End Select
        End Select
    Case "3"
        Select Case housingMaterialWet
        Case "A"
            If membraneMaterial = "5" Then
                If membraneDesign = "T" And housingDesign = "9" Then
                    GetConveyingCapacity = "704"
                Else
                    GetConveyingCapacity = "682"
                End If
            Else
                If membraneDesign = "D" Then
                    If housingDesign = "0" Then
                        GetConveyingCapacity = "867"
                    Else
                        GetConveyingCapacity = "682"
                    End If
                Else
                    If housingDesign = "0" Then
                        GetConveyingCapacity = "886"
                    Else
                        GetConveyingCapacity = "1033"
                    End If
                End If
            End If
        Case "C"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "682"
            Else
                GetConveyingCapacity = "886"
            End If
        Case "K"
            GetConveyingCapacity = "1060"
        Case "P"
            GetConveyingCapacity = "1060"
        Case Else
            If membraneMaterial = "5" Then
                If housingDesign = "0" Then
                    GetConveyingCapacity = "682"
                Else
                    If membraneDesign = "F" Then
                        GetConveyingCapacity = "682"
                    Else
                        GetConveyingCapacity = "704"
                    End If                    
                End If
            Else
                If housingDesign = "9" Then
                    If membraneDesign = "D" Then
                        GetConveyingCapacity = "954"
                    Else
                        GetConveyingCapacity = "1033"
                    End If
                Else
                    If membraneDesign = "D" Then
                        GetConveyingCapacity = "867"
                    Else
                        GetConveyingCapacity = "886"
                    End If
                End If
            End If
        End Select
    Case "4"
        Select Case housingMaterialWet
        Case "A"
            GetConveyingCapacity = "268"
        Case "C"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "242"
            Else
                GetConveyingCapacity = "268"
            End If
        Case "K"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "273"
            Else
                GetConveyingCapacity = "284"
            End If
        Case "P"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "272"
            Else
                GetConveyingCapacity = "284"
            End If
        Case "S"
            Select Case housingMaterialNotwet
            Case "A"
                If membraneMaterial = "5" Then
                    If housingDesign = "0" Then
                        GetConveyingCapacity = "242"
                    Else
                        GetConveyingCapacity = "250"
                    End If
                Else
                    If housingDesign = "0" Then
                        GetConveyingCapacity = "268"
                    Else
                        If membraneDesign = "R" Then
                            GetConveyingCapacity = "275"
                        Else
                            GetConveyingCapacity = "276"
                        End If
                    End If
                End If
            Case "J"
                If membraneMaterial = "5" Then
                    If membraneDesign = "T" Then
                        GetConveyingCapacity = "246"
                    Else
                        GetConveyingCapacity = "193"
                    End If
                Else
                    If options = "-SP" Then
                        GetConveyingCapacity = "325"
                    Else
                        GetConveyingCapacity = "268"
                    End If
                End If
            End Select
        End Select
    Case "40"
        Select Case housingMaterialNotwet
        Case "P"
            GetConveyingCapacity = "378"
        Case Else
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "390"
            Else
                GetConveyingCapacity = "465"
            End If
        End Select
    Case "5"
        Select Case housingMaterialWet
        Case "A"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "42"
            Else
                GetConveyingCapacity = "45"
            End If
        Case "K"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "38"
            Else
                GetConveyingCapacity = "42"
            End If
        Case "P"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "38"
            Else
                GetConveyingCapacity = "42"
            End If
        Case "S"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "42"
            Else
                GetConveyingCapacity = "45"
            End If
        End Select
    Case "6"
        GetConveyingCapacity = "19"
    Case "7"
        Select Case housingMaterialNotwet
        Case "A"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "42"
            Else
                GetConveyingCapacity = "45"
            End If
        Case "P"
            If membraneMaterial = "5" Then
                GetConveyingCapacity = "32"
            Else
                GetConveyingCapacity = "45"
            End If
        End Select
    Case "8"
        GetConveyingCapacity = "26"
    Case Else
        GetConveyingCapacity = "Unknown"
    End Select
End Function

' Get connection type
Private  Function GetConnectionType(connectionSize As String, materialWet As String, materialNotwet As String, housingDesign As String, options As String) As String
    Select Case connectionSize
    Case "1"
        Select Case materialWet
        Case "P"
            GetConnectionType = "Flanged - End Ported"
        Case "K"
            GetConnectionType = "Flanged - End Ported"
        Case "A"
            GetConnectionType = "NPT"
        Case Else
            Select Case options
            Case "-B"
                GetConnectionType = "BSP"
            Case "-FP"
                GetConnectionType = "TRI-CLAMP"
            Case "-CP"
                GetConnectionType = "Flanged - Center Ported"
            Case Else
                GetConnectionType = "NPT"
            End Select
        End Select
    Case "2"
        Select Case options
        Case "-B"
            GetConnectionType = "BSP"
        Case "-HD"
            GetConnectionType = "Flanged - Horizontal Discharge"
        Case "-CP"
            GetConnectionType = "Flanged - Center Ported"
        Case "-N"
            GetConnectionType = "NPT"
        Case "-F"
            GetConnectionType = "Universal ANSI/DIN Flange"
        Case "-HP"
            GetConnectionType = "NPT"
        Case "-SP"
            GetConnectionType = "TRI-CLAMP"
        Case "-FP"
            If materialNotwet = "J" Then
                GetConnectionType = "TRI-CLAMP"
            Else
                GetConnectionType = "NPT"
            End If
        Case Else
            Select Case materialWet
            Case "S"
                GetConnectionType = "Flanged - Vertical Discharge"
            Case "K"
                GetConnectionType = "Flanged - End Ported"
            Case "P"
                GetConnectionType = "Flanged - End Ported"
            End Select
        End Select
    Case "3"
        Select Case options
        Case "-B"
            GetConnectionType = "BSP"
        Case "-FP"
            GetConnectionType = "TRI-CLAMP"
        Case Else
            If housingDesign = "9" Then
                If materialWet = "A" Or materialWet = "S" Then
                    GetConnectionType = "Flanged - Horizontal Discharge"
                ElseIf materialWet = "K" Or materialWet = "P" Then
                    GetConnectionType = "Flanged - Center Ported"
                Else
                    GetConnectionType = "NPT"
                End If
            Else
                GetConnectionType = "NPT"
            End If
        End Select
    Case "4"
        If options = "-B" Then
            GetConnectionType = "BSP"
        ElseIf options = "-FP" Or options = "-SP" Or options = "-3A" Then
            GetConnectionType = "TRI-CLAMP"
        ElseIf options = "-HD" Then
            GetConnectionType = "Flanged - Horizontal Discharge"
        Else
            If housingDesign = "0" Then
                If housingMaterial = "S" Then
                    GetConnectionType = "NPT-Vertical Discharge"
                Else
                    GetConnectionType = "NPT"
                End If
            Else
                If housingMaterial = "S" Then
                    GetConnectionType = "Flanged - Vertical Discharge"
                Else
                    GetConnectionType = "Flanged - Center Ported"
                End If
            End If
        End If
    Case "40"
        If options = "-B" Then
            GetConnectionType = "BSP"
        ElseIf options = "-A" Then
            GetConnectionType = "Stainless Steel ANSI flange kit w/ nipple"
        Else
            If housingMaterial = "K" Or housingMaterial = "P" Then
                GetConnectionType = "Flanged - Center Ported"
            Else
                GetConnectionType = "NPT"
            End If
        End If
    Case "4D"
        If options = "-B" Then
            GetConnectionType = "BSP"
        Else
            GetConnectionType = "NPT"
        End If
    Case "5"
        If options = "-FP" Then
            GetConnectionType = "TRI-CLAMP"
        Else
            GetConnectionType = "NPT"
        End If
    Case "6"
        GetConnectionType = "NPT"
    Case "7"
        GetConnectionType = "NPT"
    Case "8"
        GetConnectionType = "NPT"
    End Select
End Function

' Get connection size for suction side
Private  Function GetConnectionSizeForSuction(connectionSize As String, options As String) As String
    Select Case connectionSize
    Case "1"
        If options = "-FP" Then
            GetConnectionSizeForSuction = "'1 1/2"
        Else
            GetConnectionSizeForSuction = "'1"
        End If
    Case "2"
        If options = "-FP" Or options = "-SP" Then
            GetConnectionSizeForSuction = "'2 1/2"
        Else
            GetConnectionSizeForSuction = "'2"
        End If
    Case "3"
        GetConnectionSizeForSuction = "'3"
    Case "4"
        If options = "-FP" Or options = "-SP" Then
            GetConnectionSizeForSuction = "'2"
        Else
            GetConnectionSizeForSuction = "'1 1/2"
        End If
    Case "4D"
        GetConnectionSizeForSuction = "'1 1/4"
    Case "40"
        GetConnectionSizeForSuction = "'1 1/2"
    Case "5"
        If options = "-FP" Then
            GetConnectionSizeForSuction = "'1 1/2"
        Else
            GetConnectionSizeForSuction = "'1/2"
        End If
    Case "6"
        GetConnectionSizeForSuction = "'1/4"
    Case "7"
        GetConnectionSizeForSuction = "'3/4"
    Case "8"
        GetConnectionSizeForSuction = "'3/8"
    End Select
End Function

' Get connection size for pressure
Private  Function GetConnectionSizeForPressure(connectionSize As String, options As String)
    Select Case connectionSize
    Case "1"
        If options = "-FP" Then
            GetConnectionSizeForPressure = "'1 1/2"
        Else
            GetConnectionSizeForPressure = "'1"
        End If
    Case "2"
        If options = "-FP" Or options = "-SP" Then
            GetConnectionSizeForPressure = "'2 1/2"
        Else
            GetConnectionSizeForPressure = "'2"
        End If
    Case "3"
        GetConnectionSizeForPressure = "'3"
    Case "4"
        If options = "-FP" Or options = "-SP" Then
            GetConnectionSizeForPressure = "'2"
        Else
            GetConnectionSizeForPressure = "'1 1/2"
        End If
    Case "4D"
        GetConnectionSizeForPressure = "'1 1/2"
    Case "40"
        GetConnectionSizeForPressure = "'1 1/2"
    Case "5"
        If options = "-FP" Then
            GetConnectionSizeForPressure = "'1 1/2"
        Else
            GetConnectionSizeForPressure = "'1/2"
        End If
    Case "6"
        GetConnectionSizeForPressure = "'1/4"
    Case "7"
        GetConnectionSizeForPressure = "'3/4"
    Case "8"
        GetConnectionSizeForPressure = "'3/8"
    End Select
End Function

' Get suction height wetted
Private  Function GetSuctionHeightWetted(connectionSize As String, housingMaterialWet As String, housingDesign As String, options As String) As String
    Select Case connectionSize
    Case "1"
        GetSuctionHeightWetted = "9,4"
    Case "2"
        If housingMaterialWet = "S" Then
            GetSuctionHeightWetted = "9,1"
        Else
            If options = "-FP" Or options = "-SP" Or options = "-3A" Then
                GetSuctionHeightWetted = "9,5"
            Else
                GetSuctionHeightWetted = "9,8"
            End If
        End If
    Case "3"
        If housingDesign = "0" Then
            GetSuctionHeightWetted = "9,8"
        Else
            GetSuctionHeightWetted = "9,4"
        End If
    Case "4"
        GetSuctionHeightWetted = "7,6"
    Case "40"
        GetSuctionHeightWetted = "9,4"
    Case "5"
        GetSuctionHeightWetted = "6,7"
    Case "6"
        GetSuctionHeightWetted = "n/a"
    Case "7"
        GetSuctionHeightWetted = "6,7"
    Case "8"
        GetSuctionHeightWetted = "n/a"
    End Select
End Function

' Get suction height dry
Private  Function GetSuctionHeightDry(connectionSize As String, housingMaterialWet As String, housingDesign As String, options As String) As String
    Select Case connectionSize
    Case "1"
        If options = "-ATEX" Then
            GetSuctionHeightDry = "4,6"
        Else
            GetSuctionHeightDry = "4,9"
        End If
    Case "2"
        If options = "-FP" Or options = "-SP" Or options = "-3A" Then
            GetSuctionHeightDry = "5,1"
        Else
            If housingMaterialWet = "S" Then
                GetSuctionHeightDry = "4,3"
            ElseIf housingMaterialWet = "K" Or housingMaterialWet = "P" Then
                GetSuctionHeightDry = "4,9"
            Else
                If housingDesign = "0" Then
                    GetSuctionHeightDry = "5,2"
                Else
                    GetSuctionHeightDry = "5,5"
                End If
            End If
        End If
    Case "3"
        If housingDesign = "0" Then
            GetSuctionHeightDry = "6,1"
        Else
            GetSuctionHeightDry = "4,9"
        End If
    Case "4"
        If housingDesign = "0" Then
            GetSuctionHeightDry = "4,6"
        Else
            GetSuctionHeightDry = "4,5"
        End If
    Case "40"
        GetSuctionHeightDry = "5,8"
    Case "5"
        GetSuctionHeightDry = "3,9"
    Case "6"
        GetSuctionHeightDry = "2,4"
    Case "7"
        GetSuctionHeightDry = "3,9"
    Case "8"
        GetSuctionHeightDry = "2,4"
    End Select
End Function
    
' Get air connection inlet
Private  Function GetAirConnectionInlet(connectionSize As String, housingMaterialWet As String, options As String) As String
    Select Case connectionSize
    Case "1"
        If options = "-HP" Then
            GetAirConnectionInlet = "'1"
        ElseIf options = "-FP" Then
            GetAirConnectionInlet = "'1/8"
        Else
            GetAirConnectionInlet = "'3/8"
        End If
    Case "2"
        If options = "-HP"Then
            GetAirConnectionInlet = "'3/4"
        Else
            If housingMaterialWet = "K" Or housingMaterialWet = "P" Then
                GetAirConnectionInlet = "'3/8"
            Else
                GetAirConnectionInlet = "'1/2"
            End If
        End If
    Case "3"
        If housingMaterialWet = "K" Or housingMaterialWet = "P" Then
            GetAirConnectionInlet = "'3/4"
        Else
            GetAirConnectionInlet = "'1/2"
        End If
    Case "4"
        If options = "-HD" Then
            GetAirConnectionInlet = "'3/4"
        Else
            GetAirConnectionInlet = "'1/2"
        End If
    Case "40"
        If housingMaterialWet = "K" Or housingMaterialWet = "P" Then
            GetAirConnectionInlet = "'3/4"
        Else
            GetAirConnectionInlet = "'1/2"
        End If
    Case "5"
        If options = "-FP" Then
            GetAirConnectionInlet = "'1/8"
        Else
            GetAirConnectionInlet = "'3/8"
        End If
    Case "6"
        If housingMaterialWet = "K" Or housingMaterialWet = "P" Then
            GetAirConnectionInlet = "'3/8"
        Else
            GetAirConnectionInlet = "'1/4"
        End If
    Case "7"
        GetAirConnectionInlet = "'3/8"
    Case "8"
        GetAirConnectionInlet = "'1/4"
    End Select
End Function

' Get air connection outlet
Private  Function GetAirConnectionOutlet(connectionSize As String, options As String) As String
    Select Case connectionSize
    Case "1"
        If options = "-HP" Then
            GetAirConnectionOutlet = "'1"
        Else
            GetAirConnectionOutlet = "'1/2"
        End If
    Case "2"
        If options = "-HP" Then
            GetAirConnectionOutlet = "'3/4"
        Else
            GetAirConnectionOutlet = "'1"
        End If
    Case "3"
        GetAirConnectionOutlet = "'1"
    Case "4"
        GetAirConnectionOutlet = "'3/4"
    Case "40"
        GetAirConnectionOutlet = "'1"
    Case "5"
        If options = "-FP" Then
            GetAirConnectionOutlet = "'1/8"
        Else
            GetAirConnectionOutlet = "'3/8"
        End If
    Case "6"
        GetAirConnectionOutlet = "'1/4"
    Case "7"
        GetAirConnectionOutlet = "'3/8"
    Case "8"
        GetAirConnectionOutlet = "'1/4"
    End Select
End Function

' Get Min. Temperature
Private Function GetMinTemperature(membraneMaterialChar As String) As String
    Dim returnString As String

    If MembraneMaterialMinTempDictionary.Exists(membraneMaterialChar) Then
        returnString = MembraneMaterialMinTempDictionary.Item(membraneMaterialChar)
    Else
        returnString = "Unknown"
    End If

    GetMinTemperature = returnString
End Function

' Get Max. Temperature
Private Function GetMaxTemperature(membraneMaterialChar As String) As String
    Dim returnString As String

    If MembraneMaterialMaxTempDictionary.Exists(membraneMaterialChar) Then
        returnString = MembraneMaterialMaxTempDictionary.Item(membraneMaterialChar)
    Else
        returnString = "Unknown"
    End If

    GetMaxTemperature = returnString
End Function

' Get weight, length, width, and heigth
Private  Function GetWeightLengthWidthHeight(connectionSize As String, materialWet As String, materialNotwet As String, housingDesign As String, connectionType As String, options As String) As Collection
    Dim weight As String
    Dim length As String
    Dim width As String
    Dim height As String
    Dim returnCollection As New Collection
    
    Select Case connectionSize
    Case "1"
        Select Case materialWet
        Case "A"
            If materialNotwet = "P" Then
                weight = "10"
                length = "20,5"
                width = "27,2"
                height = "36,9"
            ElseIf materialNotwet = "A" Then
                weight = "12,2"
                length = "20,5"
                width = "27,2"
                height = "36,9"
            End If
        Case "S"
            weight = "18,1"
            length = "20,5"
            width = "27,2"
            height = "36,9"
        Case "K"
            weight = "11,8"
            length = "20,6"
            width = "34,3"
            height = "43"
        Case "P"
            weight = "7,7"
            length = "20,6"
            width = "34,3"
            height = "43"
        End Select
    Case "2"
        Select Case materialWet
        Case "A"
            If housingDesign = "9" Then
                weight = "36,7"
                length = "33,4"
                width = "45"
                height = "67,2"
            Else
                weight = "29,5"
                length = "34,5"
                width = "41,6"
                height = "67,8"
            End If
        Case "C"
            weight = "51,3"
            length = "34,5"
            width = "41,6"
            height = "66,5"
        Case "K"
            If connectionType = "Flanged End Ported" Then
                weight = "42,2"
                length = "30,5"
                width = "50,3"
                height = "76,8"
            Else
                weight = "45,4"
                length = "30,5"
                width = "44,5"
                height = "76,8"
            End If
        Case "P"
            If connectionType = "Flanged Center Ported" Then
                weight = "33,1"
                length = "30,5"
                width = "44,5"
                height = "76,8"
            ElseIf connectionType = "Flanged End Ported" Then
                weight = "31,3"
                length = "30,5"
                width = "50,3"
                height = "76,8"
            Else
                weight = "26"
                length = "50,3"
                width = "30,5"
                height = "76,9"
            End If
        Case "S"
            If materialNotwet = "S" Then
                If options = "-SP" Or options = "-3A" Then
                    weight = "74,8"
                    length = "43,4"
                    width = "47,2"
                    height = "88,1"
                ElseIf options = "-FP" Then
                    weight = "66,2"
                    length = "34,5"
                    width = "43,6"
                    height = "66,6"
                End If
            ElseIf materialNotwet = "J" Then
                If options = "-SP" Then
                    weight = "60,8"
                    length = "43,4"
                    width = "47,2"
                    height = "88,1"
                ElseIf options = "-3A" Then
                    weight = "60,8"
                    length = "34,5"
                    width = "43,6"
                    height = "66,6"
                End If
            Else
                If options = "-HD" Then
                    weight = "65,8"
                    length = "30,5"
                    width = "45"
                    height = "71,1"
                Else
                    weight = "65,8"
                    length = "30,5"
                    width = "45"
                    height = "70,6"
                End If
            End If
        End Select
    Case "3"
        Select Case materialWet
        Case "A"
            If housingDesign = "0" Then
                weight = "49"
                length = "38,1"
                width = "50,8"
                height = "81,5"
            Else
                weight = "76,2"
                length = "40,9"
                width = "63,8"
                height = "92,2"
            End If
        Case "C"
            weight = "76,2"
            length = "38,1"
            width = "50,8"
            height = "83,3"
        Case "K"
            weight = "123"
            length = "46,3"
            width = "84,1"
            height = "103,2"
        Case "P"
            weight = "94"
            length = "46,3"
            width = "84,1"
            height = "103,2"
        Case "S"
            If options = "-FP" Then
                If materialNotwet = "J" Then
                    weight = "91"
                    length = "43,1"
                    width = "54,7"
                    height = "81,3"
                ElseIf materialNotwet = "A" Then
                    weight = "86"
                    length = "43,1"
                    width = "54,7"
                    height = "81,3"
                ElseIf materialNotwet = "S" Then
                    weight = "109"
                    length = "43,1"
                    width = "54,7"
                    height = "81,3"
                End If
            Else
                If materialNotwet = "A" Or materialNotwet = "S" Then
                    weight = "76,2"
                    length = "30,5"
                    width = "50,8"
                    height = "81,3"
                Else
                    weight = "76,2"
                    length = "41,2"
                    width = "56"
                    height = "92,1"
                End If
            End If
        End Select
    Case "4"
        Select Case materialWet
        Case "A"
            weight = "18,61"
            length = "29,2"
            width = "36"
            height = "43,5"
        Case "C"
            weight = "25,87"
            length = "29,2"
            width = "36,5"
            height = "42,9"
        Case "K"
            weight = "18,6"
            length = "30,9"
            width = "39,9"
            height = "54,9"
        Case "P"
            If materialNotwet = "P" Then
                weight = "18"
                length = "30,9"
                width = "39,9"
                height = "54,9"
            ElseIf materialNotwet = "A" Then
                weight = "25"
                length = "30,9"
                width = "39,9"
                height = "54,9"
            End If
        Case "S"
            If materialNotwet = "A" Then
                If options = "-HD" Then
                    weight = "29,5"
                    length = "31"
                    width = "33,5"
                    height = "51,3"
                Else
                    weight = "29,5"
                    length = "29,2"
                    width = "36,6"
                    height = "42,6"
                End If
            ElseIf materialNotwet = "J" Then
                If options = "-SP" Then
                    weight = "25,87"
                    length = "39,9"
                    width = "35,4"
                    height = "80,2"
                ElseIf options = "-3A" Then
                    weight = "25,87"
                    length = "29,2"
                    width = "36,2"
                    height = "45,7"
                ElseIf options = "-FP" Then
                    weight = "25,87"
                    length = "29,2"
                    width = "42,3"
                    height = "44"
                End If
            End If
        End Select
    Case "40"
        Select Case materialWet
        Case "A"
            weight = "25"
            length = "31"
            width = "47,1"
            height = "56,4"
        Case "P"
            weight = "37"
            length = "33"
            width = "58,4"
            height = "73"
        Case "K"
            weight = "51"
            length = "33"
            width = "58,4"
            height = "73"
        Case Else
            weight = "42"
            length = "31"
            width = "47,1"
            height = "56,4"
        End Select
    Case "5"
        Select Case materialWet
        Case "A"
            weight = "3,9"
            length = "15,9"
            width = "21,3"
            height = "25,5"
        Case "K"
            weight = "5,4"
            length = "15,9"
            width = "23,6"
            height = "29,7"
        Case "P"
            weight = "3,9"
            length = "15,9"
            width = "23,6"
            height = "29,7"
        Case "S"
            If materialNotwet = "A" Then
                weight = "17"
                length = "15,9"
                width = "21,3"
                height = "25,5"
            ElseIf materialNotwet = "P" Then
                If options = "-FP" Then
                    weight = "17"
                    length = "15,9"
                    width = "20,8"
                    height = "28,3"
                Else
                    weight = "18"
                    length = "15,10"
                    width = "21,4"
                    height = "25,6"
                End If
            End If
        End Select
    Case "6"
        If materialWet = "K" Or materialWet = "P" Then
            weight = "1,8"
            length = "17,8"
            width = "14"
            height = "19,8"
        ElseIf materialWet = "G" Then
            weight = "1,8"
            length = "13,9"
            width = "19,1"
            height = "20,1"
        End If
    Case "7"
        weight = "3,9"
        length = "15,9"
        width = "21,3"
        height = "25,5"
    Case "8"
        If materialWet = "K" Then
            weight = "2"
            length = "10,4"
            width = "14,4"
            height = "13,5"
        ElseIf materialWet = "P" Then
            weight = "1,4"
            length = "10,4"
            width = "14,4"
            height = "13,5"
        End If
    End Select

    returnCollection.Add weight
    returnCollection.Add length
    returnCollection.Add width
    returnCollection.Add height
    Set GetWeightLengthWidthHeight = returnCollection
End Function

Sub Main()
    '''''''''''''''''''''
    ' Initializing Data '
    '''''''''''''''''''''
    ' Modell
    Dim modelWS As Worksheet, modelTable As Range, modelData As Variant
    Set modelWS = Tabelle23
    Set modelTable = modelWS.Range("A3:B6")
    modelData = modelTable.Value

    ' Anschlussgröße
    Dim connSizeWS As Worksheet, connSizeTable As Range, connSizeData As Variant
    Set connSizeWS = Tabelle7
    Set connSizeTable = connSizeWS.Range("A3:G19")
    connSizeData = connSizeTable.Value

    ' Gehäusematerial (benetzt)
    Dim housingWetWS As Worksheet, housingWetTable As Range, housingWetData As Variant
    Set housingWetWS = Tabelle15
    Set housingWetTable = housingWetWS.Range("A3:C13")
    housingWetData = housingWetTable.Value

    ' Gehäusematerial (nicht benetzt)
    Dim housingNotwetWS As Worksheet, housingNotwetTable As Range, housingNotwetData As Variant
    Set housingNotwetWS = Sheet4
    Set housingNotwetTable = housingNotwetWS.Range("A3:C13")
    housingNotwetData = housingNotwetTable.Value

    ' Material der Membrane		
    Dim memMaterialWS As Worksheet, memMaterialTable As Range, memMaterialData As Variant
    Set memMaterialWS = Tabelle12
    Set memMaterialTable = memMaterialWS.Range("A3:C11")
    memMaterialData = memMaterialTable.Value

    ' Membranausführung		
    Dim memDesignWS As Worksheet, memDesignTable As Range, memDesignData As Variant
    Set memDesignWS = Tabelle24
    Set memDesignTable = memDesignWS.Range("A3:C8")
    memDesignData = memDesignTable.Value

    ' Material Rückschlagventil		
    Dim checkValveWS As Worksheet, checkValveTable As Range, checkValveData As Variant
    Set checkValveWS = Tabelle13
    Set checkValveTable = checkValveWS.Range("A3:C16")
    checkValveData = checkValveTable.Value

    ' Material Ventilsitz		
    Dim valveSeatWS As Worksheet, valveSeatTable As Range, valveSeatData As Variant
    Set valveSeatWS = Tabelle14
    Set valveSeatTable = valveSeatWS.Range("A3:C18")
    valveSeatData = valveSeatTable.Value

    ' Gehäuseausführung		
    Dim housingDesignWS As Worksheet, housingDesignTable As Range, housingDesignData As Variant
    Set housingDesignWS = Tabelle9
    Set housingDesignTable = housingDesignWS.Range("A3:C4")
    housingDesignData = housingDesignTable.Value

    ' Revisionslevel	
    Dim revisionWS As Worksheet, revisionTable As Range, revisionData As Variant
    Set revisionWS = Tabelle26
    Set revisionTable = revisionWS.Range("A3:B7")
    revisionData = revisionTable.Value

    ' Optionsen		
    Dim optionsWS As Worksheet, optionsTable As Range, optionsData As Variant
    Set optionsWS = Tabelle27
    Set optionsTable = optionsWS.Range("A3:C19")
    optionsData = optionsTable.Value

    ' FDA-Konformität												
    Dim FDAWS As Worksheet, FDATable As Range, FDAData As Variant
    Set FDAWS = Tabelle10
    Set FDATable = FDAWS.Range("A3:O102")
    FDAData = FDATable.Value

    ' Exlosionsschutz (ATEX)	
    Dim explosionWS As Worksheet, explosionTable As Range, explosionData As Variant
    Set explosionWS = Tabelle11
    Set explosionTable = explosionWS.Range("A3:F3")
    explosionData = explosionTable.Value

    ' Maximale Feststoffgröße							
    Dim maxSolidSizeWS As Worksheet, maxSolidSizeTable As Range, maxSolidSizeData As Variant
    Set maxSolidSizeWS = Tabelle8
    Set maxSolidSizeTable = maxSolidSizeWS.Range("A3:H17")
    maxSolidSizeData = maxSolidSizeTable.Value

    ' Fördermenge per Hub							
    Dim flowRateWS As Worksheet, flowRateTable As Range, flowRateData As Variant
    Set flowRateWS = Tabelle18
    Set flowRateTable = flowRateWS.Range("A3:H23")
    flowRateData = flowRateTable.Value

    ' Maximaler Förderdruck						
    Dim maxDischargePressureWS As Worksheet, maxDischargePressureTable As Range, maxDischargePressureData As Variant
    Set maxDischargePressureWS = Tabelle17
    Set maxDischargePressureTable = maxDischargePressureWS.Range("A3:G11")
    maxDischargePressureData = maxDischargePressureTable.Value

    ' Förderleistung										
    Dim conveyCapacityWS As Worksheet, conveyCapacityTable As Range, conveyCapacityData As Variant
    Set conveyCapacityWS = Tabelle16
    Set conveyCapacityTable = conveyCapacityWS.Range("A3:K205")
    conveyCapacityData = conveyCapacityTable.Value

    ' Anschlusstyp								
    Dim connectionTypeWS As Worksheet, connectionTypeTable As Range, connectionTypeData As Variant
    Set connectionTypeWS = Tabelle21
    Set connectionTypeTable = connectionTypeWS.Range("A3:I81")
    connectionTypeData = connectionTypeTable.Value

    ' Anschlussgröße (Saugseite)
    Dim connSizeSuctionWS As Worksheet, connSizeSuctionTable As Range, connSizeSuctionData As Variant
    Set connSizeSuctionWS = Sheet5
    Set connSizeSuctionTable = connSizeSuctionWS.Range("A3:F19")
    connSizeSuctionData = connSizeSuctionTable.Value

    ' Anschlussgröße (Druckseite)
    Dim connSizePressureWS As Worksheet, connSizePressureTable As Range, connSizePressureData As Variant
    Set connSizePressureWS = Sheet6
    Set connSizePressureTable = connSizePressureWS.Range("A3:F19")
    connSizePressureData = connSizePressureTable.Value

    ' Ansaughöhe (nass)
    Dim suctionHeightWetWS As Worksheet, suctionHeightWetTable As Range, suctionHeightWetData As Variant
    Set suctionHeightWetWS = Tabelle19
    Set suctionHeightWetTable = suctionHeightWetWS.Range("A3:H16")
    suctionHeightWetData = suctionHeightWetTable.Value

    ' Ansaughöhe (trocken)
    Dim suctionHeightDryWS As Worksheet, suctionHeightDryTable As Range, suctionHeightDryData As Variant
    Set suctionHeightDryWS = Tabelle29
    Set suctionHeightDryTable = suctionHeightDryWS.Range("A3:H21")
    suctionHeightDryData = suctionHeightDryTable.Value

    ' Luftanschluss (Eingang)
    Dim airConnInletWS As Worksheet, airConnInletTable As Range, airConnInletData As Variant
    Set airConnInletWS = Tabelle30
    Set airConnInletTable = airConnInletWS.Range("A3:G24")
    airConnInletData = airConnInletTable.Value

    ' Luftanschluss (Ausgang)
    Dim airConnOutletWS As Worksheet, airConnOutletTable As Range, airConnOutletData As Variant
    Set airConnOutletWS = Tabelle20
    Set airConnOutletTable = airConnOutletWS.Range("A3:F14")
    airConnOutletData = airConnOutletTable.Value

    ' Gewicht-Abmessungen
    Dim dimensionsWS As Worksheet, dimensionsTable As Range, dimensionsData As Variant
    Set dimensionsWS = Tabelle22
    Set dimensionsTable = dimensionsWS.Range("A3:L64")
    dimensionsData = dimensionsTable.Value

    ' Temperatur - Material der Membrane			
    Dim memMaterialTempWS As Worksheet, memMaterialTempTable As Range, memMaterialTempData As Variant
    Set memMaterialTempWS = Tabelle31
    Set memMaterialTempTable = memMaterialTempWS.Range("A3:D11")
    memMaterialTempData = memMaterialTempTable.Value

    ' Set worksheet references
    Dim wsInput As Worksheet, wsOutput As Worksheet, wsSeoOutput As Worksheet
    Set wsInput = ThisWorkbook.Sheets("INPUT")
    Set wsOutput = ThisWorkbook.Sheets("OUTPUT")
    Set wsSeoOutput = ThisWorkbook.Sheets("SEO OUTPUT")

    Dim lastRow As Long, i As Long
    Dim articleNum As String, remainedArticleNum As String

    Dim modelChar As String, model As String, remaindArticleNumber As String
    Dim connSizeChar As String, connSizeInch As String, connSizeMM As String
    Dim housingWetChar As String, housingWetDE As String, housingWetEN As String
    Dim housingNotwetChar As String, housingNotwetDE As String, housingNotwetEN As String
    Dim memMaterialChar As String, memMaterialDE As String, memMaterialEN As String
    Dim memDesignChar As String, memDesignDE As String, memDesignEN As String
    Dim checkValveChar As String, checkValveDE As String, checkValveEN As String
    Dim valveSeatChar As String, valveSeatDE As String, valveSeatEN As String
    Dim housingDesignChar As String, housingDesignDE As String, housingDesignEN As String
    Dim revisionChar As String, revision As String
    Dim optionOneChar As String, optionOne As String
    Dim optionTwoChar As String, optionTwo As String
    Dim FDAComplianceDE As String, FDAComplianceEN As String
    Dim explosionDE As String, explosionEN As String
    Dim maxSolidSize As String, flowRate As String, maxDischargePressure As String, conveyCapacity As String, connectionType As String, connSizeSuction As String, connSizePressure As String, suctionHeightWet As String, suctionHeightDry As String, airConnInlet As String, airConnOutlet As String
    Dim weight As String, length As String, width As String, height As String  
    Dim memMaterialTempMin As String, memMaterialTempMax As String
    Dim outputRow As Long
    
    ' Find last row in INPUT sheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    ' Loop through each article number
    outputRow = 2 ' Start from row 2 in OUTPUT sheet

    ' For i = 5 To lastRow
    For i = 5 To 5
        articleNum = wsInput.Cells(i, 1).Value ' Read article number

        ' Get parameters from article number
        modelChar = Mid(articleNum, 1, 1)
        remainedArticleNumber = Mid(articleNum, 2)

        connSizeChar = Mid(remainedArticleNumber, 1, 1)
        If connSizeChar = "4" Then
            connSizeChar = Mid(remainedArticleNumber, 1, 2)
            If connSizeChar = "4D" Or connSizeChar = "40" Then
                remainedArticleNumber = Mid(remainedArticleNumber, 3)
            Else
                connSizeChar = Mid(remainedArticleNumber, 1, 1)
                remainedArticleNumber = Mid(remainedArticleNumber, 2)
            End If
        Else
            remainedArticleNumber = Mid(remainedArticleNumber, 2)
        End If

        housingWetChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        housingNotwetChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        memMaterialChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        memDesignChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        checkValveChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        valveSeatChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        housingDesignChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        If Len(remainedArticleNumber) > 0 Then
            revisionChar = Mid(remainedArticleNumber, 1, 1)
            If revisionChar = "-" Then
                revisionChar = ""
                remainedArticleNumber = Mid(remainedArticleNumber, 1)
            Else
                remainedArticleNumber = Mid(remainedArticleNumber, 2)
            End If
        Else
            revisionChar = ""
        End If

        If Len(remainedArticleNumber) > 0 Then
            firstChar = Mid(remainedArticleNumber, 1, 1)
            If firstChar = "-" Then
                If InStr(2, remainedArticleNumber, "-") > 0 Then
                    optionOneChar = Mid(remainedArticleNumber, 1, InStr(2, remainedArticleNumber, "-") - 1)
                    optionTwoChar = Mid(remainedArticleNumber, InStr(2, remainedArticleNumber, "-"))
                Else
                    optionOneChar = remainedArticleNumber
                    optionTwoChar = ""
                End If
            Else
                optionOneChar = ""
                optionTwoChar = ""
            End If
        Else
            optionOneChar = ""
            optionTwoChar = ""
        End If

        ' MsgBox "Model: " & modelChar & vbNewLine & _
        '         "Connection Size: " & connSizeChar & vbNewLine & _
        '         "Housing Wet: " & housingWetChar & vbNewLine & _
        '         "Housing Notwet: " & housingNotwetChar & vbNewLine & _
        '         "Membrane Material: " & memMaterialChar & vbNewLine & _
        '         "Membrane Design: " & memDesignChar & vbNewLine & _
        '         "Check Valve: " & checkValveChar & vbNewLine & _
        '         "Valve Seat: " & valveSeatChar & vbNewLine & _
        '         "Housing Design: " & housingDesignChar & vbNewLine & _
        '         "Revision: " & revisionChar & vbNewLine & _
        '         "Option1: " & optionOneChar & vbNewLine & _
        '         "Option2: " & optionTwoChar

        For j = 1 To UBound(modelData, 1)
            If modelData(j, 1) = modelChar Then
                model = modelData(j, 2)
                Exit For
            End If
        Next j

        For j = 1 To UBound(connSizeData, 1)
            If connSizeData(j, 1) = connSizeChar Then
                If connSizeData(j, 2) = optionOneChar Then
                    connSizeInch = connSizeData(j, 5)
                    connSizeMM = connSizeData(j, 7)
                ElseIf connSizeData(j, 2) = "" Then
                    connSizeInch = connSizeData(j, 5)
                    connSizeMM = connSizeData(j, 7)
                End If
                Exit For
            End If
        Next j

        For j = 1 To UBound(housingWetData, 1)
            If housingWetData(j, 1) = housingWetChar Then
                housingWetDE = housingWetData(j, 2)
                housingWetEN = housingWetData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(housingNotwetData, 1)
            If housingNotwetData(j, 1) = housingNotwetChar Then
                housingNotwetDE = housingNotwetData(j, 2)
                housingNotwetEN = housingNotwetData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(memMaterialData, 1)
            If memMaterialData(j, 1) = memMaterialChar Then
                memMaterialDE = memMaterialData(j, 2)
                memMaterialEN = memMaterialData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(memDesignData, 1)
            If memDesignData(j, 1) = memDesignChar Then
                memDesignDE = memDesignData(j, 2)
                memDesignEN = memDesignData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(checkValveData, 1)
            If checkValveData(j, 1) = checkValveChar Then
                checkValveDE = checkValveData(j, 2)
                checkValveEN = checkValveData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(valveSeatData, 1)
            If valveSeatData(j, 1) = valveSeatChar Then
                valveSeatDE = valveSeatData(j, 2)
                valveSeatEN = valveSeatData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(housingDesignData, 1)
            If housingDesignData(j, 1) = housingDesignChar Then
                housingDesignDE = housingDesignData(j, 2)
                housingDesignEN = housingDesignData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(revisionData, 1)
            If revisionData(j, 1) = revisionChar Then
                revision = revisionData(j, 2)
                Exit For
            End If
        Next j

        For j = 1 To UBound(optionsData, 1)
            If optionsData(j, 1) = optionOneChar Then
                optionOne = optionsData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(optionsData, 1)
            If optionsData(j, 1) = optionTwoChar Then
                optionTwo = optionsData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(FDAData, 1)
            If FDAData(j, 1) = modelChar Then
                If FDAData(j, 3) = housingWetChar Then
                    If FDAData(j, 5) = memMaterialChar Then
                        If FDAData(j, 7) = checkValveChar Then
                            If FDAData(j, 8) = valveSeatChar Then
                                If FDAData(j, 2) = connSizeChar Then
                                    If FDAData(j, 4) = housingNotwetChar Then
                                        FDAComplianceDE = FDAData(j, 12)
                                        FDAComplianceEN = FDAData(j, 13)
                                    End If
                                ElseIf FDAData(j, 2) = "" Then
                                    FDAComplianceDE = FDAData(j, 12)
                                    FDAComplianceEN = FDAData(j, 13)
                                End If
                                Exit For
                            End If
                        End If
                    End If
                End If
            Else
                FDAComplianceDE = FDAData(3, 15)
                FDAComplianceEN = FDAData(3, 15)
            End If
        Next j

        For j = 1 To UBound(explosionData, 1)
            If explosionData(j, 1) = optionTwoChar Then
                explosionDE = explosionData(j, 2)
                explosionEN = explosionData(j, 3)
                Exit For
            Else
                explosionDE = explosionData(1, 6)
                explosionEN = explosionData(1, 6)
            End If
        Next j

        For j = 1 To UBound(maxSolidSizeData, 1)
            If maxSolidSizeData(j, 1) = connSizeChar Then
                If maxSolidSizeData(j, 2) = housingWetChar Then
                    If maxSolidSizeData(j, 3) = housingNotwetChar Or maxSolidSizeData(j, 3) = "" Then
                        If maxSolidSizeData(j, 7) = housingDesignChar Or maxSolidSizeData(j, 7) = "" Then
                            maxSolidSize = maxSolidSizeData(j, 8)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(flowRateData, 1)
            If flowRateData(j, 1) = connSizeChar Then
                If flowRateData(j, 2) = housingWetChar Or flowRateData(j, 2) = "" Then
                    If flowRateData(j, 3) = housingDesignChar Or flowRateData(j, 3) = "" Then
                        If flowRateData(j, 4) = optionOneChar Or flowRateData(j, 4) = "" Then
                            flowRate = flowRateData(j, 7)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(maxDischargePressureData, 1)
            If maxDischargePressureData(j, 1) = connSizeChar Or maxDischargePressureData(j, 1) = "" Then
                If maxDischargePressureData(j, 2) = housingWetChar Or maxDischargePressureData(j, 2) = "" Then
                    If maxDischargePressureData(j, 3) = housingNotwetChar Or maxDischargePressureData(j, 3) = "" Then
                        If maxDischargePressureData(j, 4) = optionOneChar Or maxDischargePressureData(j, 4) = "" Then
                            maxDischargePressure = maxDischargePressureData(j, 7)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(conveyCapacityData, 1)
            If conveyCapacityData(j, 1) = connSizeChar Then
                If conveyCapacityData(j, 2) = housingWetChar Then
                    If conveyCapacityData(j, 3) = housingNotwetChar Then
                        If conveyCapacityData(j, 4) = memMaterialChar Or conveyCapacityData(j, 4) = "" Then
                            If conveyCapacityData(j, 5) = memDesignChar Or conveyCapacityData(j, 5) = "" Then
                                If conveyCapacityData(j, 6) = housingDesignChar Or conveyCapacityData(j, 6) = "" Then
                                    conveyCapacity = conveyCapacityData(j, 10)
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(connectionTypeData, 1)
            If connectionTypeData(j, 1) = modelChar Then
                If connectionTypeData(j, 2) = connSizeChar Then
                    If connectionTypeData(j, 3) = housingWetChar Or connectionTypeData(j, 3) = "" Then
                        If connectionTypeData(j, 4) = housingNotwetChar Or connectionTypeData(j, 4) = "" Then
                            If connectionTypeData(j, 5) = housingDesignChar Or connectionTypeData(j, 5) = "" Then
                                If connectionTypeData(j, 6) = optionOneChar Or connectionTypeData(j, 6) = "" Then
                                    If connectionTypeData(j, 7) = optionTwoChar Or connectionTypeData(j, 7) = "" Then
                                        connectionType = connectionTypeData(j, 9)
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(connSizeSuctionData, 1)
            If connSizeSuctionData(j, 1) = connSizeChar Then
                If connSizeSuctionData(j, 2) = optionOneChar Or connSizeSuctionData(j, 2) = "" Then
                    connSizeSuction = connSizeSuctionData(j, 5)
                    Exit For
                End If
            End If
        Next j

        For j = 1 To UBound(connSizePressureData, 1)
            If connSizePressureData(j, 1) = connSizeChar Then
                If connSizePressureData(j, 2) = optionOneChar Or connSizePressureData(j, 2) = "" Then
                    connSizePressure = connSizePressureData(j, 5)
                    Exit For
                End If
            End If
        Next j

        For j = 1 To UBound(suctionHeightWetData, 1)
            If suctionHeightWetData(j, 1) = connSizeChar Then
                If suctionHeightWetData(j, 2) = housingWetChar Or suctionHeightWetData(j, 2) = "" Then
                    If suctionHeightWetData(j, 3) = housingDesignChar Or suctionHeightWetData(j, 3) = "" Then
                        If suctionHeightWetData(j, 4) = optionOneChar Or suctionHeightWetData(j, 4) = "" Then
                            suctionHeightWet = suctionHeightWetData(j, 7)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(suctionHeightDryData, 1)
            If suctionHeightDryData(j, 1) = connSizeChar Then
                If suctionHeightDryData(j, 2) = housingWetChar Or suctionHeightDryData(j, 2) = "" Then
                    If suctionHeightDryData(j, 3) = housingDesignChar Or suctionHeightDryData(j, 3) = "" Then
                        If suctionHeightDryData(j, 4) = optionOneChar Or suctionHeightDryData(j, 4) = "" Then
                            suctionHeightDry = suctionHeightDryData(j, 7)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j
        
        For j = 1 To UBound(airConnInletData, 1)
            If airConnInletData(j, 1) = connSizeChar Then
                If airConnInletData(j, 2) = housingWetChar Or airConnInletData(j, 2) = "" Then
                    If airConnInletData(j, 3) = housingDesignChar Or airConnInletData(j, 3) = "" Then
                        If airConnInletData(j, 4) = optionOneChar Or airConnInletData(j, 4) = "" Then
                            airConnInlet = airConnInletData(j, 6)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(airConnOutletData, 1)
            If airConnOutletData(j, 1) = connSizeChar Then
                If airConnOutletData(j, 2) = optionOneChar Or airConnOutletData(j, 2) = "" Then
                    airConnOutlet = airConnOutletData(j, 5)
                    Exit For
                End If
            End If
        Next j

        For j = 1 To UBound(dimensionsData, 1)
            If dimensionsData(j, 1) = connSizeChar Then
                If dimensionsData(j, 2) = housingWetChar Or dimensionsData(j, 2) = "" Then
                    If dimensionsData(j, 3) = housingNotwetChar Or dimensionsData(j, 3) = "" Then
                        If dimensionsData(j, 4) = housingDesignChar Or dimensionsData(j, 4) = "" Then
                            If dimensionsData(j, 5) = connectionType Or dimensionsData(j, 5) = "" Then
                                If dimensionsData(j, 6) = optionOneChar Or dimensionsData(j, 6) = "" Then
                                    weight = dimensionsData(j, 9)
                                    length = dimensionsData(j, 10)
                                    width = dimensionsData(j, 11)
                                    height = dimensionsData(j, 12)
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        
        For j = 1 To UBound(memMaterialTempData, 1)
            If memMaterialTempData(j, 1) = memMaterialChar Then
                memMaterialTempMin = memMaterialTempData(j, 3)
                memMaterialTempMax = memMaterialTempData(j, 4)
                Exit For
            End If
        Next j
        
        MsgBox "Model: " & connSizeInch & " | " & connSizeMM

    '     Set connSizeResult = GetConnectionSizeFromArticleNumber(Cstr(remainedArticleNum))
    '     connSizeChar = connSizeResult.Item(1)
    '     connSize = connSizeResult.Item(2)
    '     remainedArticleNum = connSizeResult.Item(3)

    '     Set housingWetResult = GetHousingMaterialWetFromArticleNumber(Cstr(remainedArticleNum))
    '     housingWetChar = housingWetResult.Item(1)
    '     housingWet = housingWetResult.Item(2)
    '     remainedArticleNum = housingWetResult.Item(3)

    '     Set housingNotwetResult = GetHousingMaterialNotwetFromArticleNumber(Cstr(remainedArticleNum))
    '     housingNotwetChar = housingNotwetResult.Item(1)
    '     housingNotwet = housingNotwetResult.Item(2)
    '     remainedArticleNum = housingNotwetResult.Item(3)

    '     Set memMaterialResult = GetMembraneMaterialFromArticleNumber(Cstr(remainedArticleNum))
    '     memMaterialChar = memMaterialResult.Item(1)
    '     memMaterial = memMaterialResult.Item(2)
    '     remainedArticleNum = memMaterialResult.Item(3)

    '     Set memDesignResult = GetMembraneDesignFromArticleNumber(Cstr(remainedArticleNum))
    '     memDesignChar = memDesignResult.Item(1)
    '     memDesign = memDesignResult.Item(2)
    '     remainedArticleNum = memDesignResult.Item(3)

    '     Set checkValveResult = GetCheckValveMaterialFromArticleNumber(Cstr(remainedArticleNum))
    '     checkValveChar = checkValveResult.Item(1)
    '     checkValve = checkValveResult.Item(2)
    '     remainedArticleNum = checkValveResult.Item(3)
        
    '     Set valveSeatResult = GetValveSeatMaterialFromArticleNumber(Cstr(remainedArticleNum))
    '     valveSeatChar = valveSeatResult.Item(1)
    '     valveSeat = valveSeatResult.Item(2)
    '     remainedArticleNum = valveSeatResult.Item(3)

    '     Set housingDesignResult = GetHousingDesignFromArticleNumber(Cstr(remainedArticleNum))
    '     housingDesignChar = housingDesignResult.Item(1)
    '     housingDesign = housingDesignResult.Item(2)
    '     remainedArticleNum = housingDesignResult.Item(3)

    '     Set revisionResult = GetRevisionLevelFromArticleNumber(Cstr(remainedArticleNum))
    '     revisionChar = revisionResult.Item(1)
    '     revision = revisionResult.Item(2)
    '     remainedArticleNum = revisionResult.Item(3)

    '     Set optionsResult = GetOptionsFromArticleNumber(Cstr(remainedArticleNum))
    '     optionsChar = optionsResult.Item(1)
    '     options = optionsResult.Item(2)

    '     ' Debug.Print("articleNum " & articleNum & ": " & "optionsChar: " & optionsChar)
    '     FDACompliance = GetFDACompliance(Cstr(modelChar), Cstr(housingWetChar), Cstr(memMaterialChar), Cstr(checkValveChar), Cstr(valveSeatChar), Cstr(optionsChar))

    '     explosionProtection = GetExplosionProtectionFromArticleNumber(articleNum)

    '     maxSolidSize = GetMaxSolidSize(connSizeChar, housingWetChar, housingDesignChar)

    '     flowRatePerStroke = GetFlowRatePerStroke(connSizeChar, housingWetChar, housingDesignChar, optionsChar)

    '     maxDischargePressure = GetMaxDischargePressure(connSizeChar, housingWetChar, housingNotwetChar, optionsChar)

    '     conveyingCapacity = GetConveyingCapacity(connSizeChar, housingWetChar, housingNotwetChar, memMaterialChar, memDesignChar, housingDesignChar, optionsChar)

    '     connectionType = GetConnectionType(connSizeChar, housingWetChar, housingNotwetChar, housingDesignChar, optionsChar)

    '     connectionSizeForSuction = GetConnectionSizeForSuction(connSizeChar, optionsChar)

    '     connectionSizeForPressure = GetConnectionSizeForPressure(connSizeChar, optionsChar)

    '     suctionHeightWetted = GetSuctionHeightWetted(connSizeChar, housingWetChar, housingDesignChar, optionsChar)

    '     suctionHeightDry = GetSuctionHeightDry(connSizeChar, housingWetChar, housingDesignChar, optionsChar)

    '     airConnectionInlet = GetAirConnectionInlet(connSizeChar, housingWetChar, optionsChar)

    '     airConnectionOutlet = GetAirConnectionOutlet(connSizeChar, optionsChar)

    '     minTemperature = GetMinTemperature(memMaterialChar)
    '     maxTemperature = GetMaxTemperature(memMaterialChar)

    '     Set weightLengthWidthHeight = GetWeightLengthWidthHeight(connSizeChar, housingWetChar, housingNotwetChar, housingDesignChar, Cstr(connectionType), optionsChar)
    '     weight = weightLengthWidthHeight.Item(1)
    '     length = weightLengthWidthHeight.Item(2)
    '     width = weightLengthWidthHeight.Item(3)
    '     height = weightLengthWidthHeight.Item(4)

    '     ' Write data to OUTPUT sheet
    '     wsOutput.Cells(outputRow, 1).Value = articleNum
    '     wsOutput.Cells(outputRow, 2).Value = model
    '     wsOutput.Cells(outputRow, 3).Value = connSize
    '     wsOutput.Cells(outputRow, 4).Value = housingWet(0)
    '     wsOutput.Cells(outputRow, 5).Value = housingNotwet(0)
    '     wsOutput.Cells(outputRow, 6).Value = memMaterial
    '     wsOutput.Cells(outputRow, 7).Value = memDesign
    '     wsOutput.Cells(outputRow, 8).Value = checkValve
    '     wsOutput.Cells(outputRow, 9).Value = valveSeat
    '     wsOutput.Cells(outputRow, 10).Value = housingDesign
    '     wsOutput.Cells(outputRow, 11).Value = FDACompliance
    '     wsOutput.Cells(outputRow, 12).Value = explosionProtection
    '     wsOutput.Cells(outputRow, 13).Value = maxSolidSize
    '     wsOutput.Cells(outputRow, 14).Value = flowRatePerStroke
    '     wsOutput.Cells(outputRow, 15).Value = maxDischargePressure
    '     wsOutput.Cells(outputRow, 16).Value = conveyingCapacity
    '     wsOutput.Cells(outputRow, 17).Value = connectionType
    '     wsOutput.Cells(outputRow, 18).Value = connectionSizeForSuction
    '     wsOutput.Cells(outputRow, 19).Value = connectionSizeForPressure
    '     wsOutput.Cells(outputRow, 20).Value = suctionHeightWetted
    '     wsOutput.Cells(outputRow, 21).Value = suctionHeightDry
    '     wsOutput.Cells(outputRow, 22).Value = airConnectionInlet
    '     wsOutput.Cells(outputRow, 23).Value = airConnectionOutlet
    '     wsOutput.Cells(outputRow, 24).Value = minTemperature
    '     wsOutput.Cells(outputRow, 25).Value = maxTemperature
    '     wsOutput.Cells(outputRow, 26).Value = weight
    '     wsOutput.Cells(outputRow, 27).Value = length
    '     wsOutput.Cells(outputRow, 28).Value = width
    '     wsOutput.Cells(outputRow, 29).Value = height
    '     ' wsOutput.Cells(outputRow, 11).Value = revision
    '     ' wsOutput.Cells(outputRow, 12).Value = options
        
    '     ' Write SEO data to SEO OUTPUT sheet
    '     wsSeoOutput.Cells(outputRow, 1).Value = articleNum
    '     wsSeoOutput.Cells(outputRow, 2).Value = "Druckluftmembranpumpe | " & connSize & " Zoll | " & articleNum
    '     wsSeoOutput.Cells(outputRow, 3).Value = articleNum

    '     wsSeoOutput.Cells(outputRow, 4).Value = "Selbstansaugende Druckluftmembranpumpe (trocken) | Anschlussgr" & ChrW(246) & ChrW(223) & "e: " & connSize & " Zoll | F" & ChrW(246) & "rderleistung: " & conveyingCapacity & " Liter pro Minute | F" & ChrW(246) & "rderdruck: max. " & maxDischargePressure & " bar | Geh" & ChrW(228) & "usematerial: " & housingWet(0) & " | Membranmaterial: " & memMaterial & " | Feststoffgr" & ChrW(246) & ChrW(223) & "e: " & maxSolidSize & " mm"

    '     wsSeoOutput.Cells(outputRow, 5).Value = "<ul>" & vbNewLine & _
    '     "<li>Selbstansaugende Druckluftmembranpumpe (trocken)</li>" & vbNewLine & _
    '     "<li>Anschlussgr" & ChrW(246) & ChrW(223) & "e: " & connSize & " Zoll</li>" & vbNewLine & _
    '     "<li>F" & ChrW(246) & "rderleistung: " & conveyingCapacity & " Liter pro Minute</li>" & vbNewLine & _
    '     "<li>F" & ChrW(246) & "rderdruck: max. " & maxDischargePressure & " bar</li>" & vbNewLine & _
    '     "<li>Geh" & ChrW(228) & "usematerial: " & housingWet(0) & "</li>" & vbNewLine & _
    '     "<li>Membranmaterial: " & memMaterial & "</li>" & vbNewLine & _
    '     "<li>Feststoffgr" & ChrW(246) & ChrW(223) & "e: " & maxSolidSize & " mm</li>" & vbNewLine & _
    '     "</ul>" & vbNewLine & _
    '     "<ul>" & vbNewLine & _
    '     "<li><strong><a href=""#tab-attributes"" title=""Weitere technische Daten"">Weitere technische Daten</a></strong></li>" & vbNewLine & _
    '     "<li><strong><a href=""#tab-cross"" title=""Kompatible Reparaturs" & ChrW(228) & "tze oder Ersatzteile"">Kompatible Reparaturs" & ChrW(228) & "tze oder Ersatzteile</a></strong></li>" & vbNewLine & _
    '     "</ul>"

    '     wsSeoOutput.Cells(outputRow, 6).Value = "Air-operated double diaphragm pump | " & connSize & " Inch | " & articleNum
    '     wsSeoOutput.Cells(outputRow, 7).Value = "versamatic-" & articleNum

    '     wsSeoOutput.Cells(outputRow, 8).Value = "Self-priming air-operated diaphragm pump (dry)  | Connection size: " & connSize & " Inch | Flow rate: " & conveyingCapacity & " Litres per minute | Delivery pressure: max. " & maxDischargePressure & " bar | Housing material: " & housingWet(1) & " | Diaphragm material: " & memMaterial & " | Solids size: " & maxSolidSize & " mm"

    '     wsSeoOutput.Cells(outputRow, 9).Value = "<ul>" & vbNewLine & _
    '     "<li>Self-priming air-operated diaphragm pump (dry)" & vbNewLine & _
    '     "<li>Connection size: " & connSize & " Inch</li>" & vbNewLine & _
    '     "<li>Flow rate: " & conveyingCapacity & " Litres per minute</li>" & vbNewLine & _
    '     "<li>Delivery pressure: max. " & maxDischargePressure & " bar</li>" & vbNewLine & _
    '     "<li>Housing material: " & housingWet(1) & "</li>" & vbNewLine & _
    '     "<li>Diaphragm material: " & memMaterial & "</li>" & vbNewLine & _
    '     "<li>Solids size: " & maxSolidSize & " mm</li>" & vbNewLine & _
    '     "</ul>" & vbNewLine & _
    '     "<ul>" & vbNewLine & _
    '     "<li><strong><a href=""#tab-attributes"" title=""Further technical data"">Further technical data</a></strong></li>" & vbNewLine & _
    '     "<li><strong><a href=""#tab-cross"" title=""Compatible repair kits or spare parts"">Compatible repair kits or spare parts</a></strong></li>" & vbNewLine & _
    '     "</ul>"
       
    '     ' Move to next row in OUTPUT sheet
    '     outputRow = outputRow + 1
    Next i
    
    ' MsgBox "Article numbers processed successfully!", vbInformation, "Done"
End Sub
