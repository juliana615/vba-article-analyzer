Attribute VB_Name = "main"    

Sub Test()
    ' MsgBox "Hello, World!"
    MsgBox GetDictionaryKeys(ModelDictionary).Item(1)
End Sub

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
        obj.Add "A", "Aluminium"
        obj.Add "B", "Aluminium B"
        obj.Add "C", "Gusseisen"
        obj.Add "G", "Leitfähiges Polypropylen (Acetal)"
        obj.Add "H", "Hastelloy C"
        obj.Add "J", "Vernickeltes Aluminium"
        obj.Add "K", "PVDF"
        obj.Add "P", "Polypropylen"
        obj.Add "Q", "Epoxidbeschichtetes Aluminium"
        obj.Add "S", "Edelstahl"
        obj.Add "Z", "PTFE-beschichtetes Aluminium"
    End If

    Set HousingMaterialWetDictionary = obj
End Function

' Gehäusematerial (nicht benetzt)
Private Function HousingMaterialNotwetDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "A", "Aluminium"
        obj.Add "B", "Aluminium B"
        obj.Add "C", "Gusseisen"
        obj.Add "G", "Leitfähiges Polypropylen (Acetal)"
        obj.Add "H", "Hastelloy C"
        obj.Add "J", "Vernickeltes Aluminium"
        obj.Add "K", "PVDF"
        obj.Add "P", "Polypropylen"
        obj.Add "Q", "Epoxidbeschichtetes Aluminium"
        obj.Add "S", "Edelstahl"
        obj.Add "Z", "PTFE-beschichtetes Aluminium"
    End If

    Set HousingMaterialNotwetDictionary = obj
End Function

' Material der Membrane		
Private Function MembraneMaterialDictionary() As Dictionary
    Static obj As Dictionary

    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "1", "Neoprene® - CR"
        obj.Add "2", "BunaN® - NBR - Nitrile"
        obj.Add "3", "FKM - Viton®"
        obj.Add "4", "EPDM - Nordel™"
        obj.Add "5", "Teflon® - PTFE"
        obj.Add "6", "Santoprene® -  TPE"
        obj.Add "7", "Hytrel® - TPC"
        obj.Add "9", "Geolast®"
        obj.Add "Y", "Santoprene® -  FDA"
    End If

    Set MembraneMaterialDictionary = obj
End Function

' Membranausführung		
Private Function MembraneDesignDictionary() As Dictionary
    Static obj As Dictionary

    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "R", "Versa-Rugged™"
        obj.Add "D", "Versa-Dome™"
        obj.Add "X", "Thermo-Matic™"
        obj.Add "T", "2-piece"
        obj.Add "B", "Versa-Tuff™"
        obj.Add "F", "FUSION™"
    End If

    Set MembraneDesignDictionary = obj
End Function

' Material Rückschlagventil		
Private Function CheckValveMaterialDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "1", "Neoprene® - CR"
        obj.Add "2", "BunaN® - NBR - Nitrile"
        obj.Add "3", "FKM - Viton®"
        obj.Add "4", "EPDM - Nordel™"
        obj.Add "5", "Teflon® - PTFE"
        obj.Add "6", "Santoprene® -  TPE"
        obj.Add "7", "Hytrel® - TPC"
        obj.Add "8", "Polyurethan"
        obj.Add "9", "Geolast®"
        obj.Add "A", "Acetal"
        obj.Add "S", "Edelstahl"
        obj.Add "Y", "Santoprene® -  FDA"
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
        obj.Add "1", "Neoprene® - CR"
        obj.Add "2", "BunaN® - NBR - Nitrile"
        obj.Add "3", "FKM - Viton®"
        obj.Add "4", "EPDM - Nordel™"
        obj.Add "5", "Teflon® - PTFE"
        obj.Add "6", "Santoprene® -  TPE"
        obj.Add "7", "Hytrel® - TPC"
        obj.Add "8", "Polyurethan"
        obj.Add "9", "Geolast®"
        obj.Add "A", "Aluminium"
        obj.Add "S", "Edelstahl"
        obj.Add "C", "Stahl"
        obj.Add "H", "Hastelloy C"
        obj.Add "T", "PTFE-ummanteltes Silikon"
        obj.Add "Y", "Santoprene® -  FDA"
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
    Dim housingMaterialWet As String
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    housingMaterialWetChar = Mid(ArticleNumber, 1, 1)
    If HousingMaterialWetDictionary.Exists(housingMaterialWetChar) Then
        housingMaterialWet = HousingMaterialWetDictionary.Item(housingMaterialWetChar)
    Else
        housingMaterialWet = "Unknown"
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
    Dim housingMaterialNotwet As String
    Dim remainedArticleNumber As String
    Dim returnCollection As New Collection

    housingMaterialNotwetChar = Mid(ArticleNumber, 1, 1)
    If HousingMaterialNotwetDictionary.Exists(housingMaterialNotwetChar) Then
        housingMaterialNotwet = HousingMaterialNotwetDictionary.Item(housingMaterialNotwetChar)
    Else
        housingMaterialNotwet = "Unknown"
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
    GetFDACompliance = "Ohne FDA-Konformität"
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

Sub BreakdownArticleName()
    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, i As Long
    Dim articleNum As String, remainedArticleNum As String

    Dim modelResult As Collection, modelChar As String, model As String
    Dim connSizeResult As Collection, connSizeChar As String, connSize As String
    Dim housingWetResult As Collection, housingWetChar As String, housingWet As String
    Dim housingNotwetResult As Collection, housingNotwetChar As String, housingNotwet As String
    Dim memMaterialResult As Collection, memMaterialChar As String, memMaterial As String
    Dim memDesignResult As Collection, memDesignChar As String, memDesign As String
    Dim checkValveResult As Collection, checkValveChar As String, checkValve As String
    Dim valveSeatResult As Collection, valveSeatChar As String, valveSeat As String
    Dim housingDesignResult As Collection, housingDesignChar As String, housingDesign As String
    Dim revisionResult As Collection, revisionChar As String, revision As String
    Dim optionsResult As Collection, optionsChar As String, options As String
    Dim FDACompliance As String
    Dim outputRow As Long
    
    ' Set worksheet references
    Set wsInput = ThisWorkbook.Sheets("INPUT")
    Set wsOutput = ThisWorkbook.Sheets("OUTPUT")

    ' Find last row in INPUT sheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    ' Loop through each article number
    outputRow = 2 ' Start from row 2 in OUTPUT sheet
    For i = 5 To lastRow
        articleNum = wsInput.Cells(i, 1).Value ' Read article number

        ' Check if article number is valid
        ' If Len(articleNum) >= 11 Then
            ' Extract components based on predefined structure
        Set modelResult = GetModelFromArticleNumber(articleNum)
        modelChar = modelResult.Item(1)
        model = modelResult.Item(2)
        remainedArticleNum = modelResult.Item(3)

        Set connSizeResult = GetConnectionSizeFromArticleNumber(Cstr(remainedArticleNum))
        connSizeChar = connSizeResult.Item(1)
        connSize = connSizeResult.Item(2)
        remainedArticleNum = connSizeResult.Item(3)

        Set housingWetResult = GetHousingMaterialWetFromArticleNumber(Cstr(remainedArticleNum))
        housingWetChar = housingWetResult.Item(1)
        housingWet = housingWetResult.Item(2)
        remainedArticleNum = housingWetResult.Item(3)

        Set housingNotwetResult = GetHousingMaterialNotwetFromArticleNumber(Cstr(remainedArticleNum))
        housingNotwetChar = housingNotwetResult.Item(1)
        housingNotwet = housingNotwetResult.Item(2)
        remainedArticleNum = housingNotwetResult.Item(3)

        Set memMaterialResult = GetMembraneMaterialFromArticleNumber(Cstr(remainedArticleNum))
        memMaterialChar = memMaterialResult.Item(1)
        memMaterial = memMaterialResult.Item(2)
        remainedArticleNum = memMaterialResult.Item(3)

        Set memDesignResult = GetMembraneDesignFromArticleNumber(Cstr(remainedArticleNum))
        memDesignChar = memDesignResult.Item(1)
        memDesign = memDesignResult.Item(2)
        remainedArticleNum = memDesignResult.Item(3)

        Set checkValveResult = GetCheckValveMaterialFromArticleNumber(Cstr(remainedArticleNum))
        checkValveChar = checkValveResult.Item(1)
        checkValve = checkValveResult.Item(2)
        remainedArticleNum = checkValveResult.Item(3)
        
        Set valveSeatResult = GetValveSeatMaterialFromArticleNumber(Cstr(remainedArticleNum))
        valveSeatChar = valveSeatResult.Item(1)
        valveSeat = valveSeatResult.Item(2)
        remainedArticleNum = valveSeatResult.Item(3)

        Set housingDesignResult = GetHousingDesignFromArticleNumber(Cstr(remainedArticleNum))
        housingDesignChar = housingDesignResult.Item(1)
        housingDesign = housingDesignResult.Item(2)
        remainedArticleNum = housingDesignResult.Item(3)

        Set revisionResult = GetRevisionLevelFromArticleNumber(Cstr(remainedArticleNum))
        revisionChar = revisionResult.Item(1)
        revision = revisionResult.Item(2)
        remainedArticleNum = revisionResult.Item(3)

        Set optionsResult = GetOptionsFromArticleNumber(Cstr(remainedArticleNum))
        optionsChar = optionsResult.Item(1)
        options = optionsResult.Item(2)

        ' Debug.Print("articleNum " & articleNum & ": " & "optionsChar: " & optionsChar)
        FDACompliance = GetFDACompliance(Cstr(modelChar), Cstr(housingWetChar), Cstr(memMaterialChar), Cstr(checkValveChar), Cstr(valveSeatChar), Cstr(optionsChar))

        explosionProtection = GetExplosionProtectionFromArticleNumber(articleNum)

        maxSolidSize = GetMaxSolidSize(connSizeChar, housingWetChar, housingDesignChar)

        flowRatePerStroke = GetFlowRatePerStroke(connSizeChar, housingWetChar, housingDesignChar, optionsChar)

        ' Write data to OUTPUT sheet
        wsOutput.Cells(outputRow, 1).Value = articleNum
        wsOutput.Cells(outputRow, 2).Value = model
        wsOutput.Cells(outputRow, 3).Value = connSize
        wsOutput.Cells(outputRow, 4).Value = housingWet
        wsOutput.Cells(outputRow, 5).Value = housingNotwet
        wsOutput.Cells(outputRow, 6).Value = memMaterial
        wsOutput.Cells(outputRow, 7).Value = memDesign
        wsOutput.Cells(outputRow, 8).Value = checkValve
        wsOutput.Cells(outputRow, 9).Value = valveSeat
        wsOutput.Cells(outputRow, 10).Value = housingDesign
        wsOutput.Cells(outputRow, 11).Value = FDACompliance
        wsOutput.Cells(outputRow, 12).Value = explosionProtection
        wsOutput.Cells(outputRow, 13).Value = maxSolidSize
        wsOutput.Cells(outputRow, 14).Value = flowRatePerStroke
        ' wsOutput.Cells(outputRow, 11).Value = revision
        ' wsOutput.Cells(outputRow, 12).Value = options
        
        ' Move to next row in OUTPUT sheet
        outputRow = outputRow + 1
    Next i
    
    MsgBox "Article numbers processed successfully!", vbInformation, "Done"
End Sub
