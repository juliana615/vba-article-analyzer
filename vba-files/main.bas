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

        maxDischargePressure = GetMaxDischargePressure(connSizeChar, housingWetChar, housingNotwet, optionsChar)

        conveyingCapacity = GetConveyingCapacity(connSizeChar, housingWetChar, housingNotwetChar, memMaterialChar, memDesignChar, housingDesignChar, optionsChar)

        connectionType = GetConnectionType(connSizeChar, housingWetChar, housingNotwetChar, housingDesignChar, optionsChar)


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
        wsOutput.Cells(outputRow, 15).Value = maxDischargePressure
        wsOutput.Cells(outputRow, 16).Value = conveyingCapacity
        ' wsOutput.Cells(outputRow, 11).Value = revision
        ' wsOutput.Cells(outputRow, 12).Value = options
        
        ' Move to next row in OUTPUT sheet
        outputRow = outputRow + 1
    Next i
    
    MsgBox "Article numbers processed successfully!", vbInformation, "Done"
End Sub
