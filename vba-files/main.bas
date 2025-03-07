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
    
' Optionen		
Private Function OptionsDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "ATEX", "ATEX Compliant"
        obj.Add "B", "BSP threaded"
        obj.Add "CP", "Center Port"
        obj.Add "DV", "Drop in Bolted Units to Replace VM Clamped (2”+3” Al & SS Only)"
        obj.Add "FP", "Food Processing"
        obj.Add "HD", "Horizontal Discharge   "
        obj.Add "SP", "Sanitary Processing "
        obj.Add "3A", "3A Sanitary"
        obj.Add "HP", "High Pressure"
        obj.Add "DW", "Drop in Bolted Units to Replace Wilden Clamped (2”+3” Al & SS Only)"
        obj.Add "F", "Flap Valve 2” Al Only)"
        obj.Add "OB", "Oil Bottle (V-Series Only"
        obj.Add "SM", "Split Mainfold"
        obj.Add "UL", "UL Listed"
        obj.Add "E4", "120VAC Coil"
        obj.Add "E0", "24VDC Coil "
        obj.Add "U", "Universal ANSI/DIN Flange"
    End If

    Set OptionsDictionary = obj
End Function

' Get the keys of given dictionary
' Private Function GetDictionaryKeys(dict As Dictionary) As Collection
'     Dim key As Variant
'     Dim keys As New Collection
    
'     For Each key In dict.Keys
'         keys.Add key
'     Next key
    
'     Set GetDictionaryKeys = keys
' End Function

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

    revisionLevelChar = Mid(ArticleNumber, 1, 1)
    If RevisionLevelDictionary.Exists(revisionLevelChar) Then
        revisionLevel = RevisionLevelDictionary.Item(revisionLevelChar)
    Else
        revisionLevel = "Unknown"
    End If
    remaindArticleNumber = Mid(ArticleNumber, 2)
    returnCollection.Add revisionLevelChar
    returnCollection.Add revisionLevel
    returnCollection.Add remaindArticleNumber
    Set GetRevisionLevelFromArticleNumber = returnCollection
End Function

Sub BreakdownArticleName()
    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, i As Long
    Dim articleNum As String
    Dim model As String, connSize As String, housingWet As String
    Dim housingNotwet As String, memMaterial As String, memDesign As String
    Dim checkValve As String, valveSeat As String, housingDesign As String
    Dim revision As String, options As String
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
        If Len(articleNum) >= 11 Then
            ' Extract components based on predefined structure
            modelChar = GetModelFromArticleNumber(articleNum).Item(1)
            model = GetModelFromArticleNumber(articleNum).Item(2)
            remainedArticleNum = GetModelFromArticleNumber(articleNum).Item(3)
            connSizeChar = GetConnectionSizeFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            connSize = GetConnectionSizeFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            remainedArticleNum = GetConnectionSizeFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            housingWetChar = GetHousingMaterialWetFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            housingWet = GetHousingMaterialWetFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            remainedArticleNum = GetHousingMaterialWetFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            housingNotwetChar = GetHousingMaterialNotwetFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            housingNotwet = GetHousingMaterialNotwetFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            remainedArticleNum = GetHousingMaterialNotwetFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            memMaterialChar = GetMembraneMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            memMaterial = GetMembraneMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            remainedArticleNum = GetMembraneMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            memDesignChar = GetMembraneDesignFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            memDesign = GetMembraneDesignFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            remainedArticleNum = GetMembraneDesignFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            checkValveChar = GetCheckValveMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            checkValve = GetCheckValveMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            remainedArticleNum = GetCheckValveMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            valveSeatChar = GetValveSeatMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            valveSeat = GetValveSeatMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            remainedArticleNum = GetValveSeatMaterialFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            housingDesignChar = GetHousingDesignFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            housingDesign = GetHousingDesignFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            remainedArticleNum = GetHousingDesignFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            revisionChar = GetRevisionLevelFromArticleNumber(Cstr(remainedArticleNum)).Item(1)
            revision = GetRevisionLevelFromArticleNumber(Cstr(remainedArticleNum)).Item(2)
            ' remainedArticleNum = GetRevisionLevelFromArticleNumber(Cstr(remainedArticleNum)).Item(3)
            
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
            ' wsOutput.Cells(outputRow, 11).Value = revision
            ' wsOutput.Cells(outputRow, 12).Value = options
            
            ' Move to next row in OUTPUT sheet
            outputRow = outputRow + 1
        End If
    Next i
    
    MsgBox "Article numbers processed successfully!", vbInformation, "Done"
End Sub
