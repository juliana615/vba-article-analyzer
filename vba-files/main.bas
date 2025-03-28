Attribute VB_Name = "main"    

' ChrW(246) → ö, ChrW(223) → ß, ChrW(228) → ä, ChrW(252) → ü, ChrW(174) → ®, ChrW(8482) → ™, ChrW(8443) → °C

' Sub GetRGBColor()
'     Dim testCell As Range
'     Dim wsSeoInput As Worksheet
'     Set wsSeoInput = ThisWorkbook.Sheets("SEO INPUT")
'     Set testCell = wsSeoInput.Range("B6")
'     Dim colorValue As Long
'     Dim redVal As Integer, greenVal As Integer, blueVal As Integer
    
'     colorValue = testCell.Interior.Color
    
'     redVal = colorValue Mod 256
'     greenVal = (colorValue \ 256) Mod 256
'     blueVal = (colorValue \ 65536) Mod 256
    
'     MsgBox "RGB: (" & redVal & ", " & greenVal & ", " & blueVal & ")"
' End Sub

Private Function GetVariableDict(variableData As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    For i = 1 To UBound(variableData, 1) ' Loop through rows
        If Not dict.exists(variableData(i, 1)) Then ' Avoid duplicates
            dict.Add variableData(i, 1), variableData(i, 2) ' Add key-value pair
        End If
    Next i
    
    Set GetVariableDict = dict ' Return dictionary
End Function
    
' Private Function GetVariableDictionary(variableData As Variant) As Dictionary
'     Static obj As Dictionary

'     If obj Is Nothing Then
'         Set obj = New Dictionary
'         Dim i As Integer
'         For i = 1 To UBound(variableData, 1)
'             obj.Add variableData(i, 1), variableData(i, 2)
'         Next i
'     End If

'     Set GetVariableDictionary = obj
' End Function
    
Sub Main()
    '''''''''''''''''''''
    ' Initializing Data '
    '''''''''''''''''''''
    ' Brand
    Dim brandWS As Worksheet, brandTable As Range, brandData As Variant
    Set brandWS = Sheet4
    Set brandTable = brandWS.Range("A3:B3")
    brandData = brandTable.Value

    ' Modell
    Dim modelWS As Worksheet, modelTable As Range, modelData As Variant
    Set modelWS = Tabelle23
    Set modelTable = modelWS.Range("A3:B7")
    modelData = modelTable.Value

    ' Anschlussgröße
    Dim connSizeWS As Worksheet, connSizeTable As Range, connSizeData As Variant
    Set connSizeWS = Tabelle7
    Set connSizeTable = connSizeWS.Range("A3:H21")
    connSizeData = connSizeTable.Value

    ' Typ des Rückschlagventils
    Dim checkValveTypeWS As Worksheet, checkValveTypeTable As Range, checkValveTypeData As Variant
    Set checkValveTypeWS = Sheet5
    Set checkValveTypeTable = checkValveTypeWS.Range("A3:D4")
    checkValveTypeData = checkValveTypeTable.Value

    ' Gehäusematerial (benetzt)
    Dim housingWetWS As Worksheet, housingWetTable As Range, housingWetData As Variant
    Set housingWetWS = Tabelle15
    Set housingWetTable = housingWetWS.Range("A3:C14")
    housingWetData = housingWetTable.Value

    ' Gehäusematerial (nicht benetzt)
    Dim housingNotwetWS As Worksheet, housingNotwetTable As Range, housingNotwetData As Variant
    Set housingNotwetWS = Sheet6
    Set housingNotwetTable = housingNotwetWS.Range("A3:C14")
    housingNotwetData = housingNotwetTable.Value

    ' Material der Membrane		
    Dim memMaterialWS As Worksheet, memMaterialTable As Range, memMaterialData As Variant
    Set memMaterialWS = Tabelle12
    Set memMaterialTable = memMaterialWS.Range("A3:C13")
    memMaterialData = memMaterialTable.Value

    ' Backup-Membran
    Dim backupMembraneWS As Worksheet, backupMembraneTable As Range, backupMembraneData As Variant
    Set backupMembraneWS = Sheet7
    Set backupMembraneTable = backupMembraneWS.Range("A3:D7")
    backupMembraneData = backupMembraneTable.Value

    ' ' Membranausführung		
    ' Dim memDesignWS As Worksheet, memDesignTable As Range, memDesignData As Variant
    ' Set memDesignWS = Tabelle24
    ' Set memDesignTable = memDesignWS.Range("A3:C8")
    ' memDesignData = memDesignTable.Value

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

    ' Luftventil
    Dim airValveWS As Worksheet, airValveTable As Range, airValveData As Variant
    Set airValveWS = Sheet8
    Set airValveTable = airValveWS.Range("A3:C4")
    airValveData = airValveTable.Value

    ' Luftventil-Option
    Dim airValveOptionWS As Worksheet, airValveOptionTable As Range, airValveOptionData As Variant
    Set airValveOptionWS = Sheet9
    Set airValveOptionTable = airValveOptionWS.Range("A3:C4")
    airValveOptionData = airValveOptionTable.Value

    ' Schalldämpfer-Option
    Dim silencerOptionWS As Worksheet, silencerOptionTable As Range, silencerOptionData As Variant
    Set silencerOptionWS = Sheet10
    Set silencerOptionTable = silencerOptionWS.Range("A3:C5")
    silencerOptionData = silencerOptionTable.Value

    
    ' Anschlussart								
    Dim connectionTypeWS As Worksheet, connectionTypeTable As Range, connectionTypeData As Variant
    Set connectionTypeWS = Sheet11
    Set connectionTypeTable = connectionTypeWS.Range("A3:D6")
    connectionTypeData = connectionTypeTable.Value

    ' Anschluss-Option
    Dim connectionOptionWS As Worksheet, connectionOptionTable As Range, connectionOptionData As Variant
    Set connectionOptionWS = Tabelle21
    Set connectionOptionTable = connectionOptionWS.Range("A3:J88")
    connectionOptionData = connectionOptionTable.Value

    ' ' Gehäuseausführung		
    ' Dim housingDesignWS As Worksheet, housingDesignTable As Range, housingDesignData As Variant
    ' Set housingDesignWS = Tabelle9
    ' Set housingDesignTable = housingDesignWS.Range("A3:C4")
    ' housingDesignData = housingDesignTable.Value

    ' Revisionslevel	
    Dim revisionWS As Worksheet, revisionTable As Range, revisionData As Variant
    Set revisionWS = Tabelle26
    Set revisionTable = revisionWS.Range("A3:B7")
    revisionData = revisionTable.Value

    ' ' Optionsen		
    ' Dim optionsWS As Worksheet, optionsTable As Range, optionsData As Variant
    ' Set optionsWS = Tabelle27
    ' Set optionsTable = optionsWS.Range("A3:C19")
    ' optionsData = optionsTable.Value

    ' ' FDA-Konformität												
    ' Dim FDAWS As Worksheet, FDATable As Range, FDAData As Variant
    ' Set FDAWS = Tabelle10
    ' Set FDATable = FDAWS.Range("A3:O102")
    ' FDAData = FDATable.Value

    ' Maximale Feststoffgröße
    Dim maxSolidSizeWS As Worksheet, maxSolidSizeTable As Range, maxSolidSizeData As Variant
    Set maxSolidSizeWS = Tabelle8
    Set maxSolidSizeTable = maxSolidSizeWS.Range("A3:I19")
    maxSolidSizeData = maxSolidSizeTable.Value

    ' Fördermenge per Hub
    Dim flowRateWS As Worksheet, flowRateTable As Range, flowRateData As Variant
    Set flowRateWS = Tabelle18
    Set flowRateTable = flowRateWS.Range("A3:I25")
    flowRateData = flowRateTable.Value

    ' Maximaler Förderdruck
    Dim maxDischargePressureWS As Worksheet, maxDischargePressureTable As Range, maxDischargePressureData As Variant
    Set maxDischargePressureWS = Tabelle17
    Set maxDischargePressureTable = maxDischargePressureWS.Range("A3:G11")
    maxDischargePressureData = maxDischargePressureTable.Value

    ' Förderleistung
    Dim conveyCapacityWS As Worksheet, conveyCapacityTable As Range, conveyCapacityData As Variant
    Set conveyCapacityWS = Tabelle16
    Set conveyCapacityTable = conveyCapacityWS.Range("A3:L207")
    conveyCapacityData = conveyCapacityTable.Value

    ' Anschlussgröße (Saugseite)
    Dim connSizeSuctionWS As Worksheet, connSizeSuctionTable As Range, connSizeSuctionData As Variant
    Set connSizeSuctionWS = Sheet12
    Set connSizeSuctionTable = connSizeSuctionWS.Range("A3:H23")
    connSizeSuctionData = connSizeSuctionTable.Value

    ' Anschlussgröße (Druckseite)
    Dim connSizePressureWS As Worksheet, connSizePressureTable As Range, connSizePressureData As Variant
    Set connSizePressureWS = Sheet13
    Set connSizePressureTable = connSizePressureWS.Range("A3:H23")
    connSizePressureData = connSizePressureTable.Value

    ' Ansaughöhe (nass)
    Dim suctionHeightWetWS As Worksheet, suctionHeightWetTable As Range, suctionHeightWetData As Variant
    Set suctionHeightWetWS = Tabelle19
    Set suctionHeightWetTable = suctionHeightWetWS.Range("A3:J18")
    suctionHeightWetData = suctionHeightWetTable.Value

    ' Ansaughöhe (trocken)
    Dim suctionHeightDryWS As Worksheet, suctionHeightDryTable As Range, suctionHeightDryData As Variant
    Set suctionHeightDryWS = Tabelle29
    Set suctionHeightDryTable = suctionHeightDryWS.Range("A3:J25")
    suctionHeightDryData = suctionHeightDryTable.Value

    ' Luftanschluss (Eingang)
    Dim airConnInletWS As Worksheet, airConnInletTable As Range, airConnInletData As Variant
    Set airConnInletWS = Tabelle30
    Set airConnInletTable = airConnInletWS.Range("A3:H26")
    airConnInletData = airConnInletTable.Value

    ' Luftanschluss (Ausgang)
    Dim airConnOutletWS As Worksheet, airConnOutletTable As Range, airConnOutletData As Variant
    Set airConnOutletWS = Tabelle20
    Set airConnOutletTable = airConnOutletWS.Range("A3:G16")
    airConnOutletData = airConnOutletTable.Value

    ' Exlosionsschutz (ATEX)	
    Dim explosionWS As Worksheet, explosionTable As Range, explosionData As Variant
    Set explosionWS = Tabelle11
    Set explosionTable = explosionWS.Range("A3:I5")
    explosionData = explosionTable.Value

    ' Gewicht-Abmessungen
    Dim dimensionsWS As Worksheet, dimensionsTable As Range, dimensionsData As Variant
    Set dimensionsWS = Tabelle22
    Set dimensionsTable = dimensionsWS.Range("A2:P75")
    dimensionsData = dimensionsTable.Value

    ' Temperatur - Material der Membrane			
    Dim memMaterialTempWS As Worksheet, memMaterialTempTable As Range, memMaterialTempData As Variant
    Set memMaterialTempWS = Tabelle31
    Set memMaterialTempTable = memMaterialTempWS.Range("A3:E12")
    memMaterialTempData = memMaterialTempTable.Value

    ' Variables
    Dim wsVariable As Worksheet, variableTable As Range, variableData As Variant, variableDictDE As Object, variableDictEN As Object
    Set wsVariable = ThisWorkbook.Sheets("Variables")
    Set variableTable = wsVariable.Range("A1:B7")
    variableData = variableTable.Value
    ' Set variableDict = GetVariableDict(variableData)

    Dim wsSeoInput As Worksheet, seoInputTable As Range, seoInputData As Variant
    Set wsSeoInput = ThisWorkbook.Sheets("SEO INPUT")
    set seoInputTable = wsSeoInput.Range("B4:I200")
    seoInputData = seoInputTable.Value

    ' Set worksheet references
    Dim wsInput As Worksheet, wsOutput As Worksheet, wsSeoOutput As Worksheet
    Set wsInput = ThisWorkbook.Sheets("INPUT")
    Set wsOutput = ThisWorkbook.Sheets("OUTPUT")
    Set wsSeoOutput = ThisWorkbook.Sheets("SEO OUTPUT")

    Dim lastRow As Long
    Dim articleNum As String, remainedArticleNum As String

    Dim brandChar As String, brand As String, remaindArticleNumber As String
    Dim modelChar As String, model As String
    Dim connSizeChar As String, connSizeInch As String, connSizeMM As String
    Dim checkValveTypeChar As String, checkValveTypeDE As String, checkValveTypeEN As String
    Dim housingWetChar As String, housingWetDE As String, housingWetEN As String
    Dim housingNotwetChar As String, housingNotwetDE As String, housingNotwetEN As String
    Dim memMaterialChar As String, memMaterialDE As String, memMaterialEN As String
    Dim replacementMemChar As String, replacementMemDE As String, replacementMemEN As String
    ' Dim memDesignChar As String, memDesignDE As String, memDesignEN As String
    Dim checkValveChar As String, checkValveDE As String, checkValveEN As String
    Dim valveSeatChar As String, valveSeatDE As String, valveSeatEN As String
    Dim airValveChar As String, airValveDE As String, airValveEN As String
    Dim airValveOptionChar As String, airValveOptionDE As String, airValveOptionEN As String
    Dim silencerOptionChar As String, silencerOptionDE As String, silencerOptionEN As String
    Dim connTypeChar As String, connTypeDE As String, connTypeEN As String
    Dim connTypeOption As String
    Dim connTypeDetailChar As String, connTypeDetailDE As String, connTypeDetailEN As String
    ' Dim housingDesignChar As String, housingDesignDE As String, housingDesignEN As String
    Dim revisionChar As String, revision As String
    ' Dim optionOneChar As String, optionOne As String
    ' Dim optionTwoChar As String, optionTwo As String
    ' Dim FDAComplianceDE As String, FDAComplianceEN As String
    Dim explosionDE As String, explosionEN As String
    Dim maxSolidSize As String, flowRate As String, maxDischargePressure As String, conveyCapacity As String, connSizeSuction As String, connSizePressure As String, suctionHeightWet As String, suctionHeightDry As String, airConnInlet As String, airConnOutlet As String
    Dim weight As String, length As String, width As String, height As String  
    Dim memMaterialTempMin As String, memMaterialTempMax As String
    Dim redColor As Long, greenColor As Long
    Dim seoArticleNameDE As String, seoUrlPathDE As String, seoMetaDescriptionDE As String, seoShortDescriptionDE As String, seoArticleNameEN As String, seoUrlPathEN As String, seoMetaDescriptionEN As String, seoShortDescriptionEN As String
    Dim seoArticleNameDECell As String, seoUrlPathDECell As String, seoMetaDescriptionDECell As String, seoShortDescriptionDECell As String, seoArticleNameENCell As String, seoUrlPathENCell As String, seoMetaDescriptionENCell As String, seoShortDescriptionENCell As String
    Dim outputRow As Long
    
    whiteColor = RGB(255, 255, 255)
    redColor = RGB(252, 228, 214)
    greenColor = RGB(226, 239, 218)

    ' Find last row in INPUT sheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row

    ' Loop through each article number
    outputRow = 2 ' Start from row 2 in OUTPUT sheet

    ' For i = 5 To lastRow
    For i = 5 To 5
        ' Initialize values each time
        brand = ""
        model = ""
        connSizeInch = ""
        checkValveTypeDE = ""
        checkValveTypeEN = ""
        housingWetDE = ""
        housingWetEN = ""
        housingNotwetDE = ""
        housingNotwetEN = ""
        memMaterialDE = ""
        memMaterialEN = ""
        replacementMemDE = ""
        replacementMemEN = ""
        memDesignDE = ""
        memDesignEN = ""
        checkValveDE = ""
        checkValveEN = ""
        valveSeatDE = ""
        valveSeatEN = ""
        airValveDE = ""
        airValveEN = ""
        airValveOptionDE = ""
        airValveOptionEN = ""
        silencerOptionDE = ""
        silencerOptionEN = ""
        connTypeDE = ""
        connTypeEN = ""
        connTypeDetailDE = ""
        connTypeDetailEN = ""
        housingDesignDE = ""
        housingDesignEN = ""
        revision = ""
        optionOne = ""
        optionTwo = ""
        FDAComplianceDE = ""
        FDAComplianceEN = ""
        explosionDE = ""
        explosionEN = ""
        maxSolidSize = ""
        flowRate = ""
        maxDischargePressure = ""
        conveyCapacity = ""
        connSizeSuction = ""
        connSizePressure = ""
        suctionHeightWet = ""
        suctionHeightDry = ""
        airConnInlet = ""
        airConnOutlet = ""
        weight = ""
        length = ""
        width = ""
        height = ""
        memMaterialTempMin = ""
        memMaterialTempMax = ""

        articleNum = wsInput.Cells(i, 1).Value ' Read article number

        ' Get parameters from article number
        brandChar = Mid(articleNum, 1, 2)
        remainedArticleNumber = Mid(articleNum, 3)

        modelChar = Mid(remainedArticleNumber, 1, 1)
        If modelChar = "R" Then
            modelChar = Mid(remainedArticleNumber, 1, 2)
            If modelChar = "RE" Then
                remainedArticleNumber = Mid(remainedArticleNumber, 3)
            Else
                modelChar = Mid(remainedArticleNumber, 1, 1)
                remainedArticleNumber = Mid(remainedArticleNumber, 2)
            End If
        Else
            remainedArticleNumber = Mid(remainedArticleNumber, 2)
        End If

        connSizeChar = Mid(remainedArticleNumber, 1, 1)
        If connSizeChar = "1" Or connSizeChar = "2" Or connSizeChar = "4" Then
            connSizeChar = Mid(remainedArticleNumber, 1, 2)
            If connSizeChar = "10" Or connSizeChar = "20" Or connSizeChar = "4D" Or connSizeChar = "40" Then
                remainedArticleNumber = Mid(remainedArticleNumber, 3)
            Else
                connSizeChar = Mid(remainedArticleNumber, 1, 1)
                remainedArticleNumber = Mid(remainedArticleNumber, 2)
            End If
        Else
            remainedArticleNumber = Mid(remainedArticleNumber, 2)
        End If

        checkValveTypeChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        housingWetChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        housingNotwetChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        memMaterialChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        replacementMemChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        ' memDesignChar = Mid(remainedArticleNumber, 1, 1)
        ' remainedArticleNumber = Mid(remainedArticleNumber, 2)

        checkValveChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        valveSeatChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        airValveChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        airValveOptionChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        silencerOptionChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        connTypeChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        connTypeDetailChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        ' housingDesignChar = Mid(remainedArticleNumber, 1, 1)
        ' remainedArticleNumber = Mid(remainedArticleNumber, 2)

        revisionChar = Mid(remainedArticleNumber, 1, 1)
        remainedArticleNumber = Mid(remainedArticleNumber, 2)

        ' If Len(remainedArticleNumber) > 0 Then
        '     revisionChar = Mid(remainedArticleNumber, 1, 1)
        '     If revisionChar = "-" Then
        '         revisionChar = ""
        '         remainedArticleNumber = Mid(remainedArticleNumber, 1)
        '     Else
        '         remainedArticleNumber = Mid(remainedArticleNumber, 2)
        '     End If
        ' Else
        '     revisionChar = ""
        ' End If

        ' If Len(remainedArticleNumber) > 0 Then
        '     firstChar = Mid(remainedArticleNumber, 1, 1)
        '     If firstChar = "-" Then
        '         If InStr(2, remainedArticleNumber, "-") > 0 Then
        '             optionOneChar = Mid(remainedArticleNumber, 1, InStr(2, remainedArticleNumber, "-") - 1)
        '             optionTwoChar = Mid(remainedArticleNumber, InStr(2, remainedArticleNumber, "-"))
        '         Else
        '             optionOneChar = remainedArticleNumber
        '             optionTwoChar = ""
        '         End If
        '     Else
        '         optionOneChar = ""
        '         optionTwoChar = ""
        '     End If
        ' Else
        '     optionOneChar = ""
        '     optionTwoChar = ""
        ' End If

        ' MsgBox "Brand: " & brandChar & vbNewLine & _
        '         "Model: " & modelChar & vbNewLine & _
        '         "Connection Size: " & connSizeChar & vbNewLine & _
        '         "Check Valve Type: " & checkValveTypeChar & vbNewLine & _
        '         "Housing Wet: " & housingWetChar & vbNewLine & _
        '         "Housing Notwet: " & housingNotwetChar & vbNewLine & _
        '         "Membrane Material: " & memMaterialChar & vbNewLine & _
        '         "Replacement Membrane: " & replacementMemChar & vbNewLine & _
        '         "Check Valve: " & checkValveChar & vbNewLine & _
        '         "Valve Seat: " & valveSeatChar & vbNewLine & _
        '         "Air Valve: " & airValveChar & vbNewLine & _
        '         "Air Valve Option: " & airValveOptionChar & vbNewLine & _
        '         "Silencer Option: " & silencerOptionChar & vbNewLine & _
        '         "Connection Type: " & connTypeChar & vbNewLine & _
        '         "Connection Type Detail: " & connTypeDetailChar & vbNewLine & _
        '         "Revision: " & revisionChar

        For j = 1 To UBound(brandData, 1)
            If brandData(j, 1) = brandChar Then
                brand = brandData(j, 2)
                Exit For
            End If
        Next j

        For j = 1 To UBound(modelData, 1)
            If modelData(j, 1) = modelChar Then
                model = modelData(j, 2)
                Exit For
            End If
        Next j

        For j = 1 To UBound(connSizeData, 1)
            If connSizeData(j, 1) = modelChar Or connSizeData(j, 1) = "" Then
                If connSizeData(j, 2) = connSizeChar Then
                    connSizeInch = connSizeData(j, 6)
                    connSizeMM = connSizeData(j, 8)
                    Exit For
                End If
            End If
        Next j

        For j = 1 To UBound(checkValveTypeData, 1)
            If checkValveTypeData(j, 1) = modelChar Then
                If checkValveTypeData(j, 2) = checkValveTypeChar Then
                    checkValveTypeDE = checkValveTypeData(j, 3)
                    checkValveTypeEN = checkValveTypeData(j, 4)
                    Exit For
                End If
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

        For j = 1 To UBound(backupMembraneData, 1)
            If backupMembraneData(j, 1) = modelChar Then
                If backupMembraneData(j, 2) = replacementMemChar Then
                    replacementMemDE = backupMembraneData(j, 3)
                    replacementMemEN = backupMembraneData(j, 4)
                    Exit For
                End If
            End If
        Next j

        ' For j = 1 To UBound(memDesignData, 1)
        '     If memDesignData(j, 1) = memDesignChar Then
        '         memDesignDE = memDesignData(j, 2)
        '         memDesignEN = memDesignData(j, 3)
        '         Exit For
        '     End If
        ' Next j

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

        For j = 1 To UBound(airValveData, 1)
            If airValveData(j, 1) = airValveChar Then
                airValveDE = airValveData(j, 2)
                airValveEN = airValveData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(airValveOptionData, 1)
            If airValveOptionData(j, 1) = airValveOptionChar Then
                airValveOptionDE = airValveOptionData(j, 2)
                airValveOptionEN = airValveOptionData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(silencerOptionData, 1)
            If silencerOptionData(j, 1) = silencerOptionChar Then
                silencerOptionDE = silencerOptionData(j, 2)
                silencerOptionEN = silencerOptionData(j, 3)
                Exit For
            End If
        Next j

        For j = 1 To UBound(connectionTypeData, 1)
            If connectionTypeData(j, 1) = modelChar Then
                If connectionTypeData(j, 2) = connTypeChar Or connectionTypeData(j, 2) = "" Then
                    connTypeDE = connectionTypeData(j, 3)
                    connTypeEN = connectionTypeData(j, 4)
                    Exit For
                End If
            End If
        Next j

        For j = 1 To UBound(connectionOptionData, 1)
            If connectionOptionData(j, 1) = modelChar Then
                If connectionOptionData(j, 2) = connSizeChar Then
                    If connectionOptionData(j, 3) = connTypeChar Or connectionOptionData(j, 3) = "" Then
                        If connectionOptionData(j, 4) = housingWetChar Or connectionOptionData(j, 4) = "" Then
                            If connectionOptionData(j, 5) = housingNotwetChar Or connectionOptionData(j, 5) = "" Then
                                connTypeOption = connectionOptionData(j, 10)
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next j

        ' For j = 1 To UBound(housingDesignData, 1)
        '     If housingDesignData(j, 1) = housingDesignChar Then
        '         housingDesignDE = housingDesignData(j, 2)
        '         housingDesignEN = housingDesignData(j, 3)
        '         Exit For
        '     End If
        ' Next j

        For j = 1 To UBound(revisionData, 1)
            If revisionData(j, 1) = revisionChar Then
                revision = revisionData(j, 2)
                Exit For
            End If
        Next j

        ' For j = 1 To UBound(optionsData, 1)
        '     If optionsData(j, 1) = optionOneChar Then
        '         optionOne = optionsData(j, 3)
        '         Exit For
        '     End If
        ' Next j

        ' For j = 1 To UBound(optionsData, 1)
        '     If optionsData(j, 1) = optionTwoChar Then
        '         optionTwo = optionsData(j, 3)
        '         Exit For
        '     End If
        ' Next j

        ' For j = 1 To UBound(FDAData, 1)
        '     If FDAData(j, 1) = modelChar And FDAData(j, 3) = housingWetChar And FDAData(j, 5) = memMaterialChar And FDAData(j, 7) = checkValveChar And FDAData(j, 8) = valveSeatChar And FDAData(j, 11) = optionOneChar Then
        '         If FDAData(j, 2) = connSizeChar Or FDAData(j, 2) = "" Then
        '             If FDAData(j, 4) = housingNotwetChar Or FDAData(j, 4) = "" Then
        '                 FDAComplianceDE = FDAData(j, 12)
        '                 FDAComplianceEN = FDAData(j, 13)
        '                 Exit For
        '             End If
        '         End If
        '     Else
        '         FDAComplianceDE = FDAData(1, 15)
        '         FDAComplianceEN = FDAData(1, 15)
        '     End If
        ' Next j

        For j = 1 To UBound(maxSolidSizeData, 1)
            If maxSolidSizeData(j, 1) = modelChar Or maxSolidSizeData(j, 1) = "" Then
                If maxSolidSizeData(j, 2) = connSizeChar Then
                    If maxSolidSizeData(j, 3) = housingWetChar Or maxSolidSizeData(j, 3) = "" Then
                        If maxSolidSizeData(j, 4) = housingNotwetChar Or maxSolidSizeData(j, 4) = "" Then
                            maxSolidSize = maxSolidSizeData(j, 9)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(flowRateData, 1)
            If flowRateData(j, 1) = modelChar Or flowRateData(j, 1) = "" Then
                If flowRateData(j, 2) = connSizeChar Then
                    If flowRateData(j, 3) = housingWetChar Or flowRateData(j, 3) = "" Then
                        flowRate = flowRateData(j, 9)
                        Exit For
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(maxDischargePressureData, 1)
            If maxDischargePressureData(j, 1) = connSizeChar Or maxDischargePressureData(j, 1) = "" Then
                If maxDischargePressureData(j, 2) = housingWetChar Or maxDischargePressureData(j, 2) = "" Then
                    If maxDischargePressureData(j, 3) = housingNotwetChar Or maxDischargePressureData(j, 3) = "" Then
                        maxDischargePressure = maxDischargePressureData(j, 7)
                        Exit For
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(conveyCapacityData, 1)
            If conveyCapacityData(j, 1) = modelChar Or conveyCapacityData(j, 1) = "" Then
                If conveyCapacityData(j, 2) = connSizeChar Then
                    If conveyCapacityData(j, 3) = housingWetChar Or conveyCapacityData(j, 3) = "" Then
                        If conveyCapacityData(j, 4) = housingNotwetChar Or conveyCapacityData(j, 4) = "" Then
                            If conveyCapacityData(j, 5) = memMaterialChar Or conveyCapacityData(j, 5) = "" Then
                                conveyCapacity = conveyCapacityData(j, 12)
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next j

        ' For j = 1 To UBound(connectionTypeData, 1)
        '     If connectionTypeData(j, 1) = modelChar Then
        '         If connectionTypeData(j, 2) = connSizeChar Then
        '             If connectionTypeData(j, 3) = housingWetChar Or connectionTypeData(j, 3) = "" Then
        '                 If connectionTypeData(j, 4) = housingNotwetChar Or connectionTypeData(j, 4) = "" Then
        '                     If connectionTypeData(j, 5) = housingDesignChar Or connectionTypeData(j, 5) = "" Then
        '                         If connectionTypeData(j, 6) = optionOneChar Or connectionTypeData(j, 6) = "" Then
        '                             If connectionTypeData(j, 7) = optionTwoChar Or connectionTypeData(j, 7) = "" Then
        '                                 connectionType = connectionTypeData(j, 9)
        '                                 Exit For
        '                             End If
        '                         End If
        '                     End If
        '                 End If
        '             End If
        '         End If
        '     End If
        ' Next j

        For j = 1 To UBound(connSizeSuctionData, 1)
            If connSizeSuctionData(j, 1) = modelChar Or connSizeSuctionData(j, 1) = "" Then
                If connSizeSuctionData(j, 2) = connSizeChar Then
                    If connSizeSuctionData(j, 3) = connTypeChar Or connSizeSuctionData(j, 3) = "" Then
                        connSizeSuction = connSizeSuctionData(j, 8)
                        Exit For
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(connSizePressureData, 1)
            If connSizePressureData(j, 1) = modelChar Or connSizePressureData(j, 1) = "" Then
                If connSizePressureData(j, 2) = connSizeChar Then
                    If connSizePressureData(j, 3) = connTypeChar Or connSizePressureData(j, 3) = "" Then
                        connSizePressure = connSizePressureData(j, 8)
                        Exit For
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(suctionHeightWetData, 1)
            If suctionHeightWetData(j, 1) = modelChar Or suctionHeightWetData(j, 1) = "" Then
                If suctionHeightWetData(j, 2) = connSizeChar Then
                    If suctionHeightWetData(j, 3) = housingWetChar Or suctionHeightWetData(j, 3) = "" Then
                        If suctionHeightWetData(j, 5) = memMaterialChar Or suctionHeightWetData(j, 3) = "" Then
                            suctionHeightWet = suctionHeightWetData(j, 10)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(suctionHeightDryData, 1)
            If suctionHeightDryData(j, 1) = modelChar Or suctionHeightDryData(j, 1) = "" Then
                If suctionHeightDryData(j, 2) = connSizeChar Then
                    If suctionHeightDryData(j, 3) = housingWetChar Or suctionHeightDryData(j, 3) = "" Then
                        If suctionHeightDryData(j, 5) = memMaterialChar Or suctionHeightDryData(j, 3) = "" Then
                            suctionHeightDry = suctionHeightDryData(j, 10)
                            Exit For
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(airConnInletData, 1)
            If airConnInletData(j, 1) = modelChar Or airConnInletData(j, 1) = "" Then
                If airConnInletData(j, 2) = connSizeChar Then
                    If airConnInletData(j, 3) = housingWetChar Or airConnInletData(j, 3) = "" Then
                        airConnInlet = airConnInletData(j, 8)
                        Exit For
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(airConnOutletData, 1)
            If airConnOutletData(j, 1) = modelChar Then
                If airConnOutletData(j, 2) = connSizeChar Then
                    airConnOutlet = airConnOutletData(j, 7)
                    Exit For
                End If
            End If
        Next j

        For j = 1 To UBound(explosionData, 1)
            If explosionData(j, 1) = modelChar Or explosionData(j, 1) = "" Then
                If explosionData(j, 2) = housingWetChar Or explosionData(j, 2) = "" Then
                    If explosionData(j, 3) = housingNotwetChar Or explosionData(j, 3) = "" Then
                        If explosionData(j, 4) = airValveChar Or explosionData(j, 4) = "" Then
                            If explosionData(j, 5) = airValveOptionChar Or explosionData(j, 5) = "" Then
                                If explosionData(j, 6) = silencerOptionChar Or explosionData(j, 6) = "" Then
                                    explosionDE = explosionData(j, 8)
                                    explosionEN = explosionData(j, 9)
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next j

        For j = 1 To UBound(dimensionsData, 1)
            If dimensionsData(j, 1) = modelChar Or dimensionsData(j, 1) = "" Then
                If dimensionsData(j, 2) = connSizeChar Then
                    If dimensionsData(j, 3) = housingWetChar Or dimensionsData(j, 3) = "" Then
                        If dimensionsData(j, 4) = housingNotwetChar Or dimensionsData(j, 4) = "" Then
                            If dimensionsData(j, 6) = connTypeChar Or dimensionsData(j, 6) = "" Then
                                If dimensionsData(j, 7) = connTypeOption Or dimensionsData(j, 7) = "" Then
                                    If dimensionsData(j, 11) = silencerOptionChar Or dimensionsData(j, 11) = "" Then
                                        If dimensionsData(j, 12) = connTypeDetailChar Or dimensionsData(j, 12) = "" Then
                                            weight = dimensionsData(j, 13)
                                            length = dimensionsData(j, 14)
                                            width = dimensionsData(j, 15)
                                            height = dimensionsData(j, 16)
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        
        For j = 1 To UBound(memMaterialTempData, 1)
            If memMaterialTempData(j, 1) = memMaterialChar Then
                memMaterialTempMin = memMaterialTempData(j, 4)
                memMaterialTempMax = memMaterialTempData(j, 5)
                Exit For
            End If
        Next j
        
        If variableDictDE Is Nothing Then
            Set variableDictDE = New Dictionary
            variableDictDE.add "model", model
            variableDictDE.add "connSizeInch", connSizeInch
            variableDictDE.add "articleNum", articleNum
            variableDictDE.add "conveyCapacity", conveyCapacity
            variableDictDE.add "maxDischargePressure", maxDischargePressure
            variableDictDE.add "housingWet", housingWetDE
            variableDictDE.add "memMaterial", memMaterialDE
            variableDictDE.add "maxSolidSize", maxSolidSize
        Else
            variableDictDE("model") = model
            variableDictDE("connSizeInch") = connSizeInch
            variableDictDE("articleNum") = articleNum
            variableDictDE("conveyCapacity") = conveyCapacity
            variableDictDE("maxDischargePressure") = maxDischargePressure
            variableDictDE("housingWet") = housingWetDE
            variableDictDE("memMaterial") = memMaterialDE
            variableDictDE("maxSolidSize") = maxSolidSize
        End If
        
        If variableDictEN Is Nothing Then
            Set variableDictEN = New Dictionary
            variableDictEN.add "model", model
            variableDictEN.add "connSizeInch", connSizeInch
            variableDictEN.add "articleNum", articleNum
            variableDictEN.add "conveyCapacity", conveyCapacity
            variableDictEN.add "maxDischargePressure", maxDischargePressure
            variableDictEN.add "housingWet", housingWetEN
            variableDictEN.add "memMaterial", memMaterialEN
            variableDictEN.add "maxSolidSize", maxSolidSize
        Else
            variableDictEN("model") = model
            variableDictEN("connSizeInch") = connSizeInch
            variableDictEN("articleNum") = articleNum
            variableDictEN("conveyCapacity") = conveyCapacity
            variableDictEN("maxDischargePressure") = maxDischargePressure
            variableDictEN("housingWet") = housingWetEN
            variableDictEN("memMaterial") = memMaterialEN
            variableDictEN("maxSolidSize") = maxSolidSize
        End If

        ' Reset OUTPUT sheet and SEO OUTPUT sheet
        For k = 1 To 29
            wsOutput.Cells(outputRow, k).Value = ""
        Next k
        For k = 1 To 9
            wsSeoOutput.Cells(outputRow, k).Value = ""
        Next k

        ' Write data to OUTPUT sheet
        If articleNum <> "" Then
            wsOutput.Cells(outputRow, 1).Value = articleNum
            wsOutput.Cells(outputRow, 1).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 1).Value = ""
            wsOutput.Cells(outputRow, 1).Interior.Color = redColor
        End If

        If brand <> "" Then
            wsOutput.Cells(outputRow, 2).Value = brand
            wsOutput.Cells(outputRow, 2).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 2).Value = ""
            wsOutput.Cells(outputRow, 2).Interior.Color = redColor
        End If

         
        If model <> "" Then
           wsOutput.Cells(outputRow, 3).Value = model
           wsOutput.Cells(outputRow, 3).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 3).Value = ""
            wsOutput.Cells(outputRow, 3).Interior.Color = redColor
        End If

        If connSizeInch <> "" Then
            wsOutput.Cells(outputRow, 4).Value = connSizeInch & " Zoll"
            wsOutput.Cells(outputRow, 4).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 4).Value = ""
            wsOutput.Cells(outputRow, 4).Interior.Color = redColor
        End If

        If checkValveTypeDE <> "" Then
            wsOutput.Cells(outputRow, 5).Value = checkValveTypeDE
            wsOutput.Cells(outputRow, 5).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 5).Value = ""
            wsOutput.Cells(outputRow, 5).Interior.Color = redColor
        End If

        If housingWetDE <> "" Then
            wsOutput.Cells(outputRow, 6).Value = housingWetDE
            wsOutput.Cells(outputRow, 6).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 6).Value = ""
            wsOutput.Cells(outputRow, 6).Interior.Color = redColor
        End If

        If housingNotwetDE <> "" Then
            wsOutput.Cells(outputRow, 7).Value = housingNotwetDE
            wsOutput.Cells(outputRow, 7).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 7).Value = ""
            wsOutput.Cells(outputRow, 7).Interior.Color = redColor
        End If

        If memMaterialDE <> "" Then
            wsOutput.Cells(outputRow, 8).Value = memMaterialDE
            wsOutput.Cells(outputRow, 8).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 8).Value = ""
            wsOutput.Cells(outputRow, 8).Interior.Color = redColor
        End If

        If replacementMemDE <> "" Then
            wsOutput.Cells(outputRow, 9).Value = replacementMemDE
            wsOutput.Cells(outputRow, 9).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 9).Value = ""
            wsOutput.Cells(outputRow, 9).Interior.Color = redColor
        End If

        ' If memDesignDE <> "" Then
        '     wsOutput.Cells(outputRow, 7).Value = memDesignDE
        '     wsOutput.Cells(outputRow, 7).Interior.ColorIndex = xlNone
        ' Else
        '     wsOutput.Cells(outputRow, 7).Value = ""
        '     wsOutput.Cells(outputRow, 7).Interior.Color = redColor
        ' End If

        If checkValveDE <> "" Then
            wsOutput.Cells(outputRow, 10).Value = checkValveDE
            wsOutput.Cells(outputRow, 10).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 10).Value = ""
            wsOutput.Cells(outputRow, 10).Interior.Color = redColor
        End If

        If valveSeatDE <> "" Then
            wsOutput.Cells(outputRow, 11).Value = valveSeatDE
            wsOutput.Cells(outputRow, 11).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 11).Value = ""
            wsOutput.Cells(outputRow, 11).Interior.Color = redColor
        End If

        If airValveDE <> "" Then
            wsOutput.Cells(outputRow, 12).Value = airValveDE
            wsOutput.Cells(outputRow, 12).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 12).Value = ""
            wsOutput.Cells(outputRow, 12).Interior.Color = redColor
        End If

        If airValveOptionDE <> "" Then
            wsOutput.Cells(outputRow, 13).Value = airValveOptionDE
            wsOutput.Cells(outputRow, 13).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 13).Value = ""
            wsOutput.Cells(outputRow, 13).Interior.Color = redColor
        End If

        If silencerOptionDE <> "" Then
            wsOutput.Cells(outputRow, 14).Value = silencerOptionDE
            wsOutput.Cells(outputRow, 14).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 14).Value = ""
            wsOutput.Cells(outputRow, 14).Interior.Color = redColor
        End If

        ' If housingDesignDE <> "" Then
        '     wsOutput.Cells(outputRow, 10).Value = housingDesignDE
        '     wsOutput.Cells(outputRow, 10).Interior.ColorIndex = xlNone
        ' Else
        '     wsOutput.Cells(outputRow, 10).Value = ""
        '     wsOutput.Cells(outputRow, 10).Interior.Color = redColor
        ' End If

        ' If FDAComplianceDE <> "" Then
        '     wsOutput.Cells(outputRow, 11).Value = FDAComplianceDE
        '     wsOutput.Cells(outputRow, 11).Interior.ColorIndex = xlNone
        ' Else
        '     wsOutput.Cells(outputRow, 11).Value = ""
        '     wsOutput.Cells(outputRow, 11).Interior.Color = redColor
        ' End If

        If connTypeDE <> "" Then
            wsOutput.Cells(outputRow, 15).Value = connTypeDE
            wsOutput.Cells(outputRow, 15).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 15).Value = ""
            wsOutput.Cells(outputRow, 15).Interior.Color = redColor
        End If

        If connTypeOption <> "" Then
            wsOutput.Cells(outputRow, 16).Value = connTypeOption
            wsOutput.Cells(outputRow, 16).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 16).Value = ""
            wsOutput.Cells(outputRow, 16).Interior.Color = redColor
        End If

        If revision <> "" Then
            wsOutput.Cells(outputRow, 17).Value = revision
            wsOutput.Cells(outputRow, 17).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 17).Value = ""
            wsOutput.Cells(outputRow, 17).Interior.Color = redColor
        End If

        If maxSolidSize <> "" Then
            wsOutput.Cells(outputRow, 18).Value = maxSolidSize & " mm"
            wsOutput.Cells(outputRow, 18).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 18).Value = ""
            wsOutput.Cells(outputRow, 18).Interior.Color = redColor
        End If

        If flowRate <> "" Then
            wsOutput.Cells(outputRow, 19).Value = flowRate & " Liter"
            wsOutput.Cells(outputRow, 19).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 19).Value = ""
            wsOutput.Cells(outputRow, 19).Interior.Color = redColor
        End If
        
        If maxDischargePressure <> "" Then
            wsOutput.Cells(outputRow, 20).Value = maxDischargePressure & " bar"
            wsOutput.Cells(outputRow, 20).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 20).Value = ""
            wsOutput.Cells(outputRow, 20).Interior.Color = redColor
        End If
        
        If conveyCapacity <> "" Then
            wsOutput.Cells(outputRow, 21).Value = conveyCapacity & " Liter pro Minute"
            wsOutput.Cells(outputRow, 21).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 21).Value = ""
            wsOutput.Cells(outputRow, 21).Interior.Color = redColor
        End If
        
        ' If connectionType <> "" Then
        '     wsOutput.Cells(outputRow, 17).Value = connectionType
        '     wsOutput.Cells(outputRow, 17).Interior.ColorIndex = xlNone
        ' Else
        '     wsOutput.Cells(outputRow, 17).Value = ""
        '     wsOutput.Cells(outputRow, 17).Interior.Color = redColor
        ' End If

        If connSizeSuction <> "" Then
            wsOutput.Cells(outputRow, 22).Value = connSizeSuction & " Zoll"
            wsOutput.Cells(outputRow, 22).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 22).Value = ""
            wsOutput.Cells(outputRow, 22).Interior.Color = redColor
        End If

        If connSizePressure <> "" Then
            wsOutput.Cells(outputRow, 23).Value = connSizePressure & " Zoll"
            wsOutput.Cells(outputRow, 23).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 23).Value = ""
            wsOutput.Cells(outputRow, 23).Interior.Color = redColor
        End If
        
        If suctionHeightWet <> "" Then
            wsOutput.Cells(outputRow, 24).Value = suctionHeightWet & " Meter"
            wsOutput.Cells(outputRow, 24).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 24).Value = ""
            wsOutput.Cells(outputRow, 24).Interior.Color = redColor
        End If

        If suctionHeightDry <> "" Then
            wsOutput.Cells(outputRow, 25).Value = suctionHeightDry & " Meter"
            wsOutput.Cells(outputRow, 25).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 25).Value = ""
            wsOutput.Cells(outputRow, 25).Interior.Color = redColor
        End If

        If airConnInlet <> "" Then
            wsOutput.Cells(outputRow, 26).Value = airConnInlet & " Zoll"
            wsOutput.Cells(outputRow, 26).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 26).Value = ""
            wsOutput.Cells(outputRow, 26).Interior.Color = redColor
        End If

        If airConnOutlet <> "" Then
            wsOutput.Cells(outputRow, 27).Value = airConnOutlet & " Zoll"
            wsOutput.Cells(outputRow, 27).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 27).Value = ""
            wsOutput.Cells(outputRow, 27).Interior.Color = redColor
        End If

        If memMaterialTempMin <> "" Then
            wsOutput.Cells(outputRow, 28).Value = memMaterialTempMin & " " & ChrW(176) & "C"
            wsOutput.Cells(outputRow, 28).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 28).Value = ""
            wsOutput.Cells(outputRow, 28).Interior.Color = redColor
        End If

        If memMaterialTempMax <> "" Then
            wsOutput.Cells(outputRow, 29).Value = memMaterialTempMax & " " & ChrW(176) & "C"
            wsOutput.Cells(outputRow, 29).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 29).Value = ""
            wsOutput.Cells(outputRow, 29).Interior.Color = redColor
        End If

        If explosionDE <> "" Then
            wsOutput.Cells(outputRow, 30).Value = explosionDE
            wsOutput.Cells(outputRow, 30).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 30).Value = ""
            wsOutput.Cells(outputRow, 30).Interior.Color = redColor
        End If
        
        If weight <> "" Then
            wsOutput.Cells(outputRow, 31).Value = weight
            wsOutput.Cells(outputRow, 31).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 31).Value = ""
            wsOutput.Cells(outputRow, 31).Interior.Color = redColor
        End If
        
        If length <> "" Then
            wsOutput.Cells(outputRow, 32).Value = length
            wsOutput.Cells(outputRow, 32).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 32).Value = ""
            wsOutput.Cells(outputRow, 32).Interior.Color = redColor
        End If
        
        If width <> "" Then
            wsOutput.Cells(outputRow, 33).Value = width
            wsOutput.Cells(outputRow, 33).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 33).Value = ""
            wsOutput.Cells(outputRow, 33).Interior.Color = redColor
        End If
        
        If height <> "" Then
            wsOutput.Cells(outputRow, 34).Value = height
            wsOutput.Cells(outputRow, 34).Interior.ColorIndex = xlNone
        Else
            wsOutput.Cells(outputRow, 34).Value = ""
            wsOutput.Cells(outputRow, 34).Interior.Color = redColor
        End If
        
        ' Create SEO fields
        For j = 1 To UBound(seoInputData, 1)
            ' Initialize
            If j = 1 Then
                seoArticleNameDE = ""
                seoUrlPathDE = ""
                seoMetaDescriptionDE = ""
                seoShortDescriptionDE = ""
                seoArticleNameEN = ""
                seoUrlPathEN = ""
                seoMetaDescriptionEN = ""
                seoShortDescriptionEN = ""
            End If
            seoArticleNameDECell = seoInputData(j, 1)
            If seoArticleNameDECell <> "" Then
                ' Debug.Print "seoArticleNameDECell: "; seoArticleNameDECell
                If seoArticleNameDECell Like "*[[]*[]]*" Then
                    seoArticleNameDECell = Mid(seoArticleNameDECell, InStr(seoArticleNameDECell, "[") + 1, InStr(seoArticleNameDECell, "]") - InStr(seoArticleNameDECell, "[") - 1)
                    ' Debug.Print "Refined seoArticleNameDECell: "; seoArticleNameDECell
                    For k = 1 To UBound(variableData, 1)
                        If variableData(k, 1) = seoArticleNameDECell Then
                            seoArticleNameDECell = variableData(k, 2)
                            ' Debug.Print "variableData: "; seoArticleNameDECell
                            If variableDictDE.Exists(seoArticleNameDECell) Then
                                seoArticleNameDECell = variableDictDE.Item(seoArticleNameDECell)
                            End If
                            Exit For
                        End If
                    Next k
                End If
                If seoArticleNameDE = "" Then
                    seoArticleNameDE = seoArticleNameDECell
                Else
                    seoArticleNameDE = seoArticleNameDE & " " & seoArticleNameDECell
                End If
            End If

            seoUrlPathDECell = seoInputData(j, 2)
            If seoUrlPathDECell <> "" Then
                ' Debug.Print "seoUrlPathDECell: "; seoUrlPathDECell
                If seoUrlPathDECell Like "*[[]*[]]*" Then
                    seoUrlPathDECell = Mid(seoUrlPathDECell, InStr(seoUrlPathDECell, "[") + 1, InStr(seoUrlPathDECell, "]") - InStr(seoUrlPathDECell, "[") - 1)
                    ' Debug.Print "Refined seoUrlPathDECell: "; seoUrlPathDECell
                    For k = 1 To UBound(variableData, 1)
                        If variableData(k, 1) = seoUrlPathDECell Then
                            seoUrlPathDECell = variableData(k, 2)
                            ' Debug.Print "variableData: "; seoUrlPathDECell
                            If variableDictDE.Exists(seoUrlPathDECell) Then
                                seoUrlPathDECell = variableDictDE.Item(seoUrlPathDECell)
                            End If
                            Exit For
                        End If
                    Next k
                End If
                If seoUrlPathDE = "" Then
                    seoUrlPathDE = seoUrlPathDECell
                Else
                    seoUrlPathDE = seoUrlPathDE & " " & seoUrlPathDECell
                End If
            End If

            seoMetaDescriptionDECell = seoInputData(j, 3)
            If seoMetaDescriptionDECell <> "" Then
                ' Debug.Print "seoMetaDescriptionDECell: "; seoMetaDescriptionDECell
                If seoMetaDescriptionDECell Like "*[[]*[]]*" Then
                    seoMetaDescriptionDECell = Mid(seoMetaDescriptionDECell, InStr(seoMetaDescriptionDECell, "[") + 1, InStr(seoMetaDescriptionDECell, "]") - InStr(seoMetaDescriptionDECell, "[") - 1)
                    ' Debug.Print "Refined seoMetaDescriptionDECell: "; seoMetaDescriptionDECell
                    For k = 1 To UBound(variableData, 1)
                        If variableData(k, 1) = seoMetaDescriptionDECell Then
                            seoMetaDescriptionDECell = variableData(k, 2)
                            ' Debug.Print "variableData: "; seoMetaDescriptionDECell
                            If variableDictDE.Exists(seoMetaDescriptionDECell) Then
                                seoMetaDescriptionDECell = variableDictDE.Item(seoMetaDescriptionDECell)
                            End If
                            Exit For
                        End If
                    Next k
                End If
                If seoMetaDescriptionDE = "" Then
                    seoMetaDescriptionDE = seoMetaDescriptionDECell
                Else
                    seoMetaDescriptionDE = seoMetaDescriptionDE & " " & seoMetaDescriptionDECell
                End If
            End If

            seoShortDescriptionDECell = seoInputData(j, 4)
            If seoShortDescriptionDECell <> "" Then
                ' Debug.Print "seoShortDescriptionDECell: "; seoShortDescriptionDECell
                If seoShortDescriptionDECell Like "*[[]*[]]*" Then
                    seoShortDescriptionDECell = Mid(seoShortDescriptionDECell, InStr(seoShortDescriptionDECell, "[") + 1, InStr(seoShortDescriptionDECell, "]") - InStr(seoShortDescriptionDECell, "[") - 1)
                    ' Debug.Print "Refined seoShortDescriptionDECell: "; seoShortDescriptionDECell
                    For k = 1 To UBound(variableData, 1)
                        If variableData(k, 1) = seoShortDescriptionDECell Then
                            seoShortDescriptionDECell = variableData(k, 2)
                            ' Debug.Print "variableData: "; seoShortDescriptionDECell
                            If variableDictDE.Exists(seoShortDescriptionDECell) Then
                                seoShortDescriptionDECell = variableDictDE.Item(seoShortDescriptionDECell)
                            End If
                            Exit For
                        End If
                    Next k
                End If
                If seoShortDescriptionDE = "" Then
                    seoShortDescriptionDE = seoShortDescriptionDECell
                Else
                    seoShortDescriptionDE = seoShortDescriptionDE & " " & seoShortDescriptionDECell
                End If
            End If

            seoArticleNameENCell = seoInputData(j, 5)
            If seoArticleNameENCell <> "" Then
                ' Debug.Print "seoArticleNameENCell: "; seoArticleNameENCell
                If seoArticleNameENCell Like "*[[]*[]]*" Then
                    seoArticleNameENCell = Mid(seoArticleNameENCell, InStr(seoArticleNameENCell, "[") + 1, InStr(seoArticleNameENCell, "]") - InStr(seoArticleNameENCell, "[") - 1)
                    ' Debug.Print "Refined seoArticleNameENCell: "; seoArticleNameENCell
                    For k = 1 To UBound(variableData, 1)
                        If variableData(k, 1) = seoArticleNameENCell Then
                            seoArticleNameENCell = variableData(k, 2)
                            ' Debug.Print "variableData: "; seoArticleNameENCell
                            If variableDictEN.Exists(seoArticleNameENCell) Then
                                seoArticleNameENCell = variableDictEN.Item(seoArticleNameENCell)
                            End If
                            Exit For
                        End If
                    Next k
                End If
                If seoArticleNameEN = "" Then
                    seoArticleNameEN = seoArticleNameENCell
                Else
                    seoArticleNameEN = seoArticleNameEN & " " & seoArticleNameENCell
                End If
            End If

            seoUrlPathENCell = seoInputData(j, 6)
            If seoUrlPathENCell <> "" Then
                ' Debug.Print "seoUrlPathENCell: "; seoUrlPathENCell
                If seoUrlPathENCell Like "*[[]*[]]*" Then
                    seoUrlPathENCell = Mid(seoUrlPathENCell, InStr(seoUrlPathENCell, "[") + 1, InStr(seoUrlPathENCell, "]") - InStr(seoUrlPathENCell, "[") - 1)
                    ' Debug.Print "Refined seoUrlPathENCell: "; seoUrlPathENCell
                    For k = 1 To UBound(variableData, 1)
                        If variableData(k, 1) = seoUrlPathENCell Then
                            seoUrlPathENCell = variableData(k, 2)
                            ' Debug.Print "variableData: "; seoUrlPathENCell
                            If variableDictEN.Exists(seoUrlPathENCell) Then
                                seoUrlPathENCell = variableDictEN.Item(seoUrlPathENCell)
                            End If
                            Exit For
                        End If
                    Next k
                End If
                If seoUrlPathEN = "" Then
                    seoUrlPathEN = seoUrlPathENCell
                Else
                    seoUrlPathEN = seoUrlPathEN & " " & seoUrlPathENCell
                End If
            End If

            seoMetaDescriptionENCell = seoInputData(j, 7)
            If seoMetaDescriptionENCell <> "" Then
                ' Debug.Print "seoMetaDescriptionENCell: "; seoMetaDescriptionENCell
                If seoMetaDescriptionENCell Like "*[[]*[]]*" Then
                    seoMetaDescriptionENCell = Mid(seoMetaDescriptionENCell, InStr(seoMetaDescriptionENCell, "[") + 1, InStr(seoMetaDescriptionENCell, "]") - InStr(seoMetaDescriptionENCell, "[") - 1)
                    ' Debug.Print "Refined seoMetaDescriptionENCell: "; seoMetaDescriptionENCell
                    For k = 1 To UBound(variableData, 1)
                        If variableData(k, 1) = seoMetaDescriptionENCell Then
                            seoMetaDescriptionENCell = variableData(k, 2)
                            ' Debug.Print "variableData: "; seoMetaDescriptionENCell
                            If variableDictEN.Exists(seoMetaDescriptionENCell) Then
                                seoMetaDescriptionENCell = variableDictEN.Item(seoMetaDescriptionENCell)
                            End If
                            Exit For
                        End If
                    Next k
                End If
                If seoMetaDescriptionEN = "" Then
                    seoMetaDescriptionEN = seoMetaDescriptionENCell
                Else
                    seoMetaDescriptionEN = seoMetaDescriptionEN & " " & seoMetaDescriptionENCell
                End If
            End If

            seoShortDescriptionENCell = seoInputData(j, 8)
            If seoShortDescriptionENCell <> "" Then
                ' Debug.Print "seoShortDescriptionENCell: "; seoShortDescriptionENCell
                If seoShortDescriptionENCell Like "*[[]*[]]*" Then
                    seoShortDescriptionENCell = Mid(seoShortDescriptionENCell, InStr(seoShortDescriptionENCell, "[") + 1, InStr(seoShortDescriptionENCell, "]") - InStr(seoShortDescriptionENCell, "[") - 1)
                    ' Debug.Print "Refined seoShortDescriptionENCell: "; seoShortDescriptionENCell
                    For k = 1 To UBound(variableData, 1)
                        If variableData(k, 1) = seoShortDescriptionENCell Then
                            seoShortDescriptionENCell = variableData(k, 2)
                            ' Debug.Print "variableData: "; seoShortDescriptionENCell
                            If variableDictEN.Exists(seoShortDescriptionENCell) Then
                                seoShortDescriptionENCell = variableDictEN.Item(seoShortDescriptionENCell)
                            End If
                            Exit For
                        End If
                    Next k
                End If
                If seoShortDescriptionEN = "" Then
                    seoShortDescriptionEN = seoShortDescriptionENCell
                Else
                    seoShortDescriptionEN = seoShortDescriptionEN & " " & seoShortDescriptionENCell
                End If
            End If
        Next j

        ' Write SEO data to SEO OUTPUT sheet
        wsSeoOutput.Cells(outputRow, 1).Value = articleNum
        wsSeoOutput.Cells(outputRow, 2).Value = seoArticleNameDE
        wsSeoOutput.Cells(outputRow, 3).Value = seoUrlPathDE
        wsSeoOutput.Cells(outputRow, 4).Value = seoMetaDescriptionDE
        wsSeoOutput.Cells(outputRow, 5).Value = seoShortDescriptionDE
        wsSeoOutput.Cells(outputRow, 6).Value = seoArticleNameEN
        wsSeoOutput.Cells(outputRow, 7).Value = seoUrlPathEN
        wsSeoOutput.Cells(outputRow, 8).Value = seoMetaDescriptionEN
        wsSeoOutput.Cells(outputRow, 9).Value = seoShortDescriptionEN

        ' Move to next row in OUTPUT sheet
        outputRow = outputRow + 1
    Next i
    
    MsgBox "Article numbers processed successfully!", vbInformation, "Done"
End Sub
