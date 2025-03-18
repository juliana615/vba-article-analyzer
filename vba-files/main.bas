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
    Dim redColor As Long, greenColor As Long
    Dim seoArticleNameDE As String, seoUrlPathDE As String, seoMetaDescriptionDE As String, seoShortDescriptionDE As String, seoArticleNameEN As String, seoUrlPathEN As String, seoMetaDescriptionEN As String, seoShortDescriptionEN As String
    Dim seoArticleNameDECell As String, seoUrlPathDECell As String, seoMetaDescriptionDECell As String, seoShortDescriptionDECell As String, seoArticleNameENCell As String, seoUrlPathENCell As String, seoMetaDescriptionENCell As String, seoShortDescriptionENCell As String
    Dim outputRow As Long
    
    redColor = RGB(252, 228, 214)
    greenColor = RGB(226, 239, 218)

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
            If FDAData(j, 1) = modelChar And FDAData(j, 3) = housingWetChar And FDAData(j, 5) = memMaterialChar And FDAData(j, 7) = checkValveChar And FDAData(j, 8) = valveSeatChar And FDAData(j, 11) = optionOneChar Then
                If FDAData(j, 2) = connSizeChar Or FDAData(j, 2) = "" Then
                    If FDAData(j, 4) = housingNotwetChar Or FDAData(j, 4) = "" Then
                        FDAComplianceDE = FDAData(j, 12)
                        FDAComplianceEN = FDAData(j, 13)
                        Exit For
                    End If
                End If
            Else
                FDAComplianceDE = FDAData(1, 15)
                FDAComplianceEN = FDAData(1, 15)
            End If
        Next j

        For j = 1 To UBound(explosionData, 1)
            If explosionData(j, 1) = optionOneChar Or explosionData(j, 1) = optionTwoChar Then
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
                If maxSolidSizeData(j, 2) = housingWetChar Or maxSolidSizeData(j, 2) = "" Then
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
        
        If variableDictDE Is Nothing Then
            Set variableDictDE = New Dictionary
            variableDictDE.add "connSizeInch", connSizeInch
            variableDictDE.add "articleNum", articleNum
            variableDictDE.add "flowRate", flowRate
            variableDictDE.add "maxDischargePressure", maxDischargePressure
            variableDictDE.add "housingWet", housingWetDE
            variableDictDE.add "memMaterial", memMaterialDE
            variableDictDE.add "maxSolidSize", maxSolidSize
        End If
        
        If variableDictEN Is Nothing Then
            Set variableDictEN = New Dictionary
            variableDictEN.add "connSizeInch", connSizeInch
            variableDictEN.add "articleNum", articleNum
            variableDictEN.add "flowRate", flowRate
            variableDictEN.add "maxDischargePressure", maxDischargePressure
            variableDictEN.add "housingWet", housingWetEN
            variableDictEN.add "memMaterial", memMaterialEN
            variableDictEN.add "maxSolidSize", maxSolidSize
        End If

        ' Reset OUTPUT sheet and SEO OUTPUT sheet
        For k = 1 To 29
            wsOutput.Cells(outputRow, k).Value = ""
        Next k
        For k = 1 To 9
            wsSeoOutput.Cells(outputRow, k).Value = ""
        Next k

        ' Write data to OUTPUT sheet
        wsOutput.Cells(outputRow, 1).Value = articleNum
        wsOutput.Cells(outputRow, 2).Value = model
        wsOutput.Cells(outputRow, 3).Value = connSizeInch & " Zoll"
        wsOutput.Cells(outputRow, 4).Value = housingWetDE
        wsOutput.Cells(outputRow, 5).Value = housingNotwetDE
        wsOutput.Cells(outputRow, 6).Value = memMaterialDE
        wsOutput.Cells(outputRow, 7).Value = memDesignDE
        wsOutput.Cells(outputRow, 8).Value = checkValveDE
        wsOutput.Cells(outputRow, 9).Value = valveSeatDE
        wsOutput.Cells(outputRow, 10).Value = housingDesignDE
        wsOutput.Cells(outputRow, 11).Value = FDAComplianceDE
        wsOutput.Cells(outputRow, 12).Value = explosionDE
        wsOutput.Cells(outputRow, 13).Value = maxSolidSize & " mm"
        wsOutput.Cells(outputRow, 14).Value = flowRate & " Liter"
        wsOutput.Cells(outputRow, 15).Value = maxDischargePressure & " bar"
        wsOutput.Cells(outputRow, 16).Value = conveyCapacity & " Liter pro Minute"
        wsOutput.Cells(outputRow, 17).Value = connectionType
        wsOutput.Cells(outputRow, 18).Value = connSizeSuction & " Zoll"
        wsOutput.Cells(outputRow, 19).Value = connSizePressure & " Zoll"
        wsOutput.Cells(outputRow, 20).Value = suctionHeightWet & " Meter"
        wsOutput.Cells(outputRow, 21).Value = suctionHeightDry & " Meter"
        wsOutput.Cells(outputRow, 22).Value = airConnInlet & " Zoll"
        wsOutput.Cells(outputRow, 23).Value = airConnOutlet & " Zoll"
        wsOutput.Cells(outputRow, 24).Value = memMaterialTempMin & " " & ChrW(176) & "C"
        wsOutput.Cells(outputRow, 25).Value = memMaterialTempMax & " " & ChrW(176) & "C"
        wsOutput.Cells(outputRow, 26).Value = weight
        wsOutput.Cells(outputRow, 27).Value = length
        wsOutput.Cells(outputRow, 28).Value = width
        wsOutput.Cells(outputRow, 29).Value = height
        
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
                            Debug.Print "variableData: "; seoMetaDescriptionDECell
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
