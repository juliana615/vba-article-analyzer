Attribute VB_Name = "main"    

Sub Test()
    MsgBox ModelDictionary().Item("E")
End Sub

' Modell
Public Function ModelDictionary() As Dictionary
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
Public Function ConnectionSizeInchDictionary() As Dictionary
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

    Set ConnectionSizeDictionary = obj
End Function

' Gehäusematerial (benetzt)
Public Function HousingMaterialWetDictionary() As Dictionary
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
Public Function HousingMaterialNonwetDictionary() As Dictionary
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

    Set HousingMaterialDryDictionary = obj
End Function

' Material der Membrane		
Public Function MembraneMaterialDictionary() As Dictionary
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
Public Function MembraneDesignDictionary() As Dictionary
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
Public Function CheckValveMaterialDictionary() As Dictionary
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
Public Function ValveSeatMaterialDictionary() As Dictionary
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
Public Function HousingDesignDictionary() As Dictionary
    Static obj As Dictionary
    
    If obj Is Nothing Then
        Set obj = New Dictionary
        obj.Add "9", "Geschraubt"
        obj.Add "0", "Geklemmt"
    End If

    Set HousingDesignDictionary = obj
End Function

' Revisionslevel	
Public Function RevisionLevelDictionary() As Dictionary
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
Public Function OptionsDictionary() As Dictionary
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

Public Function GetModelName(ByVal key As String) As String
    Dim dict As Dictionary
    Set dict = ModelDictionary()
    GetModelName = dict(key)
End Function

Sub BreakdownArticleName()
    Dim wsInput As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long, i As Long
    Dim articleNum As String
    Dim model As String, connSize As String, housingWet As String
    Dim housingDry As String, memMaterial As String, memDesign As String
    Dim checkValve As String, valveSeat As String, housingDesign As String
    Dim revision As String, options As String
    Dim outputRow As Long
    
    ' Set worksheet references
    Set wsInput = ThisWorkbook.Sheets("INPUT")
    Set wsOutput = ThisWorkbook.Sheets("OUTPUT")

    ' Find last row in INPUT sheet
    lastRow = wsInput.Cells(wsInput.Rows.Count, 1).End(xlUp).Row
    
    ' Set headers in OUTPUT sheet
    ' wsOutput.Cells(1, 1).Value = "Article Number"
    ' wsOutput.Cells(1, 2).Value = "Model"
    ' wsOutput.Cells(1, 3).Value = "Connection Size"
    ' wsOutput.Cells(1, 4).Value = "Housing Material (Wet)"
    ' wsOutput.Cells(1, 5).Value = "Housing Material (Dry)"
    ' wsOutput.Cells(1, 6).Value = "Membrane Material"
    ' wsOutput.Cells(1, 7).Value = "Membrane Design Check"
    ' wsOutput.Cells(1, 8).Value = "Check Valve Material"
    ' wsOutput.Cells(1, 9).Value = "Valve Seat Material"
    ' wsOutput.Cells(1, 10).Value = "Housing Design"
    ' wsOutput.Cells(1, 11).Value = "Revision Level"
    ' wsOutput.Cells(1, 12).Value = "Options"

    ' Loop through each article number
    outputRow = 2 ' Start from row 2 in OUTPUT sheet
    For i = 5 To lastRow
        articleNum = wsInput.Cells(i, 1).Value ' Read article number
        
        ' Check if article number is valid
        If Len(articleNum) >= 11 Then
            ' Extract components based on predefined structure
            model = Mid(articleNum, 1, 1)
            connSize = Mid(articleNum, 2, 1)
            housingWet = Mid(articleNum, 3, 1)
            housingDry = Mid(articleNum, 4, 1)
            memMaterial = Mid(articleNum, 5, 1)
            memDesign = Mid(articleNum, 6, 1)
            checkValve = Mid(articleNum, 7, 1)
            valveSeat = Mid(articleNum, 8, 1)
            housingDesign = Mid(articleNum, 9, 1)
            revision = Mid(articleNum, 10, 1)
            
            ' Extract options (if present)
            If InStr(articleNum, "-") > 0 Then
                options = Mid(articleNum, InStr(articleNum, "-"))
            Else
                options = ""
            End If
            
            ' Write data to OUTPUT sheet
            wsOutput.Cells(outputRow, 1).Value = articleNum
            wsOutput.Cells(outputRow, 2).Value = model
            wsOutput.Cells(outputRow, 3).Value = connSize
            wsOutput.Cells(outputRow, 4).Value = housingWet
            wsOutput.Cells(outputRow, 5).Value = housingDry
            wsOutput.Cells(outputRow, 6).Value = memMaterial
            wsOutput.Cells(outputRow, 7).Value = memDesign
            wsOutput.Cells(outputRow, 8).Value = checkValve
            wsOutput.Cells(outputRow, 9).Value = valveSeat
            wsOutput.Cells(outputRow, 10).Value = housingDesign
            wsOutput.Cells(outputRow, 11).Value = revision
            wsOutput.Cells(outputRow, 12).Value = options
            
            ' Move to next row in OUTPUT sheet
            outputRow = outputRow + 1
        End If
    Next i
    
    MsgBox "Article numbers processed successfully!", vbInformation, "Done"
End Sub
