Attribute VB_Name = "VendorAndPriceLists"
Option Explicit

Sub VendorMaterialListsGen(b As clsBuilding)
Dim VendorSht As Worksheet
Dim WriteCell As Range
Dim item As clsMiscItem
Dim Panel As clsPanel
Dim Trim As clsTrim
Dim n As Integer



'delete old output sheets
Application.DisplayAlerts = False
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Vendor Sheet Metal Materials" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Vendor Misc. Materials" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n
Application.DisplayAlerts = True

''' Vendor Sheet Metal Materials List
'set new output sheet
VendorSheetMetalShtTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set VendorSht = ThisWorkbook.Sheets("VendorSheetMetalShtTmp (2)")
'rename
VendorSht.Name = "Vendor Sheet Metal Materials"
VendorSht.Visible = xlSheetVisible

'populate
Set WriteCell = VendorSht.Range("MatListQtyCell1")
For Each Panel In b.PanelCollection
    WriteCell.Value = Panel.Quantity
    WriteCell.offset(0, 1).Value = Panel.PanelShape
    WriteCell.offset(0, 2).Value = Panel.PanelType
    WriteCell.offset(0, 3).Value = Panel.PanelMeasurement
    WriteCell.offset(0, 4).Value = Panel.PanelColor
    Set WriteCell = WriteCell.offset(1, 0)
Next Panel
For Each Trim In b.TrimCollection
    WriteCell.Value = Trim.Quantity
    WriteCell.offset(0, 1).Value = Trim.tShape
    WriteCell.offset(0, 2).Value = Trim.tType
    WriteCell.offset(0, 3).Value = Trim.tMeasurement
    WriteCell.offset(0, 4).Value = Trim.Color
    Set WriteCell = WriteCell.offset(1, 0)
Next Trim
'write misc items that need to be sent to the sheet metal vendor list
For Each item In b.MiscMaterialsCollection
    With item
        If InStr(1, .Name, "Formed Ridge Cap") <> 0 Or .Name = "Sculptured Gutter End Cap" Or .Name = "Gutter Strap" _
        Or .Name = "Downspout Strap" Or .Name = "Pop Rivets" Or .Name = "Tek Screws" Or .Name = "Lap Screws" _
        Or .Name = "Butyl Tape" Or .Name = "Inside Closures" Or .Name = "Outside Closures" Then
            WriteCell.Value = .Quantity
            WriteCell.offset(0, 1).Value = .Shape
            WriteCell.offset(0, 2).Value = .Name
            WriteCell.offset(0, 3).Value = .Measurement
            WriteCell.offset(0, 4).Value = .Color
            Set WriteCell = WriteCell.offset(1, 0)
        End If
    End With
Next item

    
    
'format
VendorSht.Columns.AutoFit


''' Vendor Misc Materials List
'set new output sheet
VendorMiscMaterialsShtTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set VendorSht = ThisWorkbook.Sheets("VendorMiscMaterialsShtTmp (2)")
'rename sheet
VendorSht.Name = "Vendor Misc. Materials"
VendorSht.Visible = xlSheetVisible
Set WriteCell = VendorSht.Range("MatListQtyCell1")
'write misc items that need to be sent to the misc materials vendor list
For Each item In b.MiscMaterialsCollection
    With item
        If InStr(1, .Name, "Formed Ridge Cap") = 0 And .Name <> "Sculptured Gutter End Cap" And .Name <> "Gutter Strap" _
        And .Name <> "Downspout Strap" And .Name <> "Pop Rivets" And .Name <> "Tek Screws" And .Name <> "Lap Screws" _
        And .Name <> "Butyl Tape" And .Name <> "Inside Closures" And .Name <> "Outside Closures" Then
            WriteCell.Value = .Quantity
            WriteCell.offset(0, 1).Value = .Name
            WriteCell.offset(0, 3).Value = .Measurement
            WriteCell.offset(0, 4).Value = .Color
            Set WriteCell = WriteCell.offset(1, 0)
            'merge cells
            WriteCell.offset(0, 1).Resize(1, 2).Merge
        End If
    End With
Next item

'format
VendorSht.Columns.AutoFit

End Sub

Sub PriceListGen(b As clsBuilding)
Dim item As clsMiscItem
Dim Panel As clsPanel
Dim Trim As clsTrim
Dim LookupCol As Integer
Dim PriceSht As Worksheet
Dim WriteCell As Range
Dim n As Integer
Dim PriceTbl As ListObject
Dim Row As Integer
Dim LookupName As String
Dim SectionalOHDoorPriceTbl As ListObject
Dim PricingQty As Integer
Dim PanelType As String

'set master price table
Set PriceTbl = MasterPriceSht.ListObjects("MasterPriceTbl")
Set SectionalOHDoorPriceTbl = MasterPriceSht.ListObjects("SectionalOHDoorPriceTbl")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Lookup Item Prices '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' Panels
For Each Panel In b.PanelCollection
    With Panel
        If InStr(1, .PanelType, "Skylights") = 0 And InStr(1, .PanelType, "Reverse") = 0 Then
            '''footage cost for normal panels
            'check for errors
            If IsError(Application.VLookup(.PanelType, PriceTbl.Range, 4, False)) = True Then
                .FootageCost = "Unknown"
                .UnitCost = "Unknown"
                .totalCost = "Item Not Found"
            Else    'successful lookup
                .FootageCost = Application.WorksheetFunction.VLookup(.PanelType, PriceTbl.Range, 4, False)
                .UnitCost = .FootageCost * (.PanelLength / 12)
                .totalCost = .UnitCost * .Quantity
            End If
        ElseIf InStr(1, .PanelType, "Skylights") <> 0 Then
            'unit cost for skylight panels
            If IsError(Application.WorksheetFunction.VLookup(.PanelType & ", 12'", PriceTbl.Range, 3, False)) = True Then
                .UnitCost = "Unknown"
                .totalCost = "Item Not Found"
            Else    'successful lookup
                .UnitCost = Application.WorksheetFunction.VLookup(.PanelType & ", 12'", PriceTbl.Range, 3, False)
                .totalCost = .UnitCost * .Quantity
            End If
        ElseIf InStr(1, .PanelType, "Reverse") <> 0 Then
            'unit cost for reverse panels
            'check for errors
            PanelType = Replace(.PanelType, "Reverse ", "")
            If IsError(Application.VLookup(PanelType, PriceTbl.Range, 4, False)) = True Then
                .FootageCost = "Unknown"
                .UnitCost = "Unknown"
                .totalCost = "Item Not Found"
            Else    'successful lookup
                .FootageCost = Application.WorksheetFunction.VLookup(PanelType, PriceTbl.Range, 4, False)
                .UnitCost = .FootageCost * (.PanelLength / 12)
                .totalCost = .UnitCost * .Quantity
            End If
        End If
    End With
Next Panel
''''' Trim
For Each Trim In b.TrimCollection
    With Trim
        'determine color
        Select Case .Color
        Case "Galvalume"
            LookupCol = 7
        Case "Copper Metallic"
            LookupCol = 5
        Case Else
            LookupCol = 6
        End Select
        'trim name
        Select Case True
        ''' Jamb is in collection with trim
        Case InStr(1, .tType, "Jamb W/ Deadbolt") <> 0, InStr(1, .tType, "Jamb W/O Deadbolt") <> 0
            'lookup name
            LookupName = Right(.tType, 4) & " " & Trim.tMeasurement & """" & " Jamb Kit " & _
            Right(Left(.tType, Len(.tType) - 7), Len(Left(.tType, Len(.tType) - 7)) - 5)
            .UnitCost = Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 3, False)
            .totalCost = .UnitCost * .Quantity
        ''' Door Slab are in collection with trim
        Case InStr(1, .tType, "Door Slab") <> 0
            'lookup name
            LookupName = Right(.tType, 4) & " " & Left(.tType, Len(.tType) - 7)
            .UnitCost = Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 3, False)
            'TBD Placeholder
            If .UnitCost <> "TBD" Then
                .totalCost = .UnitCost * .Quantity
            Else
                .totalCost = "TBD"
            End If
        ''' Pitch String Items
        Case InStr(1, .tType, "High-Side Eave") <> 0
            .FootageCost = Application.WorksheetFunction.VLookup("High-Side Eave", PriceTbl.Range, LookupCol, False)
            .UnitCost = .FootageCost * (.tLength / 12)
            .totalCost = .UnitCost * .Quantity
        Case InStr(1, .tType, "Short Eave") <> 0
            .FootageCost = Application.WorksheetFunction.VLookup("Short Eave", PriceTbl.Range, LookupCol, False)
            .UnitCost = .FootageCost * (.tLength / 12)
            .totalCost = .UnitCost * .Quantity
        Case InStr(1, .tType, "Sculptured Gutter Hang-On") <> 0
            .FootageCost = Application.WorksheetFunction.VLookup("Sculptured Gutter Hang-On", PriceTbl.Range, LookupCol, False)
            .UnitCost = .FootageCost * (.tLength / 12)
            .totalCost = .UnitCost * .Quantity
        ''' Head Trim (assumed to be the same cost with or without kickout)
        Case InStr(1, .tType, "Head Trim") <> 0
            .FootageCost = Application.WorksheetFunction.VLookup("Head Trim", PriceTbl.Range, LookupCol, False)
            .UnitCost = .FootageCost * (.tLength / 12)
            .totalCost = .UnitCost * .Quantity
        ''' Flat Rate Items
        Case InStr(1, .tType, "Formed Ridge Cap") <> 0, InStr(1, .tType, "Sculptured Gutter End Cap") <> 0, InStr(1, .tType, "Gutter Strap") <> 0
            'only a flat unit cost
            .UnitCost = Application.WorksheetFunction.VLookup(.tType, PriceTbl.Range, LookupCol + 3, False)
            .totalCost = .UnitCost * .Quantity
        ''' Normal Items
        Case Else
            ''' Perform vlookup as normal
            'Check for errors
            If IsError(Application.VLookup(.tType, PriceTbl.Range, LookupCol, False)) = True Then
                .FootageCost = "Unknown"
                .UnitCost = "Unknown"
                .totalCost = "Item Not Found"
            Else    ' No lookup error
                .FootageCost = Application.WorksheetFunction.VLookup(.tType, PriceTbl.Range, LookupCol, False)
                .UnitCost = .FootageCost * (.tLength / 12)
                .totalCost = .UnitCost * .Quantity
            End If
        End Select
    End With
Next Trim

'''''' Misc Items
For Each item In b.MiscMaterialsCollection
    With item
        'determine color
        Select Case .Color
        Case "Galvalume"
            LookupCol = 7
        Case "Copper Metallic"
            LookupCol = 5
        Case Else
            LookupCol = 6
        End Select
        Select Case True
        ''''''''''''''''''''''''''''''''''''''''''''''''' Items Priced by Bulk Quantity
        Case InStr(1, .Name, "Pop Rivets") <> 0, InStr(1, .Name, "Tek Screws") <> 0, InStr(1, .Name, "Lap Screws") <> 0
            If .Name = "Pop Rivets" Then
                PricingQty = .Quantity
            ElseIf .Name = "Tek Screws" Or .Name = "Lap Screws" Then
                PricingQty = Application.WorksheetFunction.RoundUp(.Quantity / 250, 0)
            End If
    
            .UnitCost = Application.WorksheetFunction.VLookup(.Name, PriceTbl.Range, 3, False)
            .totalCost = .UnitCost * PricingQty
        'Sectional OH Doors
        Case InStr(1, .Name, "Sectional OH Door") <> 0
            ''''find width, height
            On Error Resume Next
            .UnitCost = "Size Not Found"
            .totalCost = "Size Not Found"
            For Row = 1 To SectionalOHDoorPriceTbl.ListRows.Count
                If SectionalOHDoorPriceTbl.DataBodyRange(Row, 1) = .Width Then
                    .UnitCost = SectionalOHDoorPriceTbl.DataBodyRange(Row, SectionalOHDoorPriceTbl.ListColumns(CStr(.Height)).Index)
                    .totalCost = .UnitCost * .Quantity
                    Exit For
                End If
            Next Row
            'resume error handling
            On Error GoTo 0
                
            'masterpricesht.ListObjects("SectionalOHDoorPriceTbl").DataBodyRange
        'Flat Rate Colored Items
        Case InStr(1, .Name, "Formed Ridge Cap") <> 0, InStr(1, .Name, "Sculptured Gutter End Cap") <> 0, InStr(1, .Name, "Gutter Strap") <> 0
            '''only a flat unit cost
            ''' Pitch String
            If InStr(1, .Name, "Formed Ridge Cap") <> 0 Then
                .UnitCost = Application.WorksheetFunction.VLookup("Formed Ridge Cap", PriceTbl.Range, LookupCol + 3, False)
                .totalCost = .UnitCost * .Quantity
            Else    ''' Normal Items
                .UnitCost = Application.WorksheetFunction.VLookup(.Name, PriceTbl.Range, LookupCol + 3, False)
                .totalCost = .UnitCost * .Quantity
            End If
        ''' other items
        Case Else
            '''items with differing naming conventions between materials list and master price list
            If InStr(1, .Name, "Wall Insulation") <> 0 Then
                'remove "Wall" from name, lookup
                LookupName = Left(.Name, InStr(1, .Name, " Wall")) & "Insulation"
                .UnitCost = Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 4, False)
                .totalCost = .UnitCost * .Quantity
            ElseIf InStr(1, .Name, "Roof Insulation") <> 0 Then
                'remove "Wall" from name, lookup
                LookupName = Left(.Name, InStr(1, .Name, " Roof")) & "Insulation"
                .UnitCost = Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 4, False)
                .totalCost = .UnitCost * .Quantity
            ElseIf InStr(1, .Name, "High Lift") <> 0 Or InStr(1, .Name, "Door Canopy") <> 0 Or _
            InStr(1, .Name, "Exhaust Fan") <> 0 Or InStr(1, .Name, "Louver") <> 0 Or InStr(1, .Name, "Weather Hood") <> 0 Then
                'put measurement first
                LookupName = .Measurement & " " & .Name
                .UnitCost = Application.WorksheetFunction.VLookup(LookupName, PriceTbl.Range, 3, False)
                .totalCost = .UnitCost * .Quantity
            End If
            
            
            '''' Normal Items
            For Row = 1 To PriceTbl.ListRows.Count
                If PriceTbl.DataBodyRange(Row, 1) = .Name Then
                    'check for price per measurement data
                    If PriceTbl.DataBodyRange(Row, 4) <> "-" And PriceTbl.DataBodyRange(Row, 4) <> "" Then
                        'items needed to be priced by area that doesn't match to the stored quantity
                        If InStr(1, .Name, "Roll Up OH Door") <> 0 Or InStr(1, .Name, "Standard Window") <> 0 Or _
                        InStr(1, .Name, "Full Glass Panel Window") <> 0 Then
                            'footage cost is actually cost per SF but keeping var name for now
                            .FootageCost = PriceTbl.DataBodyRange(Row, 4)
                            .UnitCost = .Area * PriceTbl.DataBodyRange(Row, 4)
                            .totalCost = .UnitCost * .Quantity
                        'Items with Quantity matching measurement
                        Else
                            .UnitCost = PriceTbl.DataBodyRange(Row, 4)
                            .totalCost = .UnitCost * .Quantity
                        End If
                    'check for flat rate data
                    ElseIf PriceTbl.DataBodyRange(Row, 3) <> "-" And PriceTbl.DataBodyRange(Row, 3) <> "" Then
                        .UnitCost = PriceTbl.DataBodyRange(Row, 3)
                        'TBD Placeholder
                        If .UnitCost <> "TBD" Then
                            .totalCost = .UnitCost * .Quantity
                        Else
                            .totalCost = "TBD"
                        End If
                    End If
                    Exit For
                End If
            Next Row
               
            ''' Electric Opener and Unknown Items
            'Electric Opener
            If InStr(1, .Name, "Electric Opener") <> 0 Then
                .UnitCost = "Input Required"
                .totalCost = "Input Required"
            ' handle unknown items
            ElseIf .totalCost = "" Then
                .FootageCost = "Unknown"
                .UnitCost = "Unknown"
                .totalCost = "Item Not Found"
            End If
        End Select
    End With
Next item


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Generate Price Sheet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'delete old output sheet
Application.DisplayAlerts = False
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Materials Price List" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n
Application.DisplayAlerts = True

'set new output sheet
MaterialsPriceShtTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set PriceSht = ThisWorkbook.Sheets("MaterialsPriceShtTmp (2)")
'rename
PriceSht.Name = "Materials Price List"
PriceSht.Visible = xlSheetVisible


' Begin Output
Set WriteCell = PriceSht.Range("PriceListQtyCell1")
For Each Panel In b.PanelCollection
    WriteCell.Value = Panel.Quantity
    WriteCell.offset(0, 1).Value = Panel.PanelShape
    WriteCell.offset(0, 2).Value = Panel.PanelType
    WriteCell.offset(0, 3).Value = Panel.PanelMeasurement
    WriteCell.offset(0, 4).Value = Panel.PanelColor
    WriteCell.offset(0, 5).Value = Panel.FootageCost
    WriteCell.offset(0, 6).Value = Panel.UnitCost
    WriteCell.offset(0, 7).Value = Panel.totalCost
    Set WriteCell = WriteCell.offset(1, 0)
Next Panel
For Each Trim In b.TrimCollection
    WriteCell.Value = Trim.Quantity
    WriteCell.offset(0, 1).Value = Trim.tShape
    WriteCell.offset(0, 2).Value = Trim.tType
    WriteCell.offset(0, 3).Value = Trim.tMeasurement
    WriteCell.offset(0, 4).Value = Trim.Color
    WriteCell.offset(0, 5).Value = Trim.FootageCost
    WriteCell.offset(0, 6).Value = Trim.UnitCost
    WriteCell.offset(0, 7).Value = Trim.totalCost
    Set WriteCell = WriteCell.offset(1, 0)
Next Trim
For Each item In b.MiscMaterialsCollection
    WriteCell.Value = item.Quantity
    WriteCell.offset(0, 1).Value = item.Shape
    WriteCell.offset(0, 2).Value = item.Name
    WriteCell.offset(0, 3).Value = item.Measurement
    WriteCell.offset(0, 4).Value = item.Color
    WriteCell.offset(0, 5).Value = item.FootageCost
    WriteCell.offset(0, 6).Value = item.UnitCost
    WriteCell.offset(0, 7).Value = item.totalCost
    Set WriteCell = WriteCell.offset(1, 0)
Next item
'format
PriceSht.Columns.AutoFit


Dim LastRow As Double
Dim mCell As Range
Dim MissingPrice As Boolean
With PriceSht
    LastRow = .Cells(.Rows.Count, "H").End(xlUp).Row
    For Each mCell In .Range("H4:H" & LastRow)
        If IsNumeric(mCell.Value) = False Then
            MissingPrice = True
            Exit For
        End If
    Next mCell
End With

If MissingPrice Then MsgBox "At least one price could  not be found on the materials price list.", vbExclamation, "Missing Price"
        
End Sub

Sub CostEstimateGen(b As clsBuilding)
Dim StructuralSteel As New Collection
Dim Sheetmetal As New Collection
Dim OHDoors As New Collection
Dim Insulation As New Collection
Dim ElectricOpeners As New Collection
Dim Windows As New Collection
Dim RidgeVents As New Collection
Dim DoorCanopies As New Collection
Dim ExhaustFansLouversWeatherhoods As New Collection
Dim clsItem As clsMiscItem
Dim clsPanel As clsPanel
Dim clsTrim As clsTrim
Dim CollectionItem As Variant
'Material Name Sorting Arrays
Dim OHDoorNames() As Variant
Dim InsulationNames() As Variant
Dim CollectionTotal As Currency
Dim n As Integer
Dim CostEstimateSht As Worksheet

'Oh Door Material Collections
OHDoorNames = Array("Roll Up OH Door", "Sectional OH Door", "Chain Hoist Opener", "High Lift", "Non-Insulated Window", "Insulated Window", "Full Glass Panel Window", _
"Vinyl Backed Insulation", "Steel Backed Insulation")
InsulationNames = Array("3"" VRR", "4"" VRR", "6"" VRR", "1"" Spray Foam", "2"" Spray Foam")


'delete old output sheets
Application.DisplayAlerts = False
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Cost Estimate" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n
Application.DisplayAlerts = True

''' Vendor Sheet Metal Materials List
'set new output sheet
CostEstimateShtTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set CostEstimateSht = ThisWorkbook.Sheets("CostEstimateShtTmp (2)")
'rename
CostEstimateSht.Name = "Cost Estimate"
CostEstimateSht.Columns.AutoFit
CostEstimateSht.Visible = xlSheetVisible
'Set header info
With CostEstimateSht
    .Range("A1").Value = "Job Information - " & EstSht.Range("CustomerName").Value & b.bWidth & "' x " & b.bLength & "' x " & b.bHeight & "'"
    .Range("B3").Value = EstSht.Range("CustomerName").Value
    .Range("B4").Value = "" & b.bWidth & "' x " & b.bLength & "' x " & b.bHeight & "'"
    .Range("F3").Value = Now()
End With

'debug
Debug.Print "------------------------------------- Cost Estimate Prep ------------------------------"
Debug.Print "---------------------- Panel Collection ---------------"
'sort collections
For Each clsPanel In b.PanelCollection
    Sheetmetal.Add clsPanel
    Debug.Print clsPanel.PanelType & " added to SheetMetal"
Next clsPanel
Debug.Print "---------------------- Trim Collection ---------------"
For Each clsTrim In b.TrimCollection

    'check for OH doors
    For n = LBound(OHDoorNames) To UBound(OHDoorNames)
        If InStr(1, clsTrim.tType, OHDoorNames(n)) <> 0 Then
            OHDoors.Add clsTrim
            Debug.Print clsTrim.tType & " added to OHDoors"
            GoTo NextTrimClass
        End If
    Next n
    'Check for Insulation
    For n = LBound(InsulationNames) To UBound(InsulationNames)
        If InStr(1, clsTrim.tType, InsulationNames(n)) <> 0 Then
            Insulation.Add clsTrim
            Debug.Print clsTrim.tType & " added to Insulation"
            GoTo NextTrimClass
        End If
    Next n
    'Check for Windows
    If clsTrim.tType = "Standard Window" Then
        Windows.Add clsTrim
        Debug.Print clsTrim.tType & " added to Windows"
        GoTo NextTrimClass
    'check for electric openers
    ElseIf InStr(1, clsTrim.tType, "Electric Opener") <> 0 Then
        ElectricOpeners.Add clsTrim
        Debug.Print clsTrim.tType & " added to Electric Openers"
        GoTo NextTrimClass
    'Check for Ridge Vents
    ElseIf InStr(1, clsTrim.tType, "Ridge Vent") <> 0 Then
        RidgeVents.Add clsTrim
        Debug.Print clsTrim.tType & " added to Ridge Vents"
        GoTo NextTrimClass
    'check for door canopies
    ElseIf InStr(1, clsTrim.tType, "Door Canopy") <> 0 Then
        DoorCanopies.Add clsTrim
        Debug.Print clsTrim.tType & " added to Door Canopies"
        GoTo NextTrimClass
    'check for exhaust fans, louvers, or weather hoods
    ElseIf InStr(1, clsTrim.tType, "Exhaust Fan") <> 0 Or InStr(1, clsTrim.tType, "Louver") <> 0 Or InStr(1, clsTrim.tType, "Weather Hood") <> 0 Then
        ExhaustFansLouversWeatherhoods.Add clsTrim
        Debug.Print clsTrim.tType & " added to ExhaustFans,Louvers,Weatherhoods"
        GoTo NextTrimClass
    End If
'otherwise, add to sheet metal collection
Sheetmetal.Add clsTrim
Debug.Print clsTrim.tType & " added to Sheet Metal"

NextTrimClass:
Next clsTrim

Debug.Print "---------------------- MiscItem Collection ---------------"
For Each clsItem In b.MiscMaterialsCollection
    'Check for electric Opener
    If InStr(1, clsItem.Name, "Electric Opener") <> 0 Then
        With CostEstimateSht
            .Range("Insulation_TotalCost").EntireRow.Insert shift:=xlDown
            .Range("Insulation_TotalCost").offset(-1, -1).Value = clsItem.Name
            'cost, markup %
            .Range("Insulation_TotalCost").offset(-1, 0).Value = "<Enter Cost>"
            .Range("Insulation_TotalCost").offset(-1, 2).Value = .Range("OHDoors_TotalCost").offset(0, 2).Value
            'add formulas
            .Range("Insulation_TotalCost").offset(-1, 1).Resize(2, 1).FillUp
            .Range("Insulation_TotalCost").offset(-1, 3).Resize(2, 1).FillUp
            .Range("Insulation_TotalCost").offset(-1, 4).Resize(2, 1).FillUp
        End With
    End If
    'check for OH doors
    For n = LBound(OHDoorNames) To UBound(OHDoorNames)
        If InStr(1, clsItem.Name, OHDoorNames(n)) <> 0 Then
            OHDoors.Add clsItem
            Debug.Print clsItem.Name & " added to OHDoors"
            GoTo NextItemClass
        End If
    Next n
    'Check for Insulation
    For n = LBound(InsulationNames) To UBound(InsulationNames)
        If InStr(1, clsItem.Name, InsulationNames(n)) <> 0 Then
            Insulation.Add clsItem
            Debug.Print clsItem.Name & " added to Insulation"
            GoTo NextItemClass
        End If
    Next n
    'Check for Windows
    If clsItem.Name = "Standard Window" Then
        Windows.Add clsItem
        Debug.Print clsItem.Name & " added to Windows"
        GoTo NextItemClass
    'check for electric openers
    ElseIf InStr(1, clsItem.Name, "Electric Opener") <> 0 Then
        ElectricOpeners.Add clsItem
        Debug.Print clsItem.Name & " added to Electric Openers"
        GoTo NextItemClass
    'Check for Ridge Vents
    ElseIf InStr(1, clsItem.Name, "Ridge Vent") <> 0 Then
        RidgeVents.Add clsItem
        Debug.Print clsItem.Name & " added to Ridge Vents"
        GoTo NextItemClass
    'check for door canopies
    ElseIf InStr(1, clsItem.Name, "Door Canopy") <> 0 Then
        DoorCanopies.Add clsItem
        Debug.Print clsItem.Name & " added to Door Canopies"
        GoTo NextItemClass
    'check for exhaust fans, louvers, or weather hoods
    ElseIf InStr(1, clsItem.Name, "Exhaust Fan") <> 0 Or InStr(1, clsItem.Name, "Louver") <> 0 Or InStr(1, clsItem.Name, "Weather Hood") <> 0 Then
        ExhaustFansLouversWeatherhoods.Add clsItem
        Debug.Print clsItem.Name & " added to ExhaustFans,Louvers,Weatherhoods"
        GoTo NextItemClass
    End If
'otherwise, add to sheet metal collection
Sheetmetal.Add clsItem
Debug.Print clsItem.Name & " added to Sheet Metal"
NextItemClass:
Next clsItem


''''''''''''''''''''''''''' Total Collections and Output
With CostEstimateSht
    'Structural Steel
    .Range("StructuralSteel_TotalCost").Value = b.SSTotalCost
    'Sheet Metal
    For Each CollectionItem In Sheetmetal
        BTItems.AddItem CollectionItem, "Sheetmetal"
        If IsNumeric(CollectionItem.totalCost) = True Then CollectionTotal = CollectionTotal + CollectionItem.totalCost
    Next CollectionItem
    .Range("SheetMetal_TotalCost").Value = CollectionTotal
    CollectionTotal = 0
    'OH Doors
    For Each CollectionItem In OHDoors
        BTItems.AddItem CollectionItem, "OHDoors"
        If IsNumeric(CollectionItem.totalCost) = True Then CollectionTotal = CollectionTotal + CollectionItem.totalCost
    Next CollectionItem
    .Range("OHDoors_TotalCost").Value = CollectionTotal
    CollectionTotal = 0
    'Insulation
    For Each CollectionItem In Insulation
        BTItems.AddItem CollectionItem, "Insulation"
        If IsNumeric(CollectionItem.totalCost) = True Then CollectionTotal = CollectionTotal + CollectionItem.totalCost
    Next CollectionItem
    .Range("Insulation_TotalCost").Value = CollectionTotal
    CollectionTotal = 0
    'Windows
    For Each CollectionItem In Windows
        BTItems.AddItem CollectionItem, "Windows"
        If IsNumeric(CollectionItem.totalCost) = True Then CollectionTotal = CollectionTotal + CollectionItem.totalCost
    Next CollectionItem
    .Range("Windows_TotalCost").Value = CollectionTotal
    CollectionTotal = 0
    'Ridge Vents
    For Each CollectionItem In RidgeVents
        BTItems.AddItem CollectionItem, "GenMiscMaterials"
        If IsNumeric(CollectionItem.totalCost) = True Then CollectionTotal = CollectionTotal + CollectionItem.totalCost
    Next CollectionItem
    .Range("RidgeVents_TotalCost").Value = CollectionTotal
    CollectionTotal = 0
    'Door Canopies
    For Each CollectionItem In DoorCanopies
        BTItems.AddItem CollectionItem, "GenMiscMaterials"
        If IsNumeric(CollectionItem.totalCost) = True Then CollectionTotal = CollectionTotal + CollectionItem.totalCost
    Next CollectionItem
    .Range("DoorCanopies_TotalCost").Value = CollectionTotal
    CollectionTotal = 0
    'Exhaust Fans, Louvers, Weatherhoods
    For Each CollectionItem In ExhaustFansLouversWeatherhoods
        BTItems.AddItem CollectionItem, "ExhaustFansLouversWeatherhoods"
        If IsNumeric(CollectionItem.totalCost) = True Then CollectionTotal = CollectionTotal + CollectionItem.totalCost
    Next CollectionItem
    .Range("ExhaustFansLouversWeatherhoods_TotalCost").Value = CollectionTotal
    CollectionTotal = 0
    
    'delete empty line items (other than structural steel, misc. items and electric openers)
    If .Range("OHDoors_TotalCost").Value = 0 Then .Range("OHDoors_TotalCost").EntireRow.Delete shift:=xlUp
    If .Range("Insulation_TotalCost").Value = 0 Then .Range("Insulation_TotalCost").EntireRow.Delete shift:=xlUp
    If .Range("Windows_TotalCost").Value = 0 Then .Range("Windows_TotalCost").EntireRow.Delete shift:=xlUp
    If .Range("RidgeVents_TotalCost").Value = 0 Then .Range("RidgeVents_TotalCost").EntireRow.Delete shift:=xlUp
    If .Range("DoorCanopies_TotalCost").Value = 0 Then .Range("DoorCanopies_TotalCost").EntireRow.Delete shift:=xlUp
    If .Range("ExhaustFansLouversWeatherhoods_TotalCost").Value = 0 Then .Range("ExhaustFansLouversWeatherhoods_TotalCost").EntireRow.Delete shift:=xlUp
End With


'Populate Labor Section
Call LaborGen(CostEstimateSht, b)

End Sub

Private Sub LaborGen(costEstSht As Worksheet, b As clsBuilding)
Dim FOCell As Range
Dim ItemCount As Integer
Dim ItemLF As Double
Dim ItemSF As Double
Dim clsTrim As clsTrim
Dim clsItem As clsMiscItem
Dim Row As Integer

With costEstSht
    '''Erection
    'Building Width * Building Length
    .Range("Erection").Value = b.bLength * b.bWidth
    '''Height Premium
    'Wall Square Footage over 17'
    If b.bHeight > 17 Then
        'calculate SF for overage
        ItemSF = (b.bHeight - 17) * (b.bLength * 2 + b.bWidth * 2)
    End If
    .Range("HeightPremium").Value = ItemSF
    ItemSF = 0
    '''Pitch Premium
    'Roof Area and Endwall Area due to Pitch
    If b.rShape = "Single Slope" Then
        'additional roof area
        ItemSF = ((b.RafterLength / 12) * b.bLength) - (b.bLength * b.bWidth)
        'additional endwall area
        ItemSF = ItemSF + (b.bWidth * ((b.HighSideEaveHeight / 12) - b.bHeight))
    ElseIf b.rShape = "Gable" Then
        'additional roof area
        ItemSF = ((b.RafterLength / 12) * 2 * b.bLength) - (b.bLength * b.bWidth)
        'additional endwall area
        ItemSF = ItemSF + ((b.bWidth / 2) * (((b.bWidth / 2) * b.rPitch) / 12))
    End If
    .Range("PitchPremium").Value = ItemSF
    ItemSF = 0
    '''PDoors
    'PDoor Count
    For Each FOCell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
        If FOCell.offset(0, 1).Value <> "" Then ItemCount = ItemCount + 1
    Next FOCell
    .Range("PDoors").Value = ItemCount
    ItemCount = 0
    '''Door Canopies
    'Canopy Count
    For Each clsItem In b.MiscMaterialsCollection
        If InStr(1, clsItem.Name, "Door Canopy") <> 0 Then ItemCount = ItemCount + 1
    Next clsItem
    .Range("DoorCanopies").Value = ItemCount
    ItemCount = 0
    '''OH Doors
    'OHdoor Count
    For Each FOCell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
        If FOCell.offset(0, 1).Value <> "" Then ItemCount = ItemCount + 1
    Next FOCell
    .Range("OHDoors").Value = ItemCount
    ItemCount = 0
    '''Windows
    'window count
    For Each FOCell In Range(EstSht.Range("WindowCell1"), EstSht.Range("WindowCell12"))
        If FOCell.offset(0, 1).Value <> "" Then ItemCount = ItemCount + 1
    Next FOCell
    .Range("Windows").Value = ItemCount
    ItemCount = 0
    '''Misc FOs
    'misc FO count
    For Each FOCell In Range(EstSht.Range("MiscFOCell1"), EstSht.Range("MiscFOCell12"))
        If FOCell.offset(0, 1).Value <> "" Then ItemCount = ItemCount + 1
    Next FOCell
    .Range("MiscFOs").Value = ItemCount
    ItemCount = 0
    '''VRR Wall & Roof Insulation
    'Wall/Roof Insulation Square Footage
    For Each clsItem In b.MiscMaterialsCollection
        Select Case True
        Case clsItem.Name = "3"" VRR Wall Insulation"
            .Range("VRRWallInsulation3Inch").Value = clsItem.Quantity
            .Range("VRRWallInsulation4Inch").EntireRow.Delete
        Case clsItem.Name = "4"" VRR Wall Insulation"
            .Range("VRRWallInsulation4Inch").Value = clsItem.Quantity
            .Range("VRRWallInsulation3Inch").EntireRow.Delete
        Case clsItem.Name = "3"" VRR Roof Insulation"
            .Range("VRRRoofInsulation3Inch").Value = clsItem.Quantity
            .Range("VRRRoofInsulation4Inch").EntireRow.Delete
            .Range("VRRRoofInsulation6Inch").EntireRow.Delete
        Case clsItem.Name = "4"" VRR Roof Insulation"
            .Range("VRRRoofInsulation4Inch").Value = clsItem.Quantity
            .Range("VRRRoofInsulation3Inch").EntireRow.Delete
            .Range("VRRRoofInsulation6Inch").EntireRow.Delete
        Case clsItem.Name = "6"" VRR Roof Insulation"
            .Range("VRRRoofInsulation6Inch").Value = clsItem.Quantity
            .Range("VRRRoofInsulation3Inch").EntireRow.Delete
            .Range("VRRRoofInsulation4Inch").EntireRow.Delete
        End Select
    Next clsItem
    '''Ridge Vents
    'Ridge Vent Count
    For Each clsItem In b.MiscMaterialsCollection
        If InStr(1, clsItem.Name, "Ridge Vent") <> 0 Then .Range("RidgeVents").Value = clsItem.Quantity
    Next clsItem
    '''Gutters
    'LF of gutter hang-ons
    For Each clsTrim In b.TrimCollection
        If InStr(1, clsTrim.tType, "Sculptured Gutter Hang-On") <> 0 Then ItemCount = ItemCount + (clsTrim.tLength * clsTrim.Quantity)
    Next clsTrim
    .Range("Gutters").Value = (ItemCount / 12) 'convert to ft
    ItemCount = 0
    ''''' SF for the 8 below items
    With b
    '''Gable Overhangs
        If .rShape = "Single Slope" Then
            ItemSF = ((.s2RafterSheetLength + .s2ExtensionRafterLength + .s4ExtensionRafterLength) * (.e1Overhang + .e3Overhang)) / 144
        ElseIf .rShape = "Gable" Then
            ItemSF = ((.s2RafterSheetLength + .s2ExtensionRafterLength) * (.e1Overhang + .e3Overhang)) / 144     's2
            ItemSF = ItemSF + (((.s4RafterSheetLength + .s4ExtensionRafterLength) * (.e1Overhang + .e3Overhang)) / 144)     's4
        End If
        costEstSht.Range("GableOverhangs").Value = ItemSF
        ItemSF = 0
    '''Eave Overhangs
        ItemSF = ((.s2Overhang - 4.25) * .RoofLength) / 144     's2
        If .s4Overhang <> 4.25 And .s4Overhang <> 0 Then ItemSF = ItemSF + (((.s4Overhang - 4.25) * .RoofLength) / 144)    's4
        costEstSht.Range("EaveOverhangs").Value = ItemSF
        ItemSF = 0
    '''Gable Extensions
        If .rShape = "Single Slope" Then
            ItemSF = (.s2RafterSheetLength * (.e1Extension + .e3Extension)) / 144
        ElseIf .rShape = "Gable" Then
            ItemSF = (.s2RafterSheetLength * (.e1Extension + .e3Extension)) / 144   's2
            ItemSF = ItemSF + ((.s4RafterSheetLength * (.e1Extension + .e3Extension)) / 144)   's4
        End If
        costEstSht.Range("GableExtensions").Value = ItemSF
        ItemSF = 0
    '''Eave Extensions
        ItemSF = (.s2EaveExtensionBuildingLength * .s2ExtensionRafterLength) / 144     's2
        ItemSF = ItemSF + ((.s4EaveExtensionBuildingLength * .s4ExtensionRafterLength) / 144)     's4
        costEstSht.Range("EaveExtensions").Value = ItemSF
        ItemSF = 0
    '''Gable Overhang Soffit
        If .rShape = "Single Slope" Then
            If .e1GableOverhangSoffit = True Then ItemSF = (.s2RafterSheetLength * .e1Overhang) / 144
            If .e3GableOverhangSoffit = True Then ItemSF = ItemSF + ((.s2RafterSheetLength * .e3Overhang) / 144)
        ElseIf .rShape = "Gable" Then
            If .e1GableOverhangSoffit = True Then
                ItemSF = (.s2RafterSheetLength * .e1Overhang) / 144
                ItemSF = ItemSF + ((.s4RafterSheetLength * .e1Overhang) / 144)
            End If
            If .e3GableOverhangSoffit = True Then
                ItemSF = ItemSF = ItemSF + ((.s2RafterSheetLength * .e3Overhang) / 144)
                ItemSF = ItemSF + ((.s4RafterSheetLength * .e3Overhang) / 144)
            End If
        End If
        costEstSht.Range("GableOverhangSoffit").Value = ItemSF
        ItemSF = 0
    '''Eave Overhang Soffit
        If .s2EaveOverhangSoffit = True Then ItemSF = (.RoofLength * (.s2Overhang - 4.25)) / 144
        If .s4EaveOverhangSoffit = True Then ItemSF = ItemSF + (((.s4Overhang - 4.25) * .RoofLength) / 144)
        costEstSht.Range("EaveOverhangSoffit").Value = ItemSF
        ItemSF = 0
    '''Gable Extension Soffit
        If .rShape = "Single Slope" Then
            If .e1GableExtensionSoffit = True Then ItemSF = (.s2RafterSheetLength * .e1Extension) / 144
            If .e3GableExtensionSoffit = True Then ItemSF = ItemSF + ((.s2RafterSheetLength * .e3Extension) / 144)
        ElseIf .rShape = "Gable" Then
            If .e1GableExtensionSoffit = True Then
                ItemSF = (.s2RafterSheetLength * .e1Extension) / 144
                ItemSF = ItemSF + ((.s4RafterSheetLength * .e1Extension) / 144)
            End If
            If .e3GableExtensionSoffit = True Then
                ItemSF = ItemSF = ItemSF + ((.s2RafterSheetLength * .e3Extension) / 144)
                ItemSF = ItemSF + ((.s4RafterSheetLength * .e3Extension) / 144)
            End If
        End If
        costEstSht.Range("GableExtensionSoffit").Value = ItemSF
        ItemSF = 0
    '''Eave Extension Soffit
        If .s2EaveOverhangSoffit = True Then ItemSF = (.s2EaveExtensionBuildingLength * .s2ExtensionRafterLength) / 144     's2
        If .s4EaveOverhangSoffit = True Then ItemSF = ItemSF + ((.s4EaveExtensionBuildingLength * .s4ExtensionRafterLength) / 144)     's4
        costEstSht.Range("EaveExtensionSoffit").Value = ItemSF
        ItemSF = 0
    '''Wainscot
        If .Wainscot("e1") <> "None" Then _
        ItemLF = CDbl(Left(.Wainscot("e1"), 2)) * Application.WorksheetFunction.RoundUp(.bWidth / 3, 0)
        If .Wainscot("e3") <> "None" Then _
        ItemLF = ItemLF + (CDbl(Left(.Wainscot("e3"), 2)) * Application.WorksheetFunction.RoundUp(.bWidth / 3, 0))
        If .Wainscot("s2") <> "None" Then _
        ItemLF = ItemLF + (CDbl(Left(.Wainscot("s2"), 2)) * Application.WorksheetFunction.RoundUp(.bWidth / 3, 0))
        If .Wainscot("s4") <> "None" Then _
        ItemLF = ItemLF + (CDbl(Left(.Wainscot("s4"), 2)) * Application.WorksheetFunction.RoundUp(.bWidth / 3, 0))
        ItemLF = ItemLF / 12        ' convert to FT
        costEstSht.Range("Wainscot").Value = ItemLF
        ItemLF = 0
    '''Liner Panels
        'endwall 1
        If .LinerPanels("e1") = "8'" Then
            ItemSF = ItemSF + (8 * .bWidth)
        ElseIf .LinerPanels("e1") = "Full Height" Then
            If .rShape = "Gable" Then
                ItemSF = ItemSF + (.bWidth * .bHeight) + ((.bWidth / 2) * (((.bWidth / 2) * .rPitch) / 12))
            ElseIf .rShape = "Single Slope" Then
                ItemSF = ItemSF + (.bWidth * .bHeight) + ((.bWidth / 2) * ((.bWidth * .rPitch) / 12))
            End If
            ItemSF = ItemSF - ((8 / 12) * .bWidth)  ' subtract the 8" difference
        End If
        'sidewall 2
        If .LinerPanels("s2") = "8'" Then
            ItemSF = ItemSF + (8 * .bLength)
        ElseIf .LinerPanels("s2") = "Full Height" Then
            ItemSF = ItemSF + (.bLength * (.bHeight - (8 / 12)))
        End If
        'endwall 3
        If .LinerPanels("e3") = "8'" Then
            ItemSF = ItemSF + (8 * .bWidth)
        ElseIf .LinerPanels("e3") = "Full Height" Then
            If .rShape = "Gable" Then
                ItemSF = ItemSF + (.bWidth * .bHeight) + ((.bWidth / 2) * (((.bWidth / 2) * .rPitch) / 12))
            ElseIf .rShape = "Single Slope" Then
                ItemSF = ItemSF + (.bWidth * .bHeight) + ((.bWidth / 2) * ((.bWidth * .rPitch) / 12))
            End If
            ItemSF = ItemSF - ((8 / 12) * .bWidth)  ' subtract the 8" difference
        End If
        'sidewall 4
        If .LinerPanels("s4") = "8'" Then
            ItemSF = ItemSF + (8 * .bLength)
        ElseIf .LinerPanels("s4") = "Full Height" Then
            If .rShape = "Gable" Then
                ItemSF = ItemSF + (.bLength * (.bHeight - (8 / 12)))
            ElseIf .rShape = "Single Slope" Then
                ItemSF = ItemSF + (.bLength * ((.HighSideEaveHeight - 8) / 12))
            End If
        End If
        'Roof
        If .LinerPanels("Roof") = "Full Height" Then
            If .rShape = "Single Slope" Then
                ItemSF = ItemSF + (((.RafterLength - 8) / 12) * .bLength)
            ElseIf .rShape = "Gable" Then
                ItemSF = ItemSF + ((((.RafterLength - 8) / 12) * .bLength) * 2)
            End If
        End If
        costEstSht.Range("LinerPanels").Value = ItemSF
        
    End With
    ''' Delete Blank Line Items
    For Row = .Range("LinerPanels").Row To .Range("Erection").Row Step -1
        If .Cells(Row, 2).Value = 0 Then .Cells(Row, 2).EntireRow.Delete
    Next Row

End With



End Sub


Sub DescriptionGen(b As clsBuilding)
Dim dStr As String
Dim BayStr As String
Dim cell As Range
Dim DescriptionSht As Worksheet
Dim n As Integer
Dim i As Integer
Dim PDoorTypes(11, 1)
Dim OHDoorTypes(11, 1)
Dim WindowTypes(23, 1)
Dim FOTypes(11, 1)
Dim TempDesc As String
Dim dCell As Range
Dim RowCount As Double

'delete old output sheets
Application.DisplayAlerts = False
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Project Description" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n

'set new output sheet
DescriptionShtTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set DescriptionSht = ThisWorkbook.Sheets("DescriptionShtTmp (2)")
'rename
DescriptionSht.Name = "Project Description"
DescriptionSht.Visible = xlSheetVisible
Set dCell = DescriptionSht.Range("DescriptionCell")
RowCount = 0

With b

    'main building description
    dStr = EstSht.Range("BusinessName").Value & " agrees to provide the material, labor, and equipment to erect the following metal building: "
    dCell.Value = dStr
    RowCount = RowCount + 1
    dStr = "Width: " & b.bWidth
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    dStr = "Length: " & b.bLength
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    dStr = "Eave height: " & b.bHeight
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    dStr = "Pitch: " & b.rPitch
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    dStr = "Roof Shape: " & b.rShape
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    
    dStr = EstSht.Range("BayNum").Value & " bays: "
    For i = 1 To EstSht.Range("BayNum").Value
        If i <> 1 Then dStr = dStr & ", "
        dStr = dStr & "bay #" & i & ": " & EstSht.Range("Bay1_Length").offset(i - 1, 0).Value & "'"
    Next i
    
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    
    'Walls and Wall Status
    For i = 1 To 4
        If i = 1 Or i = 3 Then
            dStr = "Endwall " & i & ": "
        Else
            dStr = "Sidewall " & i & ": "
        End If
        'status
        If EstSht.Range("e1_WallStatus").offset(i - 1, 0).Value <> "Include" Then
            dStr = dStr & EstSht.Range("e1_WallStatus").offset(i - 1, 0).Value
        Else
            dStr = dStr & "included"
        End If
        'expandable
        If EstSht.Range("e1_WallStatus").offset(i - 1, 1).Value = "Yes" Then
            dStr = dStr & ", expandable"
        End If
        'Ft above finished floor IF PARTIAL
        If EstSht.Range("e1_WallStatus").offset(i - 1, 0).Value = "Partial" Then
            dStr = dStr & ", " & EstSht.Range("e1_WallStatus").offset(i - 1, 2).Value & "ft above finished floor"
        End If
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    Next i
        
    'Wall and Roof Panels
    dStr = "Wall panels: " & LCase(.wPanelColor) & " " & LCase(.wPanelShape)
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    dStr = "Roof panels: " & LCase(.rPanelColor) & " " & LCase(.rPanelShape)
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    
    'liner panels
    If .LinerPanels("e1") <> "None" Then
        dStr = .LinerPanels("e1") & " " & LCase(EstSht.Range("e1_LinerPanels").offset(0, 3).Value) & " " & LCase(EstSht.Range("e1_LinerPanels").offset(0, 2).Value) & " endwall #1 liner panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If .LinerPanels("s2") <> "None" Then
        dStr = .LinerPanels("s2") & " " & LCase(EstSht.Range("s2_LinerPanels").offset(0, 3).Value) & " " & LCase(EstSht.Range("s2_LinerPanels").offset(0, 2).Value) & " sidewall #2 liner panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If .LinerPanels("e3") <> "None" Then
        dStr = .LinerPanels("e3") & " " & LCase(EstSht.Range("e3_LinerPanels").offset(0, 3).Value) & " " & LCase(EstSht.Range("e3_LinerPanels").offset(0, 2).Value) & " endwall #3 liner panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If .LinerPanels("s4") <> "None" Then
        dStr = .LinerPanels("s4") & " " & LCase(EstSht.Range("s4_LinerPanels").offset(0, 3).Value) & " " & LCase(EstSht.Range("s4_LinerPanels").offset(0, 2).Value) & " sidewall #4 liner panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If .LinerPanels("Roof") <> "None" Then
        dStr = .LinerPanels("Roof") & " " & LCase(EstSht.Range("Roof_LinerPanels").offset(0, 3).Value) & " " & LCase(EstSht.Range("Roof_LinerPanels").offset(0, 2).Value) & " roof liner panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    ''''trim
    'FO trim
    dStr = "Framed opening trim: " & LCase(EstSht.Range("FO_tColor").Value)
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    'base trim
    If .BaseTrim = True Then
        dStr = "Base opening trim: " & LCase(EstSht.Range("Base_tColor").Value)
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    'rake, eave, corner trim
    dStr = "Rake trim: " & LCase(EstSht.Range("Rake_tColor").Value)
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    dStr = "Eave trim: " & LCase(EstSht.Range("Eave_tColor").Value)
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    dStr = "Corner trim: " & LCase(EstSht.Range("OutsideCorner_tColor").Value)
    dCell.offset(RowCount, 0).Value = dStr
    RowCount = RowCount + 1
    'downspouts/gutters
    If EstSht.Range("GutterAndDownspouts").Value = "Yes" Then
        dStr = "Downspouts: " & LCase(EstSht.Range("DownspoutColor").Value)
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
        dStr = "Gutters: " & LCase(EstSht.Range("GutterColor").Value)
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    'wainscot
    If .Wainscot("e1") <> "None" Then
        dStr = .Wainscot("e1") & " endwall #1 wainscot, " & LCase(EstSht.Range("e1_Wainscot").offset(0, 1).Value) & " " & LCase(EstSht.Range("e1_Wainscot").offset(0, 2).Value) & " panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If .Wainscot("s2") <> "None" Then
        dStr = .Wainscot("s2") & " sidewall #2 wainscot, " & LCase(EstSht.Range("s2_Wainscot").offset(0, 1).Value) & " " & LCase(EstSht.Range("s2_Wainscot").offset(0, 2).Value) & " panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If .Wainscot("e3") <> "None" Then
        dStr = .Wainscot("e3") & " endwall #3 wainscot, " & LCase(EstSht.Range("e3_Wainscot").offset(0, 1).Value) & " " & LCase(EstSht.Range("e3_Wainscot").offset(0, 2).Value) & " panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If .Wainscot("s4") <> "None" Then
        dStr = .Wainscot("s4") & " sidewall #4 wainscot, " & LCase(EstSht.Range("s4_Wainscot").offset(0, 1).Value) & " " & LCase(EstSht.Range("s4_Wainscot").offset(0, 2).Value) & " panels"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If EstSht.Range("Wainscot_tColor").Value <> "" Then
        dStr = "Wainscot trim: " & LCase(EstSht.Range("Wainscot_tColor").Value)
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    dStr = ""
    ''PO doors
    For Each cell In Range(EstSht.Range("pDoorCell1"), EstSht.Range("pDoorCell12"))
        If cell.offset(0, 1).Value <> "" Then
            TempDesc = TempDesc & cell.offset(0, 1).Value & " size personnel door"
            'half glass, canopy, deadbolt
            If cell.offset(0, 3).Value = "Yes" Then _
            TempDesc = TempDesc & " with half glass"
            If cell.offset(0, 4).Value <> "No" Then _
            TempDesc = TempDesc & " with " & cell.offset(0, 4).Value & " canopy"
            If cell.offset(0, 6).Value = "Yes" Then _
            TempDesc = TempDesc & " with dead bolt"
            'end of line
            TempDesc = TempDesc
            'loop through array for similar PO Doors
            For i = 0 To 11
                If Not IsEmpty(PDoorTypes(i, 0)) Then
                    If TempDesc = PDoorTypes(i, 1) Then
                        PDoorTypes(i, 0) = PDoorTypes(i, 0) + 1
                        TempDesc = ""
                    End If
                Else
                    If TempDesc <> "" Then
                        PDoorTypes(i, 0) = 1
                        PDoorTypes(i, 1) = TempDesc
                        TempDesc = ""
                    End If
                End If
            Next i
        End If
    Next cell
    'add array values to dStr
    For i = 0 To 11
        If Not IsEmpty(PDoorTypes(i, 0)) Then
            dStr = dStr & "(" & PDoorTypes(i, 0) & ") " & PDoorTypes(i, 1) & ", "
        End If
    Next i
    If dStr <> "" Then
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    TempDesc = ""
    dStr = ""
    ''OH doors
    For Each cell In Range(EstSht.Range("OHDoorCell1"), EstSht.Range("OHDoorCell12"))
        If cell.offset(0, 1).Value <> "" Then
            TempDesc = TempDesc & cell.offset(0, 1).Value & "' x " & cell.offset(0, 2).Value & "' " & LCase(cell.offset(0, 4).Value) & " overhead door"
            'insulation, operation, high lift, windows
            If cell.offset(0, 5).Value <> "None" Then
                TempDesc = TempDesc & " with " & LCase(cell.offset(0, 5).Value) & " insulation"
            Else
                TempDesc = TempDesc & " non-insulated"
            End If
            Select Case cell.offset(0, 6).Value
            Case "Manual"
                TempDesc = TempDesc & " with manual operation"
            Case "Chain Hoisr"
                TempDesc = TempDesc & " with chain hoist"
            Case "Electric Opener"
                TempDesc = TempDesc & " with electric opener"
            End Select
            
            If cell.offset(0, 7).Value <> "No" Then _
            TempDesc = TempDesc & " with high lift"
            If cell.offset(0, 8).Value <> "None" Then _
            TempDesc = TempDesc & " with " & LCase(cell.offset(0, 8).Value) & " windows"
            'end of line
            TempDesc = TempDesc
        End If
        For i = 0 To 11
            If Not IsEmpty(OHDoorTypes(i, 0)) Then
                If TempDesc = OHDoorTypes(i, 1) Then
                    OHDoorTypes(i, 0) = OHDoorTypes(i, 0) + 1
                    TempDesc = ""
                End If
            Else
                If TempDesc <> "" Then
                    OHDoorTypes(i, 0) = 1
                    OHDoorTypes(i, 1) = TempDesc
                    TempDesc = ""
                End If
            End If
        Next i
    Next cell
    'add array values to dStr
    For i = 0 To 11
        If Not IsEmpty(OHDoorTypes(i, 0)) Then
            dStr = dStr & "(" & OHDoorTypes(i, 0) & ") " & OHDoorTypes(i, 1)
        End If
    Next i
    TempDesc = ""
    If dStr <> "" Then
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    dStr = ""
    ''windows
    For Each cell In Range(EstSht.Range("WindowCell1"), EstSht.Range("WindowCell12"))
        If cell.offset(0, 1).Value <> "" Then
            TempDesc = TempDesc & cell.offset(0, 1).Value & """ x " & cell.offset(0, 2).Value & """ window, "
            'end of line
            TempDesc = TempDesc
        End If
        For i = 0 To 23
            If Not IsEmpty(WindowTypes(i, 0)) Then
                If TempDesc = WindowTypes(i, 1) Then
                    WindowTypes(i, 0) = WindowTypes(i, 0) + 1
                    TempDesc = ""
                End If
            Else
                If TempDesc <> "" Then
                    WindowTypes(i, 0) = 1
                    WindowTypes(i, 1) = TempDesc
                    TempDesc = ""
                End If
            End If
        Next i
    Next cell
    'add array values to dStr
    For i = 0 To 23
        If Not IsEmpty(WindowTypes(i, 0)) Then
            dStr = dStr & "(" & WindowTypes(i, 0) & ") " & WindowTypes(i, 1)
        End If
    Next i
    TempDesc = ""
    If dStr <> "" Then
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    dStr = ""
    ''Misc FOs
    For Each cell In Range(EstSht.Range("MiscFOCell1"), EstSht.Range("MiscFOCell12"))
        If cell.offset(0, 1).Value <> "" Then
            TempDesc = TempDesc & cell.offset(0, 1).Value & "' x " & cell.offset(0, 2).Value & """ Misc FO, "
            'exhaust fans/louvers, weather hoods
            If cell.offset(0, 4).Value <> "None" Then _
            TempDesc = TempDesc & " with " & LCase(cell.offset(0, 4).Value)
            If cell.offset(0, 5).Value <> "None" Then _
            TempDesc = TempDesc & " with " & LCase(cell.offset(0, 5).Value) & " weather hood"
            'end of line
            TempDesc = TempDesc
        End If
        For i = 0 To 11
            If Not IsEmpty(FOTypes(i, 0)) Then
                If TempDesc = FOTypes(i, 1) Then
                    FOTypes(i, 0) = FOTypes(i, 0) + 1
                    TempDesc = ""
                End If
            Else
                If TempDesc <> "" Then
                    FOTypes(i, 0) = 1
                    FOTypes(i, 1) = TempDesc
                    TempDesc = ""
                End If
            End If
        Next i
    Next cell
    'add array values to dStr
    For i = 0 To 11
        If Not IsEmpty(FOTypes(i, 0)) Then
            dStr = dStr & "(" & FOTypes(i, 0) & ") " & FOTypes(i, 1)
        End If
    Next i
    TempDesc = ""
    If dStr <> "" Then
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    dStr = ""
    'Wall, Roof Insulation
    If EstSht.Range("WallInsulation").Value <> "None" And EstSht.Range("WallInsulation").Value <> "" Then
        dStr = "Wall insulation: " & LCase(EstSht.Range("WallInsulation").Value)
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If EstSht.Range("RoofInsulation").Value <> "None" And EstSht.Range("RoofInsulation").Value <> "" Then
        dStr = "Roof insulation: " & LCase(EstSht.Range("RoofInsulation").Value)
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    'ridge vents, wall panels, skilights
    If EstSht.Range("RidgeVentQty").Value <> 0 And EstSht.Range("RidgeVentQty").Value <> "" Then
        dStr = "(" & EstSht.Range("RidgeVentQty").Value & ") " & LCase(EstSht.Range("RidgeVentType").Value) & " ridge vent(s)"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If EstSht.Range("TranslucentWallPanelQty").Value <> 0 And EstSht.Range("TranslucentWallPanelQty").Value <> "" Then
        dStr = "(" & EstSht.Range("TranslucentWallPanelQty").Value & ") " & EstSht.Range("TranslucentWallPanelLength").Value & "' translucent wall panel(s)" & vbNewLine
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    If EstSht.Range("SkylightQty").Value <> 0 And EstSht.Range("SkylightQty").Value <> "" Then
        dStr = "(" & EstSht.Range("SkylightQty").Value & ") " & EstSht.Range("SkylightLength").Value & "' skylight(s)"
        dCell.offset(RowCount, 0).Value = dStr
        RowCount = RowCount + 1
    End If
    dStr = ""
    'overhangs, soffits
    With EstSht
        'e1 overhang
        If .Range("e1_GableOverhang").Value <> "" Then
            dStr = dStr & .Range("Building_Width").Value & "' wide x " & .Range("e1_GableOverhang").Value & "' long gable overhang on endwall #1, "
            If .Range("e1_GableOverhangSoffit").Value = "Yes" Then
                dStr = dStr & LCase(.Range("e1_GableOverhangSoffit").offset(0, 3).Value) & " " & _
                LCase(.Range("e1_GableOverhangSoffit").offset(0, 2).Value) & " endwall #1 soffit panels with " & _
                LCase(.Range("e1_GableOverhangSoffit").offset(0, 4).Value) & " trim"
            End If
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = ""
        's2 overhang
        If .Range("s2_EaveOverhang").Value <> "" Then
            dStr = dStr & .Range("s2_EaveOverhang").Value & "' wide x " & .Range("Building_Length").Value & "' long eave overhang on sidewall #2, "
            If .Range("s2_EaveOverhangSoffit").Value = "Yes" Then
                dStr = dStr & LCase(.Range("s2_EaveOverhangSoffit").offset(0, 3).Value) & " " & _
                LCase(.Range("s2_EaveOverhangSoffit").offset(0, 2).Value) & " sidewall #2 soffit panels with " & _
                LCase(.Range("s2_EaveOverhangSoffit").offset(0, 4).Value) & " trim, "
            End If
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = ""
        'e3 overhang
        If .Range("e3_GableOverhang").Value <> "" Then
            dStr = dStr & .Range("Building_Width").Value & "' wide x " & .Range("e3_GableOverhang").Value & "' long gable overhang on endwall #3, "
            If .Range("e3_GableOverhangSoffit").Value = "Yes" Then
                dStr = dStr & LCase(.Range("e3_GableOverhangSoffit").offset(0, 3).Value) & " " & _
                LCase(.Range("e3_GableOverhangSoffit").offset(0, 2).Value) & " endwall #3 soffit panels with " & _
                LCase(.Range("e3_GableOverhangSoffit").offset(0, 4).Value) & " trim, "
            End If
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = ""
        's4 overhang
        If .Range("s4_EaveOverhang").Value <> "" Then
            dStr = dStr & .Range("s4_EaveOverhang").Value & "' wide x " & .Range("Building_Length").Value & "' long eave overhang on sidewall #4, "
            If .Range("s4_EaveOverhangSoffit").Value = "Yes" Then
                dStr = dStr & LCase(.Range("s4_EaveOverhangSoffit").offset(0, 3).Value) & " " & _
                LCase(.Range("s4_EaveOverhangSoffit").offset(0, 2).Value) & " sidewall #4 soffit panels with " & _
                LCase(.Range("s4_EaveOverhangSoffit").offset(0, 4).Value) & " trim, "
            End If
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = ""
        'e1 Extension
        If .Range("e1_GableExtension").Value <> "" Then
            dStr = dStr & .Range("Building_Width").Value & "' wide x " & .Range("e1_GableExtension").Value & "' long gable extension roof structure on endwall #1, "
            If .Range("e1_GableExtensionSoffit").Value = "Yes" Then
                dStr = dStr & LCase(.Range("e1_GableExtensionSoffit").offset(0, 3).Value) & " " & _
                LCase(.Range("e1_GableExtensionSoffit").offset(0, 2).Value) & " endwall #1 soffit panels with " & _
                LCase(.Range("e1_GableExtensionSoffit").offset(0, 4).Value) & " trim, "
            End If
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = ""
        's2 Extension
        If .Range("s2_EaveExtension").Value <> "" Then
            If .Range("s2_EaveExtensionPitch").Value = "Match Roof" Then
                dStr = dStr & .Range("s2_EaveExtension").Value & "' wide x " & .Range("Building_Length").Value & "' long " & b.rPitch & "/12 pitch eave extension roof structure on sidewall #2, "
            Else
                dStr = dStr & .Range("s2_EaveExtension").Value & "' wide x " & .Range("Building_Length").Value & "' long " & .Range("s2_EaveExtensionPitch").Value & "/12 pitch eave extension roof structure on sidewall #2, "
            End If
            If .Range("s2_EaveExtensionSoffit").Value = "Yes" Then _
            dStr = dStr & LCase(.Range("s2_EaveExtensionSoffit").offset(0, 3).Value) & " " & _
            LCase(.Range("s2_EaveExtensionSoffit").offset(0, 2).Value) & " sidewall #2 soffit panels with " & _
            LCase(.Range("s2_EaveExtensionSoffit").offset(0, 4).Value) & " trim, "
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = ""
        'e3 Extension
        If .Range("e3_GableExtension").Value <> "" Then
            dStr = dStr & .Range("Building_Width").Value & "' wide x " & .Range("e3_GableExtension").Value & "' long Gable extension roof structure on endwall #3, "
            If .Range("e3_GableExtensionSoffit").Value = "Yes" Then _
            dStr = dStr & LCase(.Range("e3_GableExtensionSoffit").offset(0, 3).Value) & " " & _
            LCase(.Range("e3_GableExtensionSoffit").offset(0, 2).Value) & " endwall #3 soffit panels with " & _
            LCase(.Range("e3_GableExtensionSoffit").offset(0, 4).Value) & " trim, "
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = ""
        's4 Extension
        If .Range("s4_EaveExtension").Value <> "" Then
            If .Range("s4_EaveExtensionPitch").Value = "Match Roof" Then
                dStr = dStr & .Range("s4_EaveExtension").Value & "' wide x " & .Range("Building_Length").Value & "' long " & b.rPitch & "/12 pitch eave extension roof structure on sidewall #4, "
            Else
                dStr = dStr & .Range("s4_EaveExtension").Value & "' wide x " & .Range("Building_Length").Value & "' long " & .Range("s4_EaveExtensionPitch").Value & "/12 pitch eave extension roof structure on sidewall #4, "
            End If
            If .Range("s4_EaveExtensionSoffit").Value = "Yes" Then _
            dStr = dStr & LCase(.Range("s4_EaveExtensionSoffit").offset(0, 3).Value) & " " & _
            LCase(.Range("s4_EaveExtensionSoffit").offset(0, 2).Value) & " sidewall #4 soffit panels with " & _
            LCase(.Range("s4_EaveExtensionSoffit").offset(0, 4).Value) & " trim, "
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = ""
        'intersections
        If b.s2e1ExtensionIntersection = True Then
            dStr = "Sidewall #2 and endwall #1 extension intersections"
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        If b.s2e3ExtensionIntersection = True Then
            dStr = "Sidewall #2 and endwall #3 extension intersections"
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        If b.s4e1ExtensionIntersection = True Then
            dStr = "Sidewall #4 and endwall #1 extension intersections"
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        If b.s4e3ExtensionIntersection = True Then
            dStr = "Sidewall #4 and endwall #3 extension intersections"
            dCell.offset(RowCount, 0).Value = dStr
            RowCount = RowCount + 1
        End If
        dStr = Trim(dStr)
        'remove last comma
        'dStr = Left(dStr, Len(dStr) - 1)
    End With
    DescriptionSht.Columns(1).AutoFit
    'DescriptionSht.Range("DescriptionCell").Value = dStr

End With

    
End Sub
