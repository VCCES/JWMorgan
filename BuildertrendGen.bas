Attribute VB_Name = "BuildertrendGen"
Option Explicit


'''packages and outputs all items to buildertrend sheet
Sub GenSheet()
Dim sht As Worksheet
Dim sheetExists As Boolean
Dim cell As Range
Dim item As Variant
Dim colorStr As String
Dim totalCost As Variant
Dim descriptStr As String
Dim readCollection As Collection
Dim sectionTitle As String
Dim taxRate As Double
Dim measurementStr As String
Dim qtyStr As String


'check for existing buildertrend sheet. make if needed, declare as current if not
On Error Resume Next
sheetExists = Not ThisWorkbook.Sheets("Buildertrend Estimate") Is Nothing
On Error GoTo 0

'events/updating
GeneralMod.EventsUpdating
If sheetExists = True Then ThisWorkbook.Sheets("Buildertrend Estimate").Delete

BuildertrendTmp.Copy Before:=EstSht
Set sht = EstSht.Previous
sht.Name = "Buildertrend Estimate"


sht.Visible = xlSheetVisible

'just read the other sheets now. I hate this lol
'start with structural steel sheet
Set cell = ThisWorkbook.Sheets("Structural Steel Price List").Range("a4")
'structural steel
Do While cell.Value <> ""
    Set item = New clsMember
    item.clsType = "Member"
    item.Qty = cell.Value
    item.Size = cell.offset(0, 1).Value
    item.Measurement = cell.offset(0, 3).Value
    item.totalCost = cell.offset(0, 7).Value
    'item.BuildertrendDescription = item.Quantity & " x " & item.Measurement & " " & item.Name
    BTItems.AddItem item
    Set cell = cell.offset(1, 0)
Loop
'read sheet metal materials list for doors since they're not otherwise captured
Set cell = ThisWorkbook.Sheets("Materials Price List").Range("PriceListQtyCell1")
Do While cell.Value <> ""
    'check for doors only
    If cell.offset(0, 2).Value Like "*3070*" Or cell.Value Like "*4070*" Then
        Set item = New clsMiscItem
        item.clsType = "MiscItem"
        item.Quantity = cell.Value
        item.Name = cell.offset(0, 2).Value
        item.Measurement = cell.offset(0, 3).Value
        item.totalCost = cell.offset(0, 7).Value
        'item.BuildertrendDescription = item.Quantity & " x " & item.Measurement & " " & item.Name
        BTItems.AddItem item
    End If
    Set cell = cell.offset(1, 0)
Loop


'write to categories
Set cell = sht.Range("FirstCodeCell").offset(0, 1)

'write values to template
Do While cell.Value <> ""
    totalCost = 0
    descriptStr = ""
    colorStr = ""
    measurementStr = ""
    qtyStr = ""
    sectionTitle = cell.Value
    Select Case True
    Case sectionTitle = "Purchased Structural Steel"
        Set readCollection = BTItems.PurchasedStructuralSteel
    Case sectionTitle = "Stocked Structural Steel"
        Set readCollection = BTItems.StockedStructuralSteel
    Case sectionTitle = "Eavestruts"
        Set readCollection = BTItems.Eavestruts
    Case sectionTitle = "Sheetmetal"
        Set readCollection = BTItems.Sheetmetal
    Case sectionTitle = "Purchased Personnel Doors"
        Set readCollection = BTItems.PurchasePersonnelDoors
    Case sectionTitle = "Stocked Personnel Doors"
        Set readCollection = BTItems.StockedPersonnelDoors
    Case sectionTitle = "Insulation"
        Set readCollection = BTItems.Insulation
    Case sectionTitle = "OH Doors"
        Set readCollection = BTItems.OHDoors
    Case sectionTitle = "Windows"
        Set readCollection = BTItems.Windows
    Case sectionTitle = "Misc. Materials"
        Set readCollection = BTItems.GenMiscMaterials
    Case sectionTitle = "Exhaust Fans, Louvers, Weatherhoods"
        Set readCollection = BTItems.ExhaustFansLouversWeatherhoods
    Case sectionTitle Like "*Anchors*"
        Set readCollection = BTItems.Anchors
    Case sectionTitle = "Benchmark Employee Labor"
        BTLaborAndEquipGen cell.offset(0, 2), cell.offset(0, 3), cell.offset(0, 4), "Labor"
        Set readCollection = Nothing    'avoid using labor collection in class since reading directly from sheet
    Case sectionTitle = "Equipment"
        BTLaborAndEquipGen cell.offset(0, 2), cell.offset(0, 3), cell.offset(0, 4), "Equipment"
        Set readCollection = Nothing    'avoid using equipment collection in class since reading directly from sheet
    Case sectionTitle = "Project Description"
        'cell.offset(0, 4).Value = ProjectDescriptionGen        'removed description gen for now per client request
        Set readCollection = Nothing
    End Select
    'total values and make description
    If Not readCollection Is Nothing Then
        For Each item In readCollection
            If item.totalCost <> 0 Then
                totalCost = totalCost + item.totalCost
            Else
                totalCost = totalCost + item.Quantity * item.UnitCost
            End If
            If descriptStr <> "" Then descriptStr = descriptStr & vbNewLine
            If item.Color <> "" And item.Color <> "N/A" And item.Color <> "n/a" Then colorStr = item.Color & " "
            If item.Measurement <> "" And item.Measurement <> "N/A" And item.Measurement <> "n/a" Then measurementStr = item.Measurement
            qtyStr = "(" & item.Quantity & ") "
            descriptStr = descriptStr & qtyStr & colorStr & item.Name & " - " & measurementStr
        Next item
        cell.offset(0, 2).Value = totalCost * (1 + MasterPriceSht.Range("TaxRate").Value)
        cell.offset(0, 4).Value = descriptStr
        'update standard markup (make normal value instead of percent for the buildertrend import)
        cell.offset(0, 3).Value = MasterPriceSht.Range("Markup").Value * 100
    End If
    Set cell = cell.offset(1, 0)
Loop


GeneralMod.EventsUpdating


End Sub

'populates labor and equipment lines item using the information on the cost estimate sheet
Sub BTLaborAndEquipGen(totalCell As Range, markupCell As Range, descriptCell As Range, LaborOrEquipment As String)
Dim costEstSht As Worksheet
Dim laborCell As Range
Dim descriptStr As String
Dim totalCost As Double
Dim SectionCostWeight As Double
Dim WeightedCost As Double

If LaborOrEquipment = "Labor" Then
    SectionCostWeight = 0.75
ElseIf LaborOrEquipment = "Equipment" Then
    SectionCostWeight = 0.1
End If
    

Set costEstSht = ThisWorkbook.Sheets("Cost Estimate")

For Each laborCell In Range(costEstSht.Range("FirstLaborTblCell"), costEstSht.Range("FirstLaborTblCell").End(xlDown))
    If laborCell.Value = "LABOR TOTAL:" Then Exit For       'exit when total row is found
    If IsNumeric(laborCell.offset(0, 3).Value) = True Then
        totalCost = totalCost + laborCell.offset(0, 3).Value * SectionCostWeight
        If descriptStr <> "" Then descriptStr = descriptStr & vbNewLine
        descriptStr = descriptStr & laborCell.Value & ": " & Format(laborCell.offset(0, 1).Value, laborCell.offset(0, 1).NumberFormat) _
        & ", " & Format(laborCell.offset(0, 2).Value * SectionCostWeight, laborCell.offset(0, 2).NumberFormat)
    End If
Next

'write values
totalCell.Value = totalCost
markupCell.Value = 17.65
descriptCell.Value = descriptStr

End Sub
'generates project description using text on project description sheet
Private Function ProjectDescriptionGen() As String
Dim descriptCell As Range
Dim descriptStr As String

For Each descriptCell In Range(ThisWorkbook.Sheets("Project Description").Range("DescriptionCell"), ThisWorkbook.Sheets("Project Description").Range("DescriptionCell").End(xlDown))
    If descriptStr <> "" Then descriptStr = descriptStr & vbNewLine
    descriptStr = descriptStr & descriptCell.Value
Next descriptCell

ProjectDescriptionGen = descriptStr

End Function




