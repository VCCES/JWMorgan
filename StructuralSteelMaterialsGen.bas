Attribute VB_Name = "StructuralSteelMaterialsGen"
Option Explicit

Public EaveStrutCount As Integer

Sub ColTest()
Dim col As New Collection
Dim Panel As New clsPanel

Panel.Quantity = 1
Panel.PanelLength = 45
Panel.rEdgePosition = 3
col.Add Panel
Set Panel = New clsPanel
Set Panel = col(1)
Panel.rEdgePosition = 2
col.Add Panel

Debug.Print "."
End Sub





Sub MoveExtensionOverhangMembers(b As clsBuilding)

Dim Member As clsMember

For Each Member In b.e1Rafters
    If Member.Placement Like "*Extension*" Then
        b.e1ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.e3Rafters
    If Member.Placement Like "*Extension*" Then
        b.e3ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.e1Columns
    If Member.Placement Like "*Extension*" Then
        b.e1ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.e3Columns
    If Member.Placement Like "*Extension*" Then
        b.e3ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.s2Columns
    If Member.Placement Like "*Extension*" Then
        b.s2ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.s4Columns
    If Member.Placement Like "*Extension*" Then
        b.s4ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.e1Rafters
    If Member.Placement Like "*Overhang*" Then
        b.e1ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.e3Rafters
    If Member.Placement Like "*Overhang*" Then
        b.e3ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.e1Columns
    If Member.Placement Like "*Overhang*" Then
        b.e1ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.e3Columns
    If Member.Placement Like "*Overhang*" Then
        b.e3ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.s2Columns
    If Member.Placement Like "*Overhang*" Then
        b.s2ExtensionMembers.Add Member
    End If
Next Member

For Each Member In b.s4Columns
    If Member.Placement Like "*Overhang*" Then
        b.s4ExtensionMembers.Add Member
    End If
Next Member
End Sub

Sub WeldPlateGen(RafterLine As String, b As clsBuilding)

Dim Columns As Collection
Dim Column As clsMember
Dim WeldPlate As clsMiscItem
Dim WeldPlateRng As Range
Dim mCell As Range

Set WeldPlateRng = SteelLookupSht.Range("WeldPlateTblStart", "WeldPlateTblEnd")

Select Case RafterLine
Case "e1"
    Set Columns = b.e1Columns
Case "e3"
    Set Columns = b.e3Columns
Case "int"
    Set Columns = b.InteriorColumns
End Select

For Each Column In Columns
    Set WeldPlate = New clsMiscItem
    WeldPlate.clsType = "Weld Plate"
    WeldPlate.Quantity = Column.Qty
    For Each mCell In WeldPlateRng
        If Column.Size = mCell.Value Then
            WeldPlate.Name = mCell.offset(0, 1).Value
            WeldPlate.Measurement = WeldPlate.Name
            WeldPlate.FootageCost = mCell.offset(0, 2).Value
            WeldPlate.Height = Right(WeldPlate.Name, Len(WeldPlate.Name) - InStr(1, WeldPlate.Name, "x"))
            WeldPlate.Width = Left(WeldPlate.Name, InStr(1, WeldPlate.Name, "x") - 1)
            Column.ComponentMembers.Add WeldPlate
        End If
    Next mCell
Next Column
    
End Sub

Sub CombineWeldPlates(b As clsBuilding)

Dim Column As clsMember
Dim WeldPlate As clsMiscItem


With b

For Each Column In .e1Columns
    For Each WeldPlate In Column.ComponentMembers
        If WeldPlate.clsType = "Weld Plate" Then
            .WeldPlates.Add WeldPlate
        End If
    Next WeldPlate
Next Column

For Each Column In .e3Columns
    For Each WeldPlate In Column.ComponentMembers
        If WeldPlate.clsType = "Weld Plate" Then
            .WeldPlates.Add WeldPlate
        End If
    Next WeldPlate
Next Column

For Each Column In .InteriorColumns
    For Each WeldPlate In Column.ComponentMembers
        If WeldPlate.clsType = "Weld Plate" Then
            .WeldPlates.Add WeldPlate
        End If
    Next WeldPlate
Next Column

Call DuplicateMaterialRemoval(.WeldPlates, "Misc")

End With

End Sub


' ------------------- Sub used for testing ----------------------
'currently being called from materialslistgen sub
Sub StructuralSteelMaterialsGen(b As clsBuilding)

Application.ScreenUpdating = False

EaveStrutCount = 0

Dim n As Integer
Dim SteelSht As Worksheet
Dim tempGirtsCollection As Collection
Dim manualGirtOptimization As Collection
Dim NewOptimizedCol As Collection
Dim GirtsCollection As Collection
Dim PurlinsCollection As Collection
Dim ReceiverCeeCollection As Collection
Dim EaveStrutCollection As Collection
Dim Member As clsMember
Dim Span As clsMember
Dim NewMember As clsMember
Dim RoofPurlins As Collection
Dim temp8RoofPurlins As Collection
Dim temp10RoofPurlins As Collection
Dim RoofPurlins8 As Collection
Dim manualRoofPurlinOptimization As Collection
Dim NewRoofOptimizedCol As Collection
Dim RoofPurlins10 As Collection
Dim temp8Receivers As Collection
Dim temp10Receivers As Collection
Dim Receivers8 As Collection
Dim Receivers10 As Collection
Dim EndwallColumns As Collection
Dim EndwallRafters As Collection
Dim SteelCollectionCls As clsSteelCollection
Dim NewSteelCollectionCls As clsSteelCollection
Dim IBeamClsCollection As Collection
Dim TSClsCollection As Collection
Dim IBeams As Collection
Dim TS As Collection
Dim Size As String
Dim Created As Boolean


Dim FO As clsFO
Dim item As Object
Dim i As Double
Dim NextSpanNum As Integer



If EstSht.Range("BayNum").Value > 1 Then
    IntColumnsGen b
    Call AdjustSidewallColumns(b, "s2")
    Call AdjustSidewallColumns(b, "s4")
End If

Call EndwallColumnCLCalc(b, "e1")
Call EndwallColumnCLCalc(b, "e3")

Call FOJambsCalc(b, "e1")
Call FOJambsCalc(b, "e3")
Call FOJambsCalc(b, "s2")
Call FOJambsCalc(b, "s4")


If b.ExpandableEndwall("e1") Then Call RemoveEndwallColumns(b, "e1")
If b.ExpandableEndwall("e3") Then Call RemoveEndwallColumns(b, "e3")

Call RafterGen(b, "e1")
Call RafterGen(b, "e3")
Call RafterGen(b, "int")



Call EndwallGirtLengthCalc(b, "e1")
Call EndwallGirtLengthCalc(b, "e3")
Call EndwallGirtLengthCalc(b, "s2")
Call EndwallGirtLengthCalc(b, "s4")

Call EaveStrutTypes(b, "s2")
Call EaveStrutTypes(b, "s4")

Call AdjustEndwallColumns(b, "e1")
Call AdjustEndwallColumns(b, "e3")
Call AdjustEndwallColumns(b, "Int")

Call AdjustFOMembers(b, "e1")
Call AdjustFOMembers(b, "e3")
Call AdjustFOMembers(b, "s2")
Call AdjustFOMembers(b, "s4")

Call OverhangExtensionMembersGen(b)

Call RoofPurlinGen(b)

Call WeldPlateGen("e1", b)
Call WeldPlateGen("e3", b)
Call WeldPlateGen("int", b)

Call CombineWeldPlates(b)

Call BaseAngleTrimGen(b)

Call AdditionalWeldClips(b, "e1")
Call AdditionalWeldClips(b, "e3")
Call AdditionalWeldClips(b, "s2")
Call AdditionalWeldClips(b, "s4")

Call FieldLocateFOCalc(b)

'delete old output sheets
Application.DisplayAlerts = False
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Structural Steel Price List" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Optimized Cut List" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n
For n = ThisWorkbook.Sheets.Count To 1 Step -1
    If ThisWorkbook.Sheets(n).Name = "Structural Steel Materials List" Then
        ThisWorkbook.Sheets(n).Delete
        Exit For
    End If
Next n

''' Structural Steel Lists
SteelCompleteMemberShtTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set SteelSht = ThisWorkbook.Sheets("SteelCompleteMemberShtTmp (2)")
'rename
SteelSht.Name = "Optimized Cut List"
SteelSht.Visible = xlSheetVisible

SteelMaterialsListTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set SteelSht = ThisWorkbook.Sheets("SteelMaterialsListTmp (2)")
'rename
SteelSht.Name = "Structural Steel Materials List"
SteelSht.Visible = xlSheetVisible

'set new output sheets
SteelOutputShtTmp.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set SteelSht = ThisWorkbook.Sheets("SteelOutputShtTmp (2)")
'rename
SteelSht.Name = "Structural Steel Price List"
SteelSht.Visible = xlSheetVisible

Application.DisplayAlerts = True

Call SteelMaterialOutput(b)

DrawItems b

EstSht.Activate

Set GirtsCollection = New Collection
Set tempGirtsCollection = New Collection
Set manualGirtOptimization = New Collection
Set EaveStrutCollection = New Collection
Set temp8RoofPurlins = New Collection
Set manualRoofPurlinOptimization = New Collection
Set NewOptimizedCol = New Collection
Set temp10RoofPurlins = New Collection
Set RoofPurlins8 = New Collection
Set RoofPurlins10 = New Collection
Set NewRoofOptimizedCol = New Collection
Set RoofPurlins = New Collection
Set temp8Receivers = New Collection
Set temp10Receivers = New Collection
Set Receivers8 = New Collection
Set Receivers10 = New Collection
Set EndwallRafters = New Collection
Set EndwallColumns = New Collection
Set IBeamClsCollection = New Collection
Set TSClsCollection = New Collection
Set IBeams = New Collection
Set TS = New Collection
Set SteelCollectionCls = New clsSteelCollection


'Girts
ParseGirts b, tempGirtsCollection, b.e1Girts, EaveStrutCollection, manualGirtOptimization
ParseGirts b, tempGirtsCollection, b.s2Girts, EaveStrutCollection, manualGirtOptimization
ParseGirts b, tempGirtsCollection, b.e3Girts, EaveStrutCollection, manualGirtOptimization
ParseGirts b, tempGirtsCollection, b.s4Girts, EaveStrutCollection, manualGirtOptimization

'FO Purlins
ParseFOPurlins tempGirtsCollection, temp8Receivers, b, b.e1FOs, manualGirtOptimization
ParseFOPurlins tempGirtsCollection, temp8Receivers, b, b.s2FOs, manualGirtOptimization
ParseFOPurlins tempGirtsCollection, temp8Receivers, b, b.e3FOs, manualGirtOptimization
ParseFOPurlins tempGirtsCollection, temp8Receivers, b, b.s4FOs, manualGirtOptimization
ParseFOPurlins tempGirtsCollection, temp8Receivers, b, b.fieldlocateFOs, manualGirtOptimization

            

'roof purlins - 8/10 C Purlins + Eave struts
For i = b.RoofPurlins.Count To 1 Step -1
    Set Member = b.RoofPurlins(i)
    If Member.Size = "8"" C Purlin" Then
        If Member.Length = 15 * 12 Then
            manualRoofPurlinOptimization.Add Member
        Else
            temp8RoofPurlins.Add Member
        End If
    ElseIf Member.Size = "10"" C Purlin" Then
        temp10RoofPurlins.Add Member
    ElseIf Member.mType = "Eave Strut" Then
        EaveStrutCollection.Add Member
        b.RoofPurlins.Remove (i)
    End If
Next i


'Rafters and Columns
ParseColumns tempGirtsCollection, manualGirtOptimization, temp8Receivers, temp10Receivers, EndwallColumns, IBeams, TS, b.e1Columns
ParseColumns tempGirtsCollection, manualGirtOptimization, temp8Receivers, temp10Receivers, EndwallColumns, IBeams, TS, b.e3Columns
ParseColumns tempGirtsCollection, manualGirtOptimization, temp8Receivers, temp10Receivers, EndwallColumns, IBeams, TS, b.e1Rafters
ParseColumns tempGirtsCollection, manualGirtOptimization, temp8Receivers, temp10Receivers, EndwallColumns, IBeams, TS, b.e3Rafters
ParseColumns tempGirtsCollection, manualGirtOptimization, temp8Receivers, temp10Receivers, EndwallColumns, IBeams, TS, b.InteriorColumns
ParseColumns tempGirtsCollection, manualGirtOptimization, temp8Receivers, temp10Receivers, EndwallColumns, IBeams, TS, b.intRafters


CombinePurlins b, manualGirtOptimization, tempGirtsCollection, NewOptimizedCol
CombinePurlins b, manualRoofPurlinOptimization, temp8RoofPurlins, NewRoofOptimizedCol

'all Ibeams are in IBeams
'all Tube Steel are in TS

Dim tempIBeams As Collection
Set tempIBeams = New Collection

'create class for each size Ibeam
For Each Member In IBeams
    Created = False
    Size = Member.Size
    If Member.Length <= 60 * 12 Then
        For Each SteelCollectionCls In IBeamClsCollection
            If SteelCollectionCls.Size = Size Then
                'already created
                Created = True
                SteelCollectionCls.Members.Add Member
            End If
        Next SteelCollectionCls
        If Created = False Then
            Set NewSteelCollectionCls = New clsSteelCollection
            NewSteelCollectionCls.Size = Member.Size
            NewSteelCollectionCls.Members.Add Member
            IBeamClsCollection.Add NewSteelCollectionCls
        End If
    Else
        tempIBeams.Add Member
    End If
Next Member

'create class for each size TS
For Each Member In TS
Created = False
Size = Member.Size
    For Each SteelCollectionCls In TSClsCollection
        If SteelCollectionCls.Size = Size Then
            'already created
            Created = True
            SteelCollectionCls.Members.Add Member
        End If
    Next SteelCollectionCls
    If Created = False Then
        Set NewSteelCollectionCls = New clsSteelCollection
        NewSteelCollectionCls.Size = Member.Size
        NewSteelCollectionCls.Members.Add Member
        TSClsCollection.Add NewSteelCollectionCls
    End If
Next Member

Call DuplicateMaterialRemoval(tempGirtsCollection, "Steel")
Call DuplicateMaterialRemoval(temp8RoofPurlins, "Steel")
Call DuplicateMaterialRemoval(temp10RoofPurlins, "Steel")
Call DuplicateMaterialRemoval(temp8Receivers, "Steel")
Call DuplicateMaterialRemoval(temp10Receivers, "Steel")

If tempGirtsCollection.Count > 0 Then
    Call JankyBPPSolver.BPP_Solver(GirtsCollection, tempGirtsCollection, "Girt", "Steel", "e1")
End If
If temp8RoofPurlins.Count > 0 Then
    Call JankyBPPSolver.BPP_Solver(RoofPurlins8, temp8RoofPurlins, "Roof Purlin", "Steel", "e1")
End If
If temp10RoofPurlins.Count > 0 Then
Call JankyBPPSolver.BPP_Solver(RoofPurlins10, temp10RoofPurlins, "Roof Purlin", "Steel", "e1")
End If
If temp8Receivers.Count > 0 Then
Call JankyBPPSolver.BPP_Solver(Receivers8, temp8Receivers, "Girt", "Steel", "e1")
End If
If temp10Receivers.Count > 0 Then
Call JankyBPPSolver.BPP_Solver(Receivers10, temp10Receivers, "Girt", "Steel", "e1")
End If

Set TS = New Collection
Set IBeams = tempIBeams

For Each SteelCollectionCls In TSClsCollection
    Call JankyBPPSolver.BPP_Solver(TS, SteelCollectionCls.Members, "TS", "Steel", "e1")
Next SteelCollectionCls

For Each SteelCollectionCls In IBeamClsCollection
    Call JankyBPPSolver.BPP_Solver(IBeams, SteelCollectionCls.Members, "IBeam", "Steel", "e1")
Next SteelCollectionCls

For Each Member In IBeams
    If Member.Length > 60 * 12 Then
        MsgBox ("At least one I-Beam in this building is over 60', the approved orderable length.")
        Exit For
    End If
Next Member

Call DuplicateMaterialRemoval(EaveStrutCollection, "Steel")

'rename all members passed through BPPSolver
For Each Member In GirtsCollection
    Member.Size = "8"" C Purlin"
    Member.mType = "C Purlin"
    For Each Span In Member.ComponentMembers
        Span.Size = "8"" C Purlin"
    Next Span
Next Member

For Each Member In RoofPurlins8
    Member.Size = "8"" C Purlin"
    Member.mType = "C Purlin"
    For Each Span In Member.ComponentMembers
        Span.Size = "8"" C Purlin"
    Next Span
Next Member

For Each Member In RoofPurlins10
    Member.Size = "10"" C Purlin"
    Member.mType = "C Purlin"
    For Each Span In Member.ComponentMembers
        Span.Size = "10"" C Purlin"
    Next Span
Next Member

For Each Member In Receivers8
    Member.Size = "8"" Receiver Cee"
    For Each Span In Member.ComponentMembers
        Span.Size = "8"" Receiver Cee"
    Next Span
Next Member

For Each Member In Receivers10
    Member.Size = "10"" Receiver Cee"
    For Each Span In Member.ComponentMembers
        Span.Size = "10"" Receiver Cee"
    Next Span
Next Member

'Add combined purlins back to Girts Collection
NextSpanNum = GirtsCollection.Count + 1
For Each Member In NewOptimizedCol
    Member.Placement = "8"" C Purlin Span #" & NextSpanNum
    Member.mType = "C Purlin"
    NextSpanNum = NextSpanNum + 1
    GirtsCollection.Add Member
Next Member

NextSpanNum = RoofPurlins8.Count + 1
For Each Member In NewRoofOptimizedCol
    Member.Placement = "Roof Purlin 8"" C Purlin Span #" & NextSpanNum
    Member.mType = "C Purlin"
    NextSpanNum = NextSpanNum + 1
    RoofPurlins8.Add Member
Next Member


''''''''''Cut List Output
Call CutListOutput(IBeams, "I Beams: Columns & Rafters")
Call CutListOutput(TS, "Tube Steel: Columns & Rafters")
Call CutListOutput(GirtsCollection, "Wall Girt")
Call CutListOutput(RoofPurlins8, "8"" Roof Purlin")
Call CutListOutput(RoofPurlins10, "10"" Roof Purlin")
Call CutListOutput(Receivers8, "8"" Receiver Cee")
Call CutListOutput(Receivers10, "10"" Receiver Cee")


''''''''''Price List Output
'Call SteelPriceOutput(b.e1Columns, "endwall 1 column")
'Call SteelPriceOutput(b.e3Columns, "endwall 3 column")
'Call SteelPriceOutput(EndwallColumns, "Endwall Column")
'Call SteelPriceOutput(b.InteriorColumns, "Main Rafter Line Column")
'Call SteelPriceOutput(b.e1Rafters, "endwall 1 rafter")
'Call SteelPriceOutput(b.e3Rafters, "endwall 3 rafter")
'Call SteelPriceOutput(EndwallRafters, "Endwall Rafter")
'Call SteelPriceOutput(b.intRafters, "Main Rafter")
Call SteelPriceOutput(IBeams, "I Beams: Columns & Rafters")
Call SteelPriceOutput(TS, "Tube Steel: Columns & Rafters")
Call SteelPriceOutput(GirtsCollection, "Wall Girts, Non-Load Bearing Columns, Endwall Rafters, and FO Members")
Call SteelPriceOutput(RoofPurlins8, "8"" Roof Purlins")
Call SteelPriceOutput(RoofPurlins10, "10"" Roof Purlins")
'Call SteelPriceOutput(b.e1Girts, "e1 Wall Girt")
'Call SteelPriceOutput(b.s2Girts, "s2 Wall Girt")
'Call SteelPriceOutput(b.e3Girts, "e3 Wall Girt")
'Call SteelPriceOutput(b.s4Girts, "s4 Wall Girt")
'Call SteelPriceOutput(b.RoofPurlins, "Roof Purlin")
Call SteelPriceOutput(EaveStrutCollection, "Eave Strut")
Call SteelPriceOutput(Receivers8, "8"" Endwall Rafters and FO Jambs")
Call SteelPriceOutput(Receivers10, "10"" Endwall Rafters")
'Call SteelPriceOutput(b.e1FOs, "FO", True)
'Call SteelPriceOutput(b.s2FOs, "FO", True)
'Call SteelPriceOutput(b.e3FOs, "FO", True)
'Call SteelPriceOutput(b.s4FOs, "FO", True)
Call SteelPriceOutput(b.e1OverhangMembers, "e1 gable overhang")
Call SteelPriceOutput(b.e1ExtensionMembers, "e1 gable extension")
Call SteelPriceOutput(b.s2OverhangMembers, "s2 eave overhang")
Call SteelPriceOutput(b.s2ExtensionMembers, "s2 eave extension")
Call SteelPriceOutput(b.e3OverhangMembers, "e3 gable overhang")
Call SteelPriceOutput(b.e3ExtensionMembers, "e3 gable extension")
Call SteelPriceOutput(b.s4OverhangMembers, "s4 eave overhang")
Call SteelPriceOutput(b.s4ExtensionMembers, "s4 eave extension")
Call SteelPriceOutput(b.BaseAngleTrim, "Base Angle")


Dim FullMemberSht As Worksheet
Dim LastRow As Integer
Set FullMemberSht = ThisWorkbook.Sheets("Structural Steel Price List")
Dim WeldPlate As clsMiscItem
Dim WeldPlateRng As Range

If FullMemberSht.Range("A4").Value = "" Then
    LastRow = 4
Else
    LastRow = FullMemberSht.Range("A3").End(xlDown).offset(1, 0).Row
End If


With FullMemberSht
'''''''''''''''''''Add Weld Plates to Output Sheet
    For Each WeldPlate In b.WeldPlates
        .Range("A" & LastRow).Value = WeldPlate.Quantity
        .Range("B" & LastRow).Value = WeldPlate.Name & " Weld Plate"
        .Range("C" & LastRow).Value = "n/a"
        .Range("D" & LastRow).Value = WeldPlate.Name
        .Range("E" & LastRow).Value = WeldPlate.FootageCost
        .Range("F" & LastRow).Value = "each"
        .Range("G" & LastRow).Value = "n/a"
        .Range("H" & LastRow).Value = WeldPlate.FootageCost * WeldPlate.Quantity
        LastRow = LastRow + 1
    Next WeldPlate

'''''''''''''''''''Add Weld Clips to Output Sheet
    .Range("A" & LastRow).Value = b.WeldClips
    .Range("B" & LastRow).Value = "Weld Clips"
    .Range("C" & LastRow).Value = "n/a"
    .Range("D" & LastRow).Value = "n/a"
    .Range("E" & LastRow).Value = 1.57
    .Range("F" & LastRow).Value = "each"
    .Range("G" & LastRow).Value = "n/a"
    .Range("H" & LastRow).Value = 1.57 * b.WeldClips
    
    'sum total structural steel costs
    'report later in cost estimate sub
    For i = 4 To LastRow
        If IsNumeric(.Range("H" & i).Value) Then
            b.SSTotalCost = b.SSTotalCost + .Range("H" & i).Value
        End If
    Next i
    
End With


    

Application.ScreenUpdating = True

End Sub
Private Sub ParseColumns(ByRef tempGirtsCollection As Collection, ByRef manualGirtOptimization As Collection, ByRef temp8Receivers As Collection, ByRef temp10Receivers As Collection, ByRef EndwallColumns As Collection, IBeams As Collection, TS As Collection, ByRef MemberCollection As Collection)
Dim Member As clsMember
'columns - receiver cees + C Purlins
For Each Member In MemberCollection
    If Member.Size = "8"" Receiver Cee" Then
        temp8Receivers.Add Member
    ElseIf Member.Size = "10"" Receiver Cee" Then
        temp10Receivers.Add Member
    ElseIf Member.Size = "8"" C Purlin" Then
        If Member.Length = 15 * 12 Then
            manualGirtOptimization.Add Member
        Else
            tempGirtsCollection.Add Member
        End If
    ElseIf Member.Size Like "W*" Then
        IBeams.Add Member
    ElseIf Member.Size Like "TS*" Then
        TS.Add Member
    End If
Next Member
End Sub
Sub ParseFOPurlins(ByRef tempGirtsCollection As Collection, ByRef temp8Receivers As Collection, ByRef b As clsBuilding, FOCollection As Collection, manualGirtOptimization As Collection)
'FO - Purlins and Receivers
Dim FO As clsFO
Dim Member As clsMember
Dim i As Integer

For Each FO In FOCollection
    For i = FO.FOMaterials.Count To 1 Step -1
        Set Member = FO.FOMaterials(i)
        If Member.Size = "8"" C Purlin" Then
            If Member.Length = 15 * 12 Then
                manualGirtOptimization.Add Member
            Else
                tempGirtsCollection.Add Member
            End If
        ElseIf Member.Size = "8"" Receiver Cee" Then
            temp8Receivers.Add Member
        End If
    Next i
Next FO
End Sub

Sub ParseGirts(b As clsBuilding, tempGirtsCollection As Collection, buildingGirts As Collection, EaveStrutCollection As Collection, manualGirtOptimization As Collection)
'Girts optimization collection
Dim Member As clsMember
Dim i As Integer

For i = buildingGirts.Count To 1 Step -1
    Set Member = buildingGirts(i)
    If Member.Size = "8"" C Purlin" Then
        If Member.Length = 15 * 12 Then
            manualGirtOptimization.Add Member
        Else
            tempGirtsCollection.Add Member
        End If
    Else
        EaveStrutCollection.Add Member
        buildingGirts.Remove (i)
    End If
Next i

End Sub

Sub CombinePurlins(b As clsBuilding, manualGirtOptimization As Collection, tempGirtsCollection As Collection, NewOptimizedCol As Collection)
Dim FifteenFootPurlinCount As Integer
Dim FirstMember As clsMember
Dim SecondMember As clsMember
Dim FullSpan As clsMember
Dim i As Integer
Dim Member As clsMember
FifteenFootPurlinCount = manualGirtOptimization.Count
If FifteenFootPurlinCount Mod 2 = 0 Then
    'even number, do nothing
Else
    'odd number, move 1 member to temp girt collection
    Set Member = manualGirtOptimization(1)
    tempGirtsCollection.Add Member
    manualGirtOptimization.Remove (1)
End If

'combine members
FifteenFootPurlinCount = manualGirtOptimization.Count
i = 1
While i < FifteenFootPurlinCount
    Set FirstMember = manualGirtOptimization(i)
    Set SecondMember = manualGirtOptimization(i + 1)
    Set FullSpan = New clsMember
    FullSpan.ComponentMembers.Add FirstMember
    FullSpan.ComponentMembers.Add SecondMember
    FullSpan.Size = "8"" C Purlin"
    FullSpan.Length = 30 * 12
    NewOptimizedCol.Add FullSpan
    i = i + 2
Wend
End Sub

Sub SteelMaterialOutput(b As clsBuilding)

Dim CurRow As Integer
Dim Member As clsMember
Dim SteelSht As Worksheet
Dim FullMemberSht As Worksheet
Dim FO As clsFO
Dim item As Object
Dim j As Double
Dim i As Integer
Dim UnitPrice As Double
Dim UnitMeasure As String
Dim UnitValue As Double
Dim PriceTbl As ListObject
Dim ExtColRow As Integer
Dim ExtRafterRow As Integer
Dim SortRange As Range
Dim KeyRange As Range
Dim EaveStrutRow As Integer

'''''''''''''''''''''''''''Steel Material OUtput Sheet
Set SteelSht = ThisWorkbook.Sheets("Structural Steel Materials List")
SteelSht.Activate

ExtColRow = 0

'e1 Columns
With SteelSht.Range("e1_ColumnsStart")
    i = 0
    For Each Member In b.e1Columns
        If Not Member.Placement Like "*Extension*" And Not Member.Placement Like "*Overhang*" Then
            .offset(i, 0).Value = Member.Qty
            .offset(i, 1).Value = Member.Size
            .offset(i, 2).Value = Member.Placement
            .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
            .offset(i, 4).Value = (Member.Length)
            
            'add next row
            SteelSht.Rows(.offset(i + 1, 0).Row).Insert
            i = i + 1
        Else
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 0).Value = Member.Qty
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 1).Value = Member.Size
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 2).Value = Member.Placement
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 3).Value = ImperialMeasurementFormat(Member.Length)
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 4).Value = Member.Length
            
            'add next row
            SteelSht.Rows(SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow + 1, 0).Row).Insert
            ExtColRow = ExtColRow + 1
        End If
    Next Member
End With

If i > 0 Then
    Set SortRange = SteelSht.Range("e1_ColumnsStart").Resize(i, 5)
    Set KeyRange = SteelSht.Range("e1_ColumnsStart").offset(0, 4).Resize(i, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If
'e3 Columns
With SteelSht.Range("e3_ColumnsStart")
    i = 0
    For Each Member In b.e3Columns
        If Not Member.Placement Like "*Extension*" And Not Member.Placement Like "*Overhang*" Then
            .offset(i, 0).Value = Member.Qty
            .offset(i, 1).Value = Member.Size
            .offset(i, 2).Value = Member.Placement
            .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
            .offset(i, 4).Value = (Member.Length)
            
            'add next row
            SteelSht.Rows(.offset(i + 1, 0).Row).Insert
            i = i + 1
        Else
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 0).Value = Member.Qty
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 1).Value = Member.Size
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 2).Value = Member.Placement
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 3).Value = ImperialMeasurementFormat(Member.Length)
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 4).Value = Member.Length
            
            'add next row
            SteelSht.Rows(SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow + 1, 0).Row).Insert
            ExtColRow = ExtColRow + 1
        End If
    Next Member
End With
If i > 0 Then
    Set SortRange = SteelSht.Range("e3_ColumnsStart").Resize(i, 5)
    Set KeyRange = SteelSht.Range("e3_ColumnsStart").offset(0, 4).Resize(i, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If
'Int Columns
With SteelSht.Range("Int_ColumnsStart")
    i = 0
    For Each Member In b.InteriorColumns
        If Not Member.Placement Like "*Extension*" And Not Member.Placement Like "*Overhang*" Then
            .offset(i, 0).Value = Member.Qty
            .offset(i, 1).Value = Member.Size
            .offset(i, 2).Value = Member.Placement
            .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
            .offset(i, 4).Value = (Member.Length)
            
            'add next row
            SteelSht.Rows(.offset(i + 1, 0).Row).Insert
            i = i + 1
        Else
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 0).Value = Member.Qty
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 1).Value = Member.Size
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 2).Value = Member.Placement
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 3).Value = ImperialMeasurementFormat(Member.Length)
            SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow, 4).Value = Member.Length
            
            'add next row
            SteelSht.Rows(SteelSht.Range("Ext_ColumnsStart").offset(ExtColRow + 1, 0).Row).Insert
            ExtColRow = ExtColRow + 1
        End If
    Next Member
End With
If i > 0 Then
    Set SortRange = SteelSht.Range("Int_ColumnsStart").Resize(i, 5)
    Set KeyRange = SteelSht.Range("Int_ColumnsStart").offset(0, 4).Resize(i, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

If ExtColRow > 0 Then
    Set SortRange = SteelSht.Range("Ext_ColumnsStart").Resize(ExtColRow, 5)
    Set KeyRange = SteelSht.Range("Ext_ColumnsStart").offset(0, 4).Resize(ExtColRow, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

ExtRafterRow = 0

'e1 Rafters
With SteelSht.Range("e1_RaftersStart")
    i = 0
    For Each Member In b.e1Rafters
        If Not Member.Placement Like "*Extension*" And Not Member.Placement Like "*Overhang*" And Not Member.Placement Like "*Stub*" Then
            .offset(i, 0).Value = Member.Qty
            .offset(i, 1).Value = Member.Size
            .offset(i, 2).Value = Member.Placement
            .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
            .offset(i, 4).Value = (Member.Length)
            
            'add next row
            SteelSht.Rows(.offset(i + 1, 0).Row).Insert
            i = i + 1
        Else
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 0).Value = Member.Qty
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 1).Value = Member.Size
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 2).Value = Member.Placement
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 3).Value = ImperialMeasurementFormat(Member.Length)
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 4).Value = Member.Length
            
            'add next row
            SteelSht.Rows(SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow + 1, 0).Row).Insert
            ExtRafterRow = ExtRafterRow + 1
        End If
    Next Member
End With
If i > 0 Then
    Set SortRange = SteelSht.Range("e1_RaftersStart").Resize(i, 5)
    Set KeyRange = SteelSht.Range("e1_RaftersStart").offset(0, 4).Resize(i, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If
'e3 Rafters
With SteelSht.Range("e3_RaftersStart")
    i = 0
    For Each Member In b.e3Rafters
        If Not Member.Placement Like "*Extension*" And Not Member.Placement Like "*Overhang*" And Not Member.Placement Like "*Stub*" Then
            .offset(i, 0).Value = Member.Qty
            .offset(i, 1).Value = Member.Size
            .offset(i, 2).Value = Member.Placement
            .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
            .offset(i, 4).Value = (Member.Length)
            
            'add next row
            SteelSht.Rows(.offset(i + 1, 0).Row).Insert
            i = i + 1
        Else
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 0).Value = Member.Qty
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 1).Value = Member.Size
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 2).Value = Member.Placement
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 3).Value = ImperialMeasurementFormat(Member.Length)
            SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 4).Value = Member.Length
            
            'add next row
            SteelSht.Rows(SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow + 1, 0).Row).Insert
            ExtRafterRow = ExtRafterRow + 1
        End If
    Next Member
End With
If i > 0 Then
    Set SortRange = SteelSht.Range("e3_RaftersStart").Resize(i, 5)
    Set KeyRange = SteelSht.Range("e3_RaftersStart").offset(0, 4).Resize(i, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

If b.intRafters.Count > 0 Then
    'Int Rafters
    With SteelSht.Range("Int_RaftersStart")
        i = 0
        For Each Member In b.intRafters
            If Not Member.Placement Like "*Extension*" And Not Member.Placement Like "*Overhang*" And Not Member.Placement Like "*Stub*" Then
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            Else
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 0).Value = Member.Qty
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 1).Value = Member.Size
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 2).Value = Member.Placement
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 3).Value = ImperialMeasurementFormat(Member.Length)
                SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow, 4).Value = Member.Length
                
                'add next row
                SteelSht.Rows(SteelSht.Range("Ext_RaftersStart").offset(ExtRafterRow + 1, 0).Row).Insert
                ExtRafterRow = ExtRafterRow + 1
            End If
        Next Member
    End With
    If i > 0 Then
        Set SortRange = SteelSht.Range("Int_RaftersStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("Int_RaftersStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

If ExtRafterRow > 0 Then
    Set SortRange = SteelSht.Range("Ext_RaftersStart").Resize(ExtRafterRow, 5)
    Set KeyRange = SteelSht.Range("Ext_RaftersStart").offset(0, 4).Resize(ExtRafterRow, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

EaveStrutRow = 0

If b.e1Girts.Count > 0 Then
    'e1 Girts
    With SteelSht.Range("e1_GirtsStart")
        i = 0
        For Each Member In b.e1Girts
            If Not Member.mType Like "*Eave Strut*" Then
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            Else
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length)
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length
                
                'add next row
                SteelSht.Rows(SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow + 1, 0).Row).Insert
                EaveStrutRow = EaveStrutRow + 1
            End If
        Next Member
    End With
    If i > 0 Then
        Set SortRange = SteelSht.Range("e1_GirtsStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("e1_GirtsStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

If b.s2Girts.Count > 0 Then
    's2 Girts
    With SteelSht.Range("s2_GirtsStart")
        i = 0
        For Each Member In b.s2Girts
            If Not Member.mType Like "*Eave Strut*" Then
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            Else
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length)
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length
                
                'add next row
                SteelSht.Rows(SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow + 1, 0).Row).Insert
                EaveStrutRow = EaveStrutRow + 1
            End If
        Next Member
    End With
    
    If i > 0 Then
        Set SortRange = SteelSht.Range("s2_GirtsStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("s2_GirtsStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

If b.e3Girts.Count > 0 Then
    'e3 Girts
    With SteelSht.Range("e3_GirtsStart")
        i = 0
        For Each Member In b.e3Girts
            If Not Member.mType Like "*Eave Strut*" Then
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            Else
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length)
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length
                
                'add next row
                SteelSht.Rows(SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow + 1, 0).Row).Insert
                EaveStrutRow = EaveStrutRow + 1
            End If
        Next Member
    End With
    If i > 0 Then
        Set SortRange = SteelSht.Range("e3_GirtsStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("e3_GirtsStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If


If b.s4Girts.Count > 0 Then
    's4 Girts
    With SteelSht.Range("s4_GirtsStart")
        i = 0
        For Each Member In b.s4Girts
            If Not Member.mType Like "*Eave Strut*" Then
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            Else
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length)
                SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length
                
                'add next row
                SteelSht.Rows(SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow + 1, 0).Row).Insert
                EaveStrutRow = EaveStrutRow + 1
            End If
        Next Member
    End With

    If i > 0 Then
        Set SortRange = SteelSht.Range("s4_GirtsStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("s4_GirtsStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

'roof purlins / Eave Struts
With SteelSht.Range("RoofPurlinsStart")
    i = 0
    For Each Member In b.RoofPurlins
        If Not Member.mType Like "*Eave Strut*" Then
            .offset(i, 0).Value = Member.Qty
            .offset(i, 1).Value = Member.Size
            .offset(i, 2).Value = Member.Placement
            .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
            .offset(i, 4).Value = (Member.Length)
            
            'add next row
            SteelSht.Rows(.offset(i + 1, 0).Row).Insert
            i = i + 1
        Else
            SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 0).Value = Member.Qty
            SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 1).Value = Member.Size
            SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 2).Value = Member.Placement
            SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 3).Value = ImperialMeasurementFormat(Member.Length)
            SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow, 4).Value = Member.Length
            
            'add next row
            SteelSht.Rows(SteelSht.Range("EaveStrutsStart").offset(EaveStrutRow + 1, 0).Row).Insert
            EaveStrutRow = EaveStrutRow + 1
        End If
    Next Member
End With
If i > 0 Then
    Set SortRange = SteelSht.Range("RoofPurlinsStart").Resize(i, 5)
    Set KeyRange = SteelSht.Range("RoofPurlinsStart").offset(0, 4).Resize(i, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If
If EaveStrutRow > 0 Then
    Set SortRange = SteelSht.Range("EaveStrutsStart").Resize(EaveStrutRow, 5)
    Set KeyRange = SteelSht.Range("EaveStrutsStart").offset(0, 4).Resize(EaveStrutRow, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If
If b.e1FOs.Count > 0 Then
    'e1 FOs
    With SteelSht.Range("e1_FOStart")
        i = 0
        For Each FO In b.e1FOs
        For Each item In FO.FOMaterials
            If item.clsType = "Member" Then
                Set Member = item
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            End If
        Next item
        Next FO
    End With
    If i > 0 Then
        Set SortRange = SteelSht.Range("e1_FOStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("e1_FOStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

If b.s2FOs.Count > 0 Then
    's2 FOs
    With SteelSht.Range("s2_FOStart")
        i = 0
        For Each FO In b.s2FOs
        For Each item In FO.FOMaterials
            If item.clsType = "Member" Then
                Set Member = item
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            End If
        Next item
        Next FO
    End With
    If i > 0 Then
        Set SortRange = SteelSht.Range("s2_FOStart").Resize(i + 1, 5)
        Set KeyRange = SteelSht.Range("s2_FOStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

If b.e3FOs.Count > 0 Then
    'e3 FOs
    With SteelSht.Range("e3_FOStart")
        i = 0
        For Each FO In b.e3FOs
        For Each item In FO.FOMaterials
            If item.clsType = "Member" Then
                Set Member = item
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            End If
        Next item
        Next FO
    End With
    If i > 0 Then
        Set SortRange = SteelSht.Range("e3_FOStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("e3_FOStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

If b.s4FOs.Count > 0 Then
    's4 FOs
    With SteelSht.Range("s4_FOStart")
        i = 0
        For Each FO In b.s4FOs
        For Each item In FO.FOMaterials
            If item.clsType = "Member" Then
                Set Member = item
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            End If
        Next item
        Next FO
    End With
    If i > 0 Then
        Set SortRange = SteelSht.Range("s4_FOStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("s4_FOStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

If b.fieldlocateFOs.Count > 0 Then
    's4 FOs
    With SteelSht.Range("FieldLocate_FOStart")
        i = 0
        For Each FO In b.fieldlocateFOs
        For Each item In FO.FOMaterials
            If item.clsType = "Member" Then
                Set Member = item
                .offset(i, 0).Value = Member.Qty
                .offset(i, 1).Value = Member.Size
                .offset(i, 2).Value = Member.Placement
                .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
                .offset(i, 4).Value = (Member.Length)
                
                'add next row
                SteelSht.Rows(.offset(i + 1, 0).Row).Insert
                i = i + 1
            End If
        Next item
        Next FO
    End With
    If i > 0 Then
        Set SortRange = SteelSht.Range("FieldLocate_FOStart").Resize(i, 5)
        Set KeyRange = SteelSht.Range("FieldLocate_FOStart").offset(0, 4).Resize(i, 1)
        SortRange.Select
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Clear
        ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
            Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
            .SetRange SortRange
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
End If

If b.BaseAngleTrim.Count > 0 Then
    'Base Angle
    With SteelSht.Range("BaseAngleStart")
        i = 0
        For Each Member In b.BaseAngleTrim
            .offset(i, 0).Value = Member.Qty
            .offset(i, 1).Value = Member.Size
            .offset(i, 2).Value = Member.Placement
            .offset(i, 3).Value = ImperialMeasurementFormat(Member.Length)
            .offset(i, 4).Value = (Member.Length)
            
            'add next row
            SteelSht.Rows(.offset(i + 1, 0).Row).Insert
            i = i + 1
        Next Member
    End With
    
    Set SortRange = SteelSht.Range("BaseAngleStart").Resize(i, 5)
    Set KeyRange = SteelSht.Range("BaseAngleStart").offset(0, 4).Resize(i, 1)
    SortRange.Select
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort.SortFields. _
        Add2 Key:=KeyRange, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Structural Steel Materials List").Sort
        .SetRange SortRange
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End If

'Weld Clips
SteelSht.Range("WeldClipsStart").Value = b.WeldClips

Dim WeldPlate As clsMiscItem

'Weld Plates
If b.WeldPlates.Count > 0 Then
    With SteelSht.Range("WeldPlateStart")
        i = 0
        For Each WeldPlate In b.WeldPlates
            .offset(i, 0).Value = WeldPlate.Quantity
            .offset(i, 1).Value = WeldPlate.Name
            
            'add next row
            SteelSht.Rows(.offset(i + 1, 0).Row).Insert
            i = i + 1
        Next WeldPlate
    End With
End If

End Sub

Sub AdditionalWeldClips(b As clsBuilding, eWall As String)

Dim MemberCollection As Collection
Dim FOCollection As Collection
Dim Member As clsMember
Dim Jamb As clsMember
Dim item As Object
Dim FO As clsFO
Dim WeldClips As Double
Dim RightClip As Integer
Dim LeftClip As Integer

WeldClips = b.WeldClips

Select Case eWall
Case "e1"
    Set MemberCollection = b.e1Girts
    Set FOCollection = b.e1FOs
    If b.rShape = "Gable" Then
        WeldClips = WeldClips + 1
        If Not b.ExpandableEndwall("e1") Then
            WeldClips = WeldClips + 1
        End If
    End If
Case "e3"
    Set MemberCollection = b.e3Girts
    Set FOCollection = b.e3FOs
    If Not b.ExpandableEndwall("e3") Then
        WeldClips = WeldClips + 1
    End If
Case "s2"
    Set MemberCollection = b.s2Girts
    Set FOCollection = b.s2FOs
Case "s4"
    Set MemberCollection = b.s4Girts
    Set FOCollection = b.s4FOs
End Select

WeldClips = b.WeldClips

'''''''''''''''''''''''''''''''''''roof purlin Weld Clips are added in Roof Purlin Gen'''''''''''''''''''''''''''''''''''

'Add Weld Clips to Wall Girts
For Each Member In MemberCollection
    RightClip = 1
    LeftClip = 1
    If Member.Size = "8"" C Purlin" Or Member.Size = "10"" C Purlin" Then
        For Each FO In FOCollection
            If Member.rEdgePosition >= FO.rEdgePosition And Member.lEdgePosition - Member.Length <= FO.rEdgePosition And Member.tEdgeHeight <= 30 * 12 And FO.FOType = "OHDoor" Then
                RightClip = 0
                LeftClip = 0
            End If
            For Each item In FO.FOMaterials
                If item.clsType = "Member" Then
                    If item.CL = Member.rEdgePosition And item.tEdgeHeight >= Member.tEdgeHeight And FO.rEdgePosition = Member.rEdgePosition Then
                        RightClip = 0
                    End If
                    If item.CL = Member.lEdgePosition - Member.Length And item.tEdgeHeight >= Member.tEdgeHeight And FO.lEdgePosition = Member.lEdgePosition - Member.Length Then
                        LeftClip = 0
                    End If
                End If
            Next item
        Next FO
    Else
        RightClip = 0
        LeftClip = 0
    End If
    WeldClips = WeldClips + RightClip + LeftClip
Next Member

'FO Weld Clips
For Each FO In FOCollection
    For Each item In FO.FOMaterials
        If item.clsType = "Member" Then
            If item.mType = "FO Receiver Jamb" Then
                If item.bEdgeHeight > 0 Then
                    WeldClips = WeldClips + item.Qty
                ElseIf item.tEdgeHeight = b.DistanceToRoof(eWall, item.CL) Then
                    WeldClips = WeldClips + (item.Qty * 2)
                End If
            ElseIf item.mType = "FO Header" Or item.mType = "FO Stool" Then
                WeldClips = WeldClips + (item.Qty * 2)
            End If
            WeldClips = WeldClips + 2
        End If
    Next item
Next FO
        
b.WeldClips = WeldClips

End Sub
''''''''''' only used for non-expandable endwalls, returns FO edges that meet the following conditions:
'if OHDoor or MiscFO w/ Full Height Jambs option is within the max distance, returns edge that should be used as load bearing (if within 5' of ideal span), creates jambs as necessary
'if ideal column location lands on FO, returns edge that should be used as load bearing, creates jambs as necessary
'if no FOs qualify, returns IdealSpan
'Direction used to check for FOs going towards lower side of roof.
    'e1 single slope: always positive direction (right to left)
    'e3 single slope: always negative direction (left to right)
    'gable roofs: both directions are used for both endwalls
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''StartPos, MaxDistance, IdealSpan all in inches
Function NonExpandableFOJambs(b As clsBuilding, eWall As String, StartPos As Double, MaxDistance As Double, IdealSpan As Double, Direction As Integer) As Double

Dim WallColumns As Collection
Dim FOs As Collection
Dim FO As clsFO
Dim tempLocation As Double
Dim Jamb As clsMember
Dim JambSupport As clsMember
Dim LoadBearingJamb As String
Dim lGtob As Double
Dim rGtob As Double
Dim FirstCorner As Double
Dim LastCorner As Double

If eWall = "e1" Then
    Set FOs = b.e1FOs
    Set WallColumns = b.e1Columns
ElseIf eWall = "e3" Then
    Set FOs = b.e3FOs
    Set WallColumns = b.e3Columns
End If

If Direction = 1 Then 'positive direction
    If b.bWidth * 12 - StartPos < MaxDistance Then
        NonExpandableFOJambs = b.bWidth * 12
        Exit Function
    End If
    For Each FO In FOs
        'if OHDoor or MiscFO w/ full height jambs is inside max distance and the furthest edge is at least within 5' of max distance,
        'then one of the jambs should be load bearing
        If FO.rEdgePosition > StartPos And FO.rEdgePosition < MaxDistance + StartPos And _
            FO.lEdgePosition > StartPos + MaxDistance - 60 And _
            (FO.FOType = "OHDoor" Or FO.StructuralSteelOption = "Full Height Jambs w/ Header & Stool") Then
            
            If FO.lEdgePosition < StartPos + MaxDistance Then 'left edge should be used as load bearing
                LoadBearingJamb = "Left"
            Else
                LoadBearingJamb = "Right"
            End If
            
            lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition)
            rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition)
            'left jamb
            Set Jamb = New clsMember
            Jamb.bEdgeHeight = 0
            If lGtob < 30 * 12 + 4 Then 'don't need jamb support
                Jamb.tEdgeHeight = lGtob
                If LoadBearingJamb = "Left" Then
                    Jamb.LoadBearing = True
                    NonExpandableFOJambs = FO.lEdgePosition
                End If
            Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                Jamb.tEdgeHeight = 30 * 12
                Set JambSupport = New clsMember
                JambSupport.bEdgeHeight = 0
                JambSupport.CL = FO.lEdgePosition
                If LoadBearingJamb = "Left" Then
                    JambSupport.LoadBearing = True
                    NonExpandableFOJambs = JambSupport.CL
                End If
                JambSupport.tEdgeHeight = lGtob
                JambSupport.Length = lGtob
                JambSupport.SetSize b, "Column", eWall, 30
                JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                WallColumns.Add JambSupport
            End If
            Jamb.Length = Jamb.tEdgeHeight
            Jamb.Size = "8"" Receiver Cee"
            Jamb.Width = 2.5
            Jamb.CL = FO.lEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            If Jamb.LoadBearing = False Then
                FO.FOMaterials.Add Jamb
            Else
                WallColumns.Add Jamb
            End If
            
            'right jamb
            Set Jamb = New clsMember
            Jamb.bEdgeHeight = 0
            If rGtob < 30 * 12 + 4 Then 'don't need jamb support
                Jamb.tEdgeHeight = rGtob
                If LoadBearingJamb = "Right" Then
                    Jamb.LoadBearing = True
                    NonExpandableFOJambs = FO.rEdgePosition
                End If
            Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                Jamb.tEdgeHeight = 30 * 12
                Set JambSupport = New clsMember
                JambSupport.bEdgeHeight = 0
                JambSupport.CL = FO.rEdgePosition
                If LoadBearingJamb = "Right" Then
                    JambSupport.LoadBearing = True
                    NonExpandableFOJambs = JambSupport.CL
                End If
                JambSupport.tEdgeHeight = rGtob
                JambSupport.Length = rGtob
                JambSupport.SetSize b, "Column", eWall, 30
                JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                WallColumns.Add JambSupport
            End If
            Jamb.Length = Jamb.tEdgeHeight
            Jamb.Size = "8"" Receiver Cee"
            Jamb.Width = 2.5
            Jamb.CL = FO.rEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            If Jamb.LoadBearing = False Then
                FO.FOMaterials.Add Jamb
            Else
                WallColumns.Add Jamb
            End If
            Exit Function
        ElseIf FO.lEdgePosition >= StartPos + IdealSpan And FO.rEdgePosition <= StartPos + IdealSpan And _
            FO.FOType <> "PDoor" Then
            lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition)
            rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition)
            If FO.lEdgePosition < StartPos + MaxDistance Then
                LoadBearingJamb = "Left"
                Set Jamb = New clsMember
                Jamb.bEdgeHeight = 0
                If lGtob < 30 * 12 + 4 Then 'don't need jamb support
                    Jamb.tEdgeHeight = lGtob
                    If LoadBearingJamb = "Left" Then
                        Jamb.LoadBearing = True
                        NonExpandableFOJambs = FO.lEdgePosition
                    End If
                Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = 30 * 12
                    Set JambSupport = New clsMember
                    JambSupport.bEdgeHeight = 0
                    JambSupport.CL = FO.lEdgePosition
                    If LoadBearingJamb = "Left" Then
                        JambSupport.LoadBearing = True
                        NonExpandableFOJambs = JambSupport.CL
                    End If
                    JambSupport.tEdgeHeight = lGtob
                    JambSupport.Length = lGtob
                    JambSupport.SetSize b, "Column", eWall, 30
                    JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                    WallColumns.Add JambSupport
                End If
                Jamb.Length = Jamb.tEdgeHeight
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.CL = FO.lEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                If Jamb.LoadBearing = False Then
                    FO.FOMaterials.Add Jamb
                Else
                    WallColumns.Add Jamb
                End If
            Else
                LoadBearingJamb = "Right"
                Set Jamb = New clsMember
                Jamb.bEdgeHeight = 0
                If rGtob < 30 * 12 + 4 Then 'don't need jamb support
                    Jamb.tEdgeHeight = rGtob
                    Jamb.LoadBearing = True
                    NonExpandableFOJambs = FO.rEdgePosition
                Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = 30 * 12
                    Set JambSupport = New clsMember
                    JambSupport.bEdgeHeight = 0
                    JambSupport.LoadBearing = True
                    JambSupport.tEdgeHeight = rGtob
                    JambSupport.Length = rGtob
                    JambSupport.CL = FO.rEdgePosition
                    JambSupport.SetSize b, "Column", eWall, 30
                    JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                    WallColumns.Add JambSupport
                    NonExpandableFOJambs = JambSupport.CL
                End If
                Jamb.Length = Jamb.tEdgeHeight
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.CL = FO.rEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                If Jamb.LoadBearing = False Then
                    FO.FOMaterials.Add Jamb
                Else
                    WallColumns.Add Jamb
                End If
            End If
            Exit Function
        Else
            NonExpandableFOJambs = StartPos
        End If
    Next FO
Else
    If StartPos < MaxDistance Then
        NonExpandableFOJambs = FirstCorner
        Exit Function
    End If
    For Each FO In FOs
        'if OHDoor or MiscFO w/ full height jambs is inside max distance and the furthest edge is at least within 5' of max distance,
        'then one of the jambs should be load bearing
        If FO.lEdgePosition < StartPos And FO.lEdgePosition > StartPos - MaxDistance And _
            FO.rEdgePosition < StartPos - MaxDistance + 60 And _
            (FO.FOType = "OHDoor" Or FO.StructuralSteelOption = "Full Height Jambs w/ Header & Stool") Then
            
            If FO.rEdgePosition > StartPos - MaxDistance Then 'right edge should be used as load bearing
                LoadBearingJamb = "Right"
            Else
                LoadBearingJamb = "Left"
            End If
            
            lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition)
            rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition)
            'left jamb
            Set Jamb = New clsMember
            Jamb.bEdgeHeight = 0
            If lGtob < 30 * 12 + 4 Then 'don't need jamb support
                Jamb.tEdgeHeight = lGtob
                If LoadBearingJamb = "Left" Then
                    Jamb.LoadBearing = True
                    NonExpandableFOJambs = FO.lEdgePosition
                End If
            Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                Jamb.tEdgeHeight = 30 * 12
                Set JambSupport = New clsMember
                JambSupport.bEdgeHeight = 0
                JambSupport.CL = FO.lEdgePosition
                If LoadBearingJamb = "Left" Then
                    JambSupport.LoadBearing = True
                    NonExpandableFOJambs = JambSupport.CL
                End If
                JambSupport.tEdgeHeight = lGtob
                JambSupport.Length = lGtob
                JambSupport.SetSize b, "Column", eWall, 30
                JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                WallColumns.Add JambSupport
            End If
            Jamb.Length = Jamb.tEdgeHeight
            Jamb.Size = "8"" Receiver Cee"
            Jamb.Width = 2.5
            Jamb.CL = FO.lEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            If Jamb.LoadBearing = False Then
                FO.FOMaterials.Add Jamb
            Else
                WallColumns.Add Jamb
            End If
            
            'right jamb
            Set Jamb = New clsMember
            Jamb.bEdgeHeight = 0
            If rGtob < 30 * 12 + 4 Then 'don't need jamb support
                Jamb.tEdgeHeight = rGtob
                If LoadBearingJamb = "Right" Then
                    Jamb.LoadBearing = True
                    NonExpandableFOJambs = FO.rEdgePosition
                End If
            Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                Jamb.tEdgeHeight = 30 * 12
                Set JambSupport = New clsMember
                JambSupport.bEdgeHeight = 0
                JambSupport.CL = FO.rEdgePosition
                If LoadBearingJamb = "Right" Then
                    JambSupport.LoadBearing = True
                    NonExpandableFOJambs = JambSupport.CL
                End If
                JambSupport.tEdgeHeight = rGtob
                JambSupport.Length = rGtob
                JambSupport.SetSize b, "Column", eWall, 30
                JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                WallColumns.Add JambSupport
            End If
            Jamb.Length = Jamb.tEdgeHeight
            Jamb.Size = "8"" Receiver Cee"
            Jamb.Width = 2.5
            Jamb.CL = FO.rEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            If Jamb.LoadBearing = False Then
                FO.FOMaterials.Add Jamb
            Else
                WallColumns.Add Jamb
            End If
            Exit Function
        ElseIf FO.lEdgePosition >= StartPos - IdealSpan And FO.rEdgePosition <= StartPos - IdealSpan And _
            FO.FOType <> "PDoor" Then
            lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition)
            rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition)
            If FO.rEdgePosition < StartPos - MaxDistance Then
                LoadBearingJamb = "Left"
                Set Jamb = New clsMember
                Jamb.bEdgeHeight = 0
                If lGtob < 30 * 12 + 4 Then 'don't need jamb support
                    Jamb.tEdgeHeight = lGtob
                    If LoadBearingJamb = "Left" Then
                        Jamb.LoadBearing = True
                        NonExpandableFOJambs = FO.lEdgePosition
                    End If
                Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = 30 * 12
                    Set JambSupport = New clsMember
                    JambSupport.bEdgeHeight = 0
                    JambSupport.CL = FO.lEdgePosition
                    If LoadBearingJamb = "Left" Then
                        JambSupport.LoadBearing = True
                        NonExpandableFOJambs = JambSupport.CL
                    End If
                    JambSupport.tEdgeHeight = lGtob
                    JambSupport.Length = lGtob
                    JambSupport.SetSize b, "Column", eWall, 30
                    JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                    WallColumns.Add JambSupport
                End If
                Jamb.Length = Jamb.tEdgeHeight
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.CL = FO.lEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                If Jamb.LoadBearing = False Then
                    FO.FOMaterials.Add Jamb
                Else
                    WallColumns.Add Jamb
                End If
            Else
                LoadBearingJamb = "Right"
                Set Jamb = New clsMember
                Jamb.bEdgeHeight = 0
                If rGtob < 30 * 12 + 4 Then 'don't need jamb support
                    Jamb.tEdgeHeight = rGtob
                    Jamb.LoadBearing = True
                    NonExpandableFOJambs = FO.rEdgePosition
                Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = 30 * 12
                    Set JambSupport = New clsMember
                    JambSupport.bEdgeHeight = 0
                    JambSupport.LoadBearing = True
                    JambSupport.tEdgeHeight = rGtob
                    JambSupport.Length = rGtob
                    JambSupport.CL = FO.rEdgePosition
                    JambSupport.SetSize b, "Column", eWall, 30
                    JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                    WallColumns.Add JambSupport
                    NonExpandableFOJambs = JambSupport.CL
                End If
                Jamb.Length = Jamb.tEdgeHeight
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.CL = FO.rEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                If Jamb.LoadBearing = False Then
                    FO.FOMaterials.Add Jamb
                Else
                    WallColumns.Add Jamb
                End If
            End If
            Exit Function
        Else
            NonExpandableFOJambs = StartPos
        End If
    Next FO
End If
   
End Function

'''''''''''''' Adds FO Jambs for FOs without a wall location
'Window Jambs will default to 7'2" Jambs w/ Header and Stool
Sub FieldLocateFOCalc(b As clsBuilding)

Dim FO As clsFO
Dim Jamb As clsMember
Dim Purlin As clsMember

For Each FO In b.fieldlocateFOs
    If FO.FOType = "PDoor" Then
        'Do Nothing - no additional steel for  PDoors
    ElseIf FO.FOType = "Window" Then
        'Add left and right jamb
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = "8"" Receiver Cee"
        Jamb.Width = 2.5
        Jamb.Length = 86
        Jamb.mType = "FO Receiver Jamb"
        Jamb.Placement = "FO Jamb"
        Jamb.Qty = 1
        FO.FOMaterials.Add Jamb
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = "8"" Receiver Cee"
        Jamb.Width = 2.5
        Jamb.Length = 86
        Jamb.mType = "FO Receiver Jamb"
        Jamb.Placement = "FO Jamb"
        Jamb.Qty = 1
        FO.FOMaterials.Add Jamb
        'Add Header and Stool
        'Header
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = "8"" C Purlin"
        Jamb.Width = 2.5
        Jamb.Length = FO.Width
        Jamb.mType = "FO Header"
        Jamb.Placement = Jamb.mType
        FO.FOMaterials.Add Jamb
        'Stool
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = "8"" C Purlin"
        Jamb.Width = 2.5
        Jamb.Length = FO.Width
        Jamb.mType = "FO Stool"
        Jamb.Placement = Jamb.mType
        FO.FOMaterials.Add Jamb
        
        b.WeldClips = b.WeldClips + 6
    End If
Next FO

End Sub

'''''''''''''' Adds FO Jambs if not already added in Column Calc
Sub FOJambsCalc(b As clsBuilding, eWall As String)
Dim FO As clsFO
Dim Column As clsMember
Dim AllColumnsValid As Boolean
Dim ColumnCollection As Collection
Dim FOCollection As Collection
Dim ColIndex As Integer
Dim Jamb As clsMember
Dim Purlin As clsMember
Dim MiscFOColumnReplacement As Boolean
Dim OHDoorColumnReplacement As Boolean
Dim WindowColumnReplacement As Boolean
Dim Member As clsMember
Dim WeldClips As clsMiscItem
Dim ReplacedColLocation As Double
Dim RightJambExists As Boolean
Dim LeftJambExists As Boolean
Dim RightSupportExists As Boolean
Dim LeftSupportExists As Boolean
Dim rGtob As Double
Dim lGtob As Double

Select Case eWall
Case "e1"
    Set ColumnCollection = b.e1Columns
    Set FOCollection = b.e1FOs
Case "s2"
    Set ColumnCollection = b.s2Columns
    Set FOCollection = b.s2FOs
Case "e3"
    Set ColumnCollection = b.e3Columns
    Set FOCollection = b.e3FOs
Case "s4"
    Set ColumnCollection = b.s4Columns
    Set FOCollection = b.s4FOs
End Select

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Endwalls
If eWall = "e1" Or eWall = "e3" Then
    For Each FO In FOCollection
    Select Case FO.FOType
    Case "OHDoor"
        'Check if Right Jamb already Exists
        RightJambExists = False
        For Each Member In FO.FOMaterials
            If Member.CL = FO.rEdgePosition Then
                RightJambExists = True
            End If
        Next Member
        For Each Member In ColumnCollection
            If Member.CL = FO.rEdgePosition And Member.LoadBearing = False Then
                RightJambExists = True
            End If
        Next Member
        'Check if Left Jamb already Exists
        LeftJambExists = False
        For Each Member In FO.FOMaterials
            If Member.CL = FO.rEdgePosition Then
                LeftJambExists = True
            End If
        Next Member
        For Each Member In ColumnCollection
            If Member.CL = FO.lEdgePosition And Member.LoadBearing = False Then
                LeftJambExists = True
            End If
        Next Member
        'if Right Jamb doesn't exist, create jamb
        If RightJambExists = False Then
            rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition)
            If rGtob <= 30 * 12 + 4 Then
                'create full height jamb
                Set Jamb = New clsMember
                Jamb.CL = FO.rEdgePosition
                Jamb.bEdgeHeight = 0
                Jamb.tEdgeHeight = rGtob
                Jamb.Length = rGtob
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            Else
                'create 30'4" jamb
                Set Jamb = New clsMember
                Jamb.CL = FO.rEdgePosition
                Jamb.bEdgeHeight = 0
                Jamb.tEdgeHeight = 30 * 12 + 4
                Jamb.Length = 30 * 12 + 4
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
                'Check if load bearing column already exists
                RightSupportExists = False
                For Each Column In ColumnCollection
                    If Column.CL >= FO.rEdgePosition - 12 And Column.CL <= FO.rEdgePosition + 12 Then
                        RightSupportExists = True
                    End If
                Next Column
                
                If RightSupportExists = True Then
                    'Do Nothing
                Else
                    Set Jamb = New clsMember
                    Jamb.CL = FO.rEdgePosition
                    Jamb.bEdgeHeight = 0
                    Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL)
                    Jamb.Length = Jamb.tEdgeHeight
                    Jamb.SetSize b, "Column", eWall, 30, "NonExpandable"
                    Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                    Jamb.mType = "FO Support Jamb"
                    FO.FOMaterials.Add Jamb
                End If
            End If
        End If
        'if Left Jamb doesn't exist, create jamb
        If LeftJambExists = False Then
            lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition)
            If lGtob <= 30 * 12 + 4 Then
                'create full height jamb
                Set Jamb = New clsMember
                Jamb.CL = FO.lEdgePosition
                Jamb.bEdgeHeight = 0
                Jamb.tEdgeHeight = lGtob
                Jamb.Length = lGtob
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            Else
                'create 30'4" jamb
                Set Jamb = New clsMember
                Jamb.CL = FO.lEdgePosition
                Jamb.bEdgeHeight = 0
                Jamb.tEdgeHeight = 30 * 12 + 4
                Jamb.Length = 30 * 12 + 4
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
                'Check if load bearing column already exists
                LeftSupportExists = False
                For Each Column In ColumnCollection
                    If Column.CL >= FO.lEdgePosition - 12 And Column.CL <= FO.lEdgePosition + 12 Then
                        LeftSupportExists = True
                    End If
                Next Column
                
                If LeftSupportExists = True Then
                    'Do Nothing
                Else
                    Set Jamb = New clsMember
                    Jamb.CL = FO.lEdgePosition
                    Jamb.bEdgeHeight = 0
                    Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL)
                    Jamb.Length = Jamb.tEdgeHeight
                    Jamb.SetSize b, "Column", eWall, 30, "NonExpandable"
                    Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                    Jamb.mType = "FO Support Jamb"
                    FO.FOMaterials.Add Jamb
                End If
            End If
        End If
        'Create Header
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = ("8"" C Purlin")
        Jamb.Width = 2.5
        Jamb.rEdgePosition = FO.rEdgePosition
        Jamb.Length = FO.lEdgePosition - FO.rEdgePosition
        Jamb.tEdgeHeight = FO.tEdgeHeight
        Jamb.CL = 0
        Jamb.mType = "FO Header"
        FO.FOMaterials.Add Jamb
    Case "Window"
        'Check if Right Jamb already Exists
        RightJambExists = False
        For Each Member In FO.FOMaterials
            If Member.CL = FO.rEdgePosition Then
                RightJambExists = True
            End If
        Next Member
        For Each Member In ColumnCollection
            If Member.CL = FO.rEdgePosition And Member.LoadBearing = False Then
                RightJambExists = True
            End If
        Next Member
        'Check if Left Jamb already Exists
        LeftJambExists = False
        For Each Member In FO.FOMaterials
            If Member.CL = FO.rEdgePosition Then
                LeftJambExists = True
            End If
        Next Member
        For Each Member In ColumnCollection
            If Member.CL = FO.lEdgePosition And Member.LoadBearing = False Then
                LeftJambExists = True
            End If
        Next Member
        'if Right Jamb doesn't exist, create jamb
        If RightJambExists = False Then
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = "8"" Receiver Cee"
            Jamb.Width = 2.5
            Jamb.CL = FO.rEdgePosition
            Jamb.tEdgeHeight = FO.tEdgeHeight
            Jamb.bEdgeHeight = FO.bEdgeHeight
            Jamb.Length = Jamb.tEdgeHeight - Jamb.bEdgeHeight
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.mType = "FO Receiver Jamb"
            FO.FOMaterials.Add Jamb
        End If
        'if Right Jamb doesn't exist, create jamb
        If LeftJambExists = False Then
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = "8"" Receiver Cee"
            Jamb.Width = 2.5
            Jamb.CL = FO.lEdgePosition
            Jamb.tEdgeHeight = FO.tEdgeHeight
            Jamb.bEdgeHeight = FO.bEdgeHeight
            Jamb.Length = Jamb.tEdgeHeight - Jamb.bEdgeHeight
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.mType = "FO Receiver Jamb"
            FO.FOMaterials.Add Jamb
        End If
        'Add Header and Stool
        'Header
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = "8"" C Purlin"
        Jamb.Width = 2.5
        Jamb.rEdgePosition = FO.rEdgePosition
        Jamb.Length = FO.Width
        Jamb.tEdgeHeight = FO.tEdgeHeight
        Jamb.bEdgeHeight = FO.bEdgeHeight
        Jamb.mType = "FO Header"
        FO.FOMaterials.Add Jamb
        'Stool
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = "8"" C Purlin"
        Jamb.Width = 2.5
        Jamb.rEdgePosition = FO.rEdgePosition
        Jamb.Length = FO.Width
        Jamb.tEdgeHeight = FO.bEdgeHeight
        Jamb.bEdgeHeight = FO.bEdgeHeight
        Jamb.mType = "FO Stool"
        FO.FOMaterials.Add Jamb
    Case "MiscFO"
        'Check if Right Jamb already Exists
        RightJambExists = False
        For Each Member In FO.FOMaterials
            If Member.CL = FO.rEdgePosition Then
                RightJambExists = True
            End If
        Next Member
        For Each Member In ColumnCollection
            If Member.CL = FO.rEdgePosition And Member.LoadBearing = False Then
                RightJambExists = True
            End If
        Next Member
        'Check if Left Jamb already Exists
        LeftJambExists = False
        For Each Member In FO.FOMaterials
            If Member.CL = FO.rEdgePosition Then
                LeftJambExists = True
            End If
        Next Member
        For Each Member In ColumnCollection
            If Member.CL = FO.lEdgePosition And Member.LoadBearing = False Then
                LeftJambExists = True
            End If
        Next Member
        Select Case FO.StructuralSteelOption
        Case "Full Height Jambs w/ Header & Stool"
            'if Right Jamb doesn't exist, create jamb
            If RightJambExists = False Then
                rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition)
                If rGtob <= 30 * 12 + 4 Then
                    'create full height jamb
                    Set Jamb = New clsMember
                    Jamb.CL = FO.rEdgePosition
                    Jamb.bEdgeHeight = 0
                    Jamb.tEdgeHeight = rGtob
                    Jamb.Length = rGtob
                    Jamb.Size = "8"" Receiver Cee"
                    Jamb.Width = 2.5
                    Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                    Jamb.mType = "FO Receiver Jamb"
                    FO.FOMaterials.Add Jamb
                Else
                    'create FO sized jamb w/ support
                    Set Jamb = New clsMember
                    Jamb.CL = FO.rEdgePosition
                    Jamb.bEdgeHeight = FO.bEdgeHeight
                    Jamb.tEdgeHeight = FO.tEdgeHeight
                    Jamb.Length = Jamb.tEdgeHeight - Jamb.bEdgeHeight
                    Jamb.Size = "8"" Receiver Cee"
                    Jamb.Width = 2.5
                    Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                    Jamb.mType = "FO Receiver Jamb"
                    FO.FOMaterials.Add Jamb
                    'Check if load bearing column already exists
                    RightSupportExists = False
                    For Each Column In ColumnCollection
                        If Column.CL >= FO.rEdgePosition - 12 And Column.CL <= FO.rEdgePosition + 12 Then
                            RightSupportExists = True
                        End If
                    Next Column
                    
                    If RightSupportExists = True Then
                        'Do Nothing
                    Else
                        Set Jamb = New clsMember
                        Jamb.CL = FO.rEdgePosition
                        Jamb.bEdgeHeight = 0
                        Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL)
                        Jamb.Length = Jamb.tEdgeHeight
                        Jamb.SetSize b, "Column", eWall, 30, "NonExpandable"
                        Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                        Jamb.mType = "FO Support Jamb"
                        FO.FOMaterials.Add Jamb
                    End If
                End If
            End If
            'if Left Jamb doesn't exist, create jamb
            If LeftJambExists = False Then
                lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition)
                If lGtob <= 30 * 12 + 4 Then
                    'create full height jamb
                    Set Jamb = New clsMember
                    Jamb.CL = FO.lEdgePosition
                    Jamb.bEdgeHeight = 0
                    Jamb.tEdgeHeight = lGtob
                    Jamb.Length = lGtob
                    Jamb.Size = "8"" Receiver Cee"
                    Jamb.Width = 2.5
                    Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                    Jamb.mType = "FO Receiver Jamb"
                    FO.FOMaterials.Add Jamb
                Else
                    'create FO sized jamb w/ support
                    Set Jamb = New clsMember
                    Jamb.CL = FO.lEdgePosition
                    Jamb.bEdgeHeight = FO.bEdgeHeight
                    Jamb.tEdgeHeight = FO.tEdgeHeight
                    Jamb.Length = Jamb.tEdgeHeight - Jamb.bEdgeHeight
                    Jamb.Size = "8"" Receiver Cee"
                    Jamb.Width = 2.5
                    Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                    Jamb.mType = "FO Receiver Jamb"
                    FO.FOMaterials.Add Jamb
                    'Check if load bearing column already exists
                    LeftSupportExists = False
                    For Each Column In ColumnCollection
                        If Column.CL >= FO.lEdgePosition - 12 And Column.CL <= FO.lEdgePosition + 12 Then
                            LeftSupportExists = True
                        End If
                    Next Column
                    
                    If LeftSupportExists = True Then
                        'Do Nothing
                    Else
                        Set Jamb = New clsMember
                        Jamb.CL = FO.lEdgePosition
                        Jamb.bEdgeHeight = 0
                        Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL)
                        Jamb.Length = Jamb.tEdgeHeight
                        Jamb.SetSize b, "Column", eWall, 30, "NonExpandable"
                        Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                        Jamb.mType = "FO Support Jamb"
                        FO.FOMaterials.Add Jamb
                    End If
                End If
            End If
            'header and stool
            Set Purlin = New clsMember
            Purlin.Length = FO.Width
            Purlin.Size = ("8"" C Purlin")
            Purlin.mType = "FO Material"
            If FO.tEdgeHeight = 86 Then
                Purlin.Qty = 1
            Else
                Purlin.Qty = 2
            End If
            Purlin.mType = "FO Header"
            FO.FOMaterials.Add Purlin
        Case "7'2"" Jambs w/ Header & Stool"
            If RightJambExists = False Then
                'add jambs if they weren't already added as a column replacement
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.rEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = 2 + (7 * 12)
                Jamb.tEdgeHeight = 2 + (7 * 12)
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            End If
            If LeftJambExists = False Then
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.lEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = 2 + (7 * 12)
                Jamb.tEdgeHeight = 2 + (7 * 12)
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            End If
            'header and stool
            Set Purlin = New clsMember
            Purlin.Length = FO.Width
            Purlin.Size = ("8"" C Purlin")
            Purlin.mType = "FO Material"
            If FO.tEdgeHeight = 86 Then
                Purlin.Qty = 1
            Else
                Purlin.Qty = 2
            End If
            Purlin.mType = "FO Header"
            FO.FOMaterials.Add Purlin
        Case "7'2"" Jambs w/ Stool"
            If RightJambExists = False Then
                'add jambs if they weren't already added as a column replacement
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.rEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = 2 + (7 * 12)
                Jamb.tEdgeHeight = 2 + (7 * 12)
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            End If
            If LeftJambExists = False Then
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.lEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = 2 + (7 * 12)
                Jamb.tEdgeHeight = 2 + (7 * 12)
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            End If
            'stool
            Set Purlin = New clsMember
            Purlin.Length = FO.Width
            Purlin.Size = ("8"" C Purlin")
            Purlin.mType = "FO Material"
            If FO.tEdgeHeight = 86 Then
                Purlin.Qty = 1
            Else
                Purlin.Qty = 2
            End If
            Purlin.mType = "FO Header"
            FO.FOMaterials.Add Purlin
        Case "7'2"" Jambs"
            If RightJambExists = False Then
                'add jambs if they weren't already added as a column replacement
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.rEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = 2 + (7 * 12)
                Jamb.tEdgeHeight = 2 + (7 * 12)
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            End If
            If LeftJambExists = False Then
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.lEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = 2 + (7 * 12)
                Jamb.tEdgeHeight = 2 + (7 * 12)
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            End If
        Case "5' Jambs w/ Header & Stool"
            If RightJambExists = False Then
                'add jambs if they weren't already added as a column replacement
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.rEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = 5 * 12
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            End If
            If LeftJambExists = False Then
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.lEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = 5 * 12
                Jamb.mType = "FO Receiver Jamb"
                FO.FOMaterials.Add Jamb
            End If
            'header and stool
            Set Purlin = New clsMember
            Purlin.Length = FO.Width
            Purlin.Size = ("8"" C Purlin")
            Purlin.mType = "FO Material"
            If FO.tEdgeHeight = 86 Then
                Purlin.Qty = 1
            Else
                Purlin.Qty = 2
            End If
            Purlin.mType = "FO Header"
            FO.FOMaterials.Add Purlin
        End Select
    End Select
    Next FO
End If
            
            
        
                    
        







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Sidewalls
If eWall = "s2" Or eWall = "s4" Then
    For Each FO In FOCollection
    Select Case FO.FOType
    Case "pDoor"
        '''' check column placement
        For Each Column In ColumnCollection
            'if within 6" of column CL
             If Column.CL > (FO.rEdgePosition - 6) And Column.CL < (FO.lEdgePosition + 6) Then
                'error condition
                If MsgBox(FO.Description & " has been found to be located within 6"" of a column! Relocate this personnel door before proceeding. Continue anyways?", vbOKCancel, "FO Placement Error") = 7 Then
                    End
                Else
                    Exit Sub
                End If
            End If
        Next Column
    Case "OHDoor"
        '''' check for intersecting/approaching columns
        For ColIndex = ColumnCollection.Count To 1 Step -1
            Set Column = ColumnCollection(ColIndex)
            'if within 1' of a column CL
            If (Column.CL > FO.rEdgePosition - (1 * 12) And Column.CL < FO.lEdgePosition + (1 * 12)) Then
                'error condition
                If MsgBox(FO.Description & " has been found to be located within 1' of a column! Relocate or resize this overhead door before proceeding. Continue anyways?", vbYesNo + vbCritical, "FO Placement Error") = 7 Then
                    End
                Else
                    'Exit Sub
                End If
            End If
        Next ColIndex
        'create jambs
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = ("8"" Receiver Cee")
        Jamb.Width = 2.5
        Jamb.CL = FO.rEdgePosition
        Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
        Jamb.Length = b.DistanceToRoof(eWall, Jamb.CL)
        Jamb.tEdgeHeight = Jamb.Length
        Jamb.bEdgeHeight = 0
        FO.FOMaterials.Add Jamb
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = ("8"" Receiver Cee")
        Jamb.Width = 2.5
        Jamb.CL = FO.lEdgePosition
        Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
        Jamb.Length = b.DistanceToRoof(eWall, Jamb.CL)
        Jamb.tEdgeHeight = Jamb.Length
        Jamb.bEdgeHeight = 0
        FO.FOMaterials.Add Jamb
        'OHDoor Header'''''''''''''''''''''''''''''''''
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = ("8"" C Purlin")
        Jamb.Width = 2.5
        Jamb.rEdgePosition = FO.rEdgePosition
        Jamb.Length = FO.lEdgePosition - FO.rEdgePosition
        Jamb.tEdgeHeight = FO.tEdgeHeight
        Jamb.CL = 0
        FO.FOMaterials.Add Jamb
    Case "Window"
        '''' check for intersecting columns
        For ColIndex = ColumnCollection.Count To 1 Step -1
            Set Column = ColumnCollection(ColIndex)
            'if within 1' of a column CL
            If Column.CL > FO.rEdgePosition And Column.CL < FO.lEdgePosition Then
                'error condition
                If MsgBox(FO.Description & " has been found to intersect a sidewall column! Relocate or resize this window before proceeding. Continue anyways?", vbYesNo + vbCritical, "FO Placement Error") = 7 Then
                    End
                Else
                    'Exit Sub
                End If
            End If
        Next ColIndex
        'create jambs
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = ("8"" Receiver Cee")
        Jamb.Width = 2.5
        Jamb.CL = FO.rEdgePosition
        Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
        Jamb.Length = FO.Height
        Jamb.tEdgeHeight = FO.tEdgeHeight
        Jamb.bEdgeHeight = FO.bEdgeHeight
        FO.FOMaterials.Add Jamb
        Set Jamb = New clsMember
        Jamb.mType = "FO Material"
        Jamb.Size = ("8"" Receiver Cee")
        Jamb.Width = 2.5
        Jamb.CL = FO.lEdgePosition
        Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
        Jamb.Length = FO.Height
        Jamb.tEdgeHeight = FO.tEdgeHeight
        Jamb.bEdgeHeight = FO.bEdgeHeight
        FO.FOMaterials.Add Jamb
    Case "MiscFO"
        '''' check for intersecting/approaching columns
        For ColIndex = ColumnCollection.Count To 1 Step -1
            Set Column = ColumnCollection(ColIndex)
            'if within 6" of a column CL
            If Column.CL > FO.rEdgePosition - (1 * 6) And Column.CL < FO.lEdgePosition + (1 * 6) Then
                'error condition
                If MsgBox(FO.Description & " has been found to be intersecting a sidewall column! Relocate or resize this misc. FO opening before proceeding. Continue anyways?", vbYesNo + vbCritical, "FO Placement Error") = 7 Then
                    End
                Else
                    'Exit Sub
                End If
            End If
        Next ColIndex
        'create jambs
                'add structural steel depending on options selected in input field
        Select Case FO.StructuralSteelOption
        Case "Full Height Jambs w/ Header & Stool"
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.rEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = b.DistanceToRoof(eWall, Jamb.CL)
                Jamb.tEdgeHeight = Jamb.Length
                Jamb.bEdgeHeight = 0
                FO.FOMaterials.Add Jamb
                Set Jamb = New clsMember
                Jamb.mType = "FO Material"
                Jamb.Size = ("8"" Receiver Cee")
                Jamb.Width = 2.5
                Jamb.CL = FO.lEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                Jamb.Length = b.DistanceToRoof(eWall, Jamb.CL)
                Jamb.tEdgeHeight = Jamb.Length
                Jamb.bEdgeHeight = 0
                FO.FOMaterials.Add Jamb
            'header and stool
            Set Purlin = New clsMember
            Purlin.Length = FO.Width
            Purlin.Size = ("8"" C Purlin")
            Purlin.mType = "FO Material"
            If FO.tEdgeHeight = 86 Then
                Purlin.Qty = 1
            Else
                Purlin.Qty = 2
            End If
            FO.FOMaterials.Add Purlin
            'weld clips
            'b.WeldClips = b.WeldClips + 10
            'Set WeldClips = New clsMiscItem
            'WeldClips.Quantity = 10
            'WeldClips.Name = "Weld Clips"
            'FO.FOMaterials.Add WeldClips
        Case "7'2"" Jambs w/ Header & Stool"
            'add jambs if they weren't already added as a column replacement
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = ("8"" Receiver Cee")
            Jamb.Width = 2.5
            Jamb.CL = FO.rEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.Length = 2 + (7 * 12)
            Jamb.tEdgeHeight = 2 + (7 * 12)
            FO.FOMaterials.Add Jamb
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = ("8"" Receiver Cee")
            Jamb.Width = 2.5
            Jamb.CL = FO.lEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.Length = 2 + (7 * 12)
            Jamb.tEdgeHeight = 2 + (7 * 12)
            FO.FOMaterials.Add Jamb
            'header and stool
            Set Purlin = New clsMember
            Purlin.Length = FO.Width
            Purlin.Size = ("8"" C Purlin")
            Purlin.mType = "FO Material"
            If FO.tEdgeHeight = 86 Then
                Purlin.Qty = 1
            Else
                Purlin.Qty = 2
            End If
            FO.FOMaterials.Add Purlin
            'weld clips
            'b.WeldClips = b.WeldClips + 6
            'Set WeldClips = New clsMiscItem
            'WeldClips.Quantity = 6
            'WeldClips.Name = "Weld Clips"
            'FO.FOMaterials.Add WeldClips
        Case "7'2"" Jambs w/ Stool"
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = ("8"" Receiver Cee")
            Jamb.Width = 2.5
            Jamb.CL = FO.rEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.Length = 2 + (7 * 12)
            Jamb.tEdgeHeight = 2 + (7 * 12)
            FO.FOMaterials.Add Jamb
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = ("8"" Receiver Cee")
            Jamb.Width = 2.5
            Jamb.CL = FO.lEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.Length = 2 + (7 * 12)
            Jamb.tEdgeHeight = 2 + (7 * 12)
            FO.FOMaterials.Add Jamb
            'stool
            Set Purlin = New clsMember
            Purlin.Length = FO.Width
            Purlin.Size = ("8"" C Purlin")
            Purlin.mType = "FO Material"
            If FO.tEdgeHeight = 86 Then
                Purlin.Qty = 1
            Else
                Purlin.Qty = 2
            End If
            FO.FOMaterials.Add Purlin
            'weld clips
            'b.WeldClips = b.WeldClips + 4
            'Set WeldClips = New clsMiscItem
            'WeldClips.Quantity = 4
            'WeldClips.Name = "Weld Clips"
            'FO.FOMaterials.Add WeldClips
        Case "7'2"" Jambs"
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = ("8"" Receiver Cee")
            Jamb.Width = 2.5
            Jamb.CL = FO.rEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.Length = 2 + (7 * 12)
            Jamb.tEdgeHeight = 2 + (7 * 12)
            FO.FOMaterials.Add Jamb
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = ("8"" Receiver Cee")
            Jamb.Width = 2.5
            Jamb.CL = FO.lEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.Length = 2 + (7 * 12)
            Jamb.tEdgeHeight = 2 + (7 * 12)
            FO.FOMaterials.Add Jamb
            'weld clips
            'b.WeldClips = b.WeldClips + 2
            'Set WeldClips = New clsMiscItem
            'WeldClips.Quantity = 2
            'WeldClips.Name = "Weld Clips"
            'FO.FOMaterials.Add WeldClips
        Case "5' Jambs w/ Header & Stool"
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = ("8"" Receiver Cee")
            Jamb.Width = 2.5
            Jamb.CL = FO.rEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.Length = 5 * 12
            FO.FOMaterials.Add Jamb
            Set Jamb = New clsMember
            Jamb.mType = "FO Material"
            Jamb.Size = ("8"" Receiver Cee")
            Jamb.Width = 2.5
            Jamb.CL = FO.lEdgePosition
            Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
            Jamb.Length = 5 * 12
            FO.FOMaterials.Add Jamb
            'header and stool
            Set Purlin = New clsMember
            Purlin.Length = FO.Width
            Purlin.Size = ("8"" C Purlin")
            Purlin.mType = "FO Material"
            If FO.tEdgeHeight = 86 Then
                Purlin.Qty = 1
            Else
                Purlin.Qty = 2
            End If
            FO.FOMaterials.Add Purlin
            'weld clips
            'b.WeldClips = b.WeldClips + 8
            'Set WeldClips = New clsMiscItem
            'WeldClips.Quantity = 8
            'WeldClips.Name = "Weld Clips"
            'FO.FOMaterials.Add WeldClips
        End Select
    End Select
    Next FO
End If
                



End Sub

''''''''''''''''''''' Sub used to sort arrays in ascending order. Currently only used for Endwall centerline calc as of 10/4/2021 at 8:00 PM EST --------------
Sub QuickSort(arr As Variant, first As Long, last As Long)
  Dim vCentreVal As Variant, vTemp As Variant
  Dim lTempLow As Long
  Dim lTempHi As Long
  
  lTempLow = first
  lTempHi = last
  vCentreVal = arr((first + last) \ 2)
  Do While lTempLow <= lTempHi
    Do While arr(lTempLow) < vCentreVal And lTempLow < last
      lTempLow = lTempLow + 1
    Loop
    Do While vCentreVal < arr(lTempHi) And lTempHi > first
      lTempHi = lTempHi - 1
    Loop
    If lTempLow <= lTempHi Then
        ' Swap values
        vTemp = arr(lTempLow)
        arr(lTempLow) = arr(lTempHi)
        arr(lTempHi) = vTemp
        ' Move to next positions
        lTempLow = lTempLow + 1
        lTempHi = lTempHi - 1
    End If
  Loop
  If first < lTempHi Then QuickSort arr, first, lTempHi
  If lTempLow < last Then QuickSort arr, lTempLow, last
  
End Sub

Public Sub ReverseArray(vArray As Variant)
'Reverse the order of an array, so if it's already sorted
'from smallest to largest, it will now be sorted from
'largest to smallest.
Dim vTemp As Variant
Dim i As Long
Dim iUpper As Long
Dim iMidPt As Long
iUpper = UBound(vArray)
iMidPt = (UBound(vArray) - LBound(vArray)) \ 2 + LBound(vArray)
For i = LBound(vArray) To iMidPt
    vTemp = vArray(iUpper)
    vArray(iUpper) = vArray(i)
    vArray(i) = vTemp
    iUpper = iUpper - 1
Next i
End Sub

Sub NewExpandableEndwallColumnsGen(b As clsBuilding, eWall As String, EndwallColumnCLs() As Double, Optional NewColNum As Integer, Optional Reiterate As Boolean)

Dim ColNum As Integer
Dim MaxHorizontalDistance As Double
Dim ColLocation() As Double
Dim Column As clsMember
Dim DistanceToPreviousColumn As Double
Dim DistanceToNextColumn As Double
Dim i As Integer



MaxHorizontalDistance = (60 / (Sqr((b.rPitch / 12) ^ 2 + 1)))
'create new endwall columns like interior columns, then continue with these values
If NewColNum <> 0 Then
    ColNum = NewColNum
Else
    If b.rShape = "Gable" Then
        If b.bWidth <= 80 Then
            ColNum = 0
        ElseIf b.bWidth > 80 And b.bWidth < (MaxHorizontalDistance * 2) Then
            ColNum = 1
        ElseIf b.bWidth >= MaxHorizontalDistance * 2 Then
            ColNum = Application.WorksheetFunction.RoundUp(b.bWidth / MaxHorizontalDistance, 0) - 1
        End If
    ElseIf b.rShape = "Single Slope" Then
        If b.bWidth < MaxHorizontalDistance Then
            ColNum = 0
        ElseIf b.bWidth > MaxHorizontalDistance Then
            ColNum = Application.WorksheetFunction.RoundUp(b.bWidth / MaxHorizontalDistance, 0) - 1
        End If
    End If
    'lower Col Num by 1 on first iteration to check for marginal cases
    'some column widths (to be determined) will require less columns, this will check those cases
    If ColNum > 0 Then
        ColNum = ColNum - 1
    End If
End If


'first, evenly space columns along the width of the building to adjust later; add to array
ReDim ColLocation(ColNum + 1)
'includes s2 and s4 columns along rafter lines
ColLocation(0) = 0
ColLocation(ColNum + 1) = b.bWidth * 12
Select Case ColNum
Case 1
    ColLocation(1) = b.bWidth / 2 * 12
Case 2
    ColLocation(1) = b.bWidth / 3 * 12
    ColLocation(2) = b.bWidth / 3 * 12 * 2
Case 3
    ColLocation(1) = b.bWidth / 4 * 12
    ColLocation(2) = b.bWidth / 4 * 12 * 2
    ColLocation(3) = b.bWidth / 4 * 12 * 3
Case 4
    ColLocation(1) = b.bWidth / 5 * 12
    ColLocation(2) = b.bWidth / 5 * 12 * 2
    ColLocation(3) = b.bWidth / 5 * 12 * 3
    ColLocation(4) = b.bWidth / 5 * 12 * 4
End Select

'loop through array and check if columns conflict with OHDoors; if so, move 5' away from nearest edge
For i = 1 To ColNum
    If ConflictingEndwallOHDoor(ColLocation(i), b, eWall) = True Then
        ColLocation(i) = NearestEndwallLocation(ColLocation(i), b, , eWall)
    End If
Next i

'''''''''''''''check for No Interior Columns
If ColNum = 0 Then
    ''''''''''''''Distance between Columns
    DistanceToPreviousColumn = Abs(ColLocation(0) - ColLocation(1))

    ''''''''''''''Estimate COlumn widths
    'get first width
    Set Column = New clsMember
    Column.Length = b.DistanceToRoof("e1", ColLocation(0))
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", Abs(ColLocation(0) - ColLocation(1))
    'subtract half of first width
    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
    'get second width
    Set Column = New clsMember
    Column.Length = b.DistanceToRoof("e1", ColLocation(1))
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", Abs(ColLocation(0) - ColLocation(1))
    'subtract half of second width
    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
    
    If DistanceToPreviousColumn > (MaxHorizontalDistance * 12) Then
        Erase ColLocation
        Erase EndwallColumnCLs
        Call NewExpandableEndwallColumnsGen(b, eWall, EndwallColumnCLs(), ColNum + 1, True)
        Exit Sub
    End If
End If


'''''''''''''''check Interior Columns
'check that columns are no more than MaxHorizontalDistance ft apart since they may have been moved
For i = 1 To ColNum
    'get distance to next column to make sure it does NOT exceed max rafter length
    'if the two rafters stradle the center and the roof shape is "Gable", then go only to the center
    'estimate column widths to get accurate distances
    
    ''''''''''''''Distance to PREVIOUS Column
    If ColLocation(i) > (b.bWidth * 12 / 2) And ColLocation(i - 1) < (b.bWidth * 12 / 2) And b.rShape = "Gable" Then
        DistanceToPreviousColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
    Else
        DistanceToPreviousColumn = Abs(ColLocation(i) - ColLocation(i - 1))
    End If
        ''''''''''''''Estimate COlumn widths
        'get first width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
        'subtract half of width
        DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
        'get second width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i - 1))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
        'subtract width if sidewall column, or half of width otherwise
        If i - 1 = 0 Then
            DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
        Else
            DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
        End If
    
    ''''''''''''''Distance to NEXT Column
    If ColLocation(i) < b.bWidth * 12 / 2 And ColLocation(i + 1) > b.bWidth * 12 / 2 And b.rShape = "Gable" Then
        DistanceToNextColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
    Else
        DistanceToNextColumn = Abs(ColLocation(i) - ColLocation(i + 1))
    End If
        ''''''''''''''Estimate COlumn widths
        'get first width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
        'subtract half of width
        DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
        'get second width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i + 1))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
        'subtract width if sidewall column, or half of width otherwise
        If i + 1 = UBound(ColLocation()) Then
            DistanceToNextColumn = DistanceToNextColumn - Column.Width
        Else
            DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
        End If
    
    'check if the columns are too far apart; if so, run this sub again with 1 more column (optional parameter)
    If DistanceToPreviousColumn > (MaxHorizontalDistance * 12) Or DistanceToNextColumn > (MaxHorizontalDistance * 12) Then
        'Debug.Print "columns too far apart"
        'CHECK COLUMN DISTANCES AGAIN WITH NEW COLUMN WIDTH ESTIMATES
        If NearestEndwallLocation(ColLocation(i), b, "Alternate") <> ColLocation(i) Then
            ColLocation(i) = NearestEndwallLocation(ColLocation(i), b, "Alternate")
            ''''''''''''''Distance to PREVIOUS Column
            If ColLocation(i) > b.bWidth * 12 / 2 And ColLocation(i - 1) < b.bWidth * 12 / 2 And b.rShape = "Gable" Then
                DistanceToPreviousColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
            Else
                DistanceToPreviousColumn = Abs(ColLocation(i) - ColLocation(i - 1))
            End If
                ''''''''''''''Estimate COlumn widths
                'get first width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof("e1", ColLocation(i))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
                'subtract half of width
                DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
                'get second width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof("e1", ColLocation(i - 1))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
                'subtract width if sidewall column, or half of width otherwise
                If i - 1 = 0 Then
                    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
                Else
                    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
                End If
            ''''''''''''''Distance to NEXT Column
            If ColLocation(i) < b.bWidth * 12 / 2 And ColLocation(i + 1) > b.bWidth * 12 / 2 And b.rShape = "Gable" Then
                DistanceToNextColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
            Else
                DistanceToNextColumn = Abs(ColLocation(i) - ColLocation(i + 1))
            End If
                ''''''''''''''Estimate COlumn widths
                'get first width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof("e1", ColLocation(i))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
                'subtract half of width
                DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
                'get second width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof("e1", ColLocation(i + 1))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
                'subtract width if sidewall column, or half of width otherwise
                If i + 1 = UBound(ColLocation()) Then
                    DistanceToNextColumn = DistanceToNextColumn - Column.Width
                Else
                    DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
                End If
        End If
        If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
            Erase ColLocation
            Erase EndwallColumnCLs()
            Call NewExpandableEndwallColumnsGen(b, eWall, EndwallColumnCLs(), ColNum + 1, True)
            Exit Sub
        End If
'    ElseIf DistanceToPreviousColumn <= MaxHorizontalDistance * 12 And DistanceToPreviousColumn >= MinHorizontalDistance * 12 _
'    Or DistanceToNextColumn <= MaxHorizontalDistance * 12 And DistanceToNextColumn >= MinHorizontalDistance * 12 Then
'    'if distance is between the min and max horizontal value, we need to check actual column widths and recheck.
'        EndWidth = MinimumInteriorColumnWidth(b, i + 1, ColLocation) / 2
'        StartWidth = MinimumInteriorColumnWidth(b, i, ColLocation) / 2
'        PrevWidth = MinimumInteriorColumnWidth(b, i - 1, ColLocation) / 2
'        DistanceToPreviousColumn = Abs((ColLocation(i) - StartWidth) - (ColLocation(i - 1) + PrevWidth))
'        DistanceToNextColumn = Abs((ColLocation(i) + StartWidth) - (ColLocation(i + 1) - PrevWidth))
'        If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
'            Erase ColLocation
'            Call IntColumnsGen(b, ColNum + 1)
'            Exit Sub
'        End If
    End If
Next i

ReDim EndwallColumnCLs(UBound(ColLocation))
For i = 0 To UBound(ColLocation)
    EndwallColumnCLs(i) = ColLocation(i)
Next i
    

End Sub


'''''''''''''' generates endwall column centerlines, optimizing for like groupings of girt segments, symetric spacing, and variable center column requirements.
Sub EndwallColumnCLCalc(b As clsBuilding, Optional eWall As String)
Dim TwentyFootQty As Integer
Dim TwentyFiveQty As Integer
Dim ThirtyQty As Integer
Dim MinSegs As Integer
Dim Girt As clsMember
Dim Column As clsMember
Dim tempGirtSpan As Double
Dim EndwallColumnCLs() As Double
Dim tempEndwallCLs() As Double
Dim ColCount As Integer
Dim SpanCount As Integer
Dim i As Integer
Dim GirtSpan As Double
Dim ColNum As Integer
Dim TotalSegmentGroupLength As Double
Dim HalfWallSegCount As Integer
Dim EndwallSecondHalfCLs() As Double
'Dim PartialSegmentTotal As Integer
Dim PreviousSegment As Double
Dim NextSegment As Double
Dim LargestSegmentSize As Double
Dim CenterGirtLength As Double
Dim LoadBearingColumn As Boolean
Dim EndwallGirts As New Collection
Dim FO As clsFO
Dim WallColumns As Collection
Dim FOs As Collection
Dim DistanceToS2 As Double
Dim tempColLocation As Double
Dim j As Integer
Dim LongerDistance As Double
Dim IntColumn As clsMember
Dim IdealSpan As Double
Dim StartCol As clsMember
Dim StartPos As Double
Dim EndPos As Double
Dim tempPos As Double
Dim MaxHorizontalDistance As Double
Dim DistanceToNextCol As Double
Dim DistanceToPrevCol As Double
Dim NextColumn As clsMember
Dim NewColumn As clsMember
Dim StartPosRight As Double
Dim StartPosLeft As Double
Dim CenterFO As Boolean
Dim lGtob As Double
Dim rGtob As Double
Dim Jamb As clsMember
Dim JambSupport As clsMember
Dim tempColumn As clsMember

i = 0

If eWall = "e1" Then
    Set WallColumns = b.e1Columns
    Set FOs = b.e1FOs
ElseIf eWall = "e3" Then
    Set WallColumns = b.e3Columns
    Set FOs = b.e3FOs
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Expandable Endwall Columns
'create load-bearing columns identical to interior columns
If b.ExpandableEndwall(eWall) Then
    MaxHorizontalDistance = (60 / (Sqr((b.rPitch / 12) ^ 2 + 1))) * 12
    If EstSht.Range("BayNum").Value <= 1 Then
        'single bay has no interior columns
        ReDim EndwallColumnCLs(1)
        Call NewExpandableEndwallColumnsGen(b, eWall, EndwallColumnCLs())
    Else
        ColCount = b.InteriorColumns.Count
        ReDim EndwallColumnCLs(ColCount - 1)
        For i = 1 To ColCount
            Set IntColumn = b.InteriorColumns(i)
            EndwallColumnCLs(i - 1) = IntColumn.CL
        Next i
    End If
    

    If eWall = "e1" Then
        Call QuickSort(EndwallColumnCLs, LBound(EndwallColumnCLs), UBound(EndwallColumnCLs))
    ElseIf eWall = "e3" Then
        For i = 0 To UBound(EndwallColumnCLs)
            EndwallColumnCLs(i) = b.bWidth * 12 - EndwallColumnCLs(i)
        Next i
        Call QuickSort(EndwallColumnCLs, LBound(EndwallColumnCLs), UBound(EndwallColumnCLs))
        Call ReverseArray(EndwallColumnCLs)
    End If
    
    For i = 0 To UBound(EndwallColumnCLs)
        'Check MiscFOs and Windows that interfere with Load Bearing Columns
        For Each FO In FOs
            If (FO.FOType = "Window" Or FO.FOType = "MiscFO") And _
            FO.rEdgePosition < EndwallColumnCLs(i) And FO.lEdgePosition > EndwallColumnCLs(i) Then
                'FO is in the way
                'check closest jamb
                If Abs(FO.rEdgePosition - EndwallColumnCLs(i)) < Abs(FO.lEdgePosition - EndwallColumnCLs(i)) Then
                    tempColLocation = FO.rEdgePosition - 12
                Else
                    tempColLocation = FO.lEdgePosition + 12
                End If
                DistanceToNextCol = Abs(tempColLocation - EndwallColumnCLs(i + 1))
                DistanceToPrevCol = Abs(tempColLocation - EndwallColumnCLs(i - 1))
                If DistanceToNextCol < MaxHorizontalDistance And DistanceToPrevCol < MaxHorizontalDistance Then
                    EndwallColumnCLs(i) = tempColLocation
                Else
                    'check other jamb
                    If Abs(FO.rEdgePosition - EndwallColumnCLs(i)) > Abs(FO.lEdgePosition - EndwallColumnCLs(i)) Then
                        tempColLocation = FO.rEdgePosition - 12
                    Else
                        tempColLocation = FO.lEdgePosition + 12
                    End If
                    DistanceToNextCol = Abs(tempColLocation - EndwallColumnCLs(i + 1))
                    DistanceToPrevCol = Abs(tempColLocation - EndwallColumnCLs(i - 1))
                    If DistanceToNextCol < MaxHorizontalDistance And DistanceToPrevCol < MaxHorizontalDistance Then
                        EndwallColumnCLs(i) = tempColLocation
                    Else
                        'make both jambs load bearing
                        ReDim Preserve EndwallColumnCLs(UBound(EndwallColumnCLs) + 1)
                        EndwallColumnCLs(i) = FO.rEdgePosition - 12
                        EndwallColumnCLs(UBound(EndwallColumnCLs)) = FO.lEdgePosition + 12
                        'Re-Sort
                        If eWall = "e1" Then
                            Call QuickSort(EndwallColumnCLs, LBound(EndwallColumnCLs), UBound(EndwallColumnCLs))
                        ElseIf eWall = "e3" Then
                            Call QuickSort(EndwallColumnCLs, LBound(EndwallColumnCLs), UBound(EndwallColumnCLs))
                            Call ReverseArray(EndwallColumnCLs)
                            For j = 0 To UBound(EndwallColumnCLs)
                                EndwallColumnCLs(j) = b.bWidth * 12 - EndwallColumnCLs(j)
                            Next j
                        End If
                    End If
                End If
            End If
        Next FO
    Next i
    
    'set column variables, types, sizes, etc.
    For i = 0 To UBound(EndwallColumnCLs)
        'find larger distance to neighboring columns to use in lookup tables
        's2 and s4 columns only have 1 value, all other columns have 2 neighboring columns, the one farthest away is the distance used
        If i = UBound(EndwallColumnCLs) Then
            LongerDistance = Abs(EndwallColumnCLs(i) - EndwallColumnCLs(i - 1))
        ElseIf i = 0 Then
            LongerDistance = Abs(EndwallColumnCLs(i) - EndwallColumnCLs(i + 1))
        Else
            LongerDistance = Application.WorksheetFunction.Max(Abs(EndwallColumnCLs(i) - EndwallColumnCLs(i - 1)), Abs(EndwallColumnCLs(i) - EndwallColumnCLs(i + 1)))
        End If
        
        Set Column = New clsMember
        Column.mType = "Column"
        Column.CL = EndwallColumnCLs(i)
        Column.LoadBearing = True
        If eWall = "e1" Then
            If b.rShape = "Single Slope" Then
                If i = 0 Then
                    Column.Length = ((b.bWidth * 12) * (b.rPitch / 12)) + b.bHeight * 12
                    Column.CL = 0
                ElseIf i = UBound(EndwallColumnCLs) Then
                    Column.Length = b.bHeight * 12
                    Column.CL = b.bWidth * 12
                Else
                    Column.Length = b.DistanceToRoof("e1", Column.CL)
                End If
            Else 'Gable
                If i = 0 Then
                    Column.Length = b.bHeight * 12
                    Column.CL = 0
                ElseIf i = UBound(EndwallColumnCLs) Then
                    Column.Length = b.bHeight * 12
                    Column.CL = b.bWidth * 12
                Else
                    Column.Length = b.DistanceToRoof("e1", Column.CL)
                End If
            End If
        Else 'e3
            If b.rShape = "Single Slope" Then
                If i = 0 Then
                    Column.Length = b.bHeight * 12
                    Column.CL = 0
                ElseIf i = UBound(EndwallColumnCLs) Then
                    Column.Length = ((b.bWidth * 12) * (b.rPitch / 12)) + b.bHeight * 12
                    Column.CL = b.bWidth * 12
                Else
                    Column.Length = b.DistanceToRoof(eWall, Column.CL)
                End If
            Else 'Gable roof
                If i = 0 Then
                    Column.Length = b.bHeight * 12
                    Column.CL = 0
                ElseIf i = UBound(EndwallColumnCLs) Then
                    Column.Length = b.bHeight * 12
                    Column.CL = b.bWidth * 12
                Else
                    Column.Length = b.DistanceToRoof(eWall, Column.CL)
                End If
            End If
        End If
                
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", LongerDistance
        If Column.CL = 0 Then
            Column.CL = Column.Width / 2
        ElseIf Column.CL = b.bWidth * 12 Then
            Column.CL = b.bWidth * 12 - Column.Width / 2
        End If
        Column.rEdgePosition = Column.CL - Column.Width / 2
        WallColumns.Add Column
    Next i
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Non Load Bearing Columns
    'Add non-load bearing columns in spaces greater than 30'
    ColCount = WallColumns.Count
    'first check for largest possible space spanning gable roof to add column
    If b.rShape = "Gable" Then
        For i = 1 To ColCount - 1
            Set Column = WallColumns(i)
            Set NextColumn = WallColumns(i + 1)
            If Column.CL < b.bWidth * 12 / 2 And NextColumn.CL > b.bWidth * 12 / 2 And Abs(Column.CL - NextColumn.CL) > 30 * 12 Then
                Set NewColumn = New clsMember
                NewColumn.CL = b.bWidth * 12 / 2
                NewColumn.LoadBearing = False
                NewColumn.bEdgeHeight = 0
                NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL)
                NewColumn.Length = NewColumn.tEdgeHeight
                If NewColumn.Length < 30 * 12 + 4 Then
                    NewColumn.Size = "8"" C Purlin"
                    NewColumn.Width = 2.5
                Else
                    NewColumn.SetSize b, "Column", eWall, GirtSpan, "NonExpandable"
                End If
                NewColumn.rEdgePosition = NewColumn.CL - NewColumn.Width / 2
                WallColumns.Add NewColumn, , i + 1
                ColCount = ColCount + 1
            End If
        Next i
    End If
    For i = 1 To ColCount - 1
        Set Column = WallColumns(i)
        Set NextColumn = WallColumns(i + 1)
        If b.rShape = "Single Slope" And eWall = "e1" Then
            If Abs(Column.CL - NextColumn.CL) > 30 * 12 Then
                tempGirtSpan = Abs(Column.CL - NextColumn.CL) / 2
                GirtSpan = NearestMemberSize(tempGirtSpan, 1, "C Purlin", True)
                Set NewColumn = New clsMember
                NewColumn.CL = NextColumn.CL - GirtSpan
                NewColumn.LoadBearing = False
                NewColumn.bEdgeHeight = 0
                NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL)
                NewColumn.Length = NewColumn.tEdgeHeight
                If NewColumn.Length < 30 * 12 + 4 Then
                    NewColumn.Size = "8"" C Purlin"
                    NewColumn.Width = 2.5
                Else
                    NewColumn.SetSize b, "Column", eWall, GirtSpan, "NonExpandable"
                End If
                NewColumn.rEdgePosition = NewColumn.CL - NewColumn.Width / 2
                WallColumns.Add NewColumn
            End If
        ElseIf b.rShape = "Single Slope" And eWall = "e3" Then
            If Abs(Column.CL - NextColumn.CL) > 30 * 12 Then
                tempGirtSpan = Abs(Column.CL - NextColumn.CL) / 2
                GirtSpan = NearestMemberSize(tempGirtSpan, 1, "C Purlin", True)
                Set NewColumn = New clsMember
                NewColumn.CL = Column.CL + GirtSpan
                NewColumn.LoadBearing = False
                NewColumn.bEdgeHeight = 0
                NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL)
                NewColumn.Length = NewColumn.tEdgeHeight
                If NewColumn.Length < 30 * 12 + 4 Then
                    NewColumn.Size = "8"" C Purlin"
                    NewColumn.Width = 2.5
                Else
                    NewColumn.SetSize b, "Column", eWall, GirtSpan, "NonExpandable"
                End If
                NewColumn.rEdgePosition = NewColumn.CL - NewColumn.Width / 2
                WallColumns.Add NewColumn
            End If
        Else 'Gable roofs
            Set Column = WallColumns(i)
            Set NextColumn = WallColumns(i + 1)
            If Column.CL < b.bWidth * 12 / 2 Then
                If Abs(Column.CL - NextColumn.CL) > 30 * 12 Then
                    tempGirtSpan = Abs(Column.CL - NextColumn.CL) / 2
                    GirtSpan = NearestMemberSize(tempGirtSpan, 1, "C Purlin", True)
                    Set NewColumn = New clsMember
                    NewColumn.CL = NextColumn.CL - GirtSpan
                    NewColumn.LoadBearing = False
                    NewColumn.bEdgeHeight = 0
                    NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL)
                    NewColumn.Length = NewColumn.tEdgeHeight
                    If NewColumn.Length < 30 * 12 + 4 Then
                        NewColumn.Size = "8"" C Purlin"
                        NewColumn.Width = 2.5
                    Else
                        NewColumn.SetSize b, "Column", eWall, GirtSpan, "NonExpandable"
                    End If
                    NewColumn.rEdgePosition = NewColumn.CL - NewColumn.Width / 2
                    WallColumns.Add NewColumn
                End If
            Else
                If Abs(Column.CL - NextColumn.CL) > 30 * 12 Then
                    tempGirtSpan = Abs(Column.CL - NextColumn.CL) / 2
                    GirtSpan = NearestMemberSize(tempGirtSpan, 1, "C Purlin", True)
                    Set NewColumn = New clsMember
                    NewColumn.CL = Column.CL + GirtSpan
                    NewColumn.LoadBearing = False
                    NewColumn.bEdgeHeight = 0
                    NewColumn.tEdgeHeight = b.DistanceToRoof(eWall, NewColumn.CL)
                    NewColumn.Length = NewColumn.tEdgeHeight
                    If NewColumn.Length < 30 * 12 + 4 Then
                        NewColumn.Size = "8"" C Purlin"
                        NewColumn.Width = 2.5
                    Else
                        NewColumn.SetSize b, "Column", eWall, GirtSpan, "NonExpandable"
                    End If
                    NewColumn.rEdgePosition = NewColumn.CL - NewColumn.Width / 2
                    WallColumns.Add NewColumn
                End If
            End If
        End If
    Next i
    
    Dim PrevColumn As clsMember
    Dim NextDistance As Double
    Dim PrevDistance As Double
    
    For i = 1 To WallColumns.Count
        MaxHorizontalDistance = 30 * 12
        NextDistance = b.bWidth * 12
        PrevDistance = 0
        Set Column = WallColumns(i)
        If Column.LoadBearing = False Then
            For j = 1 To WallColumns.Count
                If j <> WallColumns.Count Then
                    Set NextColumn = WallColumns(j + 1)
                End If
                If j <> 1 Then
                    Set PrevColumn = WallColumns(j - 1)
                End If
                If j = WallColumns.Count Then
                    NextDistance = b.bWidth * 12
                ElseIf Abs(Column.CL - NextColumn.CL) < Abs(Column.CL - NextDistance) And NextColumn.CL > Column.CL Then
                    NextDistance = NextColumn.CL
                End If
                If j = 1 Then
                    PrevDistance = 0
                ElseIf Abs(Column.CL - PrevColumn.CL) < Abs(Column.CL - PrevDistance) And PrevColumn.CL < Column.CL Then
                    PrevDistance = PrevColumn.CL
                End If
            Next j
        'Check MiscFOs and Windows that interfere with Load Bearing Columns
            For Each FO In FOs
                If FO.rEdgePosition < Column.CL And FO.lEdgePosition > Column.CL Then
                    'FO is in the way
                    'if OHDoor or MiscFO w/ full height jambs, remove column
                    If FO.FOType = "OHDoor" Or (FO.FOType = "MiscFO" And FO.StructuralSteelOption Like "*Full Height*") Then
                        WallColumns.Remove i
                        Exit For
                    End If
                    'check closest jamb
                    If Abs(FO.rEdgePosition - Column.CL) < Abs(FO.lEdgePosition - Column.CL) Then
                        tempColLocation = FO.rEdgePosition
                    Else
                        tempColLocation = FO.lEdgePosition
                    End If
                    DistanceToNextCol = Abs(tempColLocation - NextDistance)
                    If i <> 1 Then
                        DistanceToPrevCol = Abs(tempColLocation - PrevDistance)
                    Else
                        DistanceToPrevCol = Column.CL
                    End If
                    If DistanceToNextCol < MaxHorizontalDistance And DistanceToPrevCol < MaxHorizontalDistance Then
                        Column.CL = tempColLocation
                        Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
                        Column.Length = Column.tEdgeHeight
                        If Column.Length > 30 * 12 + 4 Then
                            Column.SetSize b, "Column", eWall, 30, "NonExpandable"
                        Else
                            Column.Size = "8"" Receiver Cee"
                            Column.Width = 2.5
                        End If
                    Else
                        'check other jamb
                        If Abs(FO.rEdgePosition - Column.CL) > Abs(FO.lEdgePosition - Column.CL) Then
                            tempColLocation = FO.rEdgePosition
                        Else
                            tempColLocation = FO.lEdgePosition
                        End If
                        DistanceToNextCol = Abs(tempColLocation - NextDistance)
                        If i <> 1 Then
                            DistanceToPrevCol = Abs(tempColLocation - PrevDistance)
                        Else
                            DistanceToPrevCol = Column.CL
                        End If
                        If DistanceToNextCol < MaxHorizontalDistance And DistanceToPrevCol < MaxHorizontalDistance Then
                            Column.CL = tempColLocation
                            Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
                            Column.Length = Column.tEdgeHeight
                            If Column.Length > 30 * 12 + 4 Then
                                Column.SetSize b, "Column", eWall, 30, "NonExpandable"
                            Else
                                Column.Size = "8"" Receiver Cee"
                                Column.Width = 2.5
                            End If
                        Else
                            'make another extra column at both edges
                            Column.CL = FO.rEdgePosition
                            Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
                            Column.Length = Column.tEdgeHeight
                            If Column.Length > 30 * 12 + 4 Then
                                Column.SetSize b, "Column", eWall, 30, "NonExpandable"
                            Else
                                Column.Size = "8"" Receiver Cee"
                                Column.Width = 2.5
                            End If
                            Set Column = New clsMember
                            Column.CL = FO.lEdgePosition
                            Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
                            Column.Length = Column.tEdgeHeight
                            If Column.Length > 30 * 12 + 4 Then
                                Column.SetSize b, "Column", eWall, 30, "NonExpandable"
                            Else
                                Column.Size = "8"" Receiver Cee"
                                Column.Width = 2.5
                            End If
                            WallColumns.Add Columns, , i + 1
                        End If
                    End If
                End If
            Next FO
        End If
    Next i

    
Else '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Non-Expandable Endwall Columns
    Dim NewHorizontalDistance As Double
    MaxHorizontalDistance = (30 / (Sqr((b.rPitch / 12) ^ 2 + 1))) * 12
    'largest/ideal girt length for wall
    MaxHorizontalDistance = MaxHorizontalDistance
    
    If b.rShape = "Single Slope" Then
        'create two temporary columns, get their widths, and adjust the max horizontal distance so that a more accurate rafter length can be estimated
        Set tempColumn = New clsMember
        If eWall = "e3" Then
            tempColumn.CL = 0
        Else
            tempColumn.CL = (b.bWidth * 12)
        End If
        tempColumn.tEdgeHeight = b.DistanceToRoof(eWall, tempColumn.CL)
        tempColumn.SetSize b, "Column", eWall, 360
        NewHorizontalDistance = MaxHorizontalDistance + tempColumn.Width
        'second column
        Set tempColumn = New clsMember
        If eWall = "e3" Then
            tempColumn.CL = 360
        Else
            tempColumn.CL = b.bWidth * 12 - 360
        End If
        tempColumn.tEdgeHeight = b.DistanceToRoof(eWall, tempColumn.CL)
        tempColumn.SetSize b, "Column", eWall, 360
        NewHorizontalDistance = NewHorizontalDistance + tempColumn.Width
        
        If NewHorizontalDistance >= (30 * 12) Then
            IdealSpan = 30 * 12
        ElseIf NewHorizontalDistance >= (25 * 12) Then
            IdealSpan = 25 * 12
        ElseIf NewHorizontalDistance >= 20 Then
            IdealSpan = 20 * 12
        Else
            IdealSpan = NewHorizontalDistance * 12
        End If
        If eWall = "e1" Then
            StartPos = 0
            EndPos = IdealSpan
        Else
            StartPos = b.bWidth * 12
            EndPos = b.bWidth * 12 - IdealSpan
        End If
        
        If eWall = "e1" Then
            While EndPos < (b.bWidth * 12)
                If (b.bWidth * 12 - StartPos) < (IdealSpan * 1.5) And (b.bWidth * 12 - StartPos) > IdealSpan Then
                    IdealSpan = IdealSpan - 60
                End If
                If FOs.Count = 0 Then
                    tempPos = StartPos
                Else
                    tempPos = NonExpandableFOJambs(b, eWall, StartPos, MaxHorizontalDistance, IdealSpan, 1)
                End If
                If tempPos = StartPos Then 'no FOs interfered with ideal location, add new column
                    Set Column = New clsMember
                    Column.bEdgeHeight = 0
                    Column.CL = StartPos + IdealSpan
                    Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
                    Column.Length = Column.tEdgeHeight
                    Column.LoadBearing = True
                    Column.SetSize b, "Column", eWall, 30
                    Column.rEdgePosition = Column.CL - Column.Width / 2
                    WallColumns.Add Column
                    tempPos = Column.CL
                End If
                StartPos = tempPos
                EndPos = tempPos + IdealSpan
            Wend
        Else
            While EndPos > 0
                If (StartPos) < (IdealSpan * 1.5) Then
                    IdealSpan = IdealSpan - 60
                End If
                If FOs.Count = 0 Then
                    tempPos = StartPos
                Else
                    tempPos = NonExpandableFOJambs(b, eWall, StartPos, MaxHorizontalDistance, IdealSpan, -1)
                End If
                If tempPos = StartPos Then 'no FOs interfered with ideal location, add new column
                    Set Column = New clsMember
                    Column.bEdgeHeight = 0
                    Column.CL = StartPos - IdealSpan
                    Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
                    Column.Length = Column.tEdgeHeight
                    Column.LoadBearing = True
                    Column.SetSize b, "Column", eWall, 30
                    Column.rEdgePosition = Column.CL - Column.Width / 2
                    WallColumns.Add Column
                    tempPos = Column.CL
                End If
                StartPos = tempPos
                EndPos = tempPos - IdealSpan
            Wend
        End If
    Else 'Gable Roof
        'create two temporary columns, get their widths, and adjust the max horizontal distance so that a more accurate rafter length can be estimated
        Set tempColumn = New clsMember
        If eWall = "e1" Then
            tempColumn.CL = b.bWidth * 12 / 2
        Else
            tempColumn.CL = b.bWidth * 12 / 2
        End If
        tempColumn.tEdgeHeight = b.DistanceToRoof(eWall, tempColumn.CL)
        tempColumn.SetSize b, "Column", eWall, 360
        NewHorizontalDistance = MaxHorizontalDistance + tempColumn.Width / 2
        'second column
        Set tempColumn = New clsMember
        If eWall = "e1" Then
            tempColumn.CL = b.bWidth * 12 / 2 + 360
        Else
            tempColumn.CL = b.bWidth * 12 / 2 - 360
        End If
        tempColumn.tEdgeHeight = b.DistanceToRoof(eWall, tempColumn.CL)
        tempColumn.SetSize b, "Column", eWall, 360
        NewHorizontalDistance = NewHorizontalDistance + tempColumn.Width
        
        'Other side of Gable roof; going to the left
        'reset ideal span
        If NewHorizontalDistance >= (30 * 12) Then
            IdealSpan = 30 * 12
        ElseIf NewHorizontalDistance >= (25 * 12) Then
            IdealSpan = 25 * 12
        ElseIf NewHorizontalDistance >= 20 Then
            IdealSpan = 20 * 12
        Else
            IdealSpan = MaxHorizontalDistance * 12
        End If
        
        For Each FO In FOs 'Check if an FO is in the center of the endwall, if so, it MUST have both full height jambs and supports if necessary since it displaces the center column
            If FO.rEdgePosition < b.bWidth * 12 / 2 And FO.lEdgePosition > b.bWidth * 12 / 2 Then
                StartPosRight = FO.rEdgePosition
                StartPosLeft = FO.lEdgePosition
                CenterFO = True
                'Create Jambs and/or Supports
                lGtob = b.DistanceToRoof(eWall, FO.lEdgePosition)
                rGtob = b.DistanceToRoof(eWall, FO.rEdgePosition)
                'left jamb
                Set Jamb = New clsMember
                Jamb.bEdgeHeight = 0
                If lGtob < 30 * 12 + 4 Then 'don't need jamb support
                    Jamb.tEdgeHeight = lGtob
                    Jamb.LoadBearing = True
                Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = 30 * 12
                    Set JambSupport = New clsMember
                    JambSupport.bEdgeHeight = 0
                    JambSupport.CL = FO.lEdgePosition
                    JambSupport.LoadBearing = True
                    JambSupport.tEdgeHeight = lGtob
                    JambSupport.Length = lGtob
                    JambSupport.SetSize b, "Column", eWall, 30
                    JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                    WallColumns.Add JambSupport
                End If
                Jamb.Length = Jamb.tEdgeHeight
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.CL = FO.lEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                FO.FOMaterials.Add Jamb
                
                'right jamb
                Set Jamb = New clsMember
                Jamb.bEdgeHeight = 0
                If rGtob < 30 * 12 + 4 Then 'don't need jamb support
                    Jamb.tEdgeHeight = rGtob
                    Jamb.LoadBearing = True
                Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''need to add jamb support
                    Jamb.tEdgeHeight = 30 * 12
                    Set JambSupport = New clsMember
                    JambSupport.bEdgeHeight = 0
                    JambSupport.CL = FO.rEdgePosition
                    JambSupport.LoadBearing = True
                    JambSupport.tEdgeHeight = rGtob
                    JambSupport.Length = rGtob
                    JambSupport.SetSize b, "Column", eWall, 30
                    JambSupport.rEdgePosition = JambSupport.CL - JambSupport.Width / 2
                    WallColumns.Add JambSupport
                End If
                Jamb.Length = Jamb.tEdgeHeight
                Jamb.Size = "8"" Receiver Cee"
                Jamb.Width = 2.5
                Jamb.CL = FO.rEdgePosition
                Jamb.rEdgePosition = Jamb.CL - Jamb.Width / 2
                FO.FOMaterials.Add Jamb
            End If
        Next FO
        If CenterFO = False Then
            Set Column = New clsMember
            Column.bEdgeHeight = 0
            Column.CL = b.bWidth * 12 / 2
            Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
            Column.Length = Column.tEdgeHeight
            Column.LoadBearing = True
            Column.SetSize b, "Column", eWall, 30
            Column.rEdgePosition = Column.CL - Column.Width / 2
            WallColumns.Add Column
            StartPosRight = b.bWidth * 12 / 2
            StartPosLeft = b.bWidth * 12 / 2
        End If
        If eWall = "e1" Then
            EndPos = StartPosRight - IdealSpan
        Else
            EndPos = StartPosRight - IdealSpan
        End If
        'First side of Gable roof; going right
        While EndPos > 0
            If (StartPosRight) < (IdealSpan * 1.5) And StartPosRight > IdealSpan Then
                IdealSpan = IdealSpan - 60
            End If
            If FOs.Count = 0 Then
                tempPos = StartPosRight
            Else
                tempPos = NonExpandableFOJambs(b, eWall, StartPosRight, MaxHorizontalDistance, IdealSpan, -1)
            End If
            If tempPos = StartPosRight Then 'no FOs interfered with ideal location, add new column
                Set Column = New clsMember
                Column.bEdgeHeight = 0
                Column.CL = StartPosRight - IdealSpan
                Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
                Column.Length = Column.tEdgeHeight
                Column.LoadBearing = True
                Column.SetSize b, "Column", eWall, 30
                Column.rEdgePosition = Column.CL - Column.Width / 2
                WallColumns.Add Column
                tempPos = Column.CL
            End If
            StartPosRight = tempPos
            EndPos = tempPos - IdealSpan
            Debug.Print "Created Column #: " & i
            i = i + 1
        Wend
        'Other side of Gable roof; going to the left
        'reset ideal span
        If NewHorizontalDistance >= (30 * 12) Then
            IdealSpan = 30 * 12
        ElseIf NewHorizontalDistance >= (25 * 12) Then
            IdealSpan = 25 * 12
        ElseIf NewHorizontalDistance >= 20 Then
            IdealSpan = 20 * 12
        Else
            IdealSpan = MaxHorizontalDistance * 12
        End If
        If eWall = "e1" Then
            EndPos = StartPosLeft + IdealSpan
        Else
            EndPos = StartPosLeft + IdealSpan
        End If
        While EndPos < b.bWidth * 12
            If (b.bWidth * 12 - StartPosLeft) < (IdealSpan * 1.5) And (b.bWidth * 12 - StartPosLeft) > IdealSpan Then
                IdealSpan = IdealSpan - 60
            End If
            If FOs.Count = 0 Then
                tempPos = StartPosLeft
            Else
                tempPos = NonExpandableFOJambs(b, eWall, StartPosLeft, MaxHorizontalDistance, IdealSpan, 1)
            End If
            If tempPos = StartPosLeft Then 'no FOs interfered with ideal location, add new column
                Set Column = New clsMember
                Column.bEdgeHeight = 0
                Column.CL = StartPosLeft + IdealSpan
                Column.tEdgeHeight = b.DistanceToRoof(eWall, Column.CL)
                Column.Length = Column.tEdgeHeight
                Column.LoadBearing = True
                Column.SetSize b, "Column", eWall, 30
                Column.rEdgePosition = Column.CL - Column.Width / 2
                WallColumns.Add Column
                tempPos = Column.CL
            End If
            StartPosLeft = tempPos
            EndPos = tempPos + IdealSpan
            Debug.Print "Created Column #: " & i
            i = i + 1
        Wend
    End If
    '''''''''''''''Create Corner columns for Non-Expandable Endwalls
    Set Column = New clsMember
    Column.bEdgeHeight = 0
    Column.CL = 0
    Column.tEdgeHeight = b.DistanceToRoof(eWall, 0)
    Column.Length = Column.tEdgeHeight
    Column.SetSize b, "Column", eWall, 30
    Column.LoadBearing = True
    Column.CL = Column.Width / 2
    Column.rEdgePosition = 0
    WallColumns.Add Column
    
    Set Column = New clsMember
    Column.bEdgeHeight = 0
    Column.CL = b.bWidth * 12
    Column.tEdgeHeight = b.DistanceToRoof(eWall, b.bWidth * 12)
    Column.Length = Column.tEdgeHeight
    Column.SetSize b, "Column", eWall, 30
    Column.LoadBearing = True
    Column.CL = (b.bWidth * 12) - (Column.Width / 2)
    Column.rEdgePosition = Column.CL - (Column.Width / 2)
    WallColumns.Add Column
End If
            
End Sub
    
    
    
    
    
    
    
    


Sub BaseAngleTrimGen(b As clsBuilding)

Dim BaseAngleCollection As Collection
Dim BaseAngleTrimLength As Integer
Dim FO As clsFO
Dim Member As clsMember
Dim StartPos As Integer
Dim EndPos As Integer
Dim NextFOEdge As Integer
Dim NextStartPos As Integer
Dim AngleNetLength As Integer
Dim ReceiverCNetLength As Integer
Dim BaseOnly As Boolean
Dim OHWidth As Integer
Dim Qty As Integer

    
'Endwall 1 - check if partial, excluded, or gable only
If b.WallStatus("e1") = "Include" Then
    If EstSht.Range("e1_LinerPanels") = "None" Then
        BaseOnly = True
    Else
        BaseOnly = False
    End If
    
    For Each FO In b.e1FOs
        If FO.FOType = "OHDoor" Then
            OHWidth = OHWidth + FO.Width
        End If
    Next FO
    
    BaseAngleTrimLength = (b.bWidth * 12) - OHWidth
    
    If BaseOnly Then
        AngleNetLength = AngleNetLength + BaseAngleTrimLength
    Else
        'AngleNetLength = AngleNetLength + BaseAngleTrimLength
        ReceiverCNetLength = ReceiverCNetLength + BaseAngleTrimLength
    End If
End If
OHWidth = 0
BaseAngleTrimLength = 0
'Endwall 3 - check if partial, excluded, or gable only
If b.WallStatus("e3") = "Include" Then
    If EstSht.Range("e3_LinerPanels") = "None" Then
        BaseOnly = True
    Else
        BaseOnly = False
    End If
    
    For Each FO In b.e3FOs
        If FO.FOType = "OHDoor" Then
            OHWidth = OHWidth + FO.Width
        End If
    Next FO
    
    BaseAngleTrimLength = (b.bWidth * 12) - OHWidth
    
    If BaseOnly Then
        AngleNetLength = AngleNetLength + BaseAngleTrimLength
    Else
        'AngleNetLength = AngleNetLength + BaseAngleTrimLength
        ReceiverCNetLength = ReceiverCNetLength + BaseAngleTrimLength
    End If
End If
OHWidth = 0
BaseAngleTrimLength = 0
'Sidewall 2 - check if partial, excluded, or gable only
If b.WallStatus("s2") = "Include" Then
    If EstSht.Range("s2_LinerPanels") = "None" Then
        BaseOnly = True
    Else
        BaseOnly = False
    End If
    
    For Each FO In b.s2FOs
        If FO.FOType = "OHDoor" Then
            OHWidth = OHWidth + FO.Width
        End If
    Next FO
    
    BaseAngleTrimLength = (b.bLength * 12) - OHWidth
    
    If BaseOnly Then
        AngleNetLength = AngleNetLength + BaseAngleTrimLength
    Else
        'AngleNetLength = AngleNetLength + BaseAngleTrimLength
        ReceiverCNetLength = ReceiverCNetLength + BaseAngleTrimLength
    End If
End If
OHWidth = 0
BaseAngleTrimLength = 0
'Sidewall 4 - check if partial, excluded, or gable only
If b.WallStatus("s4") = "Include" Then
    If EstSht.Range("s4_LinerPanels") = "None" Then
        BaseOnly = True
    Else
        BaseOnly = False
    End If
    
    For Each FO In b.s4FOs
        If FO.FOType = "OHDoor" Then
            OHWidth = OHWidth + FO.Width
        End If
    Next FO
    
    BaseAngleTrimLength = (b.bLength * 12) - OHWidth
    
    If BaseOnly Then
        AngleNetLength = AngleNetLength + BaseAngleTrimLength
    Else
        'AngleNetLength = AngleNetLength + BaseAngleTrimLength
        ReceiverCNetLength = ReceiverCNetLength + BaseAngleTrimLength
    End If
End If

If AngleNetLength > 25 * 12 Then
    Qty = Application.WorksheetFunction.RoundDown(AngleNetLength / (25 * 12), 0)
    AngleNetLength = AngleNetLength - Qty * (25 * 12)
    Set Member = New clsMember
    Member.Length = 25 * 12
    Member.Qty = Qty
    Member.Size = "2x4 Base Angle"
    b.BaseAngleTrim.Add Member
End If
If AngleNetLength > 20 * 12 Then
    Qty = Application.WorksheetFunction.RoundUp(AngleNetLength / (20 * 12), 0)
    AngleNetLength = AngleNetLength - Qty * (20 * 12)
    Set Member = New clsMember
    Member.Size = "2x4 Base Angle"
    Member.Length = 20 * 12
    Member.Qty = Qty
    b.BaseAngleTrim.Add Member
End If
If ReceiverCNetLength > 30 * 12 Then
    Qty = Application.WorksheetFunction.RoundDown(ReceiverCNetLength / (30 * 12), 0)
    ReceiverCNetLength = ReceiverCNetLength - Qty * (30 * 12)
    Set Member = New clsMember
    Member.Size = "8"" Receiver Cee"
    Member.Length = 30 * 12
    Member.Qty = Qty
    b.BaseAngleTrim.Add Member
End If
If ReceiverCNetLength > 25 * 12 Then
    Qty = Application.WorksheetFunction.RoundDown(ReceiverCNetLength / (25 * 12), 0)
    ReceiverCNetLength = ReceiverCNetLength - Qty * (25 * 12)
    Set Member = New clsMember
    Member.Size = "8"" Receiver Cee"
    Member.Length = 25 * 12
    Member.Qty = Qty
    b.BaseAngleTrim.Add Member
End If
If ReceiverCNetLength > 20 * 12 Then
    Qty = Application.WorksheetFunction.RoundUp(ReceiverCNetLength / (20 * 12), 0)
    ReceiverCNetLength = ReceiverCNetLength - Qty * (20 * 12)
    Set Member = New clsMember
    Member.Size = "8"" Receiver Cee"
    Member.Length = 20 * 12
    Member.Qty = Qty
    b.BaseAngleTrim.Add Member
End If
    
           

End Sub

Sub OverhangExtensionMembersGen(b As clsBuilding)

Dim Member As clsMember
Dim NewMember As clsMember
Dim CopyMember As clsMember
Dim LinerPanels As Boolean
Dim e1Overhang As Boolean
Dim e1Extension As Boolean
Dim e3Overhang As Boolean
Dim e3Extension As Boolean
Dim s2Overhang As Boolean
Dim s2Extension As Boolean
Dim s4Overhang As Boolean
Dim s4Extension As Boolean
Dim Rafterlines As Integer
Dim Pitch As Double
Dim Angle As Double
Dim DistanceToLower As Double
Dim DistanceToLengthen As Double
Dim ExtensionHeight As Double
Dim ExtensionWidth As Double
Dim i As Integer
Dim RafterSize As String
Dim RafterWidth As Double
Dim StartPos As Double
Dim BayLength As Double
Dim lEdgeStart As Double
Dim rEdgeMax As Double
Dim lEdgeMax As Double
Dim rEdgeStart As Double
Dim tEdgeMax As Double
Dim bEdgeStart As Double
Dim HorizontalDistance As Double
Dim TotalSlopeLength As Double
Dim RafterNum As Double
Dim Size As String
Dim Width As Double
Dim ExtensionInsideEdge As Double
Dim ColumnWidth As Double



'check for liner panels, overhangs, extension, soffit
If EstSht.Range("Roof_LinerPanels").Value <> "None" Then LinerPanels = True
If EstSht.Range("e1_GableOverhang").Value > 0 Then e1Overhang = True
If EstSht.Range("e1_GableExtension").Value > 0 Then e1Extension = True
If EstSht.Range("e3_GableOverhang").Value > 0 Then e3Overhang = True
If EstSht.Range("e3_GableExtension").Value > 0 Then e3Extension = True
If EstSht.Range("s2_EaveOverhang").Value > 0 Then s2Overhang = True
If EstSht.Range("s2_EaveExtension").Value > 0 Then s2Extension = True
If EstSht.Range("s4_EaveOverhang").Value > 0 Then s4Overhang = True
If EstSht.Range("s4_EaveExtension").Value > 0 Then s4Extension = True

'Set e1OverhangMembers = New Collection
'Set s2OverhangMembers = New Collection
'Set e3OverhangMembers = New Collection
'Set s4OverhangMembers = New Collection
'Set e1ExtensionMembers = New Collection
'Set s2ExtensionMembers = New Collection
'Set e3ExtensionMembers = New Collection
'Set s4ExtensionMembers = New Collection





''''''''' s2

If s2Extension Then ''''''''''''''''''''''''''''''''''s2 Extension
    Pitch = b.s2ExtensionPitch
    Rafterlines = EstSht.Range("BayNum").Value + 1
    ExtensionWidth = b.s2Extension
    ExtensionHeight = b.bHeight * 12 - Sqr(((b.s2ExtensionRafterLength) ^ 2 - (ExtensionWidth) ^ 2))
    For i = 1 To Rafterlines
        'EXTENSION COLUMNS
        Set NewMember = New clsMember
        If i = 1 Then
            NewMember.tEdgeHeight = ExtensionHeight
            NewMember.SetSize b, "Column", "Interior", ExtensionWidth
            NewMember.RafterLeftEdge = b.bWidth * 12 + ExtensionWidth
            NewMember.Length = ExtensionHeight
            NewMember.CL = NewMember.RafterLeftEdge - NewMember.Width / 2
            NewMember.rEdgePosition = NewMember.RafterLeftEdge - NewMember.Width
            NewMember.mType = "Extension Column"
            NewMember.Placement = "s2 Extension Column at Bay " & i
            b.e1Columns.Add NewMember
            'SET BUILDING VARIABLE
            'extension width is to inside of column
            'extension height is to top of extension column
            'ExtensionInsideEdge is the distance from the building edge to the inside of the extension column; used for rafter coordinates
            'Column Width is the width of the extension columns so that the rafter is only calculated to the inside of the column
            ColumnWidth = NewMember.Width
            ExtensionInsideEdge = ExtensionWidth - NewMember.Width
            b.s2ExtensionWidth = ExtensionWidth
            b.s2ExtensionHeight = ExtensionHeight
            If b.s2e1ExtensionIntersection Then
                Set NewMember = New clsMember
                NewMember.tEdgeHeight = ExtensionHeight
                NewMember.SetSize b, "Column", "Interior", ExtensionWidth
                NewMember.RafterLeftEdge = b.bWidth * 12 + ExtensionWidth
                NewMember.Length = ExtensionHeight
                NewMember.CL = NewMember.RafterLeftEdge - NewMember.Width / 2
                NewMember.rEdgePosition = NewMember.RafterLeftEdge - NewMember.Width
                NewMember.mType = "e1 Extension Column"
                NewMember.Placement = "s2e1 Extension Intersection Column"
                b.e1Columns.Add NewMember
            End If
        ElseIf i < Rafterlines Then
            NewMember.tEdgeHeight = ExtensionHeight
            NewMember.SetSize b, "Column", "Interior", ExtensionWidth
            NewMember.RafterLeftEdge = b.bWidth * 12 + ExtensionWidth
            NewMember.Length = ExtensionHeight
            NewMember.CL = NewMember.RafterLeftEdge - NewMember.Width / 2
            NewMember.rEdgePosition = NewMember.RafterLeftEdge - NewMember.Width
            NewMember.mType = "Extension Column"
            NewMember.Placement = "s2 Extension Column at Bay " & i
            b.InteriorColumns.Add NewMember
        ElseIf i = Rafterlines Then
            NewMember.tEdgeHeight = ExtensionHeight
            NewMember.SetSize b, "Column", "Interior", ExtensionWidth
            NewMember.rEdgePosition = 0 - ExtensionWidth
            NewMember.Length = ExtensionHeight
            NewMember.CL = NewMember.rEdgePosition + NewMember.Width / 2
            NewMember.mType = "Extension Column"
            NewMember.Placement = "s2 Extension Column at Bay " & i
            b.e3Columns.Add NewMember
            If b.s2e3ExtensionIntersection Then
                Set NewMember = New clsMember
                NewMember.tEdgeHeight = ExtensionHeight
                NewMember.SetSize b, "Column", "Interior", ExtensionWidth
                NewMember.rEdgePosition = 0 - ExtensionWidth
                NewMember.Length = ExtensionHeight
                NewMember.CL = NewMember.rEdgePosition + NewMember.Width / 2
                NewMember.mType = "e3 Extension Column"
                NewMember.Placement = "s2e3 Extension Intersection Column"
                b.e3Columns.Add NewMember
            End If
        End If
        If i = 1 Then ''''' e1 rafter line
            RafterSize = b.e1Rafters(1).Size
            RafterWidth = b.e1Rafters(1).Width
            Set NewMember = New clsMember
            NewMember.Size = RafterSize
            NewMember.Width = RafterWidth
            NewMember.rEdgePosition = b.bWidth * 12
            NewMember.Length = b.s2ExtensionRafterLength - ((ColumnWidth / 12) * Sqr((12 ^ 2) + (Pitch ^ 2)))
            Angle = Atn(Pitch / 12)
            NewMember.tEdgeHeight = b.bHeight * 12
            NewMember.bEdgeHeight = ExtensionHeight
            NewMember.RafterLeftEdge = b.bWidth * 12 + ExtensionInsideEdge
            NewMember.SetSize b, "Rafter", "interior", ExtensionWidth
            Angle = Atn(Pitch / 12) * (NewMember.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (NewMember.Width / 2) ^ 2)
            DistanceToLengthen = Sqr(DistanceToLower ^ 2 + NewMember.Width ^ 2)
            NewMember.tEdgeHeight = NewMember.tEdgeHeight - DistanceToLower
            NewMember.bEdgeHeight = NewMember.bEdgeHeight - DistanceToLower
            
            ''''''''''''''''''''''''''''''''''''''
            NewMember.Length = NewMember.Length + DistanceToLengthen
            NewMember.mType = "Extension Rafter"
            NewMember.Placement = "s2 Extension Rafter at Endwall 1"
            b.e1Rafters.Add NewMember
            If b.s2e1ExtensionIntersection Then
                Set CopyMember = New clsMember
                CopyMember.Length = NewMember.Length
                CopyMember.Size = NewMember.Size
                CopyMember.tEdgeHeight = NewMember.tEdgeHeight
                CopyMember.bEdgeHeight = NewMember.bEdgeHeight
                CopyMember.rEdgePosition = NewMember.rEdgePosition
                CopyMember.RafterLeftEdge = NewMember.RafterLeftEdge
                CopyMember.Width = NewMember.Width
                CopyMember.mType = "Extension Rafter"
                CopyMember.Placement = "s2e1 Extension Intersection Rafter"
                b.e1Rafters.Add CopyMember
            End If
        ElseIf i < Rafterlines Then 'interior rafter line
            RafterSize = b.intRafters(1).Size
            RafterWidth = b.intRafters(1).Width
            Set NewMember = New clsMember
            NewMember.Size = RafterSize
            NewMember.Width = RafterWidth
            NewMember.rEdgePosition = b.bWidth * 12
            NewMember.Length = b.s2ExtensionRafterLength - ((ColumnWidth / 12) * Sqr((12 ^ 2) + (Pitch ^ 2)))
            Angle = Atn(Pitch / 12)
            NewMember.tEdgeHeight = b.bHeight * 12
            NewMember.bEdgeHeight = ExtensionHeight
            NewMember.RafterLeftEdge = b.bWidth * 12 + ExtensionInsideEdge
            NewMember.SetSize b, "Rafter", "interior", ExtensionWidth
            Angle = Atn(Pitch / 12) * (NewMember.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (NewMember.Width / 2) ^ 2)
            DistanceToLengthen = Sqr(DistanceToLower ^ 2 + NewMember.Width ^ 2)
            NewMember.tEdgeHeight = NewMember.tEdgeHeight - DistanceToLower
            NewMember.bEdgeHeight = NewMember.bEdgeHeight - DistanceToLower
            NewMember.Length = NewMember.Length + DistanceToLengthen
            NewMember.mType = "Extension Rafter"
            NewMember.Placement = "s2 Extension Rafter at Bay " & i
            b.intRafters.Add NewMember
        ElseIf i = Rafterlines Then 'e3 rafter
            RafterSize = b.e3Rafters(1).Size
            RafterWidth = b.e3Rafters(1).Width
            Set NewMember = New clsMember
            NewMember.Size = RafterSize
            NewMember.Width = RafterWidth
            NewMember.rEdgePosition = 0 - ExtensionInsideEdge
            NewMember.Length = b.s2ExtensionRafterLength - ((ColumnWidth / 12) * Sqr((12 ^ 2) + (Pitch ^ 2)))
            Angle = Atn(Pitch / 12)
            NewMember.tEdgeHeight = b.bHeight * 12
            NewMember.bEdgeHeight = ExtensionHeight
            NewMember.RafterLeftEdge = 0
            NewMember.SetSize b, "Rafter", "interior", ExtensionWidth
            Angle = Atn(Pitch / 12) * (NewMember.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (NewMember.Width / 2) ^ 2)
            DistanceToLengthen = Sqr(DistanceToLower ^ 2 + NewMember.Width ^ 2)
            NewMember.tEdgeHeight = NewMember.tEdgeHeight - DistanceToLower
            NewMember.bEdgeHeight = NewMember.bEdgeHeight - DistanceToLower
            NewMember.Length = NewMember.Length + DistanceToLengthen
            NewMember.mType = "Extension Rafter"
            NewMember.Placement = "s2 Extension Rafter at Endwall 3"
            b.e3Rafters.Add NewMember
            If b.s2e3ExtensionIntersection Then
                Set CopyMember = New clsMember
                CopyMember.Length = NewMember.Length
                CopyMember.Size = NewMember.Size
                CopyMember.tEdgeHeight = NewMember.tEdgeHeight
                CopyMember.bEdgeHeight = NewMember.bEdgeHeight
                CopyMember.rEdgePosition = NewMember.rEdgePosition
                CopyMember.RafterLeftEdge = NewMember.RafterLeftEdge
                CopyMember.Width = NewMember.Width
                CopyMember.mType = "Extension Rafter"
                CopyMember.Placement = "s2e3 Extension Intersection Rafter"
                b.e3Rafters.Add CopyMember
            End If
        End If
        
    Next i
End If

Dim RightEdge As Double
Dim TopEdge As Double
Dim OverhangWidth As Double
Dim OverhangHeight As Double
Dim TotalRafters As Double
Dim s2e1Rafter As Integer
Dim s2e3Rafter As Integer
Dim s2e1OverhangIntersection As Integer
Dim s2e3OverhangIntersection As Integer

If s2Overhang Then ''''''''''''''''''''''''''''''''''s2 Overhang (always goes down)
    If b.s2ExtensionWidth > 0 Then
        Pitch = b.s2ExtensionPitch
        RightEdge = b.bWidth * 12 + b.s2ExtensionWidth
        TopEdge = b.s2ExtensionHeight
    Else
        Pitch = b.rPitch
        RightEdge = b.bWidth * 12
        TopEdge = b.bHeight * 12
    End If
    Rafterlines = EstSht.Range("BayNum").Value + 1
    TotalRafters = Rafterlines
    If b.s2e1ExtensionIntersection Then
        TotalRafters = TotalRafters + 1
        s2e1Rafter = TotalRafters
    End If
    If b.s2e3ExtensionIntersection Then
        TotalRafters = TotalRafters + 1
        s2e3Rafter = TotalRafters
    End If
    For i = 1 To TotalRafters
        RafterSize = "W8x10"
        RafterWidth = 8
        Set NewMember = New clsMember
        NewMember.Size = RafterSize
        NewMember.Width = RafterWidth
        NewMember.rEdgePosition = RightEdge
        'NewMember.Length = EstSht.Range("s2_EaveOverhang").Value * 12
        NewMember.Length = (EstSht.Range("s2_EaveOverhang").Value) * Sqr((12 ^ 2) + (Pitch ^ 2))
        's2ExtensionOverhangRafterLength = (s2ExtensionOverhang / 12) * Sqr((12 ^ 2) + (s2ExtensionPitch ^ 2))
        Angle = Atn(Pitch / 12)
        NewMember.tEdgeHeight = TopEdge
        NewMember.bEdgeHeight = NewMember.tEdgeHeight - (Sin(Angle) * NewMember.Length)
        NewMember.RafterLeftEdge = NewMember.rEdgePosition + Sqr(NewMember.Length ^ 2 - (NewMember.tEdgeHeight - NewMember.bEdgeHeight) ^ 2)
        OverhangWidth = NewMember.RafterLeftEdge - NewMember.rEdgePosition
        Angle = Atn(Pitch / 12) * (NewMember.Width / 2)
        DistanceToLower = Sqr(Angle ^ 2 + (NewMember.Width / 2) ^ 2)
        DistanceToLengthen = Sqr(DistanceToLower ^ 2 + NewMember.Width ^ 2)
        NewMember.tEdgeHeight = NewMember.tEdgeHeight - DistanceToLower
        NewMember.bEdgeHeight = NewMember.bEdgeHeight - DistanceToLower
        OverhangHeight = NewMember.bEdgeHeight
        NewMember.Length = NewMember.Length + DistanceToLengthen
        NewMember.mType = "Overhang Stub Rafter"
        Select Case True
        Case i = 1
            NewMember.Placement = "s2 Stub Rafter at Endwall 1"
            b.e1Rafters.Add NewMember
        Case i < Rafterlines
            NewMember.Placement = "s2 Stub Rafter at Bay " & i
            b.intRafters.Add NewMember
        Case i = Rafterlines
            NewMember.rEdgePosition = -b.s2ExtensionWidth - OverhangWidth
            NewMember.RafterLeftEdge = -b.s2Extension
            NewMember.Placement = "s2 Stub Rafter at Endwall 3"
            b.e3Rafters.Add NewMember
        Case i = s2e1Rafter
            NewMember.Placement = "s2 Stub Rafter at s2e1 intersection"
            b.e1Rafters.Add NewMember
        Case i = s2e3Rafter
            NewMember.rEdgePosition = -b.s2ExtensionWidth - OverhangWidth
            NewMember.RafterLeftEdge = -b.s2Extension
            NewMember.Placement = "s2 Stub Rafter at s2e3 intersection"
            b.e3Rafters.Add NewMember
        End Select
        
    Next i
    
    '''''''''''''''''''add eave struts for s2 eave overhang
    StartPos = 0
    For i = 1 To EstSht.Range("BayNum").Value
        BayLength = EstSht.Range("Bay1_Length").offset(i - 1, 0).Value * 12
        Set NewMember = New clsMember
        NewMember.mType = "Eave Strut"
        NewMember.rEdgePosition = StartPos
        NewMember.Length = BayLength
        NewMember.tEdgeHeight = OverhangHeight
        If Pitch = 1 Then
            NewMember.Size = "8"" C Purlin"
        Else
            If EstSht.Range("s2_EaveOverhangSoffit").Value = "Yes" Then
                NewMember.Size = "8"" " & b.rPitch & ":12 double up eave strut"
            Else
                NewMember.Size = "8"" " & b.rPitch & ":12 single up eave strut"
            End If
        End If
        NewMember.Placement = "s2 Overhang Eave Strut"
        b.RoofPurlins.Add NewMember
        StartPos = StartPos + BayLength
    Next i
    If b.s2e1ExtensionIntersection Then
        BayLength = b.e1Extension
        Set NewMember = New clsMember
        NewMember.mType = "Eave Strut"
        NewMember.rEdgePosition = -b.e1Extension
        NewMember.Length = BayLength
        NewMember.tEdgeHeight = OverhangHeight
        If Pitch = 1 Then
            NewMember.Size = "8"" C Purlin"
        Else
            If EstSht.Range("s2_EaveOverhangSoffit").Value = "Yes" Then
                NewMember.Size = "8"" " & b.rPitch & ":12 double up eave strut"
            Else
                NewMember.Size = "8"" " & b.rPitch & ":12 single up eave strut"
            End If
        End If
        NewMember.Placement = "s2e1 Intersection Overhang Eave Strut"
        b.RoofPurlins.Add NewMember
    End If
    If b.s2e3ExtensionIntersection Then
        BayLength = b.e1Extension
        Set NewMember = New clsMember
        NewMember.mType = "Eave Strut"
        NewMember.rEdgePosition = b.e3Extension
        NewMember.Length = BayLength
        NewMember.tEdgeHeight = OverhangHeight
        If Pitch = 1 Then
            NewMember.Size = "8"" C Purlin"
        Else
            If EstSht.Range("s2_EaveOverhangSoffit").Value = "Yes" Then
                NewMember.Size = "8"" " & b.rPitch & ":12 double up eave strut"
            Else
                NewMember.Size = "8"" " & b.rPitch & ":12 single up eave strut"
            End If
        End If
        NewMember.Placement = "s2e3 Intersection Overhang Eave Strut"
        b.RoofPurlins.Add NewMember
    End If
End If

Dim BottomEdge As Double
''''''''''''''''''''''''''''''''''' s4
If s4Extension Then ''''''''''''''''''''''''''''''''''s4 Extension   (!!! Goes up on Single Slope Buildings!!!)
    Pitch = b.s4ExtensionPitch
    Rafterlines = EstSht.Range("BayNum").Value + 1
    ExtensionWidth = b.s4Extension
    b.s4ExtensionWidth = ExtensionWidth
    If b.rShape = "Gable" Then
        ExtensionHeight = b.bHeight * 12 - Sqr(((b.s4ExtensionRafterLength) ^ 2 - (ExtensionWidth) ^ 2))
        b.s4ExtensionHeight = ExtensionHeight
        BottomEdge = b.bHeight * 12
    Else
        ExtensionHeight = b.HighSideEaveHeight + Sqr(((b.s4ExtensionRafterLength) ^ 2 - (ExtensionWidth) ^ 2))
        b.s4ExtensionHeight = ExtensionHeight
        BottomEdge = b.HighSideEaveHeight
    End If
    For i = 1 To Rafterlines
        'EXTENSION COLUMNS
        Set NewMember = New clsMember
        If i = 1 Then
            NewMember.tEdgeHeight = ExtensionHeight
            'NewMember.bEdgeHeight = BottomEdge
            NewMember.SetSize b, "Column", "Interior", ExtensionWidth
            NewMember.rEdgePosition = -ExtensionWidth
            NewMember.Length = ExtensionHeight
            NewMember.CL = NewMember.rEdgePosition + NewMember.Width / 2
            NewMember.mType = "Extension Column"
            NewMember.Placement = "s4 Extension Column at Endwall 1"
            b.e1Columns.Add NewMember
            'SET BUILDING VARIABLE
            'extension width is to inside of column
            'extension height is to top of extension column
            'ExtensionInsideEdge is the distance from the building edge to the inside of the extension column; used for rafter coordinates
            'Column Width
            ColumnWidth = NewMember.Width
            ExtensionInsideEdge = ExtensionWidth - NewMember.Width
            If b.s4e1ExtensionIntersection Then
                Set NewMember = New clsMember
                NewMember.tEdgeHeight = ExtensionHeight
                'NewMember.bEdgeHeight = BottomEdge
                NewMember.SetSize b, "Column", "Interior", ExtensionWidth
                NewMember.rEdgePosition = -ExtensionWidth
                NewMember.Length = ExtensionHeight
                NewMember.CL = NewMember.rEdgePosition + NewMember.Width / 2
                NewMember.mType = "e1 Extension Column"
                NewMember.Placement = "s4e1 Extension Intersection Column"
                b.e1Columns.Add NewMember
            End If
        ElseIf i < Rafterlines Then
            Set NewMember = New clsMember
            NewMember.tEdgeHeight = ExtensionHeight
            'NewMember.bEdgeHeight = BottomEdge
            NewMember.SetSize b, "Column", "Interior", ExtensionWidth
            NewMember.rEdgePosition = -ExtensionWidth
            NewMember.Length = ExtensionHeight
            NewMember.CL = NewMember.rEdgePosition + NewMember.Width / 2
            NewMember.mType = "Extension Column"
            NewMember.Placement = "s4 Extension Column at Bay " & i
            b.InteriorColumns.Add NewMember
        ElseIf i = Rafterlines Then
            NewMember.tEdgeHeight = ExtensionHeight
            'NewMember.bEdgeHeight = BottomEdge
            NewMember.SetSize b, "Column", "Interior", ExtensionWidth
            NewMember.rEdgePosition = b.bWidth * 12 + ExtensionInsideEdge
            NewMember.Length = ExtensionHeight
            NewMember.CL = NewMember.rEdgePosition + NewMember.Width / 2
            NewMember.mType = "Extension Column"
            NewMember.Placement = "s4 Extension Column at Endwall 3"
            b.e3Columns.Add NewMember
            If b.s4e3ExtensionIntersection Then
                Set NewMember = New clsMember
                NewMember.tEdgeHeight = ExtensionHeight
                'NewMember.bEdgeHeight = BottomEdge
                NewMember.SetSize b, "Column", "Interior", ExtensionWidth
                NewMember.rEdgePosition = b.bWidth * 12 + ExtensionInsideEdge
                NewMember.Length = ExtensionHeight
                NewMember.CL = NewMember.rEdgePosition + NewMember.Width / 2
                NewMember.mType = "e3 Extension Column"
                NewMember.Placement = "s4e3 Extension Intersection Column"
                b.e3Columns.Add NewMember
            End If
        End If
        If i = 1 Then ''''' e1 rafter line
            RafterSize = b.e1Rafters(1).Size
            RafterWidth = b.e1Rafters(1).Width
            Set NewMember = New clsMember
            NewMember.Size = RafterSize
            NewMember.Width = RafterWidth
            NewMember.RafterLeftEdge = 0
            NewMember.Length = b.s4ExtensionRafterLength - ((ColumnWidth / 12) * Sqr((12 ^ 2) + (Pitch ^ 2)))
            Angle = Atn(Pitch / 12)
            If b.rShape = "Gable" Then
                NewMember.bEdgeHeight = ExtensionHeight
                NewMember.tEdgeHeight = b.bHeight * 12
            Else
                NewMember.tEdgeHeight = ExtensionHeight
                NewMember.bEdgeHeight = BottomEdge
            End If
            NewMember.rEdgePosition = -ExtensionInsideEdge
            NewMember.SetSize b, "Rafter", "interior", ExtensionWidth
            Angle = Atn(Pitch / 12) * (NewMember.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (NewMember.Width / 2) ^ 2)
            DistanceToLengthen = Sqr(DistanceToLower ^ 2 + NewMember.Width ^ 2)
            NewMember.tEdgeHeight = NewMember.tEdgeHeight - DistanceToLower
            NewMember.bEdgeHeight = NewMember.bEdgeHeight - DistanceToLower
            
            ''''''''''''''''''''''''''''''''''''''
            NewMember.Length = NewMember.Length + DistanceToLengthen
            NewMember.mType = "Extension Rafter"
            NewMember.Placement = "s4 Extension Rafter at Endwall 1"
            b.e1Rafters.Add NewMember
            If b.s4e1ExtensionIntersection Then
                Set CopyMember = New clsMember
                CopyMember.Length = NewMember.Length
                CopyMember.Size = NewMember.Size
                CopyMember.tEdgeHeight = NewMember.tEdgeHeight
                CopyMember.bEdgeHeight = NewMember.bEdgeHeight
                CopyMember.rEdgePosition = NewMember.rEdgePosition
                CopyMember.RafterLeftEdge = NewMember.RafterLeftEdge
                CopyMember.Width = NewMember.Width
                CopyMember.mType = "Extension Rafter"
                CopyMember.Placement = "s4e1 Extension Intersection Rafter"
                b.e1Rafters.Add CopyMember
            End If
        ElseIf i < Rafterlines Then 'interior rafter line
            RafterSize = b.intRafters(1).Size
            RafterWidth = b.intRafters(1).Width
            Set NewMember = New clsMember
            NewMember.Size = RafterSize
            NewMember.Width = RafterWidth
            NewMember.RafterLeftEdge = 0
            NewMember.Length = b.s4ExtensionRafterLength - ((ColumnWidth / 12) * Sqr((12 ^ 2) + (Pitch ^ 2)))
            Angle = Atn(Pitch / 12)
            If b.rShape = "Gable" Then
                NewMember.bEdgeHeight = ExtensionHeight
                NewMember.tEdgeHeight = b.bHeight * 12
            Else
                NewMember.tEdgeHeight = ExtensionHeight
                NewMember.bEdgeHeight = BottomEdge
            End If
            NewMember.rEdgePosition = -ExtensionInsideEdge
            NewMember.SetSize b, "Rafter", "interior", ExtensionWidth
            Angle = Atn(Pitch / 12) * (NewMember.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (NewMember.Width / 2) ^ 2)
            DistanceToLengthen = Sqr(DistanceToLower ^ 2 + NewMember.Width ^ 2)
            NewMember.tEdgeHeight = NewMember.tEdgeHeight - DistanceToLower
            NewMember.bEdgeHeight = NewMember.bEdgeHeight - DistanceToLower
            NewMember.Length = NewMember.Length + DistanceToLengthen
            NewMember.mType = "Extension Rafter"
            NewMember.Placement = "s4 Extension Rafter at Bay " & i
            b.intRafters.Add NewMember
        ElseIf i = Rafterlines Then 'e3 rafter
            RafterSize = b.e3Rafters(1).Size
            RafterWidth = b.e3Rafters(1).Width
            Set NewMember = New clsMember
            NewMember.Size = RafterSize
            NewMember.Width = RafterWidth
            NewMember.rEdgePosition = b.bWidth * 12
            NewMember.Length = b.s4ExtensionRafterLength - ((ColumnWidth / 12) * Sqr((12 ^ 2) + (Pitch ^ 2)))
            Angle = Atn(Pitch / 12)
            If b.rShape = "Gable" Then
                NewMember.bEdgeHeight = ExtensionHeight
                NewMember.tEdgeHeight = b.bHeight * 12
            Else
                NewMember.tEdgeHeight = ExtensionHeight
                NewMember.bEdgeHeight = BottomEdge
            End If
            NewMember.RafterLeftEdge = NewMember.rEdgePosition + ExtensionWidth
            NewMember.SetSize b, "Rafter", "interior", ExtensionWidth
            Angle = Atn(Pitch / 12) * (NewMember.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (NewMember.Width / 2) ^ 2)
            DistanceToLengthen = Sqr(DistanceToLower ^ 2 + NewMember.Width ^ 2)
            NewMember.tEdgeHeight = NewMember.tEdgeHeight - DistanceToLower
            NewMember.bEdgeHeight = NewMember.bEdgeHeight - DistanceToLower
            NewMember.Length = NewMember.Length + DistanceToLengthen
            NewMember.mType = "Extension Rafter"
            NewMember.Placement = "s4 Extension Rafter at Endwall 3"
            b.e3Rafters.Add NewMember
            If b.s4e3ExtensionIntersection Then
                Set CopyMember = New clsMember
                CopyMember.Length = NewMember.Length
                CopyMember.Size = NewMember.Size
                CopyMember.tEdgeHeight = NewMember.tEdgeHeight
                CopyMember.bEdgeHeight = NewMember.bEdgeHeight
                CopyMember.rEdgePosition = NewMember.rEdgePosition
                CopyMember.RafterLeftEdge = NewMember.RafterLeftEdge
                CopyMember.Width = NewMember.Width
                CopyMember.mType = "Extension Rafter"
                CopyMember.Placement = "s4e3 Extension Intersection Rafter"
                b.e3Rafters.Add CopyMember
            End If
        End If
        
    Next i
End If

'''''''''''''''''''''''''''s4 Overhang
Dim LeftEdge As Double
Dim s4e1Rafter As Double
Dim s4e3Rafter As Double

If s4Overhang Then ''''''''''''''''''''''''''''''''''s4 Overhang (!!!DOESNT always go down!!!)
    If b.s4ExtensionWidth > 0 Then
        Pitch = b.s4ExtensionPitch
        LeftEdge = -b.s4ExtensionWidth
        If b.rShape = "Gable" Then
            TopEdge = b.s4ExtensionHeight
        Else
            BottomEdge = b.s4ExtensionHeight
        End If
    Else
        Pitch = b.rPitch
        LeftEdge = 0
        If b.rShape = "Gable" Then
            TopEdge = b.bHeight * 12
        Else
            BottomEdge = b.bHeight * 12 + ((b.bWidth * 12) * b.rPitch / 12)
        End If
    End If
    Rafterlines = EstSht.Range("BayNum").Value + 1
    TotalRafters = Rafterlines
    If b.s4e1ExtensionIntersection Then
        TotalRafters = TotalRafters + 1
        s4e1Rafter = TotalRafters
    End If
    If b.s4e3ExtensionIntersection Then
        TotalRafters = TotalRafters + 1
        s4e3Rafter = TotalRafters
    End If
    For i = 1 To TotalRafters
        RafterSize = "W8x10"
        RafterWidth = 8
        Set NewMember = New clsMember
        NewMember.Size = RafterSize
        NewMember.Width = RafterWidth
        NewMember.RafterLeftEdge = LeftEdge
        'NewMember.Length = EstSht.Range("s2_EaveOverhang").Value * 12
        NewMember.Length = (EstSht.Range("s4_EaveOverhang").Value) * Sqr((12 ^ 2) + (Pitch ^ 2))
        's2ExtensionOverhangRafterLength = (s2ExtensionOverhang / 12) * Sqr((12 ^ 2) + (s2ExtensionPitch ^ 2))
        Angle = Atn(Pitch / 12)
        If b.rShape = "Gable" Then
            NewMember.tEdgeHeight = TopEdge
            NewMember.bEdgeHeight = NewMember.tEdgeHeight - (Sin(Angle) * NewMember.Length)
            OverhangHeight = NewMember.bEdgeHeight
        Else
            NewMember.bEdgeHeight = BottomEdge
            NewMember.tEdgeHeight = NewMember.bEdgeHeight + (Sin(Angle) * NewMember.Length)
            OverhangHeight = NewMember.tEdgeHeight
        End If
        NewMember.rEdgePosition = NewMember.RafterLeftEdge - Sqr(NewMember.Length ^ 2 - (NewMember.tEdgeHeight - NewMember.bEdgeHeight) ^ 2)
        OverhangWidth = Abs(NewMember.RafterLeftEdge - NewMember.rEdgePosition)
        Angle = Atn(Pitch / 12) * (NewMember.Width / 2)
        DistanceToLower = Sqr(Angle ^ 2 + (NewMember.Width / 2) ^ 2)
        DistanceToLengthen = Sqr(DistanceToLower ^ 2 + NewMember.Width ^ 2)
        NewMember.tEdgeHeight = NewMember.tEdgeHeight - DistanceToLower
        NewMember.bEdgeHeight = NewMember.bEdgeHeight - DistanceToLower
        NewMember.Length = NewMember.Length + DistanceToLengthen
        NewMember.mType = "Overhang Stub Rafter"
        Select Case True
        Case i = 1
            NewMember.Placement = "s4 Stub Rafter at Endwall 1"
            b.e1Rafters.Add NewMember
        Case i < Rafterlines
            NewMember.Placement = "s4 Stub Rafter at Bay " & i
            b.intRafters.Add NewMember
        Case i = Rafterlines
            NewMember.rEdgePosition = b.bWidth * 12 + b.s4ExtensionWidth
            NewMember.RafterLeftEdge = NewMember.rEdgePosition + OverhangWidth
            NewMember.Placement = "s4 Stub Rafter at Endwall 3"
            b.e3Rafters.Add NewMember
        Case i = s4e1Rafter
            NewMember.Placement = "s4 Stub Rafter at s4e1 intersection"
            b.e1Rafters.Add NewMember
        Case i = s4e3Rafter
            NewMember.rEdgePosition = b.bWidth * 12 + b.s4ExtensionWidth
            NewMember.RafterLeftEdge = NewMember.rEdgePosition + OverhangWidth
            NewMember.Placement = "s4 Stub Rafter at s4e3 intersection"
            b.e3Rafters.Add NewMember
        End Select
        
    Next i
    
    '''''''''''''''''''add eave struts for s4 eave overhang
    StartPos = 0
    For i = 1 To EstSht.Range("BayNum").Value
        BayLength = EstSht.Range("Bay1_Length").offset(i - 1, 0).Value * 12
        Set NewMember = New clsMember
        NewMember.mType = "Eave Strut"
        NewMember.rEdgePosition = StartPos
        NewMember.Length = BayLength
        NewMember.tEdgeHeight = OverhangHeight
        If b.rShape = "Gable" Then
            If Pitch = 1 Then
                NewMember.Size = "8"" C Purlin"
            Else
                If EstSht.Range("s4_EaveOverhangSoffit").Value = "Yes" Then
                    NewMember.Size = "8"" " & b.rPitch & ":12 double up eave strut"
                Else
                    NewMember.Size = "8"" " & b.rPitch & ":12 single up eave strut"
                End If
            End If
        Else
            If Pitch = 1 Then
                NewMember.Size = "8"" C Purlin"
            Else
                If EstSht.Range("s2_EaveOverhangSoffit").Value = "Yes" Then
                    NewMember.Size = "8"" " & b.rPitch & ":12 double down eave strut"
                Else
                    NewMember.Size = "8"" " & b.rPitch & ":12 single down eave strut"
                End If
            End If
        End If
        NewMember.Placement = "s4 Overhang Eave Strut"
        b.RoofPurlins.Add NewMember
        StartPos = StartPos + BayLength
    Next i
    If b.s4e1ExtensionIntersection Then
        BayLength = b.e1Extension
        Set NewMember = New clsMember
        NewMember.mType = "Eave Strut"
        NewMember.rEdgePosition = -b.e1Extension
        NewMember.Length = BayLength
        NewMember.tEdgeHeight = OverhangHeight
        If b.rShape = "Gable" Then
            If Pitch = 1 Then
                NewMember.Size = "8"" C Purlin"
            Else
                If EstSht.Range("s4_EaveOverhangSoffit").Value = "Yes" Then
                    NewMember.Size = "8"" " & b.rPitch & ":12 double up eave strut"
                Else
                    NewMember.Size = "8"" " & b.rPitch & ":12 single up eave strut"
                End If
            End If
        Else
            If Pitch = 1 Then
                NewMember.Size = "8"" C Purlin"
            Else
                If EstSht.Range("s2_EaveOverhangSoffit").Value = "Yes" Then
                    NewMember.Size = "8"" " & b.rPitch & ":12 double down eave strut"
                Else
                    NewMember.Size = "8"" " & b.rPitch & ":12 single down eave strut"
                End If
            End If
        End If
        NewMember.Placement = "s4e1 Intersection Overhang Eave Strut"
        b.RoofPurlins.Add NewMember
    End If
    If b.s4e3ExtensionIntersection Then
        BayLength = b.e1Extension
        Set NewMember = New clsMember
        NewMember.mType = "Eave Strut"
        NewMember.rEdgePosition = b.e3Extension
        NewMember.Length = BayLength
        NewMember.tEdgeHeight = OverhangHeight
        If b.rShape = "Gable" Then
            If Pitch = 1 Then
                NewMember.Size = "8"" C Purlin"
            Else
                If EstSht.Range("s4_EaveOverhangSoffit").Value = "Yes" Then
                    NewMember.Size = "8"" " & b.rPitch & ":12 double up eave strut"
                Else
                    NewMember.Size = "8"" " & b.rPitch & ":12 single up eave strut"
                End If
            End If
        Else
            If Pitch = 1 Then
                NewMember.Size = "8"" C Purlin"
            Else
                If EstSht.Range("s2_EaveOverhangSoffit").Value = "Yes" Then
                    NewMember.Size = "8"" " & b.rPitch & ":12 double down eave strut"
                Else
                    NewMember.Size = "8"" " & b.rPitch & ":12 single down eave strut"
                End If
            End If
        End If
        NewMember.Placement = "s4e3 Intersection Overhang Eave Strut"
        b.RoofPurlins.Add NewMember
    End If
End If


''''''''' e1
'if the endwall was non-expandable, the rafter needs to be changed back to a C Purlin
'only rafters with extension attached need to be changed, so if not extension intersections, leave as Receiver Cees
If e1Overhang Or e1Extension Then
    For i = 1 To b.e1Rafters.Count
        Set Member = b.e1Rafters(i)
        If b.WallStatus("e1") = "No" And ((Member.rEdgePosition < 0 And b.s4e1ExtensionIntersection) Or _
        (Member.RafterLeftEdge > b.bWidth * 12 And b.s2e1ExtensionIntersection) Or _
        (Member.rEdgePosition >= 0 And Member.RafterLeftEdge <= b.bWidth * 12)) Then
            If Member.Size = "8"" Receiver Cee" Then
                Member.Size = "8"" C Purlin"
            ElseIf Member.Size = "10"" Receiver Cee" Then
                Member.Size = "10"" C Purlin"
            End If
        End If
    Next i
End If

'''''''''''''''''''''''''''''''ENDWALL OVERHANGS AND EXTENSIONS

'e1
'Add rafters and columns; essentially copied from interior columns and rafters in terms of positioning, coordinates, and size
If e1Extension Then
    'If b.InteriorColumns.Count = 0 Then
    'if no interior columns have been made, create e1 extension columns and rafters
        Call EndwallExtensionColumnsGen(b, "e1")
        Call RafterGen(b, "e1 Extension")
        For i = b.e1ExtensionMembers.Count To 1 Step -1
            b.e1Columns.Add b.e1ExtensionMembers(i)
            b.e1ExtensionMembers.Remove (i)
        Next i
        'For Each Member In b.e1ExtensionMembers
        '    b.e1Columns.Add Member
        'Next Member
    'Else
    '    For i = 1 To b.intRafters.Count
    '        Set Member = b.intRafters(i)
    '        If Member.Size = "8"" Receiver Cee" Then
    '            Member.Size = "8"" C Purlin"
    '        End If
    '        If Member.mType Like "*Overhang Stub Rafter*" Then
    '            'Do not copy
    '        Else
    '            Set NewMember = New clsMember
    '            NewMember.rEdgePosition = Member.rEdgePosition
    '            NewMember.RafterLeftEdge = Member.RafterLeftEdge
    '            NewMember.Length = Member.Length
    '            NewMember.tEdgeHeight = Member.tEdgeHeight
    '            NewMember.bEdgeHeight = Member.bEdgeHeight
    '            NewMember.mType = "e1 Extension Rafter"
    '            ''''''''''''''''''''''''''''''''''''''''''''''''''''NEED TO DEFINE EXTENSION RAFTER SIZE BASED ON EXPANDABLE, etc.
    '            NewMember.Size = Member.Size
    '            NewMember.Width = Member.Width
    '            NewMember.Placement = "e1 Extension Rafter"
    '            b.e1Rafters.Add NewMember
    '        End If
    '    Next i
        For i = 1 To b.InteriorColumns.Count
        'adding extension intersection columns; create copy of existing s2/s4 extension columns
            Set Member = b.InteriorColumns(i)
            If Member.LoadBearing = True And Member.mType Like "*Extension*" And Not (Member.mType Like "e*" & "Extension Column") Then
                Set NewMember = New clsMember
                NewMember.rEdgePosition = Member.rEdgePosition
                NewMember.Length = Member.Length
                NewMember.CL = Member.CL
                NewMember.tEdgeHeight = Member.tEdgeHeight
                NewMember.bEdgeHeight = Member.bEdgeHeight
                NewMember.mType = "e1 Extension Column"
                ''''''''''''''''''''''''''''''''''''''''''''''''''''NEED TO DEFINE EXTENSION RAFTER SIZE BASED ON EXPANDABLE, etc.
                NewMember.Size = Member.Size
                NewMember.Width = Member.Width
                NewMember.Placement = "e1 Extension Column"
                'b.e1ExtensionMembers.Add NewMember
                b.e1Columns.Add NewMember
            End If
        Next i
    'End If
End If
If e1Overhang Then
    'get bay length to determine rafter size
    'if e1 Extension, check Extension bay length
    If b.e1Extension > 25 * 12 Then
        Size = "10"" Receiver Cee"
        Width = 10
    ElseIf b.e1Extension > 0 Then
        Size = "8"" Receiver Cee"
        Width = 8
    ElseIf EstSht.Range("Bay1_Length").Value > 25 Then
        Size = "10"" Receiver Cee"
        Width = 10
    Else
        Size = "8"" Receiver Cee"
        Width = 8
    End If
    'create overhang rafters for endwall section only
    If b.rShape = "Single Slope" Then
        'set max / min temporary values
        rEdgeMax = 0 + 20
        lEdgeStart = b.bWidth * 12 - 20
        'get starting and maximum values from e1 Columns
        For Each Member In b.e1Columns
            If Member.CL < rEdgeMax And Not Member.mType Like "*Extension*" Then
                rEdgeMax = Member.lEdgePosition
                tEdgeMax = b.DistanceToRoof("e1", Member.lEdgePosition)
            ElseIf Member.CL > lEdgeStart And Not Member.mType Like "*Extension*" Then
                lEdgeStart = Member.rEdgePosition
                bEdgeStart = b.DistanceToRoof("e1", Member.rEdgePosition)
            End If
        Next Member
        For Each Member In b.e1Rafters
            If Member.rEdgePosition > 0 And Member.RafterLeftEdge < b.bWidth * 12 And Not Member.mType Like "*Extension*" And Not Member.mType Like "*Overhang*" Then
                TotalSlopeLength = TotalSlopeLength + Member.Length
            End If
        Next Member
        RafterNum = Application.WorksheetFunction.RoundUp(TotalSlopeLength / (30 * 12), 0)
    Else
        'Gable Roof
        'first slope of roof from left to right
        'get starting and maximum values from e1 Columns
        For Each Member In b.e1Columns
            If Member.CL > b.bWidth * 12 - 16 And Member.CL < b.bWidth * 12 And Not Member.mType Like "*Extension*" And Not Member.mType Like "*Overhang*" Then
                lEdgeStart = Member.rEdgePosition
                bEdgeStart = b.DistanceToRoof("e1", Member.rEdgePosition)
            End If
        Next Member
        rEdgeMax = b.bWidth * 12 / 2
        tEdgeMax = b.DistanceToRoof("e1", b.bWidth * 12 / 2)
        For Each Member In b.e1Rafters
            If Member.rEdgePosition >= b.bWidth * 12 / 2 And Member.lEdgePosition < b.bWidth * 12 And Not Member.mType Like "*Extension*" Then
                TotalSlopeLength = TotalSlopeLength + Member.Length
            End If
        Next Member
        RafterNum = Application.WorksheetFunction.RoundUp(TotalSlopeLength / (30 * 12), 0)
    End If
        'Create Overhang Rafters
        For i = 1 To RafterNum
            Set Member = New clsMember
            Member.Size = Size
            Member.Width = Width
            Member.Placement = "e1 Overhang Rafter"
            Member.bEdgeHeight = bEdgeStart
            Member.RafterLeftEdge = lEdgeStart
            If i <> RafterNum Then
                Member.Length = 30 * 12
                HorizontalDistance = (Member.Length / (Sqr((b.rPitch / 12) ^ 2 + 1)))
                Member.rEdgePosition = lEdgeStart - HorizontalDistance
                Member.tEdgeHeight = b.DistanceToRoof("e1", lEdgeStart - HorizontalDistance)
            Else
                Member.Length = TotalSlopeLength - (30 * 12 * (i - 1))
                Member.rEdgePosition = rEdgeMax
                Member.tEdgeHeight = tEdgeMax
            End If
            bEdgeStart = Member.tEdgeHeight
            lEdgeStart = Member.rEdgePosition
            Angle = Atn(b.rPitch / 12) * (Member.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (Member.Width / 2) ^ 2)
            Member.bEdgeHeight = Member.bEdgeHeight - DistanceToLower
            Member.tEdgeHeight = Member.tEdgeHeight - DistanceToLower
            b.e1Rafters.Add Member
        Next i
    If b.rShape = "Gable" Then ' do second slope of roof
        TotalSlopeLength = 0
        'get starting and maximum values from e1 Columns
        'get starting and maximum values from e1 Columns
        For Each Member In b.e1Columns
            If Member.CL > 0 And Member.CL < 16 And Not Member.mType Like "*Extension*" Then
                rEdgeStart = Member.lEdgePosition
                bEdgeStart = b.DistanceToRoof("e1", Member.lEdgePosition)
            End If
        Next Member
        lEdgeMax = b.bWidth * 12 / 2
        tEdgeMax = b.DistanceToRoof("e1", b.bWidth * 12 / 2)
        For Each Member In b.e1Rafters
            If Member.lEdgePosition <= b.bWidth * 12 / 2 And Member.rEdgePosition > 0 And Not Member.Placement Like "*Overhang*" And Not Member.mType Like "*Extension*" And Not Member.mType Like "*Overhang*" Then
                TotalSlopeLength = TotalSlopeLength + Member.Length
            End If
        Next Member
        RafterNum = Application.WorksheetFunction.RoundUp(TotalSlopeLength / (30 * 12), 0)
        For i = 1 To RafterNum
            Set Member = New clsMember
            Member.Size = Size
            Member.Width = Width
            Member.Placement = "e1 Overhang Rafter"
            Member.bEdgeHeight = bEdgeStart
            Member.rEdgePosition = rEdgeStart
            If i <> RafterNum Then
                Member.Length = 30 * 12
                HorizontalDistance = (Member.Length / (Sqr((b.rPitch / 12) ^ 2 + 1)))
                Member.RafterLeftEdge = rEdgeStart + HorizontalDistance
                Member.tEdgeHeight = b.DistanceToRoof("e1", rEdgeStart + HorizontalDistance)
            Else
                Member.Length = TotalSlopeLength - (30 * 12 * (i - 1))
                Member.RafterLeftEdge = lEdgeMax
                Member.tEdgeHeight = tEdgeMax
            End If
            bEdgeStart = Member.tEdgeHeight
            rEdgeStart = Member.RafterLeftEdge
            Angle = Atn(b.rPitch / 12) * (Member.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (Member.Width / 2) ^ 2)
            Member.bEdgeHeight = Member.bEdgeHeight - DistanceToLower
            Member.tEdgeHeight = Member.tEdgeHeight - DistanceToLower
            b.e1Rafters.Add Member
        Next i
    End If
    'create overhang rafters for extension/Overhang intersections
    's2 intersection(s); only 2 possible members:
        'if extension intersection only --> create overhang rafter at intersection
        'if overhang on s2 exists --> create overhang rafter at overhang intersection
    If b.s2e1ExtensionIntersection Then
        For Each Member In b.e1Rafters
            If Member.Placement = "s2e1 Extension Intersection Rafter" Then
                If Member.Length <= 30 Then
                    Set NewMember = New clsMember
                    NewMember.rEdgePosition = Member.rEdgePosition
                    NewMember.tEdgeHeight = Member.tEdgeHeight
                    NewMember.bEdgeHeight = Member.bEdgeHeight
                    NewMember.RafterLeftEdge = Member.RafterLeftEdge
                    NewMember.Length = Member.Length
                    NewMember.Size = Size
                    NewMember.Width = Width
                    NewMember.Placement = "s2e1 Extension Intersection e1 Overhang Rafter"
                    b.e1Rafters.Add NewMember
                Else
                    'create as many rafters as necessary to cover the space of the large Eave Extension
                    RafterNum = Application.WorksheetFunction.RoundUp(Member.Length / (30 * 12), 0)
                    Pitch = b.s2ExtensionPitch
                    For i = 1 To RafterNum
                        If i = 1 Then
                            Set NewMember = New clsMember
                            NewMember.Length = 30 * 12
                            NewMember.rEdgePosition = Member.rEdgePosition
                            NewMember.tEdgeHeight = Member.tEdgeHeight
                            HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                            NewMember.bEdgeHeight = NewMember.tEdgeHeight - Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                            NewMember.RafterLeftEdge = NewMember.rEdgePosition + HorizontalDistance
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e1 Extension Intersection e1 Overhang Rafter " & i
                            bEdgeStart = NewMember.bEdgeHeight
                            lEdgeStart = NewMember.RafterLeftEdge
                            b.e1Rafters.Add NewMember
                        ElseIf i = RafterNum Then
                            Set NewMember = New clsMember
                            NewMember.rEdgePosition = lEdgeStart
                            NewMember.tEdgeHeight = bEdgeStart
                            NewMember.bEdgeHeight = Member.bEdgeHeight
                            NewMember.RafterLeftEdge = Member.RafterLeftEdge
                            NewMember.Length = Member.Length - (30 * 12 * (i - 1))
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e1 Extension Intersection e1 Overhang Rafter " & i
                            b.e1Rafters.Add NewMember
                        Else
                            Set NewMember = New clsMember
                            NewMember.Length = 30 * 12
                            NewMember.rEdgePosition = lEdgeStart
                            NewMember.tEdgeHeight = bEdgeStart
                            HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                            NewMember.bEdgeHeight = NewMember.tEdgeHeight - Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                            NewMember.RafterLeftEdge = NewMember.rEdgePosition + HorizontalDistance
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e1 Extension Intersection e1 Overhang Rafter " & i
                            bEdgeStart = NewMember.bEdgeHeight
                            lEdgeStart = NewMember.RafterLeftEdge
                            b.e1Rafters.Add NewMember
                        End If
                    Next i
                End If
            ElseIf Member.Placement = "s2 Stub Rafter at s2e1 intersection" Then
                Set NewMember = New clsMember
                NewMember.rEdgePosition = Member.rEdgePosition
                NewMember.tEdgeHeight = Member.tEdgeHeight
                NewMember.bEdgeHeight = Member.bEdgeHeight
                NewMember.RafterLeftEdge = Member.RafterLeftEdge
                NewMember.Length = Member.Length
                NewMember.Size = Size
                NewMember.Width = Width
                NewMember.Placement = "s2e1 Overhang Intersection e1 Overhang Rafter"
                b.e1Rafters.Add NewMember
            End If
        Next Member
    End If
    's4 intection; only 2 possible members (same as above)
    If b.s4e1ExtensionIntersection Then
        For Each Member In b.e1Rafters
            If Member.Placement = "s4e1 Extension Intersection Rafter" Then
                If Member.Length <= 30 Then
                    Set NewMember = New clsMember
                    NewMember.rEdgePosition = Member.rEdgePosition
                    NewMember.tEdgeHeight = Member.tEdgeHeight
                    NewMember.bEdgeHeight = Member.bEdgeHeight
                    NewMember.RafterLeftEdge = Member.RafterLeftEdge
                    NewMember.Length = Member.Length
                    NewMember.Size = Size
                    NewMember.Width = Width
                    NewMember.Placement = "s2e1 Extension Intersection e1 Overhang Rafter"
                    b.e1Rafters.Add NewMember
                Else
                    'create as many rafters as necessary to cover the space of the large Eave Extension
                    'if Single slope, adjust bottom and top edges
                    RafterNum = Application.WorksheetFunction.RoundUp(Member.Length / (30 * 12), 0)
                    Pitch = b.s4ExtensionPitch
                    For i = 1 To RafterNum
                        If i = 1 Then
                            Set NewMember = New clsMember
                            NewMember.Length = 30 * 12
                            NewMember.RafterLeftEdge = Member.RafterLeftEdge
                            If b.rShape = "Gable" Then
                                NewMember.tEdgeHeight = Member.tEdgeHeight
                                HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                                NewMember.bEdgeHeight = NewMember.tEdgeHeight - Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                                bEdgeStart = NewMember.bEdgeHeight
                            Else
                                NewMember.bEdgeHeight = Member.bEdgeHeight
                                HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                                NewMember.tEdgeHeight = NewMember.bEdgeHeight + Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                                bEdgeStart = NewMember.tEdgeHeight
                            End If
                            NewMember.rEdgePosition = NewMember.RafterLeftEdge - HorizontalDistance
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e1 Extension Intersection e1 Overhang Rafter " & i
                            rEdgeStart = NewMember.rEdgePosition
                            b.e1Rafters.Add NewMember
                        ElseIf i = RafterNum Then
                            Set NewMember = New clsMember
                            NewMember.RafterLeftEdge = rEdgeStart
                            If b.rShape = "Gable" Then
                                NewMember.tEdgeHeight = bEdgeStart
                                NewMember.bEdgeHeight = Member.bEdgeHeight
                            Else
                                NewMember.bEdgeHeight = bEdgeStart
                                NewMember.tEdgeHeight = Member.tEdgeHeight
                            End If
                            NewMember.rEdgePosition = Member.rEdgePosition
                            NewMember.Length = Member.Length - (30 * 12 * (i - 1))
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e1 Extension Intersection e1 Overhang Rafter " & i
                            b.e1Rafters.Add NewMember
                        Else
                            Set NewMember = New clsMember
                            NewMember.Length = 30 * 12
                            NewMember.RafterLeftEdge = rEdgeStart
                            If b.rShape = "Gable" Then
                                NewMember.tEdgeHeight = bEdgeStart
                                HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                                NewMember.bEdgeHeight = NewMember.tEdgeHeight - Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                                bEdgeStart = NewMember.bEdgeHeight
                            Else
                                NewMember.bEdgeHeight = bEdgeStart
                                HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                                NewMember.tEdgeHeight = NewMember.bEdgeHeight + Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                                bEdgeStart = NewMember.tEdgeHeight
                            End If
                            NewMember.rEdgePosition = NewMember.RafterLeftEdge - HorizontalDistance
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e1 Extension Intersection e1 Overhang Rafter " & i
                            rEdgeStart = NewMember.rEdgePosition
                            b.e1Rafters.Add NewMember
                        End If
                    Next i
                End If
            ElseIf Member.Placement = "s4 Stub Rafter at s4e1 intersection" Then
                Set NewMember = New clsMember
                NewMember.rEdgePosition = Member.rEdgePosition
                NewMember.tEdgeHeight = Member.tEdgeHeight
                NewMember.bEdgeHeight = Member.bEdgeHeight
                NewMember.RafterLeftEdge = Member.RafterLeftEdge
                NewMember.Length = Member.Length
                NewMember.Size = Size
                NewMember.Width = Width
                NewMember.Placement = "s4e1 Overhang Intersection e1 Overhang Rafter"
                b.e1Rafters.Add NewMember
            End If
        Next Member
    End If
        
            
End If

''''''''' e3
'if the endwall was non-expandable, the rafter needs to be changed back to a C Purlin
If e3Overhang Or e3Extension Then
    For i = 1 To b.e3Rafters.Count
        Set Member = b.e3Rafters(i)
        If b.WallStatus("e3") = "No" And ((Member.rEdgePosition < 0 And b.s4e3ExtensionIntersection) Or _
        (Member.RafterLeftEdge > b.bWidth * 12 And b.s2e3ExtensionIntersection) Or _
        (Member.rEdgePosition >= 0 And Member.RafterLeftEdge <= b.bWidth * 12)) Then
            If Member.Size = "8"" Receiver Cee" Then
                Member.Size = "8"" C Purlin"
            ElseIf Member.Size = "10"" Receiver Cee" Then
                Member.Size = "10"" C Purlin"
            End If
        End If
    Next i
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''e3
If e3Extension Then
    'If b.InteriorColumns.Count = 0 Then
    'if no interior columns have been made, create e3 extension columns and rafters
        Call EndwallExtensionColumnsGen(b, "e3")
        Call RafterGen(b, "e3 Extension")
        For i = b.e3ExtensionMembers.Count To 1 Step -1
            b.e3Columns.Add b.e3ExtensionMembers(i)
            b.e3ExtensionMembers.Remove (i)
        Next i
    'End If
    'Else
    '    For i = 1 To b.intRafters.Count
    '        Set Member = b.intRafters(i)
    '        Set NewMember = New clsMember
    '        NewMember.rEdgePosition = (b.bWidth * 12) - Member.RafterLeftEdge
    '        NewMember.RafterLeftEdge = (b.bWidth * 12) - Member.rEdgePosition
    '        NewMember.Length = Member.Length
    '        NewMember.tEdgeHeight = Member.tEdgeHeight
    '        NewMember.bEdgeHeight = Member.bEdgeHeight
    '        NewMember.mType = "e3 Extension Rafter"
    '        NewMember.Size = Member.Size
    '        NewMember.Width = Member.Width
    '        NewMember.Placement = "e3 Extension Rafter"
    '        b.e3Rafters.Add NewMember
    '    Next i
        For i = 1 To b.InteriorColumns.Count
            Set Member = b.InteriorColumns(i)
            If Member.LoadBearing = True And Member.mType Like "*Extension*" And Not (Member.mType Like "e*" & "Extension Column") Then
                Set NewMember = New clsMember
                NewMember.CL = b.bWidth * 12 - Member.CL
                NewMember.rEdgePosition = Member.CL - Member.Width / 2
                NewMember.Length = Member.Length
                NewMember.tEdgeHeight = Member.tEdgeHeight
                NewMember.bEdgeHeight = Member.bEdgeHeight
                NewMember.mType = "e3 Extension Column"
                ''''''''''''''''''''''''''''''''''''''''''''''''''''NEED TO DEFINE EXTENSION RAFTER SIZE BASED ON EXPANDABLE, etc.
                NewMember.Size = Member.Size
                NewMember.Width = Member.Width
                NewMember.Placement = "e3 Extension Column"
                'b.e1ExtensionMembers.Add NewMember
                b.e3Columns.Add NewMember
            End If
        Next i
    'End If
End If
'e3 Overhang
'Add rafters and columns
If e3Overhang Then
    TotalSlopeLength = 0
        'get bay length to determine rafter size
    'if e1 Extension, check Extension bay length
    If b.e3Extension > 25 * 12 Then
        Size = "10"" Receiver Cee"
        Width = 10
    ElseIf b.e3Extension > 0 Then
        Size = "8"" Receiver Cee"
        Width = 8
    ElseIf EstSht.Range("Bay1_Length").offset(EstSht.Range("BayNum").Value, 0).Value > 25 Then
        Size = "10"" Receiver Cee"
        Width = 10
    Else
        Size = "8"" Receiver Cee"
        Width = 8
    End If
    If b.rShape = "Single Slope" Then
        'get starting and maximum values from e1 Columns
        For Each Member In b.e3Columns
            If Member.CL > 0 And Member.CL < 16 And Not Member.mType Like "*Extension*" Then
                rEdgeStart = Member.lEdgePosition
                bEdgeStart = b.DistanceToRoof("e3", Member.lEdgePosition)
            End If
            If Member.CL < b.bWidth * 12 And Member.CL > b.bWidth * 12 - 16 And Not Member.mType Like "*Extension*" Then
                lEdgeMax = Member.rEdgePosition
                tEdgeMax = b.DistanceToRoof("e3", Member.rEdgePosition)
            End If
        Next Member
        For Each Member In b.e3Rafters
            If Member.rEdgePosition > 0 And Member.lEdgePosition < b.bWidth * 12 And Not Member.mType Like "*Extension*" And Not Member.mType Like "*Overhang*" Then
                TotalSlopeLength = TotalSlopeLength + Member.Length
            End If
        Next Member
        RafterNum = Application.WorksheetFunction.RoundUp(TotalSlopeLength / (30 * 12), 0)
    Else
    'first slope of roof from right to left
        'get starting and maximum values from e3 Columns
        For Each Member In b.e3Columns
            If Member.CL > 0 And Member.CL < 16 And Not Member.mType Like "*Extension*" Then
                rEdgeStart = Member.lEdgePosition
                bEdgeStart = b.DistanceToRoof("e3", Member.lEdgePosition)
            End If
        Next Member
        lEdgeMax = b.bWidth * 12 / 2
        tEdgeMax = b.DistanceToRoof("e3", b.bWidth * 12 / 2)
        For Each Member In b.e3Rafters
            If Member.rEdgePosition > 0 And Member.RafterLeftEdge <= b.bWidth * 12 / 2 And Not Member.mType Like "*Extension*" And Not Member.mType Like "*Overhang*" Then
                TotalSlopeLength = TotalSlopeLength + Member.Length
            End If
        Next Member
        RafterNum = Application.WorksheetFunction.RoundUp(TotalSlopeLength / (30 * 12), 0)
    End If
        For i = 1 To RafterNum
            Set Member = New clsMember
            Member.Size = Size
            Member.Width = Width
            Member.Placement = "e3 Overhang Rafter"
            Member.bEdgeHeight = bEdgeStart
            Member.rEdgePosition = rEdgeStart
            If i <> RafterNum Then
                Member.Length = 30 * 12
                HorizontalDistance = (Member.Length / (Sqr((b.rPitch / 12) ^ 2 + 1)))
                Member.RafterLeftEdge = rEdgeStart + HorizontalDistance
                Member.tEdgeHeight = b.DistanceToRoof("e3", rEdgeStart + HorizontalDistance)
            Else
                Member.Length = TotalSlopeLength - (30 * 12 * (i - 1))
                Member.RafterLeftEdge = lEdgeMax
                Member.tEdgeHeight = tEdgeMax
            End If
            bEdgeStart = Member.tEdgeHeight
            rEdgeStart = Member.RafterLeftEdge
            Angle = Atn(b.rPitch / 12) * (Member.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (Member.Width / 2) ^ 2)
            Member.bEdgeHeight = Member.bEdgeHeight - DistanceToLower
            Member.tEdgeHeight = Member.tEdgeHeight - DistanceToLower
            b.e3Rafters.Add Member
        Next i
    If b.rShape = "Gable" Then ' do second slope of roof from left to right
        'get starting and maximum values from e1 Columns
        TotalSlopeLength = 0
        For Each Member In b.e3Columns
            If Member.CL > b.bWidth * 12 - 16 And Member.CL < b.bWidth * 12 And Not Member.mType Like "*Extension*" Then
                lEdgeStart = Member.rEdgePosition
                bEdgeStart = b.DistanceToRoof("e3", Member.rEdgePosition)
            End If
        Next Member
        rEdgeMax = b.bWidth * 12 / 2
        tEdgeMax = b.DistanceToRoof("e3", b.bWidth * 12 / 2)
        For Each Member In b.e3Rafters
            If Member.RafterLeftEdge < b.bWidth * 12 And Member.rEdgePosition >= b.bWidth * 12 / 2 And Not Member.Placement Like "*Overhang*" And Not Member.mType Like "*Extension*" Then
                TotalSlopeLength = TotalSlopeLength + Member.Length
            End If
        Next Member
        RafterNum = Application.WorksheetFunction.RoundUp(TotalSlopeLength / (30 * 12), 0)
        For i = 1 To RafterNum
            Set Member = New clsMember
            Member.Size = Size
            Member.Width = Width
            Member.Placement = "e3 Overhang Rafter"
            Member.bEdgeHeight = bEdgeStart
            Member.RafterLeftEdge = lEdgeStart
            If i <> RafterNum Then
                Member.Length = 30 * 12
                HorizontalDistance = (Member.Length / (Sqr((b.rPitch / 12) ^ 2 + 1)))
                Member.rEdgePosition = lEdgeStart - HorizontalDistance
                Member.tEdgeHeight = b.DistanceToRoof("e3", lEdgeStart - HorizontalDistance)
            Else
                Member.Length = TotalSlopeLength - (30 * 12 * (i - 1))
                Member.rEdgePosition = rEdgeMax
                Member.tEdgeHeight = tEdgeMax
            End If
            bEdgeStart = Member.tEdgeHeight
            lEdgeStart = Member.rEdgePosition
            Angle = Atn(b.rPitch / 12) * (Member.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (Member.Width / 2) ^ 2)
            Member.bEdgeHeight = Member.bEdgeHeight - DistanceToLower
            Member.tEdgeHeight = Member.tEdgeHeight - DistanceToLower
            b.e3Rafters.Add Member
        Next i
    End If
        'create overhang rafters for extension/Overhang intersections
    's2 intersection(s); only 2 possible members:
        'if extension intersection only --> create overhang rafter at intersection
        'if overhang on s2 exists --> create overhang rafter at overhang intersection
    If b.s4e3ExtensionIntersection Then
        For Each Member In b.e3Rafters
            If Member.Placement = "s4e3 Extension Intersection Rafter" Then
                If Member.Length <= 30 Then
                    Set NewMember = New clsMember
                    NewMember.rEdgePosition = Member.rEdgePosition
                    NewMember.tEdgeHeight = Member.tEdgeHeight
                    NewMember.bEdgeHeight = Member.bEdgeHeight
                    NewMember.RafterLeftEdge = Member.RafterLeftEdge
                    NewMember.Length = Member.Length
                    NewMember.Size = Size
                    NewMember.Width = Width
                    NewMember.Placement = "s4e3 Extension Intersection e3 Overhang Rafter"
                    b.e3Rafters.Add NewMember
                Else
                    'create as many rafters as necessary to cover the space of the large Eave Extension
                    RafterNum = Application.WorksheetFunction.RoundUp(Member.Length / (30 * 12), 0)
                    Pitch = b.s4ExtensionPitch
                    For i = 1 To RafterNum
                        If i = 1 Then
                            Set NewMember = New clsMember
                            NewMember.Length = 30 * 12
                            NewMember.rEdgePosition = Member.rEdgePosition
                            If b.rShape = "Gable" Then
                                NewMember.tEdgeHeight = Member.tEdgeHeight
                                HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                                NewMember.bEdgeHeight = NewMember.tEdgeHeight - Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                                bEdgeStart = NewMember.bEdgeHeight
                            Else
                                NewMember.bEdgeHeight = Member.bEdgeHeight
                                HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                                NewMember.tEdgeHeight = NewMember.bEdgeHeight + Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                                bEdgeStart = NewMember.tEdgeHeight
                            End If
                            NewMember.RafterLeftEdge = NewMember.rEdgePosition + HorizontalDistance
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s4e3 Extension Intersection e3 Overhang Rafter " & i
                            lEdgeStart = NewMember.RafterLeftEdge
                            b.e3Rafters.Add NewMember
                        ElseIf i = RafterNum Then
                            Set NewMember = New clsMember
                            NewMember.rEdgePosition = lEdgeStart
                            If b.rShape = "Gable" Then
                                NewMember.tEdgeHeight = bEdgeStart
                                NewMember.bEdgeHeight = Member.bEdgeHeight
                            Else
                                NewMember.tEdgeHeight = Member.tEdgeHeight
                                NewMember.bEdgeHeight = bEdgeStart
                            End If
                            NewMember.RafterLeftEdge = Member.RafterLeftEdge
                            NewMember.Length = Member.Length - (30 * 12 * (i - 1))
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s4e3 Extension Intersection e3 Overhang Rafter " & i
                            b.e3Rafters.Add NewMember
                        Else
                            Set NewMember = New clsMember
                            NewMember.Length = 30 * 12
                            NewMember.rEdgePosition = lEdgeStart
                            If b.rShape = "Gable" Then
                                NewMember.tEdgeHeight = bEdgeStart
                                HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                                NewMember.bEdgeHeight = NewMember.tEdgeHeight - Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                                bEdgeStart = NewMember.bEdgeHeight
                            Else
                                NewMember.bEdgeHeight = bEdgeStart
                                HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                                NewMember.tEdgeHeight = NewMember.bEdgeHeight + Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                                bEdgeStart = NewMember.tEdgeHeight
                            End If
                            NewMember.RafterLeftEdge = NewMember.rEdgePosition + HorizontalDistance
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s4e3 Extension Intersection e3 Overhang Rafter " & i
                            lEdgeStart = NewMember.RafterLeftEdge
                            b.e3Rafters.Add NewMember
                        End If
                    Next i
                End If
            ElseIf Member.Placement = "s4 Stub Rafter at s2e1 intersection" Then
                Set NewMember = New clsMember
                NewMember.rEdgePosition = Member.rEdgePosition
                NewMember.tEdgeHeight = Member.tEdgeHeight
                NewMember.bEdgeHeight = Member.bEdgeHeight
                NewMember.RafterLeftEdge = Member.RafterLeftEdge
                NewMember.Length = Member.Length
                NewMember.Size = Size
                NewMember.Width = Width
                NewMember.Placement = "s4e3 Overhang Intersection e3 Overhang Rafter"
                b.e3Rafters.Add NewMember
            End If
        Next Member
    End If
    's2 intection; only 2 possible members (same as above)
    If b.s2e3ExtensionIntersection Then
        For Each Member In b.e3Rafters
            If Member.Placement = "s2e3 Extension Intersection Rafter" Then
                If Member.Length <= 30 Then
                    Set NewMember = New clsMember
                    NewMember.rEdgePosition = Member.rEdgePosition
                    NewMember.tEdgeHeight = Member.tEdgeHeight
                    NewMember.bEdgeHeight = Member.bEdgeHeight
                    NewMember.RafterLeftEdge = Member.RafterLeftEdge
                    NewMember.Length = Member.Length
                    NewMember.Size = Size
                    NewMember.Width = Width
                    NewMember.Placement = "s2e3 Extension Intersection e3 Overhang Rafter"
                    b.e3Rafters.Add NewMember
                Else
                    'create as many rafters as necessary to cover the space of the large Eave Extension
                    RafterNum = Application.WorksheetFunction.RoundUp(Member.Length / (30 * 12), 0)
                    Pitch = b.s2ExtensionPitch
                    For i = 1 To RafterNum
                        If i = 1 Then
                            Set NewMember = New clsMember
                            NewMember.Length = 30 * 12
                            NewMember.RafterLeftEdge = Member.RafterLeftEdge
                            NewMember.tEdgeHeight = Member.tEdgeHeight
                            HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                            NewMember.bEdgeHeight = NewMember.tEdgeHeight - Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                            NewMember.rEdgePosition = NewMember.RafterLeftEdge - HorizontalDistance
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e1 Extension Intersection e3 Overhang Rafter " & i
                            bEdgeStart = NewMember.bEdgeHeight
                            rEdgeStart = NewMember.rEdgePosition
                            b.e3Rafters.Add NewMember
                        ElseIf i = RafterNum Then
                            Set NewMember = New clsMember
                            NewMember.RafterLeftEdge = rEdgeStart
                            NewMember.tEdgeHeight = bEdgeStart
                            NewMember.bEdgeHeight = Member.bEdgeHeight
                            NewMember.rEdgePosition = Member.rEdgePosition
                            NewMember.Length = Member.Length - (30 * 12 * (i - 1))
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e1 Extension Intersection e3 Overhang Rafter " & i
                            b.e3Rafters.Add NewMember
                        Else
                            Set NewMember = New clsMember
                            NewMember.Length = 30 * 12
                            NewMember.RafterLeftEdge = rEdgeStart
                            NewMember.tEdgeHeight = bEdgeStart
                            HorizontalDistance = (NewMember.Length / (Sqr((Pitch / 12) ^ 2 + 1)))
                            NewMember.bEdgeHeight = NewMember.tEdgeHeight - Sqr((NewMember.Length ^ 2) - (HorizontalDistance ^ 2))
                            NewMember.rEdgePosition = NewMember.RafterLeftEdge - HorizontalDistance
                            NewMember.Size = Size
                            NewMember.Width = Width
                            NewMember.Placement = "s2e3 Extension Intersection e3 Overhang Rafter " & i
                            bEdgeStart = NewMember.bEdgeHeight
                            rEdgeStart = NewMember.rEdgePosition
                            b.e3Rafters.Add NewMember
                        End If
                    Next i
                End If
            ElseIf Member.Placement = "s2 Stub Rafter at s4e1 intersection" Then
                Set NewMember = New clsMember
                NewMember.rEdgePosition = Member.rEdgePosition
                NewMember.tEdgeHeight = Member.tEdgeHeight
                NewMember.bEdgeHeight = Member.bEdgeHeight
                NewMember.RafterLeftEdge = Member.RafterLeftEdge
                NewMember.Length = Member.Length
                NewMember.Size = Size
                NewMember.Width = Width
                NewMember.Placement = "s2e1 Overhang Intersection e1 Overhang Rafter"
                b.e1Rafters.Add NewMember
            End If
        Next Member
    End If
End If

End Sub


Sub AdjustEndwallColumns(b As clsBuilding, eWall As String)

Dim ColumnCollection As Collection
Dim RafterCollection As Collection
Dim Column As clsMember
Dim Rafter As clsMember
Dim RafterWidth As Double
Dim tEdgeDifference As Double
Dim Angle As Double
Dim DistanceToLower As Double
Dim WedgeDistance As Double
Dim FirstColWidth As Double
Dim LastColWidth As Double

Select Case eWall
Case "e1"
    Set RafterCollection = b.e1Rafters
    Set ColumnCollection = b.e1Columns
    FirstColWidth = b.s4ColumnWidth
    LastColWidth = b.s2ColumnWidth
Case "e3"
    Set RafterCollection = b.e3Rafters
    Set ColumnCollection = b.e3Columns
    FirstColWidth = b.s2ColumnWidth
    LastColWidth = b.s4ColumnWidth
Case "Int"
    Set RafterCollection = b.intRafters
    Set ColumnCollection = b.InteriorColumns
    FirstColWidth = b.s4ColumnWidth
    LastColWidth = b.s2ColumnWidth
End Select

For Each Column In ColumnCollection
    If Column.lEdgePosition <> b.bWidth * 12 And Column.rEdgePosition <> 0 And Column.LoadBearing = True Then
        'extend each load bearing column to account for angle cut
        Column.tEdgeHeight = Column.tEdgeHeight + ((Column.Width / 2) * b.rPitch / 12)
        Column.Length = Column.tEdgeHeight
    ElseIf Column.LoadBearing = False Then
        'non load bearing columns need to be lowered to bottom of rafter
        For Each Rafter In RafterCollection
            If Rafter.rEdgePosition <= Column.CL And Rafter.RafterLeftEdge >= Column.CL Then
                RafterWidth = Rafter.Width
            End If
        Next Rafter
        Angle = Atn(b.rPitch / 12) * (RafterWidth / 2)
        tEdgeDifference = Sqr(Angle ^ 2 + (RafterWidth / 2) ^ 2)
        Column.tEdgeHeight = Column.tEdgeHeight - tEdgeDifference * 2 + ((Column.Width / 2) * b.rPitch / 12)
        Column.Length = Column.Length - tEdgeDifference * 2 + ((Column.Width / 2) * b.rPitch / 12)
        If Column.LoadBearing = True And (Column.CL = b.bWidth * 12 / 2 And b.rShape = "Gable") Then
            Column.Placement = Column.Placement & "cut Vee for center column at " & Application.WorksheetFunction.Round(Angle, 2) & " degree angles, "
        Else
            Column.Placement = Column.Placement & "cut at " & Application.WorksheetFunction.Round(Angle, 2) & " degree angle, "
        End If
    ElseIf Column.lEdgePosition = b.bWidth * 12 Then
        If (b.rShape = "Single Slope" And eWall = "e3") Then
            'skip high side eave columns
        Else
        'corner columns on expandable endwalls need to be longer to account for angle cut
        'non-expandable endwalls only need to be longer if distance is greater than 4"
            WedgeDistance = LastColWidth * b.rPitch / 12
            If eWall <> "Int" Then
                If b.ExpandableEndwall(eWall) Or WedgeDistance > 4 Then
                    Column.tEdgeHeight = Column.tEdgeHeight + WedgeDistance
                    Column.Length = Column.Length + WedgeDistance
                End If
            Else
                Column.tEdgeHeight = Column.tEdgeHeight + WedgeDistance
                Column.Length = Column.Length + WedgeDistance
            End If
        End If
    ElseIf Column.rEdgePosition = 0 Then
        If (b.rShape = "Single Slope" And (eWall = "e1" Or eWall = "Int")) Then
            'skip high side eave columns
        Else
        'corner columns on expandable endwalls need to be longer to account for angle cut
        'non-expandable endwalls only need to be longer if distance is greater than 4"
            WedgeDistance = FirstColWidth * b.rPitch / 12
            If eWall <> "Int" Then
                If b.ExpandableEndwall(eWall) Or WedgeDistance > 4 Then
                    Column.tEdgeHeight = Column.tEdgeHeight + WedgeDistance
                    Column.Length = Column.Length + WedgeDistance
                End If
            Else
                Column.tEdgeHeight = Column.tEdgeHeight + WedgeDistance
                Column.Length = Column.Length + WedgeDistance
            End If
        End If
    End If
Next Column


End Sub

Sub RemoveEndwallColumns(b As clsBuilding, eWall As String)

Dim FOCollection As Collection
Dim ColumnCollection As Collection
Dim NearestMemberRight As Double
Dim NearestMemberLeft As Double
Dim tempNearestMember As Double
Dim Column As clsMember
Dim FO As clsFO
Dim Jamb As Object
Dim ColIndex As Integer
Dim tempColumn As clsMember




Select Case eWall
Case "e1"
    Set FOCollection = b.e1FOs
    Set ColumnCollection = b.e1Columns
Case "e3"
    Set FOCollection = b.e3FOs
    Set ColumnCollection = b.e3Columns
End Select

For ColIndex = 1 To ColumnCollection.Count - 1
    Set Column = ColumnCollection(ColIndex)
    If Column.LoadBearing = False Then
        For Each FO In FOCollection
            If FO.rEdgePosition < Column.CL And FO.lEdgePosition > Column.CL Then
                Column.DeleteFlag = True
            End If
        Next FO
        If Column.DeleteFlag = False Then
            NearestMemberLeft = b.bWidth * 12
            NearestMemberRight = 0
            For Each tempColumn In ColumnCollection
                If (tempColumn.CL > Column.CL And tempColumn.CL < NearestMemberLeft) And tempColumn.DeleteFlag = False Then
                    NearestMemberLeft = tempColumn.CL
                End If
                If (tempColumn.CL < Column.CL And tempColumn.CL > NearestMemberRight) And tempColumn.DeleteFlag = False Then
                    NearestMemberRight = tempColumn.CL
                End If
            Next tempColumn
            For Each FO In FOCollection
                For Each Jamb In FO.FOMaterials
                    If Jamb.clsType = "Member" Then
                        If Jamb.CL > Column.CL And Jamb.CL < NearestMemberLeft Then
                            NearestMemberLeft = Jamb.CL
                        End If
                        If Jamb.CL < Column.CL And Jamb.CL > NearestMemberRight Then
                            NearestMemberRight = Jamb.CL
                        End If
                    End If
                Next Jamb
            Next FO
            If Abs(NearestMemberLeft - NearestMemberRight) < 30 * 12 Then
                ColumnCollection(ColIndex).DeleteFlag = True
            Else
                ColumnCollection(ColIndex).DeleteFlag = False
            End If
        End If
    End If
Next ColIndex

For ColIndex = ColumnCollection.Count To 1 Step -1
    If ColumnCollection(ColIndex).DeleteFlag = True Then
        ColumnCollection.Remove (ColIndex)
    End If
Next ColIndex

End Sub

Sub CutListOutput(Collection As Collection, Label As String)



Dim LastRow As Integer
Dim Member As clsMember
Dim SteelSht As Worksheet
Dim FullMemberSht As Worksheet
Dim FO As clsFO
Dim item As Object
Dim j As Double
Dim UnitPrice As Double
Dim UnitMeasure As String
Dim UnitValue As Double
Dim PriceTbl As ListObject


'''''''''''''''''''Full Member List Sheet
Set FullMemberSht = ThisWorkbook.Sheets("Optimized Cut List")

If FullMemberSht.Range("E4").Value = "" Then
    LastRow = 4
Else
    LastRow = FullMemberSht.Range("E3").End(xlDown).offset(1, 0).Row
End If

j = 1


    'Call DuplicateMaterialRemoval(Collection, "Steel")
    For Each Member In Collection
        With FullMemberSht
            .Range("A" & LastRow).Value = Member.Placement
            .Range("C" & LastRow).Value = "Total Span Length:"
            .Range("E" & LastRow).Value = ImperialMeasurementFormat(Member.Length)
            'Formatting
            .Range("A" & LastRow, "E" & LastRow).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            .Range("A" & LastRow, "E" & LastRow).Font.Bold = True
            .Range("C" & LastRow).HorizontalAlignment = xlRight
            .Range("E" & LastRow).HorizontalAlignment = xlLeft
            .Rows(LastRow).RowHeight = 30
            j = j + 1
            LastRow = LastRow + 1
        End With
        For Each item In Member.ComponentMembers
        With FullMemberSht
            .Range("B" & LastRow).Value = item.Qty
            .Range("C" & LastRow).Value = item.Size
            If item.Placement = "" Then
                .Range("D" & LastRow).Value = item.mType
            Else
                .Range("D" & LastRow).Value = item.Placement
            End If
            .Range("E" & LastRow).Value = ImperialMeasurementFormat(item.Length)
        End With
        LastRow = LastRow + 1
        Next item
    Next Member


End Sub

'' function returns string of the nearest available Member Size
Function NearestMemberSize(Length As Variant, Optional Direction As Integer, Optional MemberType As String, Optional NumericOutput As Boolean) As Variant
'DESCRIPTION: Function returns the nearest value to a target
'INPUT: Pass the function a range of cells, a target value that you want to find a number closest to
' and an optional direction variable described below.
'OPTIONS: Set the optional variable Direction equal to 0 or blank to find the closest value
' Set equal to -1 to find the closest value below your target
' set equal to 1 to find the closest value above your target
Dim t As Variant
Dim u As Variant
Dim Members() As Variant
Dim Member As Variant
Dim mSize As Double
Dim NearestMemberSizeString As String
Dim UniqueMemberType As String

'
If MemberType = "C Purlin" Then
    Members = Array(20 * 12, 25 * 12, 30 * 12)
ElseIf MemberType = "TS" Then
    Members = Array(20 * 12, 30 * 12, 40 * 12)
ElseIf MemberType = "W Beam" Then
    Members = Array(20 * 12, 25 * 12, 30 * 12, 35 * 12, 40 * 12, 45 * 12, 50 * 12, 60 * 12)
End If

t = 1.79769313486231E+308 'initialize
For Each Member In Members
    If IsNumeric(Member) Then
        u = Abs(Member - Length)
        If Direction > 0 And Member >= Length Then
            'only report if closer number is greater than the target
            If u < t Then
                t = u
                mSize = Member
            End If
        ElseIf Direction < 0 And Member <= Length Then
            'only report if closer number is less than the target
            If u < t Then
                t = u
                mSize = Member
            End If
        ElseIf Direction = 0 Then
            If u < t Then
                t = u
                mSize = Member
            End If
        End If
    End If
Next Member

'return available Member name
NearestMemberSizeString = MaterialsListGen.ImperialMeasurementFormat(mSize)
'output
If NumericOutput = False Then
    NearestMemberSize = NearestMemberSizeString
ElseIf NumericOutput = True Then
    NearestMemberSize = mSize
End If
 
End Function

 Sub SteelPriceOutput(Collection As Collection, Label As String, Optional FOMode As Boolean)

Dim LastRow As Integer
Dim Member As clsMember
Dim SteelSht As Worksheet
Dim FullMemberSht As Worksheet
Dim FO As clsFO
Dim item As Object
Dim j As Double
Dim UnitPrice As Double
Dim UnitMeasure As String
Dim UnitValue As Double
Dim PriceTbl As ListObject

'''''''''''''''''''''''''''Steel Material OUtput Sheet
Set SteelSht = ThisWorkbook.Sheets("Structural Steel Price List")
Set PriceTbl = ThisWorkbook.Worksheets("Master Price List").ListObjects("SteelPriceListTbl")

If SteelSht.Range("A4").Value = "" Then
    LastRow = 4
Else
    LastRow = SteelSht.Range("A3").End(xlDown).offset(1, 0).Row
End If

If FOMode = False Then
    Call DuplicateMaterialRemoval(Collection, "Steel")
    For Each Member In Collection
        'lookup price information in Price Table
        'check for errors
        If IsError(Application.VLookup(Member.Size, PriceTbl.Range, 2, False)) = True And Not Member.Size Like "W*" And Not Member.Size Like "*eave strut*" Then
            UnitPrice = 0
            UnitMeasure = "Unknown"
            UnitValue = 0
        Else    'successful lookup
            If Member.Size Like "W*" Then
                UnitPrice = Application.WorksheetFunction.VLookup("W--x--", PriceTbl.Range, 2, False)
                UnitMeasure = "per lb"
                UnitValue = (Member.Length / 12) * Right(Member.Size, 2) * UnitPrice
            ElseIf Member.Size Like "*eave strut*" Then
                UnitPrice = Application.WorksheetFunction.VLookup("Eave Struts", PriceTbl.Range, 2, False)
                UnitMeasure = "per ft"
                UnitValue = (Member.Length / 12) * UnitPrice
            Else
                UnitPrice = Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 2, False)
                UnitMeasure = Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 3, False)
                If UnitMeasure = "per ft" Then
                    UnitValue = UnitPrice * (Member.Length / 12)
                ElseIf UnitMeasure = "per lb" Then
                    UnitValue = UnitPrice * (Member.Length / 12) * Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 4, False)
                End If
            End If
        End If
        With SteelSht
            .Range("A" & LastRow).Value = Member.Qty
            .Range("B" & LastRow).Value = Member.Size
            .Range("C" & LastRow).Value = Label
            .Range("D" & LastRow).Value = ImperialMeasurementFormat(Member.Length)
            If UnitPrice = 0 Or UnitValue = 0 Then
                .Range("E" & LastRow).Value = "Unknown"
                .Range("F" & LastRow).Value = "Unknown"
                .Range("G" & LastRow).Value = "Unknown"
                .Range("H" & LastRow).Value = "Item not found"
            Else
                .Range("E" & LastRow).Value = UnitPrice
                .Range("F" & LastRow).Value = UnitMeasure
                .Range("G" & LastRow).Value = UnitValue
                .Range("H" & LastRow).Value = UnitValue * Member.Qty
            End If
        End With
        LastRow = LastRow + 1
        'write cost to member
        Member.ItemCost = UnitValue
    Next Member
Else
    For Each FO In Collection
        Label = FO.FOType & " " & FO.Wall
        Call DuplicateMaterialRemoval(FO.FOMaterials, "Steel")
        For Each item In FO.FOMaterials
            If item.clsType = "Member" Then
                Set Member = item
                    'check for errors
                    If IsError(Application.VLookup(Member.Size, PriceTbl.Range, 2, False)) = True Then
                        UnitPrice = "Unknown"
                        UnitMeasure = "Unknown"
                        UnitValue = "Item Not Found"
                    Else    'successful lookup
                        If Member.Size Like "W*" Then
                            UnitPrice = Application.WorksheetFunction.VLookup("W--x--", PriceTbl.Range, 3, False)
                            UnitMeasure = "per lb"
                            UnitValue = (Member.Length / 12) * Right(Member.Size, 2) * UnitPrice
                        Else
                            UnitPrice = Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 2, False)
                            UnitMeasure = Application.WorksheetFunction.VLookup(Member.Size, PriceTbl.Range, 3, False)
                            If UnitMeasure = "per ft" Then
                                UnitValue = UnitPrice * (Member.Length / 12)
                            ElseIf UnitMeasure = "per lb" Then
                                UnitValue = UnitPrice * (Member.Length / 12)
                            End If
                        End If
                    End If
                    With SteelSht
                        .Range("A" & LastRow).Value = Member.Qty
                        .Range("B" & LastRow).Value = Member.Size
                        .Range("C" & LastRow).Value = Label
                        .Range("D" & LastRow).Value = ImperialMeasurementFormat(Member.Length)
                        .Range("E" & LastRow).Value = UnitPrice
                        .Range("F" & LastRow).Value = UnitMeasure
                        .Range("G" & LastRow).Value = UnitValue
                        .Range("H" & LastRow).Value = UnitValue * Member.Qty
                    End With
                LastRow = LastRow + 1
                'add cost
                item.ItemCost = UnitValue
            End If
        Next item
    Next FO
End If

Dim mCell As Range
Dim MissingPrice As Boolean
'Check SS Price List for missing prices
With SteelSht
    LastRow = .Cells(.Rows.Count, "H").End(xlUp).Row
    If LastRow = 3 Then LastRow = 4
    For Each mCell In .Range("H4:H" & LastRow)
        If IsNumeric(mCell.Value) = False Then
            MissingPrice = True
            Exit For
        End If
    Next mCell
End With


If MissingPrice Then MsgBox "At least one price could  not be found on the structrual steel list.", vbExclamation, "Missing Price"

                
End Sub

 Sub RoofPurlinGen(b As clsBuilding)

Dim RafterCollection As Collection
Dim InteriorColumnCollection As Collection
Dim s2EaveStrutCollection As Collection
Dim s4EaveStrutCollection As Collection
Dim RoofPurlinCollection As Collection
Dim Purlins() As Double
Dim Rafter As clsMember
Dim IntColumn As clsMember
Dim EaveStrut As clsMember
Dim Purlin As clsMember
Dim RafterNum As Integer
Dim BayNum As Integer
Dim i As Integer
Dim j As Integer
Dim StartPos As Double
Dim MaxPos As Double
Dim BayLength As Double
Dim Overhang As Boolean

Set RoofPurlinCollection = b.RoofPurlins
Set InteriorColumnCollection = b.InteriorColumns
Set RafterCollection = New Collection
Set s2EaveStrutCollection = New Collection
Set s4EaveStrutCollection = New Collection

'combine rafter collections
For Each Rafter In b.e1Rafters
    RafterCollection.Add Rafter
Next Rafter
For Each Rafter In b.e3Rafters
    RafterCollection.Add Rafter
Next Rafter
For Each Rafter In b.intRafters
    RafterCollection.Add Rafter
Next Rafter

'get Eave Struts into new collection
For Each EaveStrut In b.s2Girts
    If EaveStrut.mType = "Eave Strut" Then
        s2EaveStrutCollection.Add EaveStrut
    End If
Next EaveStrut
For Each EaveStrut In b.s4Girts
    If EaveStrut.mType = "Eave Strut" Then
        s4EaveStrutCollection.Add EaveStrut
    End If
Next EaveStrut

If b.rShape = "Gable" Then
    'calculate 1 side first, then duplicate it.
    RafterNum = Application.WorksheetFunction.RoundUp((b.RafterLength - 12) / 60, 0)
    ReDim Purlins(RafterNum - 1) As Double 'Eave Strut (1) is already made w/ girts
Else
    RafterNum = Application.WorksheetFunction.RoundUp((b.RafterLength / 60) - 1, 0)
    ReDim Purlins(RafterNum - 2) As Double 'Eave Struts (2) are already made w/ girts
End If

Purlins(0) = 0
For i = 1 To UBound(Purlins)
    If i <> UBound(Purlins) Then
        Purlins(i) = 60 + (60 * i)
    Else
        Purlins(i) = b.RafterLength - 12
    End If
Next i

BayNum = EstSht.Range("BayNum").Value
StartPos = 0

'extend roof through endwall overhangs and extension
'eave overhangs and extensions will be handled separately, along with any intersections

'Notes:

'First, Set start pos either at 0 or (negative) e1 extension
    'set end pos either to b.width or e3 extension
'Next, if overhang is present, adjust starting pos or ending pos
'Create roof purlin
    'if eave strut, handle case
'next purlin



For i = 0 To UBound(Purlins) 'for each row of purlins
    For j = 1 To BayNum 'for each bay
        Overhang = False
        BayLength = EstSht.Range("Bay1_Length").offset(j - 1, 0).Value * 12
        'define bay length
            '^^^ orientation is from the relative "perspective" of s2
        'adjust based on overhangs and extensions
        If j = 1 Then
            If b.e1Extension > 0 Then
                'Create extra purlin for e1 Extension section
                BayLength = EstSht.Range("e1_GableExtension").Value * 12
                StartPos = -EstSht.Range("e1_GableExtension").Value * 12
                If b.e1Overhang > 0 Then
                    BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value * 12
                    StartPos = StartPos - EstSht.Range("e1_GableOverhang").Value * 12
                End If
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = StartPos
                Purlin.tEdgeHeight = Purlins(i)
                If i <> 0 Then 'normal purlins
                    If BayLength > 25 * 12 Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                Else 'eave strut at s2 e1 gable extension
                    Purlin.mType = "Eave Strut"
                    If b.rPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    ElseIf b.e1GableExtensionSoffit Then
                        Purlin.Size = "8"" " & b.rPitch & ":12 double up eave strut"
                    Else
                        Purlin.Size = "8"" " & b.rPitch & ":12 single up eave strut"
                    End If
                    Purlin.tEdgeHeight = b.bHeight * 12
                    Purlin.bEdgeHeight = b.bHeight * 12
                End If
                Purlin.Placement = "endwall 1 extension roof purlin, " & Application.WorksheetFunction.RoundUp(Purlins(i), 2) & """ above eave strut"
                If i = 0 Then
                    b.s2Girts.Add Purlin
                Else
                    RoofPurlinCollection.Add Purlin
                    b.WeldClips = b.WeldClips + 2
                End If
                BayLength = 0
                StartPos = 0
            ElseIf EstSht.Range("e1_GableOverhang").Value > 0 Then
                BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value * 12
                StartPos = -EstSht.Range("e1_GableOverhang").Value * 12
                Overhang = True
            End If
        ElseIf j = BayNum Then
            If EstSht.Range("e3_GableExtension").Value > 0 Then
                'create extra purlin for e3 Extension Section
                BayLength = EstSht.Range("e3_GableExtension").Value * 12
                If b.e1Overhang > 0 Then
                    BayLength = BayLength + EstSht.Range("e3_GableOverhang").Value * 12
                End If
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = b.bLength * 12
                Purlin.tEdgeHeight = Purlins(i)
                If i <> 0 Then 'normal purlins
                    If BayLength > 25 * 12 Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                Else 'eave strut at s2 e3 gable extension
                    Purlin.mType = "Eave Strut"
                    If b.rPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    ElseIf b.e3GableExtensionSoffit Then
                        Purlin.Size = "8"" " & b.rPitch & ":12 double up eave strut"
                    Else
                        Purlin.Size = "8"" " & b.rPitch & ":12 single up eave strut"
                    End If
                    Purlin.tEdgeHeight = b.bHeight * 12
                    Purlin.bEdgeHeight = b.bHeight * 12
                End If
                Purlin.Placement = "endwall 3 extension roof purlin, " & Application.WorksheetFunction.RoundUp(Purlins(i), 2) & """ above eave strut"
                If i = 0 Then
                    b.s2Girts.Add Purlin
                Else
                    RoofPurlinCollection.Add Purlin
                    b.WeldClips = b.WeldClips + 2
                End If
                BayLength = 0
            ElseIf EstSht.Range("e3_GableOverhang").Value > 0 Then
                BayLength = BayLength + EstSht.Range("e3_GableOverhang").Value * 12
                Overhang = True
            End If
        End If
        If BayLength = 0 Then
            BayLength = EstSht.Range("Bay1_Length").offset(j - 1, 0).Value * 12
        End If
        Set Purlin = New clsMember
        Purlin.mType = "Roof Purlin"
        Purlin.Length = BayLength
        Purlin.rEdgePosition = StartPos
        Purlin.tEdgeHeight = Purlins(i)
        If BayLength > 25 * 12 And Overhang = False Then
            Purlin.Size = "10"" C Purlin"
        Else
            Purlin.Size = "8"" C Purlin"
        End If
        Purlin.Placement = "sidewall 2 roof purlin, " & Application.WorksheetFunction.RoundUp(Purlins(i), 2) & """ above eave strut, bay number " & j
        RoofPurlinCollection.Add Purlin
        b.WeldClips = b.WeldClips + 2
        StartPos = StartPos + BayLength
    Next j
Next i

StartPos = 0

If b.rShape = "Gable" Then
'duplicate purlins going opposite direction
For i = 0 To UBound(Purlins) 'for each row of purlins
    For j = BayNum To 1 Step -1 'for each bay, going from e3 to e1
        Overhang = False
        BayLength = EstSht.Range("Bay1_Length").offset(j - 1, 0).Value * 12
        'define bay length
            '^^^ orientation is from the relative "perspective" of s4
        'adjust based on overhangs and extensions
        If j = BayNum Then
            If EstSht.Range("e3_GableExtension").Value > 0 Then
                'Create extra purlin for e3 extension
                BayLength = EstSht.Range("e3_GableExtension").Value * 12
                StartPos = -EstSht.Range("e3_GableExtension").Value * 12
                If EstSht.Range("e3_GableOverhang").Value > 0 Then
                    BayLength = BayLength + EstSht.Range("e3_GableOverhang").Value * 12
                    StartPos = StartPos - EstSht.Range("e3_GableOverhang").Value * 12
                End If
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = StartPos
                Purlin.tEdgeHeight = Purlins(i)
                If i <> 0 Then 'normal purlins
                    If BayLength > 25 * 12 Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                Else 'eave strut at s4 e3 gable extension
                    Purlin.mType = "Eave Strut"
                    If b.rPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    ElseIf b.e3GableExtensionSoffit Then
                        Purlin.Size = "8"" " & b.rPitch & ":12 double up eave strut"
                    Else
                        Purlin.Size = "8"" " & b.rPitch & ":12 single up eave strut"
                    End If
                    Purlin.tEdgeHeight = b.bHeight * 12
                    Purlin.bEdgeHeight = b.bHeight * 12
                End If
                Purlin.Placement = "endwall 3 extension roof purlin, " & Application.WorksheetFunction.RoundUp(Purlins(i), 2) & """ above eave strut"
                If i = 0 Then
                    b.s4Girts.Add Purlin
                Else
                    RoofPurlinCollection.Add Purlin
                    b.WeldClips = b.WeldClips + 2
                End If
                BayLength = 0
                StartPos = 0
            ElseIf EstSht.Range("e3_GableOverhang").Value > 0 Then
                BayLength = BayLength + EstSht.Range("e3_GableOverhang").Value * 12
                StartPos = -EstSht.Range("e3_GableOverhang").Value * 12
                Overhang = True
            End If
        ElseIf j = 1 Then
            If EstSht.Range("e1_GableExtension").Value > 0 Then
                'Create extra purlin for e1 Extension
                BayLength = EstSht.Range("e1_GableExtension").Value * 12
                If EstSht.Range("e1_GableOverhang").Value > 0 Then
                    BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value * 12
                End If
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = b.bLength * 12
                Purlin.tEdgeHeight = Purlins(i)
                If i <> 0 Then 'normal purlins
                    If BayLength > 25 * 12 Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                Else 'eave strut at s4 e1 gable extension
                    Purlin.mType = "Eave Strut"
                    If b.rPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    ElseIf b.e1GableExtensionSoffit Then
                        Purlin.Size = "8"" " & b.rPitch & ":12 double up eave strut"
                    Else
                        Purlin.Size = "8"" " & b.rPitch & ":12 single up eave strut"
                    End If
                    Purlin.tEdgeHeight = b.bHeight * 12
                    Purlin.bEdgeHeight = b.bHeight * 12
                End If
                Purlin.Placement = "endwall 1 extension roof purlin, " & Application.WorksheetFunction.RoundUp(Purlins(i), 2) & """ above eave strut"
                If i = 0 Then
                    b.s4Girts.Add Purlin
                Else
                    RoofPurlinCollection.Add Purlin
                    b.WeldClips = b.WeldClips + 2
                End If
                BayLength = 0
            ElseIf EstSht.Range("e1_GableOverhang").Value > 0 Then
                BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value * 12
                Overhang = True
            End If
        End If
        If BayLength = 0 Then 'start from wall line if overhang hasn't been defined above
            BayLength = EstSht.Range("Bay1_Length").offset(j - 1, 0).Value * 12
        End If
        Set Purlin = New clsMember
        Purlin.mType = "Roof Purlin"
        Purlin.Length = BayLength
        Purlin.rEdgePosition = StartPos
        Purlin.tEdgeHeight = Purlins(i)
        If BayLength > 25 * 12 And Overhang = False Then
            Purlin.Size = "10"" C Purlin"
        Else
            Purlin.Size = "8"" C Purlin"
        End If
        Purlin.Placement = "sidewall 2 roof purlin, " & Application.WorksheetFunction.RoundUp(Purlins(i), 2) & """ above eave strut, bay number " & j
        RoofPurlinCollection.Add Purlin
        b.WeldClips = b.WeldClips + 2
        StartPos = StartPos + BayLength
    Next j
Next i
End If


'''''''''''''''''''''''''''''''''''''''''''Special Case:
'Single slope building w/ e1 or e3 Gable Extension needs Eave struts for Extended section
If b.rShape = "Single Slope" Then
    If b.e1Extension > 0 Then
        ''''''''''''''''''''''''''s4 high side eave gable e1 extension strut
        Set Purlin = New clsMember
        Purlin.mType = "Eave Strut"
        Purlin.Length = EstSht.Range("e1_GableExtension").Value * 12
        If b.e1Overhang > 0 Then
            Purlin.Length = Purlin.Length + EstSht.Range("e1_GableOverhang").Value * 12
        End If
        Purlin.rEdgePosition = b.bLength * 12
        Purlin.tEdgeHeight = b.bHeight * 12 + (b.bWidth * 12 * b.rPitch / 12)
        Purlin.bEdgeHeight = b.bHeight * 12 + (b.bWidth * 12 * b.rPitch / 12)
        If b.rPitch = 1 Then
            Purlin.Size = "8"" C Purlin"
        ElseIf b.e1GableExtensionSoffit Then
            Purlin.Size = "8"" " & b.rPitch & ":12 double down eave strut"
        Else
            Purlin.Size = "8"" " & b.rPitch & ":12 single down eave strut"
        End If
        Purlin.Placement = "sidewall 4 e1 gable extension, " & Application.WorksheetFunction.RoundUp(Purlin.tEdgeHeight, 2) & """ above eave strut "
        Debug.Print "single slope s4 high side eave e1 gable extension strut created"
        Debug.Print EaveStrutCount + 1
        'RoofPurlinCollection.Add Purlin
        b.s4Girts.Add Purlin
    End If
    If b.e3Extension > 0 Then
        ''''''''''''''''''''''''''''''''''''''''''s4 high side eave gable e3 extension strut
        Set Purlin = New clsMember
        Purlin.mType = "Eave Strut"
        Purlin.Length = EstSht.Range("e3_GableExtension").Value * 12
        If b.e3Overhang > 0 Then
            Purlin.Length = Purlin.Length + EstSht.Range("e3_GableOverhang").Value * 12
        End If
        Purlin.rEdgePosition = -Purlin.Length
        Purlin.tEdgeHeight = b.bHeight * 12 + (b.bWidth * 12 * b.rPitch / 12)
        Purlin.bEdgeHeight = b.bHeight * 12 + (b.bWidth * 12 * b.rPitch / 12)
        If b.rPitch = 1 Then
            Purlin.Size = "8"" C Purlin"
        ElseIf b.e3GableExtensionSoffit Then
            Purlin.Size = "8"" " & b.rPitch & ":12 double down eave strut"
        Else
            Purlin.Size = "8"" " & b.rPitch & ":12 single down eave strut"
        End If
        Purlin.Placement = "sidewall 2 e3 extension, " & Application.WorksheetFunction.RoundUp(Purlin.tEdgeHeight, 2) & """ above eave strut "
        Debug.Print "single slope s4 high side eave e4 gable extension strut created"
        Debug.Print EaveStrutCount + 1
        'RoofPurlinCollection.Add Purlin
        b.s4Girts.Add Purlin
    End If
End If

Dim s2ExtensionPurlinNum As Integer
Dim s2ExtensionPurlins() As Double
Dim s2ExtnesionLength As Double
Dim s4ExtensionPurlinNum As Integer
Dim s4ExtensionPurlins() As Double
Dim s4ExtnesionLength As Double

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Eave Extensions
''''''''s2 extension
If EstSht.Range("s2_EaveExtension").Value > 0 Then
    s2ExtensionPurlinNum = Application.WorksheetFunction.RoundUp((b.s2ExtensionRafterLength) / 60 - 1, 0)
    ReDim s2ExtensionPurlins(s2ExtensionPurlinNum) As Double  'Eave Strut (1) is already made
    's2ExtensionLength = b.s2EaveExtensionBuildingLength
    
    s2ExtensionPurlins(0) = 0 'extension eave strut
    For i = 1 To UBound(s2ExtensionPurlins)
        If i <> UBound(s2ExtensionPurlins) Then
            s2ExtensionPurlins(i) = 60 + (60 * i)
        Else
            s2ExtensionPurlins(i) = b.s2ExtensionRafterLength - 12
        End If
    Next i
    For i = 0 To UBound(s2ExtensionPurlins)
        StartPos = 0
        MaxPos = b.bLength * 12
        For j = 1 To BayNum
            BayLength = EstSht.Range("Bay1_Length").offset(j - 1, 0).Value * 12
'            If EstSht.Range("e1_GableOverhang").Value > 0 And j = 1 Then
'                BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value
'                StartPos = -EstSht.Range("e1_GableOverhang").Value
'            ElseIf EstSht.Range("s4_EaveOverhang").Value > 0 And j = BayNum Then
'                BayLength = BayLength + EstSht.Range("s4_EaveOverhang").Value
'            End If
            If i <> 0 Then 'normal purlins
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = StartPos
                Purlin.tEdgeHeight = s2ExtensionPurlins(i)
                If BayLength > 25 * 12 And Overhang = False Then
                    Purlin.Size = "10"" C Purlin"
                Else
                    Purlin.Size = "8"" C Purlin"
                End If
                Purlin.Placement = "sidewall 2 eave extension roof purlin, " & Application.WorksheetFunction.RoundUp(s2ExtensionPurlins(i), 2) & """ above eave strut, bay number " & j
                RoofPurlinCollection.Add Purlin
                b.WeldClips = b.WeldClips + 2
            Else
                Set Purlin = New clsMember
                Purlin.mType = "Eave Strut"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = StartPos
                Purlin.tEdgeHeight = s2ExtensionPurlins(i)
                If b.s2ExtensionPitch = 1 Then
                    Purlin.Size = "8"" C Purlin"
                Else
                    Purlin.Size = "8"" " & b.s2ExtensionPitch & ": 12 double up eave strut"
                End If
                Purlin.Placement = "sidewall 2 eave extension eave strut, " & Application.WorksheetFunction.RoundUp(s2ExtensionPurlins(i), 2) & """ above eave strut, bay number " & j
                RoofPurlinCollection.Add Purlin
            End If
            StartPos = StartPos + BayLength
        Next j
        If EstSht.Range("e1_GableExtension").Value > 0 And b.s2e1ExtensionIntersection = True Then
            BayLength = EstSht.Range("e1_GableExtension").Value * 12
            Set Purlin = New clsMember
            Purlin.mType = "Roof Purlin"
            Purlin.Length = BayLength
            Purlin.rEdgePosition = -EstSht.Range("e1_GableExtension").Value * 12
            Purlin.tEdgeHeight = s2ExtensionPurlins(i)
            If i <> 0 Then 'normal purlins
                If BayLength > 25 * 12 And Overhang = False Then
                    Purlin.Size = "10"" C Purlin"
                Else
                    Purlin.Size = "8"" C Purlin"
                End If
                b.WeldClips = b.WeldClips + 2
            Else
                Purlin.mType = "Eave Strut"
                If b.s2ExtensionPitch = 1 Then
                    Purlin.Size = "8"" C Purlin"
                Else
                    Purlin.Size = "8"" " & b.s2ExtensionPitch & ": 12 double up eave strut"
                End If
                Debug.Print "eave strut e1/s2 intersection created"
                Debug.Print EaveStrutCount + 1
            End If
            Purlin.Placement = "sidewall 2 eave extension e1 intersection roof purlin, " & Application.WorksheetFunction.RoundUp(s2ExtensionPurlins(i), 2) & """ above extension eave strut"
            RoofPurlinCollection.Add Purlin
        End If
        If EstSht.Range("e3_GableExtension").Value > 0 And b.s2e3ExtensionIntersection = True Then
            BayLength = EstSht.Range("e3_GableExtension").Value * 12
            Set Purlin = New clsMember
            Purlin.mType = "Roof Purlin"
            Purlin.Length = BayLength
            Purlin.rEdgePosition = b.bLength * 12
            Purlin.tEdgeHeight = s2ExtensionPurlins(i)
            If i <> 0 Then 'normal purlins
                If BayLength > 25 * 12 And Overhang = False Then
                    Purlin.Size = "10"" C Purlin"
                Else
                    Purlin.Size = "8"" C Purlin"
                End If
                b.WeldClips = b.WeldClips + 2
            Else
                Purlin.mType = "Eave Strut"
                If b.s2ExtensionPitch = 1 Then
                    Purlin.Size = "8"" C Purlin"
                Else
                    Purlin.Size = "8"" " & b.s2ExtensionPitch & ": 12 double up eave strut"
                End If
                Debug.Print "eave strut e3/s2 intersection created"
                Debug.Print EaveStrutCount + 1
            End If
            Purlin.Placement = "sidewall 2 eave extension e3 intersection roof purlin, " & Application.WorksheetFunction.RoundUp(s2ExtensionPurlins(i), 2) & """ above extension eave strut"
            RoofPurlinCollection.Add Purlin
        End If
    Next i
End If
''''''''s4 extension
If b.rShape = "Gable" Then
    If EstSht.Range("s4_EaveExtension").Value > 0 Then
        s4ExtensionPurlinNum = Application.WorksheetFunction.RoundUp((b.s4ExtensionRafterLength) / 60 - 1, 0)
        ReDim s4ExtensionPurlins(s4ExtensionPurlinNum - 1) As Double 'Eave Strut (1) is already made w/ girts
        's4ExtensionLength = b.s4EaveExtensionBuildingLength
        s4ExtensionPurlins(0) = 0 'extension eave strut
        For i = 1 To UBound(s4ExtensionPurlins)
            If i <> UBound(s4ExtensionPurlins) Then
                s4ExtensionPurlins(i) = 60 + (60 * i)
            Else
                s4ExtensionPurlins(i) = b.s4ExtensionRafterLength - 12
            End If
        Next i
        For i = 0 To UBound(s4ExtensionPurlins)
            StartPos = 0
            MaxPos = b.bLength
            For j = BayNum To 1 Step -1
                BayLength = EstSht.Range("Bay1_Length").offset(j - 1, 0).Value * 12
    '            If EstSht.Range("e1_GableOverhang").Value > 0 And j = BayNum Then
    '                BayLength = BayLength + EstSht.Range("s4_EaveOverhang").Value
    '                StartPos = -EstSht.Range("s4_EaveOverhang").Value
    '            ElseIf EstSht.Range("s4_EaveOverhang").Value > 0 And j = 1 Then
    '                BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value
    '            End If
                If i <> 0 Then
                    Set Purlin = New clsMember
                    Purlin.mType = "Roof Purlin"
                    Purlin.Length = BayLength
                    Purlin.rEdgePosition = StartPos
                    Purlin.tEdgeHeight = s4ExtensionPurlins(i)
                    If BayLength > 25 * 12 And Overhang = False Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                    Purlin.Placement = "sidewall 4 eave extension roof purlin, " & Application.WorksheetFunction.RoundUp(s4ExtensionPurlins(i), 2) & """ above eave strut, bay number " & j
                    RoofPurlinCollection.Add Purlin
                    b.WeldClips = b.WeldClips + 2
                Else
                    Set Purlin = New clsMember
                    Purlin.mType = "Eave Strut"
                    Purlin.Length = BayLength
                    Purlin.rEdgePosition = StartPos
                    Purlin.tEdgeHeight = s4ExtensionPurlins(i)
                    If b.s4ExtensionPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    Else
                        Purlin.Size = "8"" " & b.s4ExtensionPitch & ": 12 double up eave strut"
                    End If
                    Purlin.Placement = "sidewall 2 eave extension eave strut, " & Application.WorksheetFunction.RoundUp(s4ExtensionPurlins(i), 2) & """ above eave strut, bay number " & j
                    RoofPurlinCollection.Add Purlin
                End If
                StartPos = StartPos + BayLength
            Next j
            If EstSht.Range("e1_GableExtension").Value > 0 And b.s4e1ExtensionIntersection = True Then
                BayLength = EstSht.Range("e1_GableExtension").Value * 12
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = b.bLength * 12
                Purlin.tEdgeHeight = s4ExtensionPurlins(i)
                If i <> 0 Then 'normal purlins
                    If BayLength > 25 * 12 And Overhang = False Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                    b.WeldClips = b.WeldClips + 2
                Else
                    Purlin.mType = "Eave Strut"
                    If b.s4ExtensionPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    Else
                        Purlin.Size = "8"" " & b.s4ExtensionPitch & ": 12 double up eave strut"
                    End If
                    Debug.Print "eave strut e1/s4 intersection created"
                    Debug.Print EaveStrutCount + 1
                End If
                Purlin.Placement = "sidewall 4 eave extension e1 intersection roof purlin, " & Application.WorksheetFunction.RoundUp(s4ExtensionPurlins(i), 2) & """ above extension eave strut"
                RoofPurlinCollection.Add Purlin
            End If
            If EstSht.Range("e3_GableExtension").Value > 0 And b.s4e3ExtensionIntersection = True Then
                BayLength = EstSht.Range("e3_GableExtension").Value * 12
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = -EstSht.Range("e3_GableExtension").Value * 12
                Purlin.tEdgeHeight = s4ExtensionPurlins(i)
                If i <> 0 Then 'normal purlins
                    If BayLength > 25 * 12 And Overhang = False Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                    b.WeldClips = b.WeldClips + 2
                Else
                    Purlin.mType = "Eave Strut"
                    If b.s4ExtensionPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    Else
                        Purlin.Size = "8"" " & b.s4ExtensionPitch & ": 12 double up eave strut"
                    End If
                    Debug.Print "eave strut e1/s4 intersection created"
                    Debug.Print EaveStrutCount + 1
                End If
                Purlin.Placement = "sidewall 4 eave extension e3 intersection roof purlin, " & Application.WorksheetFunction.RoundUp(s4ExtensionPurlins(i), 2) & """ above extension eave strut"
                RoofPurlinCollection.Add Purlin
            End If
        Next i
    End If
Else '''''''''''''''''''''''''''''''''''''''''''''''''''''s4 extension single slope
    If EstSht.Range("s4_EaveExtension").Value > 0 Then
        s4ExtensionPurlinNum = Application.WorksheetFunction.RoundUp((b.s4ExtensionRafterLength) / 60 - 1, 0)
        ReDim s4ExtensionPurlins(s4ExtensionPurlinNum - 1) As Double 'Eave Strut (1) is already made w/ girts
        's4ExtensionLength = b.s4EaveExtensionBuildingLength
        s4ExtensionPurlins(0) = b.s4ExtensionRafterLength 'extension eave strut
        For i = 1 To UBound(s4ExtensionPurlins)
            If i <> UBound(s4ExtensionPurlins) Then
                s4ExtensionPurlins(i) = b.s4ExtensionRafterLength - (60 + (60 * i))
            Else
                s4ExtensionPurlins(i) = 12
            End If
        Next i
        For i = 0 To UBound(s4ExtensionPurlins)
            StartPos = 0
            MaxPos = b.bLength
            For j = BayNum To 1 Step -1
                BayLength = EstSht.Range("Bay1_Length").offset(j - 1, 0).Value * 12
    '            If EstSht.Range("e1_GableOverhang").Value > 0 And j = BayNum Then
    '                BayLength = BayLength + EstSht.Range("s4_EaveOverhang").Value
    '                StartPos = -EstSht.Range("s4_EaveOverhang").Value
    '            ElseIf EstSht.Range("s4_EaveOverhang").Value > 0 And j = 1 Then
    '                BayLength = BayLength + EstSht.Range("e1_GableOverhang").Value
    '            End If
                If i <> 0 Then
                    Set Purlin = New clsMember
                    Purlin.mType = "Roof Purlin"
                    Purlin.Length = BayLength
                    Purlin.rEdgePosition = StartPos
                    Purlin.tEdgeHeight = s4ExtensionPurlins(i)
                    If BayLength > 25 * 12 And Overhang = False Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                    Purlin.Placement = "sidewall 4 eave extension roof purlin, " & Application.WorksheetFunction.RoundUp(s4ExtensionPurlins(i), 2) & """ above high side eave, bay number " & j
                    RoofPurlinCollection.Add Purlin
                    b.WeldClips = b.WeldClips + 2
                Else
                    Set Purlin = New clsMember
                    Purlin.mType = "Eave Strut"
                    Purlin.Length = BayLength
                    Purlin.rEdgePosition = StartPos
                    Purlin.tEdgeHeight = s4ExtensionPurlins(i)
                    If b.s4ExtensionPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    Else
                        Purlin.Size = "8"" " & b.s4ExtensionPitch & ": 12 double down eave strut"
                    End If
                    Purlin.Placement = "sidewall 2 eave extension eave strut, " & Application.WorksheetFunction.RoundUp(s4ExtensionPurlins(i), 2) & """ above high side eave, bay number " & j
                    RoofPurlinCollection.Add Purlin
                End If
                StartPos = StartPos + BayLength
            Next j
            If EstSht.Range("e1_GableExtension").Value > 0 And b.s4e1ExtensionIntersection = True Then
                BayLength = EstSht.Range("e1_GableExtension").Value * 12
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = b.bLength * 12
                Purlin.tEdgeHeight = s4ExtensionPurlins(i)
                If i <> 0 Then 'normal purlins
                    If BayLength > 25 * 12 And Overhang = False Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                    b.WeldClips = b.WeldClips + 2
                Else
                    Purlin.mType = "Eave Strut"
                    If b.s4ExtensionPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    Else
                        Purlin.Size = "8"" " & b.s4ExtensionPitch & ": 12 double down eave strut"
                    End If
                    Debug.Print "eave strut e1/s4 intersection created"
                    Debug.Print EaveStrutCount + 1
                End If
                Purlin.Placement = "sidewall 4 eave extension e1 intersection roof purlin, " & Application.WorksheetFunction.RoundUp(s4ExtensionPurlins(i), 2) & """ above high side eave"
                RoofPurlinCollection.Add Purlin
            End If
            If EstSht.Range("e3_GableExtension").Value > 0 And b.s4e3ExtensionIntersection = True Then
                BayLength = EstSht.Range("e3_GableExtension").Value * 12
                Set Purlin = New clsMember
                Purlin.mType = "Roof Purlin"
                Purlin.Length = BayLength
                Purlin.rEdgePosition = -EstSht.Range("e3_GableExtension").Value * 12
                Purlin.tEdgeHeight = s4ExtensionPurlins(i)
                If i <> 0 Then 'normal purlins
                    If BayLength > 25 * 12 And Overhang = False Then
                        Purlin.Size = "10"" C Purlin"
                    Else
                        Purlin.Size = "8"" C Purlin"
                    End If
                    b.WeldClips = b.WeldClips + 2
                Else
                    Purlin.mType = "Eave Strut"
                    If b.s4ExtensionPitch = 1 Then
                        Purlin.Size = "8"" C Purlin"
                    Else
                        Purlin.Size = "8"" " & b.s4ExtensionPitch & ": 12 double down eave strut"
                    End If
                    Debug.Print "eave strut e3/s4 intersection created"
                    Debug.Print EaveStrutCount + 1
                End If
                Purlin.Placement = "sidewall 4 eave extension e3 intersection roof purlin, " & Application.WorksheetFunction.RoundUp(s4ExtensionPurlins(i), 2) & """ above high side eave"
                RoofPurlinCollection.Add Purlin
            End If
        Next i
    End If
End If

End Sub

'adjusts heights and lengths for FO Jambs depending on girts and rafters
 Sub AdjustFOMembers(b As clsBuilding, eWall As String)
Dim FOCollection As Collection
Dim GirtCollection As Collection
Dim RafterCollection As Collection
Dim FO As clsFO
Dim item As Object
Dim Jamb As clsMember
Dim Girt As clsMember
Dim tempNearestGirtAbove As Double
Dim tempNearestGirtBelow As Double
Dim i As Integer
Dim Rafter As clsMember
Dim RafterWidth As Double
Dim DistanceToRafter As Double
Dim tEdgeDifference As Double

Select Case eWall
Case "e1"
    Set FOCollection = b.e1FOs
    Set GirtCollection = b.e1Girts
    Set RafterCollection = b.e1Rafters
Case "s2"
    Set FOCollection = b.s2FOs
    Set GirtCollection = b.s2Girts
Case "e3"
    Set FOCollection = b.e3FOs
    Set GirtCollection = b.e3Girts
    Set RafterCollection = b.e3Rafters
Case "s4"
    Set FOCollection = b.s4FOs
    Set GirtCollection = b.s4Girts
End Select

For Each FO In FOCollection
    For Each item In FO.FOMaterials
        'Extend jambs to nearest horizontal girt so that windows and MiscFOs aren't floating
        If item.clsType = "Member" And (FO.FOType = "Window" Or FO.FOType = "MiscFO") Then
            Set Jamb = item
            If Jamb.CL <> 0 And Jamb.Length <> b.DistanceToRoof(eWall, Jamb.CL) Then ''horizontal jambs weren't given CenterLines
                'if Jamb Lenght isn't already full height, check that it touches the nearest girt; adjust if necessary
                If Jamb.tEdgeHeight = 2 + (7 * 12) And FO.tEdgeHeight <= 2 + (7 * 12) Then
                'if jamb and FO are less than or equal to 7'2", don't extend above girt
                'check that bottom edge goes to building slab
                    Jamb.bEdgeHeight = 0
                    Jamb.Length = Jamb.tEdgeHeight
                Else
                    DistanceToRafter = b.DistanceToRoof(eWall, Jamb.CL)
                    tempNearestGirtAbove = DistanceToRafter
                    tempNearestGirtBelow = 0
                    For Each Girt In GirtCollection
                        If Girt.tEdgeHeight > FO.tEdgeHeight And Girt.tEdgeHeight - FO.tEdgeHeight < tempNearestGirtAbove - FO.tEdgeHeight Then
                            tempNearestGirtAbove = Girt.tEdgeHeight
                        End If
                        If Girt.bEdgeHeight < FO.bEdgeHeight And FO.bEdgeHeight - Girt.bEdgeHeight < FO.bEdgeHeight - tempNearestGirtBelow Then
                            tempNearestGirtBelow = Girt.bEdgeHeight
                        End If
                    Next Girt
                    If tempNearestGirtAbove - tempNearestGirtBelow <= 30 * 12 + 4 Then
                        Jamb.tEdgeHeight = tempNearestGirtAbove
                        Jamb.bEdgeHeight = tempNearestGirtBelow
                        Jamb.Length = tempNearestGirtAbove - tempNearestGirtBelow
                    End If
                End If
            End If
        End If
    Next item
Next FO

'Reduce full height jambs or jambs that connect to rafters
If eWall = "e1" Or eWall = "e3" Then
    For Each FO In FOCollection
        For Each item In FO.FOMaterials
            If item.clsType = "Member" Then
                Set Jamb = item
                If Jamb.CL <> 0 And Jamb.tEdgeHeight = b.DistanceToRoof(eWall, Jamb.CL) Then
                    'if jamb goes all the way to the ceiling and isn't load bearing, then it needs to be reduced in length to the centerline of the new rafter position
                    For Each Rafter In RafterCollection
                        If Rafter.rEdgePosition < Jamb.CL And Rafter.RafterLeftEdge > Jamb.CL Then
                            RafterWidth = Rafter.Width
                        End If
                    Next Rafter
                    tEdgeDifference = Sqr(((RafterWidth / 2) * (b.rPitch / 12)) ^ 2 + (RafterWidth / 2) ^ 2) '+ ((RafterWidth / 2) * (b.rPitch / 12))
                    'Jamb.tEdgeHeight = Jamb.tEdgeHeight - tEdgeDifference
                    Jamb.tEdgeHeight = Jamb.tEdgeHeight - tEdgeDifference * 2 + (Jamb.Width / 2) * (b.rPitch / 12)
                    Jamb.Length = Jamb.Length - tEdgeDifference * 2 + (Jamb.Width / 2) * (b.rPitch / 12)
                    Jamb.Placement = Jamb.Placement & "cut at " & Application.WorksheetFunction.Round(Atn(b.rPitch / 12), 2) & " degree angle, "
                End If
            End If
        Next item
    Next FO
End If
            
End Sub



Sub RafterGen(b As clsBuilding, originalWall As String)
Dim ColumnCollection As Collection
Dim FOCollection As Collection
Dim RafterCollection As Collection
Dim item As Object
Dim Rafters(25) As Double
Dim ColIndex As Integer
Dim Column As clsMember
Dim Member As clsMember
Dim RafterMember As clsMember
Dim FO As clsFO
Dim StartPos As Double
Dim EndPos As Double
Dim MidPos As Double
Dim NextStartPos As Double
Dim MaxDistance As Double
Dim tempNearestColumn As Double
Dim i As Integer
Dim PrevLocation As Double
Dim IntColumnMode As Boolean
Dim RafterType As String
Dim RafterPlacement As String
Dim largestWidth As Double
Dim largestSize As String
Dim SecondDimension As Double
Dim tempSecondDimension As Double
Dim DistanceToLower As Double
Dim DistanceToLengthen As Double
Dim AngleCut As Boolean
Dim Angle As Double
Dim FirstColWidth As Double
Dim LastColWidth As Double
Dim eWall As String
Dim tempRafterSize As String

eWall = originalWall


Select Case eWall
Case "e1"
    Set ColumnCollection = b.e1Columns
    Set FOCollection = b.e1FOs
    Set RafterCollection = b.e1Rafters
    IntColumnMode = False
    If b.ExpandableEndwall(eWall) Then
        RafterType = "W-Beam"
    ElseIf EstSht.Range("e1_GableOverhang").Value > 0 Then
        RafterType = "8"" C Purlin"
    ElseIf EstSht.Range("Bay1_Length").Value > 25 Then
        tempRafterSize = SteelLookupSht.Range("NonExpandableEndwallRaftersWithLargeBay").Value
    Else
        tempRafterSize = SteelLookupSht.Range("NonExpandableEndwallRaftersWithNormalBay").Value
    End If
    
    RafterPlacement = eWall & " endwall rafter"
Case "e3"
    Set ColumnCollection = b.e3Columns
    Set FOCollection = b.e3FOs
    Set RafterCollection = b.e3Rafters
    IntColumnMode = False
    If b.ExpandableEndwall(eWall) Then
        RafterType = "W-Beam"
    ElseIf EstSht.Range("e3_GableOverhang").Value > 0 Then
        RafterType = "8"" C Purlin"
    ElseIf EstSht.Range("Bay1_Length").offset(EstSht.Range("BayNum").Value - 1, 0).Value > 25 Then
        tempRafterSize = SteelLookupSht.Range("NonExpandableEndwallRaftersWithLargeBay").Value
    Else
        tempRafterSize = SteelLookupSht.Range("NonExpandableEndwallRaftersWithNormalBay").Value
    End If
    RafterPlacement = eWall & " endwall rafter"
Case "int"
    Set ColumnCollection = b.InteriorColumns
    Set RafterCollection = b.intRafters
    eWall = "e1"
    RafterType = "W-Beam"
    IntColumnMode = True
    RafterPlacement = "main rafter line"
Case "e1 Extension"
    Set ColumnCollection = b.e1ExtensionMembers
    Set RafterCollection = b.e1Rafters
    eWall = "e1"
    RafterType = "W-Beam"
    IntColumnMode = True
    RafterPlacement = "e1 Extension Rafter"
Case "e3 Extension"
    Set ColumnCollection = b.e3ExtensionMembers
    Set RafterCollection = b.e3Rafters
    eWall = "e3"
    RafterType = "W-Beam"
    IntColumnMode = True
    RafterPlacement = "e3 Extension Rafter"
End Select

If RafterType = "" Then
    If tempRafterSize Like "W8x10" Then
        RafterType = "W-Beam"
    Else
        RafterType = tempRafterSize
    End If
End If

'set actual start and end points for building using corner column widths
For Each Column In ColumnCollection
    If Column.rEdgePosition = 0 Then
        StartPos = Column.Width
        FirstColWidth = Column.Width
    ElseIf Column.lEdgePosition = b.bWidth * 12 Then
        MaxDistance = b.bWidth * 12 - Column.Width
        LastColWidth = Column.Width
    End If
Next Column


'Rafters(0) = StartPos
For i = 0 To 25 'not used, I just need to loop enough times to generate all the rafters. there should be a better way to do this. NEEDS FIX
    If EndPos < MaxDistance Then
        tempNearestColumn = 1.79769313486231E+308
        'loop through columns, find nearest rEdgePosition
        For Each Column In ColumnCollection
            If Abs(Column.rEdgePosition - StartPos) < Abs(tempNearestColumn - StartPos) And Column.rEdgePosition > StartPos And Column.LoadBearing = True Then
                tempNearestColumn = Column.rEdgePosition
                NextStartPos = Column.lEdgePosition
            End If
        Next Column
        If IntColumnMode = False Then
        'Only Endwalls have FOs
            For Each FO In FOCollection
                For Each item In FO.FOMaterials
                    If item.clsType = "Member" Then
                        Set Member = item
                        If Member.LoadBearing = True Then
                            If Abs(Member.CL - StartPos) < Abs(tempNearestColumn - StartPos) _
                                        And Member.rEdgePosition > StartPos _
                                        And Member.tEdgeHeight = b.DistanceToRoof(eWall, Member.CL) Then
                                tempNearestColumn = Member.rEdgePosition
                                NextStartPos = Member.lEdgePosition
                            End If
                        End If
                    End If
                Next item
            Next FO
        End If
        'check that girt edge does not exceed building width
        If tempNearestColumn <= MaxDistance Then
            EndPos = tempNearestColumn
        Else
            EndPos = MaxDistance
        End If
        If EndPos > b.bWidth * 12 / 2 And StartPos < b.bWidth * 12 / 2 And b.rShape = "Gable" Then
        'if the next position is on the other side of the Gable Roof, add a midpoint and create a rafter that goes to midpoint
            Set RafterMember = New clsMember
            RafterMember.mType = RafterPlacement & " Rafter"
            MidPos = b.bWidth * 12 / 2
            RafterMember.bEdgeHeight = b.DistanceToRoof(eWall, StartPos)
            RafterMember.tEdgeHeight = b.DistanceToRoof(eWall, MidPos)
            RafterMember.rEdgePosition = StartPos
            RafterMember.RafterLeftEdge = MidPos
            RafterMember.Length = Sqr((MidPos - StartPos) ^ 2 + (RafterMember.tEdgeHeight - RafterMember.bEdgeHeight) ^ 2)
            'Set Size to nearest COLUMN (not FO jambs)
            If RafterType <> "W-Beam" Then
                RafterMember.Width = Left(RafterType, InStr(1, RafterType, " ") - 2)
                RafterMember.Size = RafterType
            Else
                'To find size, use the distance between lEdge (higher) and rEdge (lower)
                RafterMember.SetSize b, "Rafter", eWall, Abs(EndPos - StartPos)
            End If
            RafterMember.Placement = RafterPlacement & ", " & RafterMember.Length & "' long"
            RafterCollection.Add RafterMember
            'Second Rafter across gable peak
            Set RafterMember = New clsMember
            RafterMember.mType = RafterPlacement & " Rafter"
            RafterMember.bEdgeHeight = b.DistanceToRoof(eWall, EndPos)
            RafterMember.tEdgeHeight = b.DistanceToRoof(eWall, MidPos)
            RafterMember.rEdgePosition = MidPos
            RafterMember.RafterLeftEdge = EndPos
            RafterMember.Length = Sqr((EndPos - MidPos) ^ 2 + (RafterMember.tEdgeHeight - RafterMember.bEdgeHeight) ^ 2)
            'Set Size to nearest COLUMN (not FO jambs)
            If RafterType <> "W-Beam" Then
                RafterMember.Width = Left(RafterType, InStr(1, RafterType, " ") - 2)
                RafterMember.Size = RafterType
            Else
                'To find size, use the distance between lEdge (higher) and rEdge (lower)
                RafterMember.SetSize b, "Rafter", eWall, Abs(EndPos - StartPos)
            End If
            RafterMember.Placement = RafterPlacement & ", " & RafterMember.Length & "' long"
            RafterCollection.Add RafterMember
            StartPos = NextStartPos
        Else
            Set RafterMember = New clsMember
            RafterMember.mType = RafterPlacement & " Rafter"
            RafterMember.rEdgePosition = StartPos
            RafterMember.RafterLeftEdge = EndPos
            'If StartPos <= 27 Then StartPos = 0
            'If EndPos >= b.bWidth * 12 - 27 Then EndPos = b.bWidth * 12
            If (EndPos > b.bWidth * 12 / 2 And b.rShape = "Gable") Or (b.rShape = "Single Slope" And eWall = "e1") Then
            'on the far side of a Gable Roof, tEdge and bEdge are switched
                RafterMember.tEdgeHeight = b.DistanceToRoof(eWall, StartPos)
                RafterMember.bEdgeHeight = b.DistanceToRoof(eWall, EndPos)
            Else
                RafterMember.bEdgeHeight = b.DistanceToRoof(eWall, StartPos)
                RafterMember.tEdgeHeight = b.DistanceToRoof(eWall, EndPos)
            End If
            'Set Size to nearest COLUMN (not FO jambs)
            If RafterType = "10"" Receiver Cee" Then
                RafterMember.Width = 10
                RafterMember.Size = RafterType
            ElseIf RafterType = "8"" Receiver Cee" Then
                RafterMember.Width = 8
                RafterMember.Size = RafterType
            ElseIf RafterType = "8"" C Purlin" Then
                RafterMember.Width = 8
                RafterMember.Size = RafterType
            ElseIf RafterType = "10"" C Purlin" Then
                RafterMember.Width = 10
                RafterMember.Size = RafterType
            Else
                RafterMember.SetSize b, "Rafter", eWall, Abs(StartPos - EndPos)
            End If
            
            RafterMember.Length = Sqr((EndPos - StartPos) ^ 2 + (RafterMember.tEdgeHeight - RafterMember.bEdgeHeight) ^ 2)
            RafterMember.Placement = RafterPlacement & ", " & RafterMember.Length & "' long"
            RafterCollection.Add RafterMember
            StartPos = NextStartPos
            '''''''' lower all rafters so that the top edge will be the actual building height
            '''''''' AND '''''' Make rafters that connect to corners or center of gable building longer
            'formulas:
            'Distance to lower rafters = SQUARE ROOT(  ((width/2)(pitch/12))^2    *    (width/2)^2    )
            'Distance to lengthen rafters = (width/2)(pitch/12) per corner/peak
            Angle = Atn(b.rPitch / 12) * (RafterMember.Width / 2)
            DistanceToLower = Sqr(Angle ^ 2 + (RafterMember.Width / 2) ^ 2)
            DistanceToLengthen = Sqr(DistanceToLower ^ 2 - (RafterMember.Width / 2) ^ 2)
            RafterMember.bEdgeHeight = RafterMember.bEdgeHeight - DistanceToLower
            RafterMember.tEdgeHeight = RafterMember.tEdgeHeight - DistanceToLower
            If RafterMember.bEdgeHeight < b.bHeight * 12 Then
                RafterMember.Length = RafterMember.Length + DistanceToLengthen
                AngleCut = True
            End If
            'Single Slope
            If RafterMember.tEdgeHeight >= b.bHeight * 12 + b.bWidth * b.rPitch - DistanceToLower Then
                RafterMember.Length = RafterMember.Length + DistanceToLengthen
                AngleCut = True
            'Gable
            ElseIf RafterMember.tEdgeHeight >= b.bHeight * 12 + b.bWidth / 2 * b.rPitch - DistanceToLower Then
                RafterMember.Length = RafterMember.Length + DistanceToLengthen
                AngleCut = True
            End If
            If AngleCut Then RafterMember.Placement = RafterMember.Placement & ", cut at " & Application.WorksheetFunction.Round(Atn(b.rPitch / 12), 2) & " angle, "
        End If
    End If
Next i

If originalWall = "int" Then
    For Each RafterMember In RafterCollection
        RafterMember.Qty = EstSht.Range("BayNum").Value - 1
        'Debug.Print "x Start: " & RafterMember.rEdgePosition & ", y Start: " & RafterMember.bEdgeHeight & ", y End: " & RafterMember.tEdgeHeight & ", length: " & RafterMember.Length
    Next RafterMember
End If
    
End Sub

'Specify Eave Strut types
 Sub EaveStrutTypes(b As clsBuilding, eWall As String)

Dim Eavestruts As Collection
Dim Member As clsMember
Dim LinerPanels As Boolean
Dim e1Overhang As Boolean
Dim e1Extension As Boolean
Dim e3Overhang As Boolean
Dim e3Extension As Boolean
Dim s2Overhang As Boolean
Dim s2Extension As Boolean
Dim s4Overhang As Boolean
Dim s4Extension As Boolean

'check for liner panels, overhangs, extension, soffit
If EstSht.Range("Roof_LinerPanels").Value <> "None" Then LinerPanels = True
If EstSht.Range("e1_GableOverhang").Value > 0 Then e1Overhang = True
If EstSht.Range("e1_GableExtension").Value > 0 Then e1Extension = True
If EstSht.Range("e3_GableOverhang").Value > 0 Then e3Overhang = True
If EstSht.Range("e3_GableExtension").Value > 0 Then e3Extension = True
If EstSht.Range("s2_EaveOverhang").Value > 0 Then s2Overhang = True
If EstSht.Range("s2_EaveExtension").Value > 0 Then s2Extension = True
If EstSht.Range("s4_EaveOverhang").Value > 0 Then s4Overhang = True
If EstSht.Range("s4_EaveExtension").Value > 0 Then s4Extension = True

Select Case eWall
Case "s2"
    Set Eavestruts = b.s2Girts
    For Each Member In Eavestruts
        If Member.mType = "Eave Strut" Then
                'if Overhang, extend eave strut
                If Member.rEdgePosition + Member.Length > b.bLength * 12 - 15 And e3Overhang = True And e3Extension = False Then
                    Member.Length = Member.Length + EstSht.Range("e3_GableOverhang").Value * 12
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If b.e3GableOverhangSoffit Or EstSht.Range("Roof_LinerPanels").Value <> "None" And b.rPitch <> 1 Then
                        Member.Size = Member.Size & "double up eave strut"
                    Else
                        Member.Size = Member.Size & "single up eave strut"
                    End If
                    Debug.Print "s2 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                ElseIf Member.rEdgePosition < 15 And e1Overhang = True And e1Extension = False Then
                    Member.rEdgePosition = -EstSht.Range("e1_GableOverhang").Value * 12
                    Member.Length = Member.Length + EstSht.Range("e1_GableOverhang").Value * 12
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If b.e1GableOverhangSoffit Or EstSht.Range("Roof_LinerPanels").Value <> "None" And b.rPitch <> 1 Then
                        Member.Size = Member.Size & "double up eave strut"
                    ElseIf b.rPitch = 1 Then
                        'do nothing
                    Else
                        Member.Size = Member.Size & "single up eave strut"
                    End If
                    Debug.Print "s2 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                Else 'no overhangs, eave strut determined only by normal factors
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If LinerPanels Then
                        Member.Size = Member.Size & "double up eave strut"
                    ElseIf b.rPitch <> 1 Then
                        Member.Size = Member.Size & "single up eave strut"
                    End If
                    Debug.Print "s2 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                End If
        End If
    Next Member
Case "s4"
    Set Eavestruts = b.s4Girts
        For Each Member In Eavestruts
        If Member.mType = "Eave Strut" Then
            If b.rShape = "Gable" Then
                'if Overhang, extend eave strut
                If Member.rEdgePosition + Member.Length > b.bLength * 12 - 15 And e1Overhang = True And e1Extension = False Then
                    Member.Length = Member.Length + EstSht.Range("e1_GableOverhang").Value * 12
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If b.e1GableOverhangSoffit Or EstSht.Range("Roof_LinerPanels").Value <> "None" And b.rPitch <> 1 Then
                        Member.Size = Member.Size & "double up eave strut"
                    Else
                        Member.Size = Member.Size & "single up eave strut"
                    End If
                    Debug.Print "s4 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                ElseIf Member.rEdgePosition < 15 And e3Overhang = True And e1Extension = False Then
                    Member.rEdgePosition = -EstSht.Range("e3_GableOverhang").Value * 12
                    Member.Length = Member.Length + EstSht.Range("e3_GableOverhang").Value * 12
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If b.e3GableOverhangSoffit Or EstSht.Range("Roof_LinerPanels").Value <> "None" And b.rPitch <> 1 Then
                        Member.Size = Member.Size & "double up eave strut"
                    Else
                        Member.Size = Member.Size & "single up eave strut"
                    End If
                    Debug.Print "s4 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                Else 'no overhangs, eave strut determined only by normal factors
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If LinerPanels Then
                        Member.Size = Member.Size & "double up eave strut"
                    ElseIf b.rPitch <> 1 Then
                        Member.Size = Member.Size & "single up eave strut"
                    End If
                    Debug.Print "s4 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                End If
            Else 'Single Slope --> struts will be single/double down
                'if Overhang, extend eave strut
                If Member.rEdgePosition + Member.Length > b.bLength * 12 - 15 And e1Overhang = True And e1Extension = False Then
                    Member.Length = Member.Length + EstSht.Range("e1_GableOverhang").Value * 12
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If b.e1GableOverhangSoffit Or EstSht.Range("Roof_LinerPanels").Value <> "None" And b.rPitch <> 1 Then
                        Member.Size = Member.Size & "double down eave strut"
                    Else
                        Member.Size = Member.Size & "single down eave strut"
                    End If
                    Debug.Print "s4 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                ElseIf Member.rEdgePosition < 15 And e3Overhang = True And e1Extension = False Then
                    Member.rEdgePosition = -EstSht.Range("e3_GableOverhang").Value * 12
                    Member.Length = Member.Length + EstSht.Range("e3_GableOverhang").Value * 12
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If b.e3GableOverhangSoffit Or EstSht.Range("Roof_LinerPanels").Value <> "None" And b.rPitch <> 1 Then
                        Member.Size = Member.Size & "double down eave strut"
                    Else
                        Member.Size = Member.Size & "single down eave strut"
                    End If
                    Debug.Print "s4 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                Else 'no overhangs, eave strut determined only by normal factors
                    If b.rPitch = 1 Then
                        Member.Size = "8"" C Purlin"
                    Else
                        Member.Size = "8"" " & b.rPitch & ":12 "
                    End If
                    If LinerPanels Then
                        Member.Size = Member.Size & "double down eave strut"
                    ElseIf b.rPitch <> 1 Then
                        Member.Size = Member.Size & "single down eave strut"
                    End If
                    Debug.Print "s4 regular eave strut created"
                    Debug.Print EaveStrutCount + 1
                End If
            End If
        End If
    Next Member
End Select
End Sub

Function ClosestWallGirt(Height As Variant, Optional Direction As Integer) As Double
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function returns the nearest value to a target
'INPUT: Pass the function a range of cells, a target value that you want to find a number closest to
' and an optional direction variable described below.
'OPTIONS: Set the optional variable Direction equal to 0 or blank to find the closest value
' Set equal to -1 to find the closest value below your target
' set equal to 1 to find the closest value above your target
'OUTPUT: The output is the number in the range closest to your target value.
' Because the output is a variant, the address of the closest number can also be returned when
' calling this function from another VBA macro.
Dim t As Variant
Dim u As Variant
Dim Purlins() As Variant
Dim Purlin As Variant
Dim pAbove As Integer
Dim pBelow As Integer
Purlins = Array(86, 146, 206, 266, 326, 386, 446, 506, 566, 626, 686, 746, 806, 866, 926, 986, 1046, 1106, 1166)

    t = 1.79769313486231E+308 'initialize
    'ClosestWallPurlin = "No value found"
    For Each Purlin In Purlins
        If IsNumeric(Purlin) Then
            u = Abs(Purlin - Height)
            If Direction > 0 And Purlin >= Height Then
                'only report if closer number is greater than the target
                If u < t Then
                    t = u
                    ClosestWallGirt = Purlin
                End If
            ElseIf Direction < 0 And Purlin <= Height Then
                'only report if closer number is less than the target
                If u < t Then
                    t = u
                    ClosestWallGirt = Purlin
                End If
            ElseIf Direction = 0 Then
                If u < t Then
                    t = u
                    ClosestWallGirt = Purlin
                End If
            End If
        End If
    Next Purlin


End Function

'''''''''' Creates Wall Collection with all relevant items, calculates girt lengths
''''''''''''''not finished
Sub EndwallGirtLengthCalc(b As clsBuilding, Optional eWall As String)

Dim ColumnCollection As Collection
Dim FOCollection As Collection
Dim GirtsCollection As Collection
Dim RafterCollection As Collection
Dim Column As clsMember
Dim Girt As clsMember
Dim FO As clsFO
Dim FOMaterial As Collection
Dim Member As clsMember
Dim Rafter As clsMember
Dim item As Variant
Dim Points() As Double
Dim Girts() As Double
Dim GirtNum As Integer
Dim LowestPoint As Integer
Dim WallStatus As String
Dim TotalHeight As Double
Dim StartPos As Double
Dim EndPos As Double
Dim MaxDistance As Double
Dim NextIntersection As Double
Dim GirtIndex As Integer
Dim GirtMiddle As Integer
Dim GirtHeight As Integer
Dim WallLength As Double
Dim DistanceToLower As Double
Dim RafterGirtAdjustment As Double
Dim RafterWidth As Double
Dim i As Integer
Dim Angle As Double



ReDim Girts(20)

Select Case eWall
Case "e1"
    Set ColumnCollection = b.e1Columns
    Set FOCollection = b.e1FOs
    Set GirtsCollection = b.e1Girts
    Set RafterCollection = b.e1Rafters
Case "s2"
    Set ColumnCollection = b.s2Columns
    Set FOCollection = b.s2FOs
    Set GirtsCollection = b.s2Girts
Case "e3"
    Set ColumnCollection = b.e3Columns
    Set FOCollection = b.e3FOs
    Set GirtsCollection = b.e3Girts
    Set RafterCollection = b.e1Rafters
Case "s4"
    Set ColumnCollection = b.s4Columns
    Set FOCollection = b.s4FOs
    Set GirtsCollection = b.s4Girts
End Select

WallStatus = b.WallStatus(eWall)
LowestPoint = b.LengthAboveFinishedFloor(eWall) * 12  'in

'check for excluded wall
If WallStatus = "Exclude" Then
    If (eWall = "e1" Or eWall = "e3") Then
        Exit Sub
    Else
        LowestPoint = b.bHeight * 12
    End If
End If

'get highest point of wall
If eWall = "s2" Or (eWall = "s4" And b.rShape = "Gable") Then
    TotalHeight = b.bHeight * 12
ElseIf eWall = "s4" And b.rShape = "Single Slope" Then
    TotalHeight = b.bHeight * 12 + b.bWidth * b.rPitch
Else
    If b.rShape = "Single Slope" Then   'in
        TotalHeight = (b.bHeight * 12) + (b.bWidth * b.rPitch)
    Else
        TotalHeight = (b.bHeight * 12) + (b.bWidth / 2 * b.rPitch)
    End If
End If

'base case, normal building in 5' increments after 12'
If LowestPoint = 0 Then
    Girts(0) = 86
    Girts(1) = 146
    i = 2
Else 'Partial Walls or "gable only" walls starting above 86"
    Girts(0) = LowestPoint
    i = 1
    'partial walls starting lower than 86"
    If LowestPoint < 86 Then
        Girts(1) = 86
        Girts(2) = 146
        i = 3
    ElseIf LowestPoint > 86 And LowestPoint < 146 Then
        'partial walls starting above 86" but below 146"
        Girts(1) = 146
        i = 2
    ElseIf WallStatus <> "Exclude" Then
        'partial walls above 146"
        Girts(1) = ClosestWallGirt(LowestPoint + 60, -1)
        i = 2
    End If
End If

GirtNum = i

'Add girt heights to array, if taller than building, value = 0
For i = i To 20
    If Girts(i - 1) + 60 < TotalHeight And Girts(i - 1) > 0 Then
        Girts(i) = Girts(i - 1) + 60
        GirtNum = GirtNum + 1
    End If
Next i

If eWall = "s2" Or eWall = "s4" Then
    If WallStatus = "Exclude" Then
        GirtNum = 0
        ReDim Preserve Girts(GirtNum)
    Else
        Girts(GirtNum) = TotalHeight
        ReDim Preserve Girts(GirtNum)
    End If
Else
    ReDim Preserve Girts(GirtNum - 1)
    GirtNum = GirtNum - 1
End If

'get length of wall depending on wall selection
If eWall = "e1" Or eWall = "e3" Then
    WallLength = b.bWidth * 12
Else
    WallLength = b.bLength * 12
End If

'get rafter width and distance rafters were lowered
If eWall = "e1" Or eWall = "e3" Then
    RafterWidth = 20
    For Each Rafter In RafterCollection
        If Rafter.Width < RafterWidth Then RafterWidth = Rafter.Width
    Next Rafter
    Angle = (Atn(b.rPitch / 12) * 180 / 3.14159265358979)
    DistanceToLower = ((RafterWidth / 2) / (Cos(Angle) * 180 / 3.14159265358979))  '(RafterWidth / 2) /
    Angle = (Atn(b.rPitch / 12)) * (RafterWidth / 2)
    DistanceToLower = Sqr(Angle ^ 2 + (RafterWidth / 2) ^ 2)
    RafterGirtAdjustment = (DistanceToLower / b.rPitch) * 12
Else
    RafterGirtAdjustment = 0
End If


'Run through each girt row;
'set starting value = 0 for begining of row;
'set ending value = first intersecting wall item;
'create member, add to girt collection
For i = 0 To GirtNum
    EndPos = 0
    'girts under eave height start at 0
    's2 will only have girts up to eave height
    's4 will have girts above eave height on single slope
    If Girts(i) <= b.bHeight * 12 Or (Girts(i) <= b.DistanceToRoof(eWall, 0) And b.rShape = "Single Slope" And eWall = "s4") Then
        StartPos = 0
        While EndPos < WallLength
            StartPos = EndPos
            EndPos = NextHorizontalGirtIntersection(b, ColumnCollection, FOCollection, StartPos, eWall, Girts(i))
            If EndPos > WallLength Then 'for s2 and s4 since sidewall column collection doesn't include endpoints
                EndPos = WallLength
            End If
            Set Girt = New clsMember
            Girt.mType = "Girt"
            Girt.Size = "8"" C Purlin"
            Girt.bEdgeHeight = Girts(i)
            Girt.tEdgeHeight = Girts(i)
            Girt.rEdgePosition = StartPos
            Girt.Length = EndPos - StartPos
            If (Girts(i) = b.bHeight * 12 And eWall = "s2") Or (eWall = "s4" And Girts(i) = b.bHeight * 12 And b.rShape = "Gable") Then
                Girt.mType = "Eave Strut"
            ElseIf eWall = "s4" And Girts(i) = (b.bHeight * 12) + (b.bWidth * 12 * b.rPitch / 12) And b.rShape = "Single Slope" Then
                Girt.mType = "Eave Strut"
            End If
            Girt.Placement = "girt screwline row " & i + 1 & " at " & Girts(i) & " inches, wall " & eWall & ", start " & StartPos & " inches, end " & EndPos & " inches, length " & Girt.Length
            GirtsCollection.Add Girt
        Wend
    ElseIf eWall = "e1" Or eWall = "e3" Then 'for endwalls girts above eave height start at roof location at height
        Select Case b.rShape
        Case "Single Slope"
            'e1 single slope: 0' is highest point
            If eWall = "e1" Then
                StartPos = 0
                EndPos = 0
                MaxDistance = b.DistanceFromCorner(eWall, Girts(i)) - RafterGirtAdjustment
                While EndPos < MaxDistance
                    StartPos = EndPos
                    NextIntersection = NextHorizontalGirtIntersection(b, ColumnCollection, FOCollection, StartPos, eWall, Girts(i))
                    'for endpoing on gable roof, columns past max distance should be ignored and replaced with max distance across gable
                    If NextIntersection > MaxDistance Then
                        EndPos = MaxDistance
                    Else
                        EndPos = NextIntersection
                    End If
                    Set Girt = New clsMember
                    Girt.mType = "Girt"
                    Girt.Size = "8"" C Purlin"
                    Girt.bEdgeHeight = Girts(i)
                    Girt.tEdgeHeight = Girts(i) 'NOT SURE IF THIS IS HOW WE WANT TO USE TOP/BOT for HORIZONTAL PCS.
                    Girt.rEdgePosition = StartPos
                    Girt.Length = EndPos - StartPos
                    'Girt.Width = Girt.Length 'ONE OF THESE SHOULD BE REMOVED
                    Girt.Placement = "girt screwline row " & i + 1 & " at " & Girts(i) & " inches, wall " & eWall & ", start " & StartPos & " inches, end " & EndPos & " inches, length " & Girt.Length
                    GirtsCollection.Add Girt
                Wend
            'e3 single slope: 0' is lowest point
            ElseIf eWall = "e3" Then
                StartPos = b.DistanceFromCorner(eWall, Girts(i)) + RafterGirtAdjustment
                EndPos = StartPos
                MaxDistance = b.bWidth * 12
                While EndPos < WallLength
                    StartPos = EndPos
                    NextIntersection = NextHorizontalGirtIntersection(b, ColumnCollection, FOCollection, StartPos, eWall, Girts(i))
                    If NextIntersection > MaxDistance Then
                        EndPos = MaxDistance
                    Else
                        EndPos = NextIntersection
                    End If
                    Set Girt = New clsMember
                    Girt.mType = "Girt"
                    Girt.Size = "8"" C Purlin"
                    Girt.bEdgeHeight = Girts(i)
                    Girt.tEdgeHeight = Girts(i) 'NOT SURE IF THIS IS HOW WE WANT TO USE TOP/BOT for HORIZONTAL PCS.
                    Girt.rEdgePosition = StartPos
                    Girt.Length = EndPos - StartPos
                    'Girt.Width = Girt.Length 'ONE OF THESE SHOULD BE REMOVED
                    Girt.Placement = "girt screwline row " & i + 1 & " at " & Girts(i) & " inches, wall " & eWall & ", start " & StartPos & " inches, end " & EndPos & " inches, length " & Girt.Length
                    GirtsCollection.Add Girt
                Wend
            End If
        Case "Gable"
            'Gable Roofs are symmetrical between e1 and e3
            StartPos = b.DistanceFromCorner(eWall, Girts(i)) + RafterGirtAdjustment
            EndPos = StartPos
            MaxDistance = WallLength - b.DistanceFromCorner(eWall, Girts(i)) - RafterGirtAdjustment
            While EndPos < MaxDistance
                StartPos = EndPos
                EndPos = NextHorizontalGirtIntersection(b, ColumnCollection, FOCollection, StartPos, eWall, Girts(i))
                If EndPos > MaxDistance Then
                    EndPos = MaxDistance
                End If
                Set Girt = New clsMember
                Girt.mType = "Girt"
                Girt.Size = "8"" C Purlin"
                Girt.bEdgeHeight = Girts(i)
                Girt.tEdgeHeight = Girts(i) 'NOT SURE IF THIS IS HOW WE WANT TO USE TOP/BOT for HORIZONTAL PCS.
                Girt.rEdgePosition = StartPos
                Girt.Length = EndPos - StartPos
                Girt.Width = Girt.Length 'ONE OF THESE SHOULD BE REMOVED
                Girt.Placement = "girt screwline row " & i & " at " & Girts(i) & " inches, wall " & eWall & ", start " & StartPos & " inches, end " & EndPos & " inches, length " & Girt.Length
                GirtsCollection.Add Girt
            Wend
        End Select
    End If
Next i

'Step through girts, check if it's in the middle of an FO, then remove
For GirtIndex = GirtsCollection.Count To 1 Step -1
    Set Girt = GirtsCollection(GirtIndex)
    GirtMiddle = Girt.rEdgePosition + Girt.Length / 2 'since each girt starts and ends at a column/jamb, middle points are either valid or invalid
    GirtHeight = Girt.tEdgeHeight
    For Each FO In FOCollection
        If GirtHeight > FO.bEdgeHeight And GirtHeight < FO.tEdgeHeight And GirtMiddle > FO.rEdgePosition And GirtMiddle < FO.lEdgePosition Then
            GirtsCollection.Remove (GirtIndex)
        End If
    Next FO
Next GirtIndex

'---------DEBUG--------
'For Each Girt In GirtsCollection
    'Debug.Print Girt.Placement
'Next Girt

End Sub

 Function NextHorizontalGirtIntersection(b As clsBuilding, Columns As Collection, FOs As Collection, start As Double, Wall As String, Height As Double) As Double
Dim Member As clsMember
Dim FO As clsFO
Dim item As Object
Dim tempNearestIntersection As Double

tempNearestIntersection = 1.79769313486231E+308
If Wall = "e1" Or Wall = "e3" Then
    For Each Member In Columns
        If (Member.CL - start) < (tempNearestIntersection - start) And Member.CL > start And Member.CL > 15 And Member.CL < b.bWidth * 12 - 15 Then
            tempNearestIntersection = Member.CL
        End If
    Next Member
Else 's2 and s4 use bLength
    For Each Member In Columns
        If (Member.CL - start) < (tempNearestIntersection - start) And Member.CL > start And Member.CL > 15 And Member.CL < b.bLength * 12 - 15 Then
            tempNearestIntersection = Member.CL
        End If
    Next Member
End If
    
For Each FO In FOs
    For Each item In FO.FOMaterials
        If item.clsType = "Member" Then
            Set Member = item
            If (Member.CL - start) < (tempNearestIntersection - start) And Member.CL > start Then
                'check if FO is along girt height OR if member intersects with girt height
                If (FO.bEdgeHeight < Height And FO.tEdgeHeight > Height) Or (Member.bEdgeHeight < Height And Member.tEdgeHeight > Height) Then
                    tempNearestIntersection = Member.CL
                End If
            End If
        End If
    Next item
    If (FO.rEdgePosition - start) < (tempNearestIntersection - start) And FO.rEdgePosition > start Then
        'check if FO is along girt height OR if member intersects with girt height
        If (FO.bEdgeHeight < Height And FO.tEdgeHeight > Height) Then
            tempNearestIntersection = FO.rEdgePosition
        End If
    ElseIf (FO.lEdgePosition - start) < (tempNearestIntersection - start) And FO.lEdgePosition > start Then
        'check if FO is along girt height OR if member intersects with girt height
        If (FO.bEdgeHeight < Height And FO.tEdgeHeight > Height) Then
            tempNearestIntersection = FO.lEdgePosition
        End If
    End If
Next FO


NextHorizontalGirtIntersection = tempNearestIntersection
    
    
End Function

Sub TestGirGen(b As clsBuilding)

'start w/ column collection
'start w/ FO collection
'start w/ building

'find out how many girt lines exist
'# of gerts = 2 + RoundDown((TotalHeight - 12)/5)
'create array of column intersections for each girt line
    'turn array into collection of spans
'create array of all FO intersections for each girt line
    'remove negative space for FODoors and MISCFOs from collection of spans
    'MAYBE: add skirts/headers & jambs for MISCFOs and Windows
'RESULT: collection of spans includes:
    'girts
    'MAYBE: skirts/headers/jambs
    'placement description
    
    
'PLACEMENT NAMING CONVENTION:
    '"Screwline height etc.
    '"Segment #1, #2, #3, etc.
    'try to describe startpoint and endpoint
    '"Window #1 Skirt, Window #1 Header, Window #1 jamb(s), etc.
    
'combine collection of spans for each wall, pass to BPP
    'RETURNS: collection of ORDERED members, with "combinedmembers" collection inside each
        'CombinedMembers includes:
            'placement description
            'span length

'questions:
    'are skirts/headers/jambs the same material as the girts?
    'do gerts extend above the bHeight? - yes
    'Single Slope Building ~ 70':
        'Endwall Columns = 4; 2 corners, 2 inside
        'Need 1 interior column
            'Should it line up with the shorter of the 2 endwall columns? or the longer? or neither?
    




End Sub

Sub EndwallExtensionColumnsGen(b As clsBuilding, eWall As String, Optional NewColNum As Integer, Optional Reiterate As Boolean)

'This sub is ONLY used in buildings with 1 bay which also have endwall Extensions, since there are no columns to copy
'The calculation is the same as the Interior Column gen

Dim ColumnCollection As Collection
Dim Column As clsMember
Dim e1CenterColumn As Boolean
Dim e3CenterColumn As Boolean
Dim RafterNum As Integer
Dim i As Integer
Dim j As Integer
Dim ColLocation() As Double
Dim MaxHorizontalDistance As Double
Dim MinHorizontalDistance As Double
Dim StartWidth As Double
Dim EndWidth As Double
Dim PrevWidth As Double
Dim DistanceToPreviousColumn As Double
Dim DistanceToNextColumn As Double
Dim LargerDistance As Double
Dim ColNum As Integer

If eWall = "e1" Then
    Set ColumnCollection = b.e1ExtensionMembers
ElseIf eWall = "e3" Then
    Set ColumnCollection = b.e3ExtensionMembers
End If

RafterNum = b.s2Columns.Count

'find horizontal distance equal to 60' rafter for this building plus maximum and minimum column thicknesses
MaxHorizontalDistance = (60 / (Sqr((b.rPitch / 12) ^ 2 + 1)))
MinHorizontalDistance = (60 / (Sqr((b.rPitch / 12) ^ 2 + 1)))

'Maximum building width = 300'
'maximum interior column # = 4

'if ColNum was not passed, determine ColNum
'Col Num might be passed if this is the second iteration, possibly increasing the number of columns
If Reiterate = True Then
    ColNum = NewColNum
Else
    If ColNum = 0 Then
        If b.rShape = "Gable" Then
            If b.bWidth <= 80 Then
                ColNum = 0
            ElseIf b.bWidth > 80 And b.bWidth < (MaxHorizontalDistance * 2) Then
                ColNum = 1
            ElseIf b.bWidth >= MaxHorizontalDistance * 2 Then
                ColNum = Application.WorksheetFunction.RoundUp(b.bWidth / MaxHorizontalDistance, 0) - 1
            End If
        ElseIf b.rShape = "Single Slope" Then
            If b.bWidth < MaxHorizontalDistance Then
                ColNum = 0
            ElseIf b.bWidth > MaxHorizontalDistance Then
                ColNum = Application.WorksheetFunction.RoundUp(b.bWidth / MaxHorizontalDistance, 0) - 1
            End If
        End If
    End If
    'lower Col Num by 1 on first iteration to check for marginal cases
    'some column widths (to be determined) will require less columns, this will check those cases
    If ColNum > 0 Then
        ColNum = ColNum - 1
    End If
End If

'first, evenly space columns along the width of the building to adjust later; add to array
ReDim ColLocation(ColNum + 1)
'includes s2 and s4 columns along rafter lines
ColLocation(0) = 0
ColLocation(ColNum + 1) = b.bWidth * 12
Select Case ColNum
Case 1
    ColLocation(1) = b.bWidth / 2 * 12
Case 2
    ColLocation(1) = b.bWidth / 3 * 12
    ColLocation(2) = b.bWidth / 3 * 12 * 2
Case 3
    ColLocation(1) = b.bWidth / 4 * 12
    ColLocation(2) = b.bWidth / 4 * 12 * 2
    ColLocation(3) = b.bWidth / 4 * 12 * 3
Case 4
    ColLocation(1) = b.bWidth / 5 * 12
    ColLocation(2) = b.bWidth / 5 * 12 * 2
    ColLocation(3) = b.bWidth / 5 * 12 * 3
    ColLocation(4) = b.bWidth / 5 * 12 * 4
End Select

If eWall = "e3" Then
    For i = 0 To UBound(ColLocation())
        ColLocation(i) = b.bWidth * 12 - ColLocation(i)
    Next i
End If

'loop through array and check if columns conflict with OHDoors; if so, move 5' away from nearest edge
For i = 1 To ColNum
    If ConflictingEndwallOHDoor(ColLocation(i), b) = True Then
        ColLocation(i) = NearestEndwallLocation(ColLocation(i), b)
    End If
Next i

'''''''''''''''check for No Interior Columns
If ColNum = 0 Then
    ''''''''''''''Distance between Columns
    If b.rShape = "Single Slope" Then
        DistanceToPreviousColumn = Abs(ColLocation(0) - ColLocation(1))
    
    ''''''''''''''Estimate COlumn widths
        'get first width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(0))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(0) - ColLocation(1))
        'subtractfirst'subtract half of first width
        DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
        'get second width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(1))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(0) - ColLocation(1))
        'subtract half of second width
        DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
    Else
        DistanceToPreviousColumn = (b.bWidth * 12 / 2)
    End If

    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
    
    If DistanceToPreviousColumn > (MaxHorizontalDistance * 12) Then
        Erase ColLocation
        Call EndwallExtensionColumnsGen(b, eWall, ColNum + 1, True)
        Exit Sub
    End If
End If


'''''''''''''''check Interior Columns
'check that columns are no more than MaxHorizontalDistance ft apart since they may have been moved
For i = 1 To ColNum
    'get distance to next column to make sure it does NOT exceed max rafter length
    'if the two rafters stradle the center and the roof shape is "Gable", then go only to the center
    'estimate column widths to get accurate distances
    
    ''''''''''''''Distance to PREVIOUS Column
    If ColLocation(i) > (b.bWidth * 12 / 2) And ColLocation(i - 1) < (b.bWidth * 12 / 2) And b.rShape = "Gable" Then
        DistanceToPreviousColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
    Else
        DistanceToPreviousColumn = Abs(ColLocation(i) - ColLocation(i - 1))
    End If
        ''''''''''''''Estimate COlumn widths
        'get first width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof(eWall, ColLocation(i))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
        'subtract half of width
        DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
        'get second width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof(eWall, ColLocation(i - 1))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
        'subtract width if sidewall column, or half of width otherwise
        If i - 1 = 0 Then
            DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
        Else
            DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
        End If
    
    ''''''''''''''Distance to NEXT Column
    If ColLocation(i) < b.bWidth * 12 / 2 And ColLocation(i + 1) > b.bWidth * 12 / 2 And b.rShape = "Gable" Then
        DistanceToNextColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
    Else
        DistanceToNextColumn = Abs(ColLocation(i) - ColLocation(i + 1))
    End If
        ''''''''''''''Estimate COlumn widths
        'get first width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
        'subtract half of width
        DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
        'get second width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof(eWall, ColLocation(i + 1))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
        'subtract width if sidewall column, or half of width otherwise
        If i + 1 = UBound(ColLocation()) Then
            DistanceToNextColumn = DistanceToNextColumn - Column.Width
        Else
            DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
        End If
    
    'check if the columns are too far apart; if so, run this sub again with 1 more column (optional parameter)
    If DistanceToPreviousColumn > (MaxHorizontalDistance * 12) Or DistanceToNextColumn > (MaxHorizontalDistance * 12) Then
        'Debug.Print "columns too far apart"
        'CHECK COLUMN DISTANCES AGAIN WITH NEW COLUMN WIDTH ESTIMATES
        If NearestEndwallLocation(ColLocation(i), b, "Alternate") <> ColLocation(i) Then
            ColLocation(i) = NearestEndwallLocation(ColLocation(i), b, "Alternate")
            ''''''''''''''Distance to PREVIOUS Column
            If ColLocation(i) > b.bWidth * 12 / 2 And ColLocation(i - 1) < b.bWidth * 12 / 2 And b.rShape = "Gable" Then
                DistanceToPreviousColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
            Else
                DistanceToPreviousColumn = Abs(ColLocation(i) - ColLocation(i - 1))
            End If
                ''''''''''''''Estimate COlumn widths
                'get first width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof(eWall, ColLocation(i))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
                'subtract half of width
                DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
                'get second width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof(eWall, ColLocation(i - 1))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
                'subtract width if sidewall column, or half of width otherwise
                If i - 1 = 0 Then
                    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
                Else
                    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
                End If
            ''''''''''''''Distance to NEXT Column
            If ColLocation(i) < b.bWidth * 12 / 2 And ColLocation(i + 1) > b.bWidth * 12 / 2 And b.rShape = "Gable" Then
                DistanceToNextColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
            Else
                DistanceToNextColumn = Abs(ColLocation(i) - ColLocation(i + 1))
            End If
                ''''''''''''''Estimate COlumn widths
                'get first width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof(eWall, ColLocation(i))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
                'subtract half of width
                DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
                'get second width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof(eWall, ColLocation(i + 1))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
                'subtract width if sidewall column, or half of width otherwise
                If i + 1 = UBound(ColLocation()) Then
                    DistanceToNextColumn = DistanceToNextColumn - Column.Width
                Else
                    DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
                End If
        End If
        If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
            Erase ColLocation
            Call IntColumnsGen(b, ColNum + 1, True)
            Exit Sub
        End If
'    ElseIf DistanceToPreviousColumn <= MaxHorizontalDistance * 12 And DistanceToPreviousColumn >= MinHorizontalDistance * 12 _
'    Or DistanceToNextColumn <= MaxHorizontalDistance * 12 And DistanceToNextColumn >= MinHorizontalDistance * 12 Then
'    'if distance is between the min and max horizontal value, we need to check actual column widths and recheck.
'        EndWidth = MinimumInteriorColumnWidth(b, i + 1, ColLocation) / 2
'        StartWidth = MinimumInteriorColumnWidth(b, i, ColLocation) / 2
'        PrevWidth = MinimumInteriorColumnWidth(b, i - 1, ColLocation) / 2
'        DistanceToPreviousColumn = Abs((ColLocation(i) - StartWidth) - (ColLocation(i - 1) + PrevWidth))
'        DistanceToNextColumn = Abs((ColLocation(i) + StartWidth) - (ColLocation(i + 1) - PrevWidth))
'        If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
'            Erase ColLocation
'            Call IntColumnsGen(b, ColNum + 1)
'            Exit Sub
'        End If
    End If
Next i

'debugging
For i = 0 To ColNum + 1
    Debug.Print "Column #: " & i & ", " & ColLocation(i) / 12 & "' from s2, Rafter Line " & i
Next i





'set column variables, types, sizes, etc.
For i = 0 To ColNum + 1
    'find larger distance to neighboring columns to use in lookup tables
    's2 and s4 columns only have 1 value, all other columns have 2 neighboring columns, the one farthest away is the distance used
    If i = ColNum + 1 Then
        LargerDistance = Abs(ColLocation(i) - ColLocation(i - 1))
    ElseIf i = 0 Then
        LargerDistance = Abs(ColLocation(i) - ColLocation(i + 1))
    Else
        LargerDistance = Application.WorksheetFunction.Max(Abs(ColLocation(i) - ColLocation(i - 1)), Abs(ColLocation(i) - ColLocation(i + 1)))
    End If
    
    Set Column = New clsMember
    Column.mType = eWall & " Extension Column"
    Column.CL = ColLocation(i)
    Column.LoadBearing = True
    Column.Qty = 1
    Column.Placement = eWall & " Extension Column"
    If b.rShape = "Single Slope" Then
        If i = 0 Then
            Column.Length = ((b.bWidth * 12) * (b.rPitch / 12)) + b.bHeight * 12
        ElseIf i = ColNum + 1 Then
            Column.Length = b.bHeight * 12
        Else
            Column.Length = b.DistanceToRoof(eWall, Column.CL)
        End If
    Else 'Gable
        If i = 0 Then
            Column.Length = b.bHeight * 12
        ElseIf i = ColNum + 1 Then
            Column.Length = b.bHeight * 12
        Else
            Column.Length = b.DistanceToRoof(eWall, Column.CL)
        End If
    End If
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", LargerDistance
    If Column.CL = 0 Then
        Column.CL = Column.Width / 2
    ElseIf Column.CL = b.bWidth * 12 Then
        Column.CL = b.bWidth * 12 - Column.Width / 2
    End If
    Column.rEdgePosition = Column.CL - Column.Width / 2
    Column.Placement = eWall & " Extension Column"
    ColumnCollection.Add Column
Next i

End Sub


Sub IntColumnsGen(b As clsBuilding, Optional NewColNum As Integer, Optional Reiterate As Boolean)

Dim e1ColumnCollection As Collection
Dim e3ColumnCollection As Collection
Dim Column As clsMember
Dim e1CenterColumn As Boolean
Dim e3CenterColumn As Boolean
Dim RafterNum As Integer
Dim i As Integer
Dim j As Integer
Dim ColLocation() As Double
Dim MaxHorizontalDistance As Double
Dim MinHorizontalDistance As Double
Dim StartWidth As Double
Dim EndWidth As Double
Dim PrevWidth As Double
Dim DistanceToPreviousColumn As Double
Dim DistanceToNextColumn As Double
Dim LargerDistance As Double
Dim ColNum As Integer

'check for building with only 1 bay (no main rafter lines)
If EstSht.Range("BayNum").Value = 1 And NewColNum = 0 Then
    Exit Sub
End If

RafterNum = b.s2Columns.Count

'find horizontal distance equal to 60' rafter for this building plus maximum and minimum column thicknesses
MaxHorizontalDistance = (60 / (Sqr((b.rPitch / 12) ^ 2 + 1)))
MinHorizontalDistance = (60 / (Sqr((b.rPitch / 12) ^ 2 + 1)))

'Maximum building width = 300'
'maximum interior column # = 4

'if ColNum was not passed, determine ColNum
'Col Num might be passed if this is the second iteration, possibly increasing the number of columns
If Reiterate = True Then
    ColNum = NewColNum
Else
    If ColNum = 0 Then
        If b.rShape = "Gable" Then
            If b.bWidth <= 80 Then
                ColNum = 0
            ElseIf b.bWidth > 80 And b.bWidth < (MaxHorizontalDistance * 2) Then
                ColNum = 1
            ElseIf b.bWidth >= MaxHorizontalDistance * 2 Then
                ColNum = Application.WorksheetFunction.RoundUp(b.bWidth / MaxHorizontalDistance, 0) - 1
            End If
        ElseIf b.rShape = "Single Slope" Then
            If b.bWidth < MaxHorizontalDistance Then
                ColNum = 0
            ElseIf b.bWidth > MaxHorizontalDistance Then
                ColNum = Application.WorksheetFunction.RoundUp(b.bWidth / MaxHorizontalDistance, 0) - 1
            End If
        End If
    End If
    'lower Col Num by 1 on first iteration to check for marginal cases
    'some column widths (to be determined) will require less columns, this will check those cases
    If ColNum > 0 Then
        ColNum = ColNum - 1
    End If
End If

'first, evenly space columns along the width of the building to adjust later; add to array
ReDim ColLocation(ColNum + 1)
'includes s2 and s4 columns along rafter lines
ColLocation(0) = 0
ColLocation(ColNum + 1) = b.bWidth * 12
Select Case ColNum
Case 1
    ColLocation(1) = b.bWidth / 2 * 12
Case 2
    ColLocation(1) = b.bWidth / 3 * 12
    ColLocation(2) = b.bWidth / 3 * 12 * 2
Case 3
    ColLocation(1) = b.bWidth / 4 * 12
    ColLocation(2) = b.bWidth / 4 * 12 * 2
    ColLocation(3) = b.bWidth / 4 * 12 * 3
Case 4
    ColLocation(1) = b.bWidth / 5 * 12
    ColLocation(2) = b.bWidth / 5 * 12 * 2
    ColLocation(3) = b.bWidth / 5 * 12 * 3
    ColLocation(4) = b.bWidth / 5 * 12 * 4
End Select

'loop through array and check if columns conflict with OHDoors; if so, move 5' away from nearest edge
For i = 1 To ColNum
    If ConflictingEndwallOHDoor(ColLocation(i), b) = True Then
        ColLocation(i) = NearestEndwallLocation(ColLocation(i), b)
    End If
Next i

'''''''''''''''check for No Interior Columns
If ColNum = 0 Then
    ''''''''''''''Distance between Columns
    DistanceToPreviousColumn = Abs(ColLocation(0) - ColLocation(1))

    ''''''''''''''Estimate COlumn widths
    'get first width
    Set Column = New clsMember
    Column.Length = b.DistanceToRoof("e1", ColLocation(0))
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", Abs(ColLocation(0) - ColLocation(1))
    'subtract half of first width
    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
    'get second width
    Set Column = New clsMember
    Column.Length = b.DistanceToRoof("e1", ColLocation(1))
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", Abs(ColLocation(0) - ColLocation(1))
    'subtract half of second width
    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
    'check for freespan exception before reiterating
    If Not (b.rShape = "Gable" And b.bWidth <= 80) Then
        If DistanceToPreviousColumn > (MaxHorizontalDistance * 12) Then
            Erase ColLocation
            Call IntColumnsGen(b, ColNum + 1, True)
            Exit Sub
        End If
    End If
End If


'''''''''''''''check Interior Columns
'check that columns are no more than MaxHorizontalDistance ft apart since they may have been moved
For i = 1 To ColNum
    'get distance to next column to make sure it does NOT exceed max rafter length
    'if the two rafters stradle the center and the roof shape is "Gable", then go only to the center
    'estimate column widths to get accurate distances
    
    ''''''''''''''Distance to PREVIOUS Column
    If ColLocation(i) > (b.bWidth * 12 / 2) And ColLocation(i - 1) < (b.bWidth * 12 / 2) And b.rShape = "Gable" Then
        DistanceToPreviousColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
    Else
        DistanceToPreviousColumn = Abs(ColLocation(i) - ColLocation(i - 1))
    End If
        ''''''''''''''Estimate COlumn widths
        'get first width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
        'subtract half of width
        DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
        'get second width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i - 1))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
        'subtract width if sidewall column, or half of width otherwise
        If i - 1 = 0 Then
            DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
        Else
            DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
        End If
    
    ''''''''''''''Distance to NEXT Column
    If ColLocation(i) < b.bWidth * 12 / 2 And ColLocation(i + 1) > b.bWidth * 12 / 2 And b.rShape = "Gable" Then
        DistanceToNextColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
    Else
        DistanceToNextColumn = Abs(ColLocation(i) - ColLocation(i + 1))
    End If
        ''''''''''''''Estimate COlumn widths
        'get first width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
        'subtract half of width
        DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
        'get second width
        Set Column = New clsMember
        Column.Length = b.DistanceToRoof("e1", ColLocation(i + 1))
        Column.tEdgeHeight = Column.Length
        Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
        'subtract width if sidewall column, or half of width otherwise
        If i + 1 = UBound(ColLocation()) Then
            DistanceToNextColumn = DistanceToNextColumn - Column.Width
        Else
            DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
        End If
    
    'check if the columns are too far apart; if so, run this sub again with 1 more column (optional parameter)
    If DistanceToPreviousColumn > (MaxHorizontalDistance * 12) Or DistanceToNextColumn > (MaxHorizontalDistance * 12) Then
        'Debug.Print "columns too far apart"
        'CHECK COLUMN DISTANCES AGAIN WITH NEW COLUMN WIDTH ESTIMATES
        If NearestEndwallLocation(ColLocation(i), b, "Alternate") <> ColLocation(i) Then
            ColLocation(i) = NearestEndwallLocation(ColLocation(i), b, "Alternate")
            ''''''''''''''Distance to PREVIOUS Column
            If ColLocation(i) > b.bWidth * 12 / 2 And ColLocation(i - 1) < b.bWidth * 12 / 2 And b.rShape = "Gable" Then
                DistanceToPreviousColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
            Else
                DistanceToPreviousColumn = Abs(ColLocation(i) - ColLocation(i - 1))
            End If
                ''''''''''''''Estimate COlumn widths
                'get first width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof("e1", ColLocation(i))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
                'subtract half of width
                DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
                'get second width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof("e1", ColLocation(i - 1))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i - 1) - ColLocation(i))
                'subtract width if sidewall column, or half of width otherwise
                If i - 1 = 0 Then
                    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width
                Else
                    DistanceToPreviousColumn = DistanceToPreviousColumn - Column.Width / 2
                End If
            ''''''''''''''Distance to NEXT Column
            If ColLocation(i) < b.bWidth * 12 / 2 And ColLocation(i + 1) > b.bWidth * 12 / 2 And b.rShape = "Gable" Then
                DistanceToNextColumn = Abs(b.bWidth * 12 / 2 - ColLocation(i))
            Else
                DistanceToNextColumn = Abs(ColLocation(i) - ColLocation(i + 1))
            End If
                ''''''''''''''Estimate COlumn widths
                'get first width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof("e1", ColLocation(i))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
                'subtract half of width
                DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
                'get second width
                Set Column = New clsMember
                Column.Length = b.DistanceToRoof("e1", ColLocation(i + 1))
                Column.tEdgeHeight = Column.Length
                Column.SetSize b, "Column", "Interior", Abs(ColLocation(i + 1) - ColLocation(i))
                'subtract width if sidewall column, or half of width otherwise
                If i + 1 = UBound(ColLocation()) Then
                    DistanceToNextColumn = DistanceToNextColumn - Column.Width
                Else
                    DistanceToNextColumn = DistanceToNextColumn - Column.Width / 2
                End If
        End If
        If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
            Erase ColLocation
            Call IntColumnsGen(b, ColNum + 1, True)
            Exit Sub
        End If
'    ElseIf DistanceToPreviousColumn <= MaxHorizontalDistance * 12 And DistanceToPreviousColumn >= MinHorizontalDistance * 12 _
'    Or DistanceToNextColumn <= MaxHorizontalDistance * 12 And DistanceToNextColumn >= MinHorizontalDistance * 12 Then
'    'if distance is between the min and max horizontal value, we need to check actual column widths and recheck.
'        EndWidth = MinimumInteriorColumnWidth(b, i + 1, ColLocation) / 2
'        StartWidth = MinimumInteriorColumnWidth(b, i, ColLocation) / 2
'        PrevWidth = MinimumInteriorColumnWidth(b, i - 1, ColLocation) / 2
'        DistanceToPreviousColumn = Abs((ColLocation(i) - StartWidth) - (ColLocation(i - 1) + PrevWidth))
'        DistanceToNextColumn = Abs((ColLocation(i) + StartWidth) - (ColLocation(i + 1) - PrevWidth))
'        If DistanceToPreviousColumn > MaxHorizontalDistance * 12 Or DistanceToNextColumn > MaxHorizontalDistance * 12 Then
'            Erase ColLocation
'            Call IntColumnsGen(b, ColNum + 1)
'            Exit Sub
'        End If
    End If
Next i

'debugging
For i = 0 To ColNum + 1
    Debug.Print "Column #: " & i & ", " & ColLocation(i) / 12 & "' from s2, Rafter Line " & i
Next i

'use temporary columns to find sidewall col widths first
If b.rShape = "Single Slope" Then
    'Sidewall 2 Column Width:
    Set Column = New clsMember
    Column.Length = b.bHeight * 12
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", Abs(ColLocation(ColNum) - ColLocation(ColNum + 1))
    b.s2ColumnWidth = Column.Width
    'Sidewall 4 Column Width:
    Set Column = New clsMember
    Column.Length = b.bWidth * b.rPitch + b.bHeight * 12
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", Abs(ColLocation(0) - ColLocation(1))
    b.s4ColumnWidth = Column.Width
Else
    'Sidewall 2 Column Width:
    Set Column = New clsMember
    Column.Length = b.bHeight * 12
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", Abs(ColLocation(ColNum) - ColLocation(ColNum + 1))
    b.s2ColumnWidth = Column.Width
    'Sidewall 4 Column Width:
    Set Column = New clsMember
    Column.Length = b.bHeight * 12
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", Abs(ColLocation(0) - ColLocation(1))
    b.s4ColumnWidth = Column.Width
End If



'set column variables, types, sizes, etc.
For i = 0 To ColNum + 1
    'find larger distance to neighboring columns to use in lookup tables
    's2 and s4 columns only have 1 value, all other columns have 2 neighboring columns, the one farthest away is the distance used
    If i = ColNum + 1 Then
        LargerDistance = Abs(ColLocation(i) - ColLocation(i - 1))
    ElseIf i = 0 Then
        LargerDistance = Abs(ColLocation(i) - ColLocation(i + 1))
    Else
        LargerDistance = Application.WorksheetFunction.Max(Abs(ColLocation(i) - ColLocation(i - 1)), Abs(ColLocation(i) - ColLocation(i + 1)))
    End If
    
    Set Column = New clsMember
    Column.mType = "Column"
    Column.CL = ColLocation(i)
    Column.LoadBearing = True
    Column.Qty = RafterNum
    Column.Placement = "main rafter line interior column number " & i
    If b.rShape = "Single Slope" Then
        If i = 0 Then
            Column.Length = ((b.bWidth * 12) * (b.rPitch / 12)) + b.bHeight * 12
        ElseIf i = ColNum + 1 Then
            Column.Length = b.bHeight * 12
        Else
            Column.Length = b.DistanceToRoof("e1", Column.CL)
        End If
    Else 'Gable
        If i = 0 Then
            Column.Length = b.bHeight * 12
        ElseIf i = ColNum + 1 Then
            Column.Length = b.bHeight * 12
        Else
            Column.Length = b.DistanceToRoof("e1", Column.CL)
        End If
    End If
    Column.tEdgeHeight = Column.Length
    Column.SetSize b, "Column", "Interior", LargerDistance
    If Column.CL = 0 Then
        Column.CL = Column.Width / 2
    ElseIf Column.CL = b.bWidth * 12 Then
        Column.CL = b.bWidth * 12 - Column.Width / 2
    End If
    Column.rEdgePosition = Column.CL - Column.Width / 2
    Column.Placement = Column.Size & " interior column, " & Column.Length & "' long"
    b.InteriorColumns.Add Column
Next i

End Sub

'given column along rafterline, return the estimated width of column based on height and distance to nearest columns
 Function MinimumInteriorColumnWidth(b As clsBuilding, ColIndex As Integer, Columns() As Double) As Double

Dim ColHeight As Double
Dim PrevColDistance As Double
Dim NextColDistance As Double
Dim MaxColDistance As Double
Dim ColumnType As String
Dim WidthCell As Range
Dim HeightCell As Range
Dim Depth As String
Dim Width As String
Dim i As Integer
Dim LookupTbl As ListObject

Set LookupTbl = SteelLookupSht.ListObjects("MainColumnAndExpandableEndwallColumnTbl")

If ColIndex = UBound(Columns) Then
    NextColDistance = 0
Else
    NextColDistance = Columns(ColIndex + 1) - Columns(ColIndex)
End If

ColHeight = b.DistanceToRoof("e1", Columns(ColIndex))

For Each WidthCell In LookupTbl.ListColumns(1).DataBodyRange
    If (NextColDistance / 12 >= WidthCell.Value And NextColDistance / 12 < WidthCell.offset(1, 0).Value) Or NextColDistance < 30 * 12 Then
        Exit For
    End If
Next WidthCell

For i = 2 To LookupTbl.HeaderRowRange.Cells.Count - 2
    If (ColHeight / 12 >= LookupTbl.HeaderRowRange(1, i).Value And ColHeight / 12 < LookupTbl.HeaderRowRange(1, i + 1).Value) Or ColHeight / 12 <= 30 Then
        Exit For
    End If
Next i

ColumnType = WidthCell.offset(0, i - 1).Value

Width = Right(Left(ColumnType, InStr(1, ColumnType, "x") - 1), Len(Left(ColumnType, InStr(1, ColumnType, "x") - 1)) - 1)

MinimumInteriorColumnWidth = CDbl(Width)
    

End Function

'function returns nearest endwall location that does not conflict with an OHDoor
 Function NearestEndwallLocation(Location As Double, b As clsBuilding, Optional Alternate As String, Optional eWall As String) As Double


Dim Column As clsMember
Dim e1Column As clsMember
Dim e3Column As clsMember
Dim e1ColLocation As Double
Dim e3ColLocation As Double
Dim e1tempNearestLocation As Double
Dim e3tempNearestLocation As Double
Dim FO As clsFO
Dim iterationLocation As Double
Dim AlternateEdge As Boolean

If Alternate = "Alternate" Then
    AlternateEdge = True
Else
    AlternateEdge = False
End If

e1tempNearestLocation = 1.79769313486231E+308 'initialize
e3tempNearestLocation = 1.79769313486231E+308 'initialize

If eWall <> "e3" Then
    'get closest OHDoor Edge for e1; only called if columns don't work
    For Each FO In b.e1FOs
        If FO.FOType = "OHDoor" Then
            If AlternateEdge = False Then
                If (FO.rEdgePosition < Location And FO.lEdgePosition > Location) Then
                    If Abs(FO.rEdgePosition - Location) < Abs(FO.lEdgePosition - Location) Then
                        e1tempNearestLocation = FO.rEdgePosition
                    Else
                        e1tempNearestLocation = FO.lEdgePosition
                    End If
                End If
            ElseIf AlternateEdge = True Then
                If (FO.rEdgePosition - 1 < Location And FO.lEdgePosition + 1 > Location) Then
                    If Abs(FO.rEdgePosition - Location) < Abs(FO.lEdgePosition - Location) Then
                        e1tempNearestLocation = FO.lEdgePosition
                    Else
                        e1tempNearestLocation = FO.rEdgePosition
                    End If
                End If
            End If
        End If
    Next FO
End If
If eWall <> "e1" Then
    'get closest OHDoor Edge for e3; only called if columns don't work
    For Each FO In b.e3FOs
        If FO.FOType = "OHDoor" Then
            If AlternateEdge = False Then
                If (b.bWidth * 12 - FO.rEdgePosition > Location And b.bWidth * 12 - FO.lEdgePosition < Location) Then
                    If Abs(b.bWidth - FO.rEdgePosition - Location) < Abs(b.bWidth - FO.lEdgePosition - Location) Then
                        e3tempNearestLocation = b.bWidth * 12 - FO.rEdgePosition
                    Else
                        e3tempNearestLocation = b.bWidth * 12 - FO.lEdgePosition
                    End If
                End If
            ElseIf AlternateEdge = True Then
                If (b.bWidth * 12 - FO.rEdgePosition + 1 > Location And b.bWidth * 12 - FO.lEdgePosition - 1 < Location) Then
                    If Abs(b.bWidth * 12 - FO.rEdgePosition - Location) < Abs(b.bWidth * 12 - FO.lEdgePosition - Location) Then
                        e3tempNearestLocation = b.bWidth * 12 - FO.lEdgePosition
                    Else
                        e3tempNearestLocation = b.bWidth * 12 - FO.rEdgePosition
                    End If
                End If
            End If
        End If
    Next FO
End If

If eWall = "e1" Then
    NearestEndwallLocation = e1tempNearestLocation
ElseIf eWall = "e3" Then
    NearestEndwallLocation = e3tempNearestLocation
Else
    If ConflictingEndwallOHDoor(e1tempNearestLocation, b) Then
        NearestEndwallLocation = e3tempNearestLocation
    ElseIf ConflictingEndwallOHDoor(e3tempNearestLocation, b) Then
        NearestEndwallLocation = e1tempNearestLocation
    ElseIf Abs(e1tempNearestLocation - Location) < Abs(e3tempNearestLocation - Location) Then
        NearestEndwallLocation = e1tempNearestLocation
    Else
        NearestEndwallLocation = e3tempNearestLocation
    End If
End If

If NearestEndwallLocation = 1.79769313486231E+308 Then
    NearestEndwallLocation = Location
End If

End Function

'Returns TRUE if a location has matching endwall columns on BOTH ends of the building
 Function MatchingEndwallColumn(Location As Double, b As clsBuilding) As Boolean

Dim Column As clsMember
Set Column = New clsMember
Dim e1MatchingEndwallColumn As Boolean
Dim e3MatchingEndwallColumn As Boolean

e1MatchingEndwallColumn = False
e3MatchingEndwallColumn = False

For Each Column In b.e1Columns
    If Column.CL = Location Then
        e1MatchingEndwallColumn = True
    End If
Next Column

For Each Column In b.e3Columns
    If Column.CL = b.bWidth * 12 - Location Then
        e3MatchingEndwallColumn = True
    End If
Next Column

If e1MatchingEndwallColumn = True And e3MatchingEndwallColumn = True Then
    MatchingEndwallColumn = True
Else
    MatchingEndwallColumn = False
End If

End Function

'Returns TRUE if location conflicts with an OHDoor on either Endwall
 Function ConflictingEndwallOHDoor(Location As Double, b As clsBuilding, Optional eWall As String) As Boolean

Dim FO As clsFO

Dim e1Conflict As Boolean
Dim e3Conflict As Boolean

If eWall <> "e3" Then
    For Each FO In b.e1FOs
        If FO.FOType = "OHDoor" Then
            If Location > FO.rEdgePosition And Location < FO.lEdgePosition Then
                e1Conflict = True
            Else
                e1Conflict = False
            End If
        End If
    Next FO
End If

If eWall <> "e1" Then
    For Each FO In b.e3FOs
        If FO.FOType = "OHDoor" Then
            If Location < b.bWidth * 12 - FO.rEdgePosition And Location > b.bWidth * 12 - FO.lEdgePosition Then
                e3Conflict = True
            Else
                e3Conflict = False
            End If
        End If
    Next FO
End If

If e1Conflict = True Or e3Conflict = True Then
    ConflictingEndwallOHDoor = True
Else
    ConflictingEndwallOHDoor = False
End If

End Function

 Sub AdjustSidewallColumns(b As clsBuilding, eWall As String)

Dim ColumnCollection As Collection
Dim IntColumnCollection As Collection
Dim NearestColumn As clsMember
Dim Column As clsMember
Dim IntColumn As clsMember
Dim Index As Integer
Dim WedgeDistance As Double


Select Case eWall
Case "s2"
    'reset s2 Columns to be the same size/type as the interior columns along s2
    Set ColumnCollection = b.s2Columns
    Set IntColumnCollection = b.InteriorColumns
    For Index = 1 To IntColumnCollection.Count
        Set IntColumn = IntColumnCollection(Index)
        If IntColumn.CL > b.bWidth * 12 - 15 Then '15 is half the widest possible column
            Set NearestColumn = IntColumn
            'IntColumnCollection.Remove (Index)
            Exit For
        End If
    Next Index
    b.s2ColumnWidth = NearestColumn.Width
    'Calculate angle cut for s2 columns
    'increase size of each column to account for angle cut
    WedgeDistance = b.s2ColumnWidth * b.rPitch / 12
    For Each Column In ColumnCollection
        Column.Size = NearestColumn.Size
        Column.Width = NearestColumn.Width
        Column.tEdgeHeight = Column.tEdgeHeight + WedgeDistance
        Column.Length = Column.Length + WedgeDistance
    Next Column
    'Calculate angle cut for s2 columns
    WedgeDistance = b.s2ColumnWidth * b.rPitch / 12
    
Case "s4"
    Set ColumnCollection = b.s4Columns
    Set IntColumnCollection = b.InteriorColumns
    For Index = 1 To IntColumnCollection.Count
        Set IntColumn = IntColumnCollection(Index)
        If IntColumn.CL < 15 Then '15 is half the widest column possible
            Set NearestColumn = IntColumn
            'IntColumnCollection.Remove (Index)
            Exit For
        End If
    Next Index
    b.s4ColumnWidth = NearestColumn.Width
    'Calculate angle cut for s2 columns
    'IF GABLE: increase size of each column to account for angle cut
    WedgeDistance = b.s2ColumnWidth * b.rPitch / 12
    For Each Column In ColumnCollection
        Column.Size = NearestColumn.Size
        Column.Width = NearestColumn.Width
        If b.rShape = "Gable" Then
            Column.tEdgeHeight = Column.tEdgeHeight + WedgeDistance
            Column.Length = Column.Length + WedgeDistance
        End If
    Next Column
End Select

End Sub

 Sub DisplayDrawingInfo(Placement As Double)

'Dim CallingShape As String

'CallingShape = Application.Caller

'If CallingShape Like "*" & "Straight Connector" & "*" Then
'    Exit Sub
'End If

'MsgBox "Length: " & Application.Round(CDbl(CallingShape), 2) & """"
MsgBox "Length: " & ImperialMeasurementFormat(Placement) & "'"

'ActiveSheet.Shapes(Application.Caller).Select

End Sub

Sub DrawDimension(b As clsBuilding, xLeft As Double, yTop As Double, Width As Double, Height As Double, Direction As String, Font As Integer, Dimensions() As Variant, Optional Label As String)
Dim MyShape As Shape
Dim TextDistance As Double
Dim i As Double

Call QuickSort(Dimensions, LBound(Dimensions), UBound(Dimensions))

If Direction = "Horizontal" Then
    'Draw full line
    Set MyShape = ThisWorkbook.Sheets("Wall Drawings").Shapes.AddLine(xLeft, yTop + Height / 2, xLeft + Width, yTop + Height / 2)
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
        .BeginArrowheadLength = msoArrowheadLong
        .BeginArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadWidth = msoArrowheadWidthMedium
        .EndArrowheadLength = msoArrowheadLong
        .EndArrowheadStyle = msoArrowheadTriangle
        .EndArrowheadWidth = msoArrowheadWidthMedium
    End With
    MyShape.Select
    For i = 0 To UBound(Dimensions)
    If Not IsEmpty(Dimensions(i)) Then
        'Draw vertical lines
        Set MyShape = ThisWorkbook.Sheets("Wall Drawings").Shapes.AddLine(xLeft + Dimensions(i), yTop, xLeft + Dimensions(i), yTop + Height)
        With MyShape.Line
            .ForeColor.RGB = RGB(0, 20, 132)
            .Weight = 0.5
            .DashStyle = msoLineDash
        End With
        If i < UBound(Dimensions) Then
            TextDistance = TextDistance + Dimensions(i + 1) - Dimensions(i)
            Set MyShape = ThisWorkbook.Sheets("Wall Drawings").Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                If (Dimensions(i + 1) - Dimensions(i)) < 60 Then
                    .Left = xLeft + TextDistance - ((Dimensions(i + 1) - Dimensions(i)) / 2) - 15
                    .Top = yTop - 45
                    .Width = 30
                Else
                    .Left = xLeft + TextDistance - ((Dimensions(i + 1) - Dimensions(i)) / 2) - 30
                    .Top = yTop
                    .Width = 60
                End If
                .Height = Height
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Fill.Transparency = 0
                .Line.Transparency = 1
            End With
            With MyShape.TextFrame
                .Characters.Text = ImperialMeasurementFormat(Dimensions(i + 1) - Dimensions(i))
                .Characters.Font.Bold = True
                .Characters.Font.Size = Font
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .Characters.Font.Color = RGB(0, 20, 132)
                
            End With
        End If
    End If
    Next i
    
Else 'Vertical orientation
    'Draw full line
    Set MyShape = ThisWorkbook.Sheets("Wall Drawings").Shapes.AddLine(xLeft + Width / 2, yTop, xLeft + Width / 2, yTop + Height)
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
        .BeginArrowheadLength = msoArrowheadLong
        .BeginArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadWidth = msoArrowheadWidthMedium
        .EndArrowheadLength = msoArrowheadLong
        .EndArrowheadStyle = msoArrowheadTriangle
        .EndArrowheadWidth = msoArrowheadWidthMedium
    End With
    For i = 0 To UBound(Dimensions)
        'Draw horizontal lines
        Set MyShape = ThisWorkbook.Sheets("Wall Drawings").Shapes.AddLine(xLeft, yTop + Dimensions(i), xLeft + Width, yTop + Dimensions(i))
        With MyShape.Line
            .ForeColor.RGB = RGB(0, 20, 132)
            .Weight = 0.5
            .DashStyle = msoLineDash
        End With
        If i < UBound(Dimensions) Then
            TextDistance = TextDistance + Dimensions(i + 1) - Dimensions(i)
            Set MyShape = ThisWorkbook.Sheets("Wall Drawings").Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                If Dimensions(i + 1) - Dimensions(i) < 60 Then
                    .Left = xLeft + 25
                    .Top = yTop + TextDistance - ((Dimensions(i + 1) - Dimensions(i)) / 2) - 30
                Else
                    .Left = xLeft
                    .Top = yTop + TextDistance - ((Dimensions(i + 1) - Dimensions(i)) / 2) - 30
                End If
                .Width = Width
                .Height = 60
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Fill.Transparency = 0
                .Line.Transparency = 1
            End With
            With MyShape.TextFrame
                .Characters.Text = ImperialMeasurementFormat(Dimensions(i + 1) - Dimensions(i))
                .Characters.Font.Bold = True
                .Characters.Font.Size = Font
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .Characters.Font.Color = RGB(0, 20, 132)
            End With
        End If
    Next i
End If
        
End Sub

Function ArrayRemoveDups(MyArray As Variant) As Variant
    Dim nFirst As Long, nLast As Long, i As Long
    Dim item As String
    
    Dim arrTemp() As Variant
    Dim Coll As New Collection
 
    'Get First and Last Array Positions
    nFirst = LBound(MyArray)
    nLast = UBound(MyArray)
    ReDim arrTemp(nFirst To nLast)
 
    'Convert Array to String
    For i = nFirst To nLast
        arrTemp(i) = (MyArray(i))
    Next i
    
    'Populate Temporary Collection
    On Error Resume Next
    For i = nFirst To nLast
        Coll.Add arrTemp(i), CStr(arrTemp(i))
    Next i
    Err.Clear
    On Error GoTo 0
 
    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)
    
    'Populate Array
    For i = nFirst To nLast
        arrTemp(i) = Coll(i - nFirst + 1)
    Next i
    
    'Output Array
    ArrayRemoveDups = arrTemp
 
End Function


Sub DrawItems(b As clsBuilding)
Dim ColumnCollection As Collection
Dim FOCollection As Collection
Dim GirtsCollection As Collection
Dim RafterCollection As Collection
Dim IntColumnCollection As Collection
Dim OverhangCollection As Collection
Dim ExtensionCollection As Collection
Dim Member As clsMember
Dim FO As clsFO
Dim item As Object
Dim eWall As String
Dim TotalHeight As Double
Dim MaxHeight As Double
Dim lEdgePosition As Double
Dim x1 As Integer
Dim y1 As Integer
Dim i As Integer
Dim ColumnWidth As Double
Dim mString As String
Dim Length As String
Dim FloorplanHeight As Double
Dim xf As Double
Dim yf As Double
Dim BayNum As Integer
Dim BayStart As Double
Dim j As Integer
Dim WeldPlate As clsMiscItem
Dim Plate As clsMiscItem
Dim IntDimensionHeight As Double
Dim MyShape As Shape
Dim s2ExtensionEdge As Double
Dim s4ExtensionEdge As Double
Dim DrawSht As Worksheet
Dim DimensionOffset As Double

'Call TestDimension(b)

'delete and remake sheet
Application.DisplayAlerts = False
If sheetExists("Wall Drawings", ThisWorkbook) Then
    ThisWorkbook.Worksheets("Wall Drawings").Delete
End If
Application.DisplayAlerts = True
'create new Drawing Sheet
ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Structural Steel Price List")).Name = "Wall Drawings"
ThisWorkbook.Worksheets("Wall Drawings").Activate
ActiveWindow.DisplayGridlines = False
ActiveWindow.Zoom = 40

Set DrawSht = ThisWorkbook.ActiveSheet


'Set Max Building height
If b.rShape = "Single Slope" Then   'in
    MaxHeight = (b.bHeight * 12) + (b.bWidth * b.rPitch)
Else
    MaxHeight = (b.bHeight * 12) + (b.bWidth / 2 * b.rPitch)
End If

'Set floorplan adjustment
FloorplanHeight = b.bLength * 12 + b.e3Extension
xf = b.bWidth * 12 + b.s2Extension + 350 '350 as buffer
yf = FloorplanHeight + 350 ' 350 as buffer

DrawSht.Range("A1:A12").RowHeight = ((350 + yf + b.e1Extension) / 12)
DrawSht.Range("A1:Z1").ColumnWidth = ((xf + b.s2Extension) / 26)
DrawSht.Range("A13:A22").RowHeight = (MaxHeight + 350) / 10
DrawSht.Range("A23:A32").RowHeight = (MaxHeight + 350) / 10
DrawSht.Range("A33:A42").RowHeight = (MaxHeight + 350) / 10
DrawSht.Range("A43:A52").RowHeight = (MaxHeight + 350) / 10



'Draw Floorplan label
Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
With MyShape
    '.Name = eWall
    .Left = 12.5
    .Top = yf - FloorplanHeight - 100
    .Width = 200
    .Height = 75
    .Fill.ForeColor.RGB = RGB(0, 0, 0)
    .Line.ForeColor.RGB = RGB(150, 150, 150)
    .Line.Weight = 3
End With

With MyShape.TextFrame
    .Characters.Text = "Floorplan"
    .Characters.Font.Bold = True
    .Characters.Font.Size = 36
    .HorizontalAlignment = xlHAlignCenter
    .VerticalAlignment = xlVAlignCenter
End With

'draw floorplan outline
IntDimensionHeight = 0
Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
With MyShape
    .Left = xf - b.bWidth * 12
    .Top = yf - b.bLength * 12
    .Height = (b.bLength * 12)
    .Width = b.bWidth * 12
    .Line.Transparency = 1
    .Fill.ForeColor.RGB = RGB(200, 200, 200)
    .ZOrder msoSendToBack
    .Fill.Transparency = 0.5
End With
'e1 Wall
Set MyShape = DrawSht.Shapes.AddLine(xf, yf, -b.bWidth * 12 + xf, yf)
With MyShape
    .Line.ForeColor.RGB = RGB(75, 75, 75)
    .Line.Weight = 0.5
End With
    'e1 Dimension
    If b.e1Extension + b.e1Overhang > 80 And b.e1Extension + b.e1Overhang < 150 Then
        DimensionOffset = -75
    End If
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = -b.bWidth * 12 + xf
        .Top = yf + (100 + DimensionOffset)
        .Height = 100
        .Width = b.bWidth * 12
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 1
        .Line.Transparency = 1
    End With
    With MyShape.TextFrame
        .Characters.Text = "Endwall 1 " & vbNewLine & ImperialMeasurementFormat(b.bWidth * 12)
        .Characters.Font.Bold = True
        .Characters.Font.Size = 24
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Characters.Font.Color = RGB(0, 20, 132)
    End With
    'vertical dimension lines
    Set MyShape = DrawSht.Shapes.AddLine(xf, yf + (100 + DimensionOffset), xf, yf + (150 + DimensionOffset))
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
    End With
    Set MyShape = DrawSht.Shapes.AddLine(-b.bWidth * 12 + xf, yf + (100 + DimensionOffset), -b.bWidth * 12 + xf, yf + (150 + DimensionOffset))
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
    End With
    'horizontal dimension lines
    Set MyShape = DrawSht.Shapes.AddLine(xf, yf + (125 + DimensionOffset), xf - (b.bWidth * 12 / 3), yf + (125 + DimensionOffset))
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
        .BeginArrowheadLength = msoArrowheadLong
        .BeginArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadWidth = msoArrowheadWide
    End With
    Set MyShape = DrawSht.Shapes.AddLine(-b.bWidth * 12 + xf, yf + (125 + DimensionOffset), -b.bWidth * 12 + xf + (b.bWidth * 12 / 3), yf + (125 + DimensionOffset))
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
        .BeginArrowheadLength = msoArrowheadLong
        .BeginArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadWidth = msoArrowheadWide
    End With
'e3 Wall
Set MyShape = DrawSht.Shapes.AddLine(xf, yf - b.bLength * 12, -b.bWidth * 12 + xf, yf - b.bLength * 12)
With MyShape.Line
    .ForeColor.RGB = RGB(75, 75, 75)
    .Weight = 0.5
End With
's4 Wall
Set MyShape = DrawSht.Shapes.AddLine(xf, yf, xf, yf - b.bLength * 12)
With MyShape.Line
    .ForeColor.RGB = RGB(75, 75, 75)
    .Weight = 0.5
End With
    's2 Dimension
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = -b.bWidth * 12 + xf - 300
        .Top = yf - b.bLength * 12
        .Height = b.bLength * 12
        .Width = 190
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 1
        .Line.Transparency = 1
    End With
    With MyShape.TextFrame
        .Characters.Text = "Sidewall 2 " & vbNewLine & ImperialMeasurementFormat(b.bLength * 12)
        .Characters.Font.Bold = True
        .Characters.Font.Size = 24
        .HorizontalAlignment = xlHAlignRight
        .VerticalAlignment = xlVAlignCenter
        .Characters.Font.Color = RGB(0, 20, 132)
    End With
    'vertical dimension lines
    Set MyShape = DrawSht.Shapes.AddLine(-b.bWidth * 12 + xf - 150, yf, -b.bWidth * 12 + xf - 150, yf - (b.bLength * 12 / 3))
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
        .BeginArrowheadLength = msoArrowheadLong
        .BeginArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadWidth = msoArrowheadWide
    End With
    Set MyShape = DrawSht.Shapes.AddLine(-b.bWidth * 12 + xf - 150, -b.bLength * 12 + yf, -b.bWidth * 12 + xf - 150, -b.bLength * 12 + yf + (b.bLength * 12 / 3))
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
        .BeginArrowheadLength = msoArrowheadLong
        .BeginArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadWidth = msoArrowheadWide
    End With
    'horizontal dimension lines
    Set MyShape = DrawSht.Shapes.AddLine(-b.bWidth * 12 + xf - 175, yf, -b.bWidth * 12 + xf - 125, yf)
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
    End With
    Set MyShape = DrawSht.Shapes.AddLine(-b.bWidth * 12 + xf - 175, -b.bLength * 12 + yf, -b.bWidth * 12 + xf - 125, -b.bLength * 12 + yf)
    With MyShape.Line
        .ForeColor.RGB = RGB(0, 20, 132)
        .Weight = 0.5
        .DashStyle = msoLineDash
    End With
's2 Wall
Set MyShape = DrawSht.Shapes.AddLine(-b.bWidth * 12 + xf, yf, -b.bWidth * 12 + xf, yf - b.bLength * 12)
With MyShape.Line
    .ForeColor.RGB = RGB(75, 75, 75)
    .Weight = 0.5
End With

'get Extension start/end points
For Each Member In b.e1Columns
    If Member.CL < 0 Then
        s4ExtensionEdge = -Member.rEdgePosition
    ElseIf Member.CL > b.bWidth * 12 Then
        s2ExtensionEdge = (Member.lEdgePosition - b.bWidth * 12)
    End If
Next Member

'Interior Columns
If b.InteriorColumns.Count > 0 Then
    Dim Bay1Start As Double
    Set IntColumnCollection = b.InteriorColumns
    BayNum = EstSht.Range("BayNum").Value - 1
    BayStart = EstSht.Range("Bay1_Length").Value
    Bay1Start = BayStart
    Dim DimensionsArr() As Variant
    ReDim DimensionsArr(b.InteriorColumns.Count - 1)
    For i = IntColumnCollection.Count To 1 Step -1
        Set Member = IntColumnCollection(i)
        For j = 1 To BayNum
            If j = 1 Then
                BayStart = EstSht.Range("Bay1_Length").Value
            Else
                BayStart = BayStart + EstSht.Range("Bay1_Length").offset(j - 1, 0).Value
            End If
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = (xf - (Member.lEdgePosition))
                If Member.Size Like "*TS*" Then
                    .Top = yf - (BayStart * 12 + 2)
                    .Height = 4
                Else
                    .Top = yf - (BayStart * 12 + 4)
                    .Height = 8
                End If
                .Width = Member.Width
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Line.ForeColor.RGB = RGB(0, 0, 0)
                .Line.Weight = 1
            End With
            If Member.CL < 0 Or Member.CL > b.bWidth * 12 Then
                MyShape.Fill.ForeColor.RGB = RGB(0, 230, 0)
                MyShape.Line.ForeColor.RGB = RGB(0, 230, 0)
            End If
            'Draw Weld Plate
            For Each WeldPlate In Member.ComponentMembers
                If WeldPlate.clsType = "Weld Plate" Then
                    Set Plate = WeldPlate
                    Exit For
                End If
            Next WeldPlate
            'label Weld Plate
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = (xf - (Member.rEdgePosition + Member.Width))
                .Top = yf - (BayStart * 12) + 15
                .Height = 25
                .Width = 75
                .Fill.Transparency = 1
                .Line.Transparency = 1
            End With
            With MyShape.TextFrame
                .Characters.Text = Plate.Width & """x" & Plate.Height & """"
                .Characters.Font.Bold = True
                .Characters.Font.Size = 14
                .HorizontalAlignment = xlHAlignRight
                .VerticalAlignment = xlVAlignCenter
                .Characters.Font.Color = RGB(0, 20, 132)
            End With
            If j = 1 Then
                If b.InteriorColumns(i).CL < 15 And b.InteriorColumns(i).CL > 0 Then
                    'ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
                    DimensionsArr(i - 1) = Member.rEdgePosition
                ElseIf b.InteriorColumns(i).CL > b.bWidth * 12 - 15 And b.InteriorColumns(i).CL < b.bWidth * 12 Then
                    'ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
                    DimensionsArr(i - 1) = Member.lEdgePosition
                ElseIf b.InteriorColumns(i).CL < 0 Then
                    'DimensionsArr(i - 1) = -Member.rEdgePosition + b.bWidth * 12
                    s4ExtensionEdge = -Member.rEdgePosition
                ElseIf b.InteriorColumns(i).CL > b.bWidth * 12 Then
                    'DimensionsArr(i - 1) = -(Member.lEdgePosition - b.bWidth * 12)
                    s2ExtensionEdge = (Member.lEdgePosition - b.bWidth * 12)
                Else
                    'ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
                    DimensionsArr(i - 1) = b.bWidth * 12 - b.InteriorColumns(i).CL
                End If
                
            End If
        Next j
    Next i
    'Interior column dimension
    DimensionsArr = ArrayRemoveDups(DimensionsArr)
    Call DrawDimension(b, xf - b.bWidth * 12, yf - (Bay1Start * 12 + 50), b.bWidth * 12, 50, "Horizontal", 18, DimensionsArr)
End If
         
'e1 column dimensions
ReDim DimensionsArr(b.e1Columns.Count - 1)
For i = b.e1Columns.Count To 1 Step -1
    If b.e1Columns(i).CL < 15 And b.e1Columns(i).CL > 0 Then
        DimensionsArr(i - 1) = 0
    ElseIf b.e1Columns(i).CL > b.bWidth * 12 - 15 And b.e1Columns(i).CL < b.bWidth * 12 Then
        DimensionsArr(i - 1) = b.bWidth * 12
    ElseIf b.e1Columns(i).CL < 0 Then
        DimensionsArr(i - 1) = -b.e1Columns(i).rEdgePosition + b.bWidth * 12
    ElseIf b.e1Columns(i).CL > b.bWidth * 12 Then
        DimensionsArr(i - 1) = -(b.e1Columns(i).lEdgePosition - b.bWidth * 12)
    ElseIf b.e1Columns(i).mType Like "*Extension*" Then
        'do not add
    Else
        DimensionsArr(i - 1) = b.bWidth * 12 - b.e1Columns(i).CL
    End If
Next i

For i = b.e1FOs.Count To 1 Step -1
    If b.e1FOs(i).FOType = "OHDoor" Then
        ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 2)
        DimensionsArr(UBound(DimensionsArr) - 1) = b.bWidth * 12 - (b.e1FOs(i).rEdgePosition + b.e1FOs(i).Width)
        DimensionsArr(UBound(DimensionsArr)) = b.bWidth * 12 - (b.e1FOs(i).rEdgePosition)
    End If
Next i

'if s2 extension, add extension width to all CLs
If b.s2Extension > 0 Then
    For i = 0 To UBound(DimensionsArr)
        DimensionsArr(i) = DimensionsArr(i) + s2ExtensionEdge
    Next i
End If

DimensionsArr = ArrayRemoveDups(DimensionsArr)
Call DrawDimension(b, xf - b.bWidth * 12 - s2ExtensionEdge, yf + b.e1Extension + b.e1Overhang + 25, b.bWidth * 12 + s2ExtensionEdge + s4ExtensionEdge, 50, "Horizontal", 18, DimensionsArr)

'e3 column dimensions
ReDim DimensionsArr(b.e3Columns.Count - 1)
For i = b.e3Columns.Count To 1 Step -1
    If b.e3Columns(i).CL < 15 And b.e3Columns(i).CL > 0 Then
        DimensionsArr(i - 1) = 0
    ElseIf b.e3Columns(i).CL > b.bWidth * 12 - 15 And b.e3Columns(i).CL < b.bWidth * 12 Then
        DimensionsArr(i - 1) = b.bWidth * 12
    ElseIf b.e3Columns(i).CL < 0 Then
        DimensionsArr(i - 1) = b.e3Columns(i).rEdgePosition
    ElseIf b.e3Columns(i).CL > b.bWidth * 12 Then
        DimensionsArr(i - 1) = (b.e3Columns(i).lEdgePosition)
    ElseIf b.e3Columns(i).mType Like "*Extension*" Then
        'do not add
    Else
        DimensionsArr(i - 1) = b.e3Columns(i).CL
    End If
Next i

For i = b.e3FOs.Count To 1 Step -1
    If b.e3FOs(i).FOType = "OHDoor" Then
        ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 2)
        DimensionsArr(UBound(DimensionsArr) - 1) = (b.e3FOs(i).rEdgePosition + b.e3FOs(i).Width)
        DimensionsArr(UBound(DimensionsArr)) = (b.e3FOs(i).rEdgePosition)
    End If
Next i

'if s2 extension, add extension width to all CLs
If b.s2ExtensionWidth > 0 Then
    For i = 0 To UBound(DimensionsArr)
        DimensionsArr(i) = DimensionsArr(i) + s2ExtensionEdge
    Next i
End If

DimensionsArr = ArrayRemoveDups(DimensionsArr)
Call DrawDimension(b, xf - b.bWidth * 12 - s2ExtensionEdge, yf - (b.bLength * 12) - b.e3Extension - b.e3Overhang - 75, b.bWidth * 12 + s2ExtensionEdge + s4ExtensionEdge, 50, "Horizontal", 18, DimensionsArr)

's4 column dimensions
Dim DimIndex As Integer
Dim BayTotal As Double
i = 1
DimIndex = EstSht.Range("BayNum").Value
ReDim DimensionsArr(DimIndex)

If b.e1Extension > 0 Then
    ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
    DimensionsArr(UBound(DimensionsArr)) = b.bLength * 12 + b.e1Extension
End If
For i = i To EstSht.Range("BayNum").Value
    BayTotal = BayTotal + EstSht.Range("Bay1_Length").offset(i - 1, 0).Value * 12
    DimensionsArr(i - 1) = b.bLength * 12 - BayTotal
Next i
DimensionsArr(EstSht.Range("BayNum").Value) = b.bLength * 12
For Each FO In b.s4FOs
    If FO.FOType = "OHDoor" Then
        ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 2)
        DimensionsArr(UBound(DimensionsArr) - 1) = (FO.rEdgePosition + FO.Width)
        DimensionsArr(UBound(DimensionsArr)) = (FO.rEdgePosition)
    End If
Next FO
If b.e3Extension > 0 Then
    ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
    DimensionsArr(UBound(DimensionsArr)) = -b.e3Extension
    For i = 0 To UBound(DimensionsArr)
        DimensionsArr(i) = DimensionsArr(i) + b.e3Extension
    Next i
End If
Call DrawDimension(b, xf + b.s4Extension + b.s4Overhang + 75, yf - b.e3Extension - (b.bLength * 12), 50, b.bLength * 12 + b.e1Extension + b.e3Extension, "Vertical", 18, DimensionsArr)

's2 column dimensions
BayTotal = 0
i = 1
DimIndex = EstSht.Range("BayNum").Value
ReDim DimensionsArr(DimIndex)

If b.e1Extension > 0 Then
    ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
    DimensionsArr(UBound(DimensionsArr)) = b.bLength * 12 + b.e1Extension
End If
For i = 1 To EstSht.Range("BayNum").Value
    BayTotal = BayTotal + EstSht.Range("Bay1_Length").offset(i - 1, 0).Value * 12
    DimensionsArr(i - 1) = b.bLength * 12 - BayTotal
Next i
DimensionsArr(EstSht.Range("BayNum").Value) = b.bLength * 12
For Each FO In b.s2FOs
    If FO.FOType = "OHDoor" Then
        ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 2)
        DimensionsArr(UBound(DimensionsArr) - 1) = (b.bLength * 12 - FO.lEdgePosition + FO.Width)
        DimensionsArr(UBound(DimensionsArr)) = (b.bLength * 12 - FO.lEdgePosition)
    End If
Next FO
If b.e3Extension > 0 Then
    ReDim Preserve DimensionsArr(UBound(DimensionsArr) + 1)
    DimensionsArr(UBound(DimensionsArr)) = -b.e3Extension
    For i = 0 To UBound(DimensionsArr)
        DimensionsArr(i) = DimensionsArr(i) + b.e3Extension
    Next i
End If
'Call DrawDimension(b, xf - b.bWidth * 12 - b.s2Extension - b.s2Overhang - 75, yf - b.e3Extension - (b.bLength * 12), 50, b.bLength * 12 + b.e1Extension + b.e3Extension, "Vertical", 18, DimensionsArr)

'''''''''''''''''''''Extension and Overhang Shaded Areas
'e1 Extension
If b.e1Extension > 0 Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = -b.bWidth * 12 + xf
        .Top = yf
        .Height = b.e1Extension
        .Width = b.bWidth * 12
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightVertical
        .ZOrder msoSendToBack
    End With
End If
'e1 Overhang
If EstSht.Range("e1_GableOverhang").Value > 0 Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = -b.bWidth * 12 + xf
        .Top = yf + b.e1Extension
        .Height = EstSht.Range("e1_GableOverhang").Value * 12
        .Width = b.bWidth * 12
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightUpwardDiagonal
        .ZOrder msoSendToBack
    End With
End If
'e3 Extension
If b.e3Extension > 0 Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = -b.bWidth * 12 + xf
        .Top = yf - (b.bLength * 12) - b.e3Extension
        .Height = b.e3Extension
        .Width = b.bWidth * 12
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightVertical
        .ZOrder msoSendToBack
    End With
End If
'e3 Overhang
If EstSht.Range("e3_GableOverhang").Value > 0 Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = -b.bWidth * 12 + xf
        .Top = yf - (b.bLength * 12) - b.e3Extension - b.e3Overhang
        .Height = EstSht.Range("e3_GableOverhang").Value * 12
        .Width = b.bWidth * 12
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightUpwardDiagonal
        .ZOrder msoSendToBack
    End With
End If
's2 Extension
If b.s2ExtensionWidth > 0 Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth
        .Top = yf - (b.bLength * 12)
        .Height = b.bLength * 12
        .Width = b.s2ExtensionWidth
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightHorizontal
        .ZOrder msoSendToBack
    End With
End If
's2 Overhang
If EstSht.Range("s2_EaveOverhang").Value > 0 Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth - (EstSht.Range("s2_EaveOverhang").Value * 12)
        .Top = yf - (b.bLength * 12)
        .Height = b.bLength * 12
        .Width = EstSht.Range("s2_EaveOverhang").Value * 12
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightUpwardDiagonal
        .ZOrder msoSendToBack
    End With
End If
's4 Extension
If b.s4ExtensionWidth > 0 Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = xf
        .Top = yf - (b.bLength * 12)
        .Height = b.bLength * 12
        .Width = b.s4ExtensionWidth
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightHorizontal
        .ZOrder msoSendToBack
    End With
End If
's4 Overhang
If EstSht.Range("s4_EaveOverhang").Value > 0 Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = xf + b.s4ExtensionWidth
        .Top = yf - (b.bLength * 12)
        .Height = b.bLength * 12
        .Width = EstSht.Range("s4_EaveOverhang").Value * 12
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightUpwardDiagonal
        .ZOrder msoSendToBack
    End With
End If
'e1s2 Extension Intersection
If b.s2e1ExtensionIntersection = True Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth
        .Top = yf
        .Height = b.e1Extension
        .Width = b.s2ExtensionWidth
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightHorizontal
        .ZOrder msoSendToBack
    End With
    If EstSht.Range("e1_GableOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth
            .Top = yf + b.e1Extension
            .Height = EstSht.Range("e1_GableOverhang").Value * 12
            .Width = b.s2ExtensionWidth
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
    If EstSht.Range("s2_EaveOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth - EstSht.Range("s2_EaveOverhang").Value * 12
            .Top = yf
            .Height = b.e1Extension
            .Width = EstSht.Range("s2_EaveOverhang").Value * 12
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
    If EstSht.Range("s2_EaveOverhang").Value > 0 And EstSht.Range("e1_GableOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth - EstSht.Range("s2_EaveOverhang").Value * 12
            .Top = yf + b.e1Extension
            .Height = EstSht.Range("e1_GableOverhang").Value * 12
            .Width = EstSht.Range("s2_EaveOverhang").Value * 12
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
End If
'e1s4 Extension Intersection
If b.s4e1ExtensionIntersection = True Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = xf
        .Top = yf
        .Height = b.e1Extension
        .Width = b.s4ExtensionWidth
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightHorizontal
        .ZOrder msoSendToBack
    End With
    If EstSht.Range("e1_GableOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf
            .Top = yf + b.e1Extension
            .Height = EstSht.Range("e1_GableOverhang").Value * 12
            .Width = b.s4ExtensionWidth
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
    If EstSht.Range("s4_EaveOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf + b.s4ExtensionWidth
            .Top = yf
            .Height = b.e1Extension
            .Width = EstSht.Range("s4_EaveOverhang").Value * 12
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
    If EstSht.Range("s4_EaveOverhang").Value > 0 And EstSht.Range("e1_GableOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf + b.s4ExtensionWidth
            .Top = yf + b.e1Extension
            .Height = EstSht.Range("e1_GableOverhang").Value * 12
            .Width = EstSht.Range("s4_EaveOverhang").Value * 12
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
End If
'e3s2 Extension Intersection
If b.s2e3ExtensionIntersection = True Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth
        .Top = yf - b.bLength * 12 - b.e3Extension
        .Height = b.e3Extension
        .Width = b.s2ExtensionWidth
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightHorizontal
        .ZOrder msoSendToBack
    End With
    If EstSht.Range("e3_GableOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth
            .Top = yf - b.bLength * 12 - b.e3Extension - EstSht.Range("e3_GableOverhang").Value * 12
            .Height = EstSht.Range("e3_GableOverhang").Value * 12
            .Width = b.s2ExtensionWidth
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
    If EstSht.Range("s2_EaveOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth - EstSht.Range("s2_EaveOverhang").Value * 12
            .Top = yf - b.bLength * 12 - b.e3Extension
            .Height = b.e3Extension
            .Width = EstSht.Range("s2_EaveOverhang").Value * 12
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
    If EstSht.Range("s2_EaveOverhang").Value > 0 And EstSht.Range("e3_GableOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf - b.bWidth * 12 - b.s2ExtensionWidth - EstSht.Range("s2_EaveOverhang").Value * 12
            .Top = yf - (b.bLength * 12) - b.e3Extension - EstSht.Range("e3_GableOverhang").Value * 12
            .Height = EstSht.Range("e1_GableOverhang").Value * 12
            .Width = EstSht.Range("s4_EaveOverhang").Value * 12
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
End If
'e3s4 Extension Intersection
If b.s4e3ExtensionIntersection = True Then
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        .Left = xf
        .Top = yf - b.bLength * 12 - b.e3Extension
        .Height = b.e3Extension
        .Width = b.s4ExtensionWidth
        .Fill.ForeColor.RGB = RGB(0, 230, 0)
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Weight = 0.5
        .Fill.Transparency = 0.75
        .Line.Transparency = 1
        .Fill.Patterned msoPatternLightHorizontal
        .ZOrder msoSendToBack
    End With
    If EstSht.Range("e3_GableOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf
            .Top = yf - b.bLength * 12 - b.e3Extension - EstSht.Range("e3_GableOverhang").Value * 12
            .Height = EstSht.Range("e3_GableOverhang").Value * 12
            .Width = b.s4ExtensionWidth
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
    If EstSht.Range("s4_EaveOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf + b.s4ExtensionWidth
            .Top = yf - b.bLength * 12 - b.e3Extension
            .Height = b.e3Extension
            .Width = EstSht.Range("s4_EaveOverhang").Value * 12
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
    If EstSht.Range("s2_EaveOverhang").Value > 0 And EstSht.Range("e3_GableOverhang").Value > 0 Then
        Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
        With MyShape
            .Left = xf + b.s4Extension
            .Top = yf - b.bLength * 12 - b.e3Extension - EstSht.Range("e3_GableOverhang").Value * 12
            .Height = EstSht.Range("e1_GableOverhang").Value * 12
            .Width = EstSht.Range("s4_EaveOverhang").Value * 12
            .Fill.ForeColor.RGB = RGB(0, 230, 0)
            .Line.ForeColor.RGB = RGB(255, 255, 255)
            .Line.Weight = 0.5
            .Fill.Transparency = 0.75
            .Line.Transparency = 1
            .Fill.Patterned msoPatternLightUpwardDiagonal
            .ZOrder msoSendToBack
        End With
    End If
End If


For i = 1 To 4
    If i = 1 Then
        eWall = "e1"
    ElseIf i = 2 Then
        eWall = "s2"
    ElseIf i = 3 Then
        eWall = "e3"
    ElseIf i = 4 Then
        eWall = "s4"
    End If

    Select Case eWall
    Case "e1"
        Set ColumnCollection = b.e1Columns
        Set FOCollection = b.e1FOs
        Set GirtsCollection = b.e1Girts
        Set RafterCollection = b.e1Rafters
        Set OverhangCollection = b.e1OverhangMembers
        Set ExtensionCollection = b.e1ExtensionMembers
        x1 = (b.bWidth * 12 + 350) + b.s2Extension
        y1 = (350 + yf + b.e1Extension) + (MaxHeight + 350)
        'ThisWorkbook.Sheets(Sheets.Count).Range("A13:A22").RowHeight = (MaxHeight + 350) / 10
        
    Case "s2"
        Set ColumnCollection = b.s2Columns
        Set FOCollection = b.s2FOs
        Set GirtsCollection = b.s2Girts
        Set OverhangCollection = b.s2OverhangMembers
        Set ExtensionCollection = b.s2ExtensionMembers
        x1 = (b.bLength * 12 + 350) + b.s2Extension
        y1 = y1 + (MaxHeight + 350)
        'ThisWorkbook.Sheets(Sheets.Count).Range("A23:A32").RowHeight = (MaxHeight + 350) / 10
        
    Case "e3"
        Set ColumnCollection = b.e3Columns
        Set FOCollection = b.e3FOs
        Set GirtsCollection = b.e3Girts
        Set RafterCollection = b.e3Rafters
        Set OverhangCollection = b.e3OverhangMembers
        Set ExtensionCollection = b.e3ExtensionMembers
        x1 = (b.bWidth * 12 + 350) + b.s2Extension
        y1 = y1 + (MaxHeight + 350)
        'ThisWorkbook.Sheets(Sheets.Count).Range("A33:A42").RowHeight = (MaxHeight + 350) / 10

    Case "s4"
        Set ColumnCollection = b.s4Columns
        Set FOCollection = b.s4FOs
        Set GirtsCollection = b.s4Girts
        Set OverhangCollection = b.s4OverhangMembers
        Set ExtensionCollection = b.s4ExtensionMembers
        x1 = (b.bLength * 12 + 350) + b.s2Extension
        y1 = y1 + (MaxHeight + 350)
        'ThisWorkbook.Sheets(Sheets.Count).Range("A43:A53").RowHeight = (MaxHeight + 350) / 10

    End Select
    
    'get highest point of wall
    If eWall = "s2" Or (eWall = "s4" And b.rShape = "Gable") Then
        TotalHeight = b.bHeight * 12
    ElseIf eWall = "s4" And b.rShape = "Single Slope" Then
        TotalHeight = b.bHeight * 12 + b.bWidth * b.rPitch
    Else
        If b.rShape = "Single Slope" Then   'in
            TotalHeight = (b.bHeight * 12) + (b.bWidth * b.rPitch)
        Else
            TotalHeight = (b.bHeight * 12) + (b.bWidth / 2 * b.rPitch)
        End If
    End If
    
    
    Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
    With MyShape
        '.Name = eWall
        .Left = 12.5
        .Top = y1 - MaxHeight
        .Width = 75
        .Height = 75
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
        .Line.Weight = 3
    End With
    
    With MyShape.TextFrame
        .Characters.Text = eWall
        .Characters.Font.Bold = True
        .Characters.Font.Size = 36
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    
    For Each Member In ColumnCollection
    If eWall = "e1" Or eWall = "e3" Then
        Set MyShape = DrawSht.Shapes.AddLine(-Member.CL + x1, -Member.bEdgeHeight + y1, -Member.CL + x1, -Member.tEdgeHeight + y1)
        With MyShape.Line
            If Member.Placement Like "*Extension*" Or Member.Placement Like "*Overhang*" Then
                .ForeColor.RGB = RGB(0, 230, 0)
                .Transparency = 0.4
            Else
                .ForeColor.RGB = RGB(75, 75, 75)
            End If
            .Weight = Member.Width
        End With
        MyShape.Select
        If Member.Length <> 0 Then
            'Selection.Name = Member.Placement
            MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
        End If
        Dim ExtensionLength As Double
        'Floorplan Columns
        If eWall = "e1" Then
            If Member.mType = "e1 Extension Column" Then
                ExtensionLength = b.e1Extension
            Else
                ExtensionLength = 0
            End If
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = (xf - (Member.lEdgePosition))
                If Member.Size Like "*TS*" Then
                    .Top = yf - 4 + ExtensionLength
                    .Height = 4
                Else
                    .Top = yf - 8 + ExtensionLength
                    .Height = 8
                End If
                .Width = Member.Width
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Line.ForeColor.RGB = RGB(0, 0, 0)
                .Line.Weight = 1
            End With
            If Member.CL < 0 Or Member.CL > b.bWidth * 12 Or Member.mType = "e1 Extension Column" Then
                MyShape.Fill.ForeColor.RGB = RGB(0, 230, 0)
                MyShape.Line.ForeColor.RGB = RGB(0, 230, 0)
            End If
            'Draw Dimension
            
            'get Weld Plate
            For Each WeldPlate In Member.ComponentMembers
                If WeldPlate.clsType = "Weld Plate" Then
                    Set Plate = WeldPlate
                    Exit For
                End If
            Next WeldPlate
            'label Weld Plate
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = (xf - (Member.rEdgePosition + Member.Width))
                .Top = yf + 15 + ExtensionLength
                .Height = 25
                .Width = 75
                .Fill.Transparency = 1
                .Line.Transparency = 1
            End With
            With MyShape.TextFrame
                .Characters.Text = Plate.Width & """x" & Plate.Height & """"
                .Characters.Font.Bold = True
                .Characters.Font.Size = 14
                .HorizontalAlignment = xlHAlignRight
                .VerticalAlignment = xlVAlignCenter
                .Characters.Font.Color = RGB(0, 20, 132)
            End With
        ElseIf eWall = "e3" Then
            If Member.mType = "e3 Extension Column" Then
                ExtensionLength = b.e3Extension
            Else
                ExtensionLength = 0
            End If
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = xf - (b.bWidth * 12 - (Member.rEdgePosition))
                .Top = yf - b.bLength * 12 - ExtensionLength
                If Member.Size Like "*TS*" Then
                    .Height = 4
                Else
                    .Height = 8
                End If
                .Width = Member.Width
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Line.ForeColor.RGB = RGB(150, 150, 150)
                .Line.Weight = 1
            End With
            If Member.CL < 0 Or Member.CL > b.bWidth * 12 Or Member.mType = "e3 Extension Column" Then
                MyShape.Fill.ForeColor.RGB = RGB(0, 230, 0)
                MyShape.Line.ForeColor.RGB = RGB(0, 230, 0)
            End If
            'get Weld Plate
            For Each WeldPlate In Member.ComponentMembers
                If WeldPlate.clsType = "Weld Plate" Then
                    Set Plate = WeldPlate
                    Exit For
                End If
            Next WeldPlate
            'label Weld Plate
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = xf - ((b.bWidth * 12) - Member.rEdgePosition)
                .Top = yf - (b.bLength * 12) + 15 - ExtensionLength
                .Height = 25
                .Width = 75
                .Fill.Transparency = 1
                .Line.Transparency = 1
            End With
            With MyShape.TextFrame
                .Characters.Text = Plate.Width & """x" & Plate.Height & """"
                .Characters.Font.Bold = True
                .Characters.Font.Size = 14
                .HorizontalAlignment = xlHAlignRight
                .VerticalAlignment = xlVAlignCenter
                .Characters.Font.Color = RGB(0, 20, 132)
            End With
        End If
        
    Else
        Set MyShape = DrawSht.Shapes.AddLine(-Member.CL + x1, -Member.bEdgeHeight + y1, -Member.CL + x1, -Member.tEdgeHeight + y1)
        With MyShape.Line
            If Member.Placement Like "*Extension*" Or Member.Placement Like "*Overhang*" Then
                .ForeColor.RGB = RGB(0, 230, 0)
                .Transparency = 0.4
            Else
                .ForeColor.RGB = RGB(75, 75, 75)
            End If
            .Weight = Member.Width
        End With
        MyShape.Select
        If Member.Length <> 0 Then
            'Selection.Name = Member.Placement
        End If
        MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
        ColumnWidth = Member.Width
        If eWall = "s2" Then
            'Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            'With MyShape
            '    .Left = xf - (b.bWidth * 12)
            '    .Top = yf - Member.CL - Member.Width
            '    .Height = Member.Width * 2
            '    .Width = Member.Width * 2
            '    .Fill.ForeColor.RGB = RGB(0, 0, 0)
            '    .Line.ForeColor.RGB = RGB(150, 150, 150)
            '    .Line.Weight = 1
            'End With
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = xf - (b.bWidth * 12)
                .Top = yf - Member.CL - Member.Width
                .Height = Member.Width * 2
                .Width = b.bWidth * 12
                .Fill.ForeColor.RGB = RGB(230, 0, 0)
                .Line.ForeColor.RGB = RGB(0, 0, 0)
                .Line.Weight = 0.5
                .Line.DashStyle = msoLineDash
                .Fill.Transparency = 0.4
            End With
            With MyShape.TextFrame
                .Characters.Text = "Main Rafter Line"
                .Characters.Font.Bold = True
                .Characters.Font.Size = 16
                .Characters.Font.ColorIndex = 2
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
            End With
        ElseIf eWall = "s4" Then
            'Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            'With MyShape
            '    .Left = xf - Member.Width * 2
            '    .Top = yf - b.bLength * 12 + Member.CL - Member.Width
            '    .Height = Member.Width * 2
            '    .Width = Member.Width * 2
            '    .Fill.ForeColor.RGB = RGB(0, 0, 0)
            '    .Line.ForeColor.RGB = RGB(150, 150, 150)
            '    .Line.Weight = 1
            'End With
        End If
    End If
    Next Member
    'endwall columns on s2 and s4
    If eWall = "s2" Then
        With DrawSht.Shapes.AddLine(-b.bLength * 12 + x1, -0 + y1, -b.bLength * 12 + x1, -TotalHeight + y1).Line
            .ForeColor.RGB = RGB(75, 75, 75)
            .Weight = b.e3Columns(1).Width
        End With
        With DrawSht.Shapes.AddLine(-0 * 12 + x1, -0 + y1, -0 + x1, -TotalHeight + y1).Line
            .ForeColor.RGB = RGB(75, 75, 75)
            .Weight = b.e1Columns(1).Width
        End With
    ElseIf eWall = "s4" Then
        With DrawSht.Shapes.AddLine(-b.bLength * 12 + x1, -0 + y1, -b.bLength * 12 + x1, -TotalHeight + y1).Line
            .ForeColor.RGB = RGB(75, 75, 75)
            .Weight = b.e1Columns(1).Width
        End With
        With DrawSht.Shapes.AddLine(-0 * 12 + x1, -0 + y1, -0 + x1, -TotalHeight + y1).Line
            .ForeColor.RGB = RGB(75, 75, 75)
            .Weight = b.e3Columns(1).Width
        End With
    End If
    
    For Each Member In GirtsCollection
        Set MyShape = DrawSht.Shapes.AddLine(-Member.rEdgePosition + x1, -Member.bEdgeHeight + y1, -Member.rEdgePosition - Member.Length + x1, -Member.tEdgeHeight + y1)
        With MyShape.Line
            If eWall = "e1" Or eWall = "e3" Then
                If Member.rEdgePosition < 0 Or Member.rEdgePosition + Member.Length > b.bWidth * 12 Then
                    .ForeColor.RGB = RGB(0, 230, 0)
                Else
                    .ForeColor.RGB = RGB(150, 150, 150)
                End If
            Else
                If Member.rEdgePosition < 0 Or Member.rEdgePosition + Member.Length > b.bLength * 12 Then
                    .ForeColor.RGB = RGB(0, 230, 0)
                Else
                    .ForeColor.RGB = RGB(150, 150, 150)
                End If
            End If
            .Weight = 2.5
        End With
        If Member.bEdgeHeight = 86 Then
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = -Member.rEdgePosition - Member.Length + x1
                .Top = -Member.bEdgeHeight + y1 - 50
                .Height = 50
                .Width = Member.Length
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Line.ForeColor.RGB = RGB(255, 255, 255)
                .Line.Weight = 1
                .Fill.Transparency = 1
                .ZOrder msoSendToBack
            End With
            With MyShape.TextFrame
                .Characters.Text = ImperialMeasurementFormat(Member.Length)
                .Characters.Font.Bold = True
                .Characters.Font.Size = 24
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignBottom
                .Characters.Font.Color = RGB(0, 20, 132)
            End With
        Else
            MyShape.Select
            Length = ImperialMeasurementFormat(Member.Length)
            'If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
            MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
        End If
    Next Member
    
    For Each FO In FOCollection
        With DrawSht.Shapes.AddLine(-FO.rEdgePosition + x1, -FO.bEdgeHeight + y1, -FO.rEdgePosition - FO.Width + x1, -FO.bEdgeHeight + y1).Line
            .ForeColor.RGB = RGB(240, 240, 0)
            .Weight = 2.5
        End With
            With DrawSht.Shapes.AddLine(-FO.rEdgePosition - FO.Width + x1, -FO.bEdgeHeight + y1, -FO.rEdgePosition - FO.Width + x1, -FO.tEdgeHeight + y1).Line
            .ForeColor.RGB = RGB(240, 240, 0)
            .Weight = 2.5
        End With
            With DrawSht.Shapes.AddLine(-FO.rEdgePosition - FO.Width + x1, -FO.tEdgeHeight + y1, -FO.rEdgePosition + x1, -FO.tEdgeHeight + y1).Line
            .ForeColor.RGB = RGB(240, 240, 0)
            .Weight = 2.5
        End With
            With DrawSht.Shapes.AddLine(-FO.rEdgePosition + x1, -FO.bEdgeHeight + y1, -FO.rEdgePosition + x1, -FO.tEdgeHeight + y1).Line
            .ForeColor.RGB = RGB(240, 240, 0)
            .Weight = 2.5
        End With
        
        'Draw Dimension (width x height) of FO
            Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
            With MyShape
                .Left = -FO.rEdgePosition - FO.Width + x1
                .Top = -FO.tEdgeHeight + y1
                .Height = FO.Height
                .Width = FO.Width
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Line.ForeColor.RGB = RGB(255, 255, 255)
                .Line.Weight = 1
                .Fill.Transparency = 0#
                .ZOrder msoSendToBack
            End With
            With MyShape.TextFrame
                .Characters.Text = "W" & ImperialMeasurementFormat(FO.Width) & " x H" & ImperialMeasurementFormat(FO.Height)
                .Characters.Font.Bold = True
                .Characters.Font.Size = 18
                .HorizontalAlignment = xlHAlignCenter
                .VerticalAlignment = xlVAlignCenter
                .Characters.Font.Color = RGB(0, 20, 132)
            End With
        
        'Floorplan View
        If eWall = "e1" Then
            Set MyShape = DrawSht.Shapes.AddLine(xf - FO.rEdgePosition, yf, xf - FO.lEdgePosition, yf)
            With MyShape.Line
                .ForeColor.RGB = RGB(240, 240, 0)
                .Weight = 5
            End With
            If FO.FOType = "OHDoor" Then
                Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
                With MyShape
                    .Top = yf - FO.Height
                    .Left = xf - FO.lEdgePosition
                    .Height = FO.Height
                    .Width = FO.Width
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .Fill.Transparency = 1
                    .Line.ForeColor.RGB = RGB(240, 240, 0)
                    .Line.Weight = 1
                    .Line.DashStyle = msoLineDash
                    .ZOrder (msoSendToBack)
                End With
                With MyShape.TextFrame
                    .Characters.Text = FO.FOType & vbNewLine & ImperialMeasurementFormat(FO.Width) & "x" & ImperialMeasurementFormat(FO.Height)
                    .Characters.Font.Bold = True
                    .Characters.Font.Size = 16
                    .Characters.Font.ColorIndex = 1
                    .HorizontalAlignment = xlHAlignCenter
                    .VerticalAlignment = xlVAlignTop
                End With
                'Draw Dimension
                
            End If
        ElseIf eWall = "e3" Then
            Set MyShape = DrawSht.Shapes.AddLine(xf - (b.bWidth * 12 - FO.rEdgePosition), yf - b.bLength * 12, xf - (b.bWidth * 12 - FO.lEdgePosition), yf - b.bLength * 12)
            With MyShape.Line
                .ForeColor.RGB = RGB(240, 240, 0)
                .Weight = 5
            End With
            If FO.FOType = "OHDoor" Then
                Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
                With MyShape
                    .Top = yf - b.bLength * 12
                    .Left = xf - (b.bWidth * 12 - FO.rEdgePosition)
                    .Height = FO.Height
                    .Width = FO.Width
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .Fill.Transparency = 1
                    .Line.ForeColor.RGB = RGB(240, 240, 0)
                    .Line.Weight = 1
                    .Line.DashStyle = msoLineDash
                    .ZOrder (msoSendToBack)
                End With
                With MyShape.TextFrame
                    .Characters.Text = FO.FOType & vbNewLine & ImperialMeasurementFormat(FO.Width) & "x" & ImperialMeasurementFormat(FO.Height)
                    .Characters.Font.Bold = True
                    .Characters.Font.Size = 16
                    .Characters.Font.ColorIndex = 1
                    .HorizontalAlignment = xlHAlignCenter
                    .VerticalAlignment = xlVAlignBottom
                End With
            End If
        ElseIf eWall = "s2" Then
            Set MyShape = DrawSht.Shapes.AddLine(xf - b.bWidth * 12, yf - FO.rEdgePosition, xf - b.bWidth * 12, yf - FO.lEdgePosition)
            With MyShape.Line
                .ForeColor.RGB = RGB(240, 240, 0)
                .Weight = 5
            End With
            If FO.FOType = "OHDoor" Then
                Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
                With MyShape
                    .Top = yf - FO.lEdgePosition
                    .Left = -b.bWidth * 12 + xf
                    .Height = FO.Width
                    .Width = FO.Height
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .Fill.Transparency = 1
                    .Line.ForeColor.RGB = RGB(240, 240, 0)
                    .Line.Weight = 1
                    .Line.DashStyle = msoLineDash
                    .ZOrder (msoSendToBack)
                End With
                With MyShape.TextFrame
                    .Characters.Text = FO.FOType & vbNewLine & ImperialMeasurementFormat(FO.Width) & "x" & ImperialMeasurementFormat(FO.Height)
                    .Characters.Font.Bold = True
                    .Characters.Font.Size = 16
                    .Characters.Font.ColorIndex = 1
                    .HorizontalAlignment = xlHAlignRight
                    .VerticalAlignment = xlVAlignCenter
                End With
            End If
        ElseIf eWall = "s4" Then
            Set MyShape = DrawSht.Shapes.AddLine(xf, yf - (b.bLength * 12) + FO.rEdgePosition, xf, yf - (b.bLength * 12) + FO.lEdgePosition)
            With MyShape.Line
                .ForeColor.RGB = RGB(240, 240, 0)
                .Weight = 5
            End With
            If FO.FOType = "OHDoor" Then
                Set MyShape = DrawSht.Shapes.AddShape(msoShapeRectangle, 1, 1, 1, 1)
                With MyShape
                    .Top = yf - (b.bLength * 12) + FO.rEdgePosition
                    .Left = xf - (FO.Height)
                    .Height = FO.Width
                    .Width = FO.Height
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .Fill.Transparency = 1
                    .Line.ForeColor.RGB = RGB(240, 240, 0)
                    .Line.Weight = 1
                    .Line.DashStyle = msoLineDash
                    .ZOrder (msoSendToBack)
                End With
                With MyShape.TextFrame
                    .Characters.Text = FO.FOType & vbNewLine & ImperialMeasurementFormat(FO.Width) & "x" & ImperialMeasurementFormat(FO.Height)
                    .Characters.Font.Bold = True
                    .Characters.Font.Size = 16
                    .Characters.Font.ColorIndex = 1
                    .HorizontalAlignment = xlHAlignLeft
                    .VerticalAlignment = xlVAlignCenter
                End With
            End If
            
        End If
        
        For Each item In FO.FOMaterials
            If item.clsType = "Member" Then
                Set Member = item
                If Member.CL <> 0 Then
                    Set MyShape = DrawSht.Shapes.AddLine(-Member.CL + x1, -Member.bEdgeHeight + y1, -Member.CL + x1, -Member.tEdgeHeight + y1)
                    With MyShape.Line
                        .ForeColor.RGB = RGB(55, 86, 35)
                        .Weight = 2.5
                    End With
                    MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
                End If
            End If
        Next item
    Next FO
    
    If eWall = "e1" Or eWall = "e3" Then
        For Each Member In RafterCollection
            If (Member.rEdgePosition < b.bWidth * 12 / 2 And b.rShape = "Gable") Then
            'b = c * cos(a)
            'a = atn(b.rPitch/12)
                lEdgePosition = Member.rEdgePosition + ((Member.tEdgeHeight - Member.bEdgeHeight) / b.rPitch * 12)
                lEdgePosition = Member.RafterLeftEdge
                Set MyShape = DrawSht.Shapes.AddLine(-Member.rEdgePosition + x1, -Member.bEdgeHeight + y1, -lEdgePosition + x1, -Member.tEdgeHeight + y1)
                With MyShape.Line
                    If Member.Placement Like "*Extension*" Or Member.Placement Like "*Overhang*" Or Member.Placement Like "*Stub*" Then
                        .ForeColor.RGB = RGB(0, 230, 0)
                        .Transparency = 0.4
                    Else
                        .ForeColor.RGB = RGB(75, 75, 75)
                    End If
                    .Weight = Member.Width
                    '.DashStyle = msoLineDashDotDot
                End With
                MyShape.Select
                'If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
            ElseIf b.rShape = "Gable" Or (b.rShape = "Single Slope" And eWall = "e1") Then
                lEdgePosition = Member.rEdgePosition + (Abs(Member.tEdgeHeight - Member.bEdgeHeight) / b.rPitch * 12)
                lEdgePosition = Member.RafterLeftEdge
                Set MyShape = DrawSht.Shapes.AddLine(-lEdgePosition + x1, -Member.bEdgeHeight + y1, -Member.rEdgePosition + x1, -Member.tEdgeHeight + y1)
                With MyShape.Line
                    If Member.Placement Like "*Extension*" Or Member.Placement Like "*Overhang*" Or Member.Placement Like "*Stub*" Then
                        .ForeColor.RGB = RGB(0, 230, 0)
                        .Transparency = 0.4
                    Else
                        .ForeColor.RGB = RGB(75, 75, 75)
                    End If
                    .Weight = Member.Width
                    '.DashStyle = msoLineDashDotDot
                End With
                MyShape.Select
                'If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
            ElseIf b.rShape = "Single Slope" Then
                lEdgePosition = Member.rEdgePosition + ((Member.tEdgeHeight - Member.bEdgeHeight) / b.rPitch * 12)
                lEdgePosition = Member.RafterLeftEdge
                Set MyShape = DrawSht.Shapes.AddLine(-Member.rEdgePosition + x1, -Member.bEdgeHeight + y1, -lEdgePosition + x1, -Member.tEdgeHeight + y1)
                With MyShape.Line
                    If Member.Placement Like "*Extension*" Or Member.Placement Like "*Overhang*" Or Member.Placement Like "*Stub*" Then
                        .ForeColor.RGB = RGB(0, 230, 0)
                        .Transparency = 0.4
                    Else
                        .ForeColor.RGB = RGB(75, 75, 75)
                    End If
                    .Weight = Member.Width
                    '.DashStyle = msoLineDashDotDot
                End With
                MyShape.Select
                'If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
            End If
        Next Member
        
        'Interior Columns
        If b.InteriorColumns.Count > 0 Then
            For Each Member In IntColumnCollection
                    If eWall = "e1" Then
                        Set MyShape = DrawSht.Shapes.AddLine(-Member.CL + x1, -Member.bEdgeHeight + y1, -Member.CL + x1, -Member.tEdgeHeight + y1)
                        With MyShape.Line
                            If Member.Placement Like "*Extension*" Or Member.Placement Like "*Overhang*" Then
                                .ForeColor.RGB = RGB(230, 0, 0)
                                .Transparency = 0.4
                            Else
                                .ForeColor.RGB = RGB(230, 0, 0)
                                .Transparency = 0.4
                            End If
                            .Weight = Member.Width
                            .DashStyle = msoLineDash
                        End With
                        MyShape.ZOrder msoSendToBack
                        MyShape.Select
                        'If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                        MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
                    Else
                       Set MyShape = DrawSht.Shapes.AddLine(-(b.bWidth * 12 - Member.CL) + x1, -Member.bEdgeHeight + y1, -(b.bWidth * 12 - Member.CL) + x1, -Member.tEdgeHeight + y1)
                       With MyShape.Line
                            If Member.Placement Like "*Extension*" Or Member.Placement Like "*Overhang*" Then
                                .ForeColor.RGB = RGB(230, 0, 0)
                                .Transparency = 0.4
                            Else
                                .ForeColor.RGB = RGB(230, 0, 0)
                                .Transparency = 0.4
                            End If
                            .Weight = Member.Width
                            .DashStyle = msoLineDash
                        End With
                        MyShape.ZOrder msoSendToBack
                        MyShape.Select
                        'If Member.Placement <> "" Then 'MyShape.Name = Member.Placement
                        MyShape.OnAction = "'DisplayDrawingInfo " & Application.WorksheetFunction.Round(Member.Length, 4) & "'"
                    End If
            Next Member
        End If
    End If
Next i

Debug.Print "drawing done"

DrawSht.PageSetup.FitToPagesWide = 1
    Application.PrintCommunication = False
    With DrawSht.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.2)
        .FooterMargin = Application.InchesToPoints(0.2)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    On Error Resume Next
    Application.PrintCommunication = True
    
    DrawSht.ResetAllPageBreaks
    DrawSht.HPageBreaks.Add Before:=DrawSht.Rows(14)
    DrawSht.Rows(14).PageBreak = xlPageBreakManual
    DrawSht.HPageBreaks.Add Before:=DrawSht.Rows(24)
    DrawSht.Rows(24).PageBreak = xlPageBreakManual
    DrawSht.HPageBreaks.Add Before:=DrawSht.Rows(34)
    DrawSht.Rows(34).PageBreak = xlPageBreakManual
    DrawSht.HPageBreaks.Add Before:=DrawSht.Rows(45)
    DrawSht.Rows(34).PageBreak = xlPageBreakManual
    DrawSht.HPageBreaks.Add Before:=DrawSht.Rows(56)
    DrawSht.Rows(34).PageBreak = xlPageBreakManual

EstSht.Activate

End Sub





