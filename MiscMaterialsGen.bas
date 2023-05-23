Attribute VB_Name = "MiscMaterialsGen"
Option Explicit

Sub MiscMaterialCalc(MiscMaterials As Collection, WriteCell As Range, b As clsBuilding)
Dim FOCell As Range
Dim NewMiscMaterials As Collection
Dim MiscMaterial As clsMiscItem
Dim FOArea As Integer
Dim WallArea As Double     'sf
Dim RoofArea As Double     'sf


'initalize new materials collection
Set NewMiscMaterials = New Collection


With EstSht
    'additional PDoor Misc Materials
    For Each FOCell In Range(.Range("pDoorCell1"), .Range("pDoorCell12"))
        'check that cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'check for a canopy
            If FOCell.offset(0, 4).Value = "4' x 4'6""" Then
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Measurement = "4' x 4'6"""
            ElseIf FOCell.offset(0, 4).Value = "4' x 7'6""" Then
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Measurement = "4' x 7'6"""
            End If
            If Not MiscMaterial Is Nothing Then
                MiscMaterial.Quantity = 1
                MiscMaterial.Name = "Door Canopy"
                NewMiscMaterials.Add MiscMaterial
                Set MiscMaterial = Nothing
            End If
        End If
    Next FOCell
    
    'additional OHDoor materials
    For Each FOCell In Range(.Range("OHDoorCell1"), .Range("OHDoorCell12"))
        'check that cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'calculate FO area
            FOArea = FOCell.offset(0, 1).Value * FOCell.offset(0, 2).Value
            
            'add OH door
            If FOCell.offset(0, 4).Value = "Sectional" Then
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Sectional OH Door"
            ElseIf FOCell.offset(0, 4).Value = "RUD" Then
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Roll Up OH Door"
            End If
            If Not MiscMaterial Is Nothing Then
                MiscMaterial.Quantity = 1
                MiscMaterial.Width = FOCell.offset(0, 1).Value
                MiscMaterial.Height = FOCell.offset(0, 2).Value
                MiscMaterial.Area = MiscMaterial.Width * MiscMaterial.Height
                MiscMaterial.Measurement = FOCell.offset(0, 1).Text & " x " & FOCell.offset(0, 2).Text
                NewMiscMaterials.Add MiscMaterial
                Set MiscMaterial = Nothing
            End If
            
            ''''check for insulation
            If FOCell.offset(0, 5).Value = "Vinyl Backed" Then
                Set MiscMaterial = New clsMiscItem
                '1 SF pieces
                MiscMaterial.Quantity = FOArea
                MiscMaterial.Name = "Vinyl Backed Insulation"
            ElseIf FOCell.offset(0, 5).Value = "Steel Backed" Then
                Set MiscMaterial = New clsMiscItem
                '1 SF pieces
                MiscMaterial.Quantity = FOArea
                MiscMaterial.Name = "Steel Backed Insulation"
            End If
            If Not MiscMaterial Is Nothing Then
                MiscMaterial.Measurement = "1 ft" & ChrW(&HB2)
                NewMiscMaterials.Add MiscMaterial
                Set MiscMaterial = Nothing
            End If
            
            ''''Door Operation
            Select Case FOCell.offset(0, 6).Value
            Case "Manual"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Manual Opener"
            Case "Chain Hoist"
                'Only add for sectional since Roll-Up Door price includes chain hoist opener
                If FOCell.offset(0, 4).Value <> "RUD" Then
                    Set MiscMaterial = New clsMiscItem
                    MiscMaterial.Name = "Chain Hoist Opener"
                End If
            Case "Electric Opener"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Electric Opener - OH Door #" & FOCell.Value
            End Select
            If Not MiscMaterial Is Nothing Then
                MiscMaterial.Quantity = 1
                NewMiscMaterials.Add MiscMaterial
                Set MiscMaterial = Nothing
            End If
            
            ''''High Lift
            If FOCell.offset(0, 7).Value = "Yes" Then
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Quantity = 1
                MiscMaterial.Name = "High Lift"
                MiscMaterial.Measurement = HighLiftSize(b.bHeight - FOCell.offset(0, 2).Value, -1)
                NewMiscMaterials.Add MiscMaterial
                Set MiscMaterial = Nothing
            End If
            
            ''''OH Door Windows
            Select Case FOCell.offset(0, 8).Value
            Case "Non-Insulated"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Non-Insulated Window"
                MiscMaterial.Measurement = "4'"
                MiscMaterial.Quantity = Application.WorksheetFunction.RoundDown(FOCell.offset(0, 1).Value / 4, 0)
              Case "Insulated"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Insulated Window"
                MiscMaterial.Measurement = "4'"
                MiscMaterial.Quantity = Application.WorksheetFunction.RoundDown(FOCell.offset(0, 1).Value / 4, 0)
              Case "Full Glass Panel"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Full Glass Panel Window"
                MiscMaterial.Measurement = "1'"
                MiscMaterial.Quantity = FOCell.offset(0, 1).Value
            End Select
            If Not MiscMaterial Is Nothing Then
                NewMiscMaterials.Add MiscMaterial
                Set MiscMaterial = Nothing
            End If
            
        End If
    Next FOCell
    
    'additional Window materials
    For Each FOCell In Range(.Range("WindowCell1"), .Range("WindowCell12"))
        'check that cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'add window
            Set MiscMaterial = New clsMiscItem
            MiscMaterial.Quantity = 1
            MiscMaterial.Name = "Standard Window"
            MiscMaterial.Measurement = FOCell.offset(0, 1).Text & " x " & FOCell.offset(0, 2).Text
            'area in square feet
            MiscMaterial.Area = (FOCell.offset(0, 1).Value * FOCell.offset(0, 2).Value) / (144)
            'add to collection
            NewMiscMaterials.Add MiscMaterial
            Set MiscMaterial = Nothing
        End If
    Next FOCell
    
    'additional Misc FO materials
    For Each FOCell In Range(.Range("MiscFOCell1"), .Range("MiscFOCell12"))
        'check that cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            '''' Exhaust Fans/Louvers
            Select Case FOCell.offset(0, 4).Value
            Case "24"" Exhaust Fan"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Exhaust Fan"
                MiscMaterial.Measurement = "24"""
            Case "30"" Exhaust Fan"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Exhaust Fan"
                MiscMaterial.Measurement = "30"""
            Case "36"" Exhaust Fan"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Exhaust Fan"
                MiscMaterial.Measurement = "36"""
            Case "24"" Louver"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Louver"
                MiscMaterial.Measurement = "24"""
            Case "30"" Louver"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Louver"
                MiscMaterial.Measurement = "30"""
            Case "36"" Louver"
                Set MiscMaterial = New clsMiscItem
                MiscMaterial.Name = "Louver"
                MiscMaterial.Measurement = "36"""
            End Select
            If Not MiscMaterial Is Nothing Then
                MiscMaterial.Quantity = 1
                NewMiscMaterials.Add MiscMaterial
                Set MiscMaterial = Nothing
            End If
            
            '''' Weather Hoods
            Select Case FOCell.offset(0, 5).Value
            Case "24""", "30""", "36"""
                Set MiscMaterial = New clsMiscItem
            End Select
            If Not MiscMaterial Is Nothing Then
                MiscMaterial.Quantity = 1
                MiscMaterial.Name = "Weather Hood"
                MiscMaterial.Measurement = FOCell.offset(0, 5).Text
                NewMiscMaterials.Add MiscMaterial
                Set MiscMaterial = Nothing
            End If
          
        End If
    Next FOCell
            
    ''''''Insulation
    'Building Areas
    With b
        If .rShape = "Gable" Then
            WallArea = Application.WorksheetFunction.RoundUp((2 * .bHeight * .bLength) + (2 * .bHeight * .bWidth) + (.bWidth * ((.bWidth * .rPitch) / 12)), 0)
            RoofArea = Application.WorksheetFunction.RoundUp(.bLength * (.RafterLength / 12) * 2, 0)
        ElseIf .rShape = "Single Slope" Then
            WallArea = Application.WorksheetFunction.RoundUp((.bHeight * .bLength) + (.bHeight * (.HighSideEaveHeight / 12)) + _
            (2 * .bHeight * .bWidth) + (.bWidth * ((.bWidth * .rPitch) / 12)), 0)
            RoofArea = Application.WorksheetFunction.RoundUp(.bLength * (.RafterLength / 12), 0)
        End If
    End With
    'subtract OH door area from wall area
    For Each MiscMaterial In NewMiscMaterials
        If InStr(1, MiscMaterial.Name, "OH Door") <> 0 Then
            'subtract area
            WallArea = WallArea - (MiscMaterial.Width * MiscMaterial.Height)
        End If
    Next MiscMaterial
    'reset object
    ''' Wall Insulation '''
    Set MiscMaterial = Nothing
    Select Case .Range("WallInsulation").Value
    Case "3"" VRR"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "3"" VRR Wall Insulation"
    Case "4"" VRR"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "4"" VRR Wall Insulation"
    Case "6"" VRR"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "6"" VRR Wall Insulation"
    Case "1"" Spray Foam"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "1"" Spray Foam Wall Insulation"
    Case "2"" Spray Foam"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "2"" Spray Foam Wall Insulation"
    End Select
    If Not MiscMaterial Is Nothing Then
        MiscMaterial.Measurement = "1 ft" & ChrW(&HB2)
        MiscMaterial.Quantity = WallArea
        NewMiscMaterials.Add MiscMaterial
        Set MiscMaterial = Nothing
    End If
    ''' Roof Insulation '''
    Select Case .Range("RoofInsulation").Value
    Case "3"" VRR"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "3"" VRR Roof Insulation"
    Case "4"" VRR"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "4"" VRR Roof Insulation"
    Case "6"" VRR"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "6"" VRR Roof Insulation"
    Case "1"" Spray Foam"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "1"" Spray Foam Roof Insulation"
    Case "2"" Spray Foam"
        Set MiscMaterial = New clsMiscItem
        MiscMaterial.Name = "2"" Spray Foam Roof Insulation"
    End Select
    If Not MiscMaterial Is Nothing Then
        MiscMaterial.Measurement = "1 ft" & ChrW(&HB2)
        MiscMaterial.Quantity = RoofArea
        NewMiscMaterials.Add MiscMaterial
        Set MiscMaterial = Nothing
    End If
    
    '''' Ridge Vents
    If .Range("RidgeVentQty").Value <> 0 Then
        Select Case .Range("RidgeVentType").Value
        Case "Standard", "Low Profile"
            Set MiscMaterial = New clsMiscItem
        End Select
        If Not MiscMaterial Is Nothing Then
            MiscMaterial.Quantity = .Range("RidgeVentQty").Value
            MiscMaterial.Measurement = "10'"
            MiscMaterial.Name = .Range("RidgeVentType").Value & " Ridge Vent"
            NewMiscMaterials.Add MiscMaterial
            Set MiscMaterial = Nothing
        End If
    End If
End With

'remove duplicates
Call MaterialsListGen.DuplicateMaterialRemoval(NewMiscMaterials, "Misc")
'''' Output to employee materials list, add to master misc materials collection
For Each MiscMaterial In NewMiscMaterials
    'output to employee materials list
    WriteCell.Value = MiscMaterial.Quantity
    WriteCell.offset(0, 1).Value = MiscMaterial.Name
    WriteCell.offset(0, 3).Value = MiscMaterial.Measurement
    WriteCell.offset(0, 4).Value = MiscMaterial.Color
    Set WriteCell = WriteCell.offset(1, 0)
    'add to master collection
    MiscMaterials.Add MiscMaterial
Next MiscMaterial
    
            
End Sub

'' function returns string of the nearest available rake HighLift size
Function HighLiftSize(FtLength As Variant, Optional Direction As Integer) As Variant

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
Dim HighLifts() As Variant
Dim HighLift As Variant
Dim hSize As Variant
Dim NearestHighLiftSizeString As String
Dim Length As Variant


HighLifts = Array(36, 54, 72, 96, 120)

'Convert Ft Length to inches
Length = FtLength * 12


t = 1.79769313486231E+308 'initialize
For Each HighLift In HighLifts
    If IsNumeric(HighLift) Then
        u = Abs(HighLift - Length)
        If Direction > 0 And HighLift >= Length Then
            'only report if closer number is greater than the target
            If u < t Then
                t = u
                hSize = HighLift
            End If
        ElseIf Direction < 0 And HighLift <= Length Then
            'only report if closer number is less than the target
            If u < t Then
                t = u
                hSize = HighLift
            End If
        ElseIf Direction = 0 Then
            If u < t Then
                t = u
                hSize = HighLift
            End If
        End If
    End If
Next


'return available High Lift size
Select Case hSize
Case 36
    NearestHighLiftSizeString = "36"""
Case 54
    NearestHighLiftSizeString = "54"""
Case 72
    NearestHighLiftSizeString = "72"""
Case 96
    NearestHighLiftSizeString = "96"""
Case 120
    NearestHighLiftSizeString = "120"""
End Select

HighLiftSize = NearestHighLiftSizeString


End Function
