Attribute VB_Name = "GeneralMod"
Option Explicit

Public BTItems As clsBuildertrend
Sub EventsEnable()
Call EstSht.UpdatesEventsProtection(True)
End Sub

Function WorksheetExists(shtName As String) As Boolean
    Dim sht As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Sub ClearTemplate()
Dim mCell As Range


'On Error Resume Next


If WorksheetExists("Cost Estimate") Then
    Call DeleteSheets
End If

Call EstSht.UpdatesEventsProtection(False)

With EstSht
    .Range("CustomerName").Value = ""
    .Range("B5").Value = ""
    .Range("E4").Value = ""
    .Range("E5").Value = ""
    .Range("B9:B12").Value = ""
    .Range("E9:E11").Value = ""
    .Range("Wall_pShape").Value = ""
    .Range("Wall_pType").Value = ""
    .Range("Wall_Color").Value = ""
    .Range("Roof_pShape").Value = ""
    .Range("Roof_pType").Value = ""
    .Range("Roof_Color").Value = ""
    .Range("All_tColors").Value = ""
    .Range("FO_tColor").Value = ""
    .Range("Base_tColor").Value = ""
    .Range("Rake_tColor").Value = ""
    .Range("Eave_tColor").Value = ""
    .Range("OutsideCorner_tColor").Value = ""
    .Range("WallInsulation").Value = ""
    .Range("RoofInsulation").Value = ""
    .Range("RidgeVentQty").Value = ""
    .Range("TranslucentWallPanelQty").Value = ""
    .Range("SkylightQty").Value = ""
    .Range("RidgeVentType").Value = ""
    .Range("TranslucentWallPanelLength").Value = ""
    .Range("SkylightLength").Value = ""
    
    'enable change events/protections to help clear out possibly visible cells
    Call .UpdatesEventsProtection(True)
    .Range("Building_Width").Value = ""
    .Range("Roof_Pitch").Value = ""
    .Range("Building_Height").Value = ""
    .Range("Building_Length").Value = ""
    .Range("BayNum").Value = ""
    .Range("Roof_Shape").Value = ""
    .Range("LinerPanels").Value = ""
    .Range("Wainscot").Value = ""
    .Range("GutterAndDownspouts").Value = ""
    .Range("DownspoutColor").Value = ""
    .Range("GutterColor").Value = ""
    .Range("PDoorNum").Value = ""
    .Range("OHDoorNum").Value = ""
    .Range("WindowNum").Value = ""
    .Range("MiscFONum").Value = ""
    .Range("AlterWalls").Value = ""
    .Range("e1_GableOverhang").Value = ""
    .Range("s2_EaveOverhang").Value = ""
    .Range("e3_GableOverhang").Value = ""
    .Range("s4_EaveOverhang").Value = ""
    .Range("e1_GableExtension").Value = ""
    .Range("s2_EaveExtension").Value = ""
    .Range("e3_GableExtension").Value = ""
    .Range("s4_EaveExtension").Value = ""
End With




End Sub

Sub DeleteSheets()
Dim CostEstimate As Worksheet
Dim Description As Worksheet
Dim MaterialsPrice As Worksheet
Dim MatSht1 As Worksheet
Dim SteelMember As Worksheet
Dim SteelMaterials As Worksheet
Dim SteelOutput As Worksheet
Dim VendorMisc As Worksheet
Dim VendorSheetMetal As Worksheet
Dim WallDrawings As Worksheet

Set CostEstimate = CostEstimateShtTmp1
Set Description = DescriptionShtTmp1
Set MaterialsPrice = MaterialsPriceShtTmp1
Set MatSht1 = MatShtTmp3
Set SteelMember = SteelCompleteMemberShtTmp1
Set SteelMaterials = SteelMaterialsListTmp1
Set SteelOutput = SteelOutputShtTmp1
Set VendorMisc = VendorMiscMaterialsShtTmp1
Set VendorSheetMetal = VendorSheetMetalShtTmp1
Set WallDrawings = ThisWorkbook.Worksheets("Wall Drawings")


'delete temporary sheets
Application.DisplayAlerts = False
     CostEstimate.Delete
     Description.Delete
     MaterialsPrice.Delete
     MatSht1.Delete
     SteelMember.Delete
     SteelMaterials.Delete
     SteelOutput.Delete
     VendorMisc.Delete
     VendorSheetMetal.Delete
     WallDrawings.Delete
     BuildertrendTmp1.Delete    'buildertrend sheet
Application.DisplayAlerts = True

End Sub

Sub ImportOpen(ImportFile As Workbook)
   
Dim FileToOpen As Variant

'User selects file to import
FileToOpen = Application.GetOpenFilename(Title:="Browse for your File  & Import Data", FileFilter:="Excel Files (*.xlsx),*.xlsx")

'check valid file
If FileToOpen <> False Then
    Set ImportFile = Application.Workbooks.Open(FileToOpen)
    'Set Sourcews = Sourcewb.Worksheets(1)
Else
    Exit Sub
End If

End Sub

Sub ImportOHDoorPriceSht()

Dim ImportFile As Workbook
Dim ImportSht As Worksheet
Dim ImportTbl As ListObject
Dim SectionalOHDoorPriceTbl As ListObject
Dim mCell As Range
Dim i As Integer
Dim j As Integer
Dim Vendor As String

Application.ScreenUpdating = False

Call ImportOpen(ImportFile)

If ImportFile Is Nothing Then
    Exit Sub
End If

Set ImportSht = ImportFile.Worksheets(1)
Set ImportTbl = ImportSht.ListObjects(1)
Set SectionalOHDoorPriceTbl = MasterPriceSht.ListObjects("SectionalOHDoorPriceTbl")

MasterPriceSht.Unprotect "WhiteTruckMafia"
ImportSht.Unprotect "WhiteTruckMafia"

SectionalOHDoorPriceTbl.AutoFilter.ShowAllData
MasterPriceSht.UsedRange.Rows.Hidden = False
ImportTbl.AutoFilter.ShowAllData

If ImportTbl.DataBodyRange.Cells.Count <> SectionalOHDoorPriceTbl.DataBodyRange.Cells.Count Then
    MsgBox "The import data table does not match the Sectional OH Door Price Table, please enter the new data manually or choose a different file."
    Exit Sub
End If



For i = 1 To ImportTbl.ListColumns.Count
    If i <> 1 Then
        For j = 1 To ImportTbl.ListRows.Count
            Set mCell = ImportTbl.DataBodyRange(j, i)
            If mCell.Value <> "-" And mCell.Value <> 0 And mCell.EntireRow.Hidden = False Then
                Debug.Print mCell.Value
                Debug.Print mCell.Address
                SectionalOHDoorPriceTbl.DataBodyRange(j, i).Value = mCell.Value
            End If
        Next j
    End If
Next i
    
SectionalOHDoorPriceTbl.AutoFilter.ShowAllData

MasterPriceSht.Protect "WhiteTruckMafia"
ImportSht.Protect "WhiteTruckMafia"

ImportFile.Close False
Application.ScreenUpdating = True

End Sub


Sub ImportSheetMetalPriceSht()

Dim ImportFile As Workbook
Dim ImportSht As Worksheet
Dim ImportTbl As ListObject
Dim MasterPriceTbl As ListObject
Dim mCell As Range
Dim i As Integer
Dim j As Integer
Dim Vendor As String

Application.ScreenUpdating = False

Call ImportOpen(ImportFile)

If ImportFile Is Nothing Then
    Exit Sub
End If

Set ImportSht = ImportFile.Worksheets(1)
Set ImportTbl = ImportSht.ListObjects(1)
Set MasterPriceTbl = MasterPriceSht.ListObjects("MasterPriceTbl")

MasterPriceTbl.AutoFilter.ShowAllData
ImportTbl.AutoFilter.ShowAllData

If ImportTbl.DataBodyRange.Cells.Count <> MasterPriceTbl.DataBodyRange.Cells.Count Then
    MsgBox "The import data table does not match the Master Price Table, please enter the new data manually or choose a different file."
    Exit Sub
End If

Vendor = ImportSht.Range("B1").Value

MasterPriceSht.Unprotect "WhiteTruckMafia"
ImportSht.Unprotect "WhiteTruckMafia"

MasterPriceTbl.DataBodyRange.AutoFilter MasterPriceTbl.ListColumns.Count, Vendor
ImportTbl.DataBodyRange.AutoFilter MasterPriceTbl.ListColumns.Count, Vendor


For i = 1 To ImportTbl.ListColumns.Count
    If i <> 1 And i <> 2 And i < 11 Then
        For j = 1 To ImportTbl.ListRows.Count
            Set mCell = ImportTbl.DataBodyRange(j, i)
            If mCell.Value <> "-" And mCell.Value <> 0 And mCell.EntireRow.Hidden = False Then
                Debug.Print mCell.Value
                Debug.Print mCell.Address
                MasterPriceTbl.DataBodyRange(j, i).Value = mCell.Value
            End If
        Next j
    End If
Next i
    

MasterPriceTbl.AutoFilter.ShowAllData

MasterPriceSht.Protect "WhiteTruckMafia"
ImportSht.Protect "WhiteTruckMafia"

ImportFile.Close False
Application.ScreenUpdating = True

End Sub

Sub CreateSheetMetalPriceExport()

Application.ScreenUpdating = False

Dim MasterPriceTbl As ListObject
Dim ExportSht As Worksheet
Dim NewName As String
Dim NewPriceSht As Worksheet
Dim NewExportFile As Workbook
Dim FirstCol As Integer
Dim LastCol As Integer
Dim FirstRow As Integer
Dim i As Double
Dim j As Double
Dim mCell As Range
Dim Vendor As String


Set ExportSht = MasterPriceSht
Vendor = ExportSht.Range("SelectedVendor").Value
NewName = "SheetMetalPriceSht_" & Vendor & "_" & Format(Date, "mmddyy")

Call SaveAs(NewName, ExportSht)

'If Dir(ThisWorkbook.path & "\" & NewName & ".xlsx") = "" Then
'    Exit Sub
'End If



Set NewExportFile = Workbooks(NewName & ".xlsx")
Set NewPriceSht = NewExportFile.Worksheets(1)
Set MasterPriceTbl = NewPriceSht.ListObjects("MasterPriceTbl")
FirstCol = MasterPriceTbl.ListColumns(1).DataBodyRange.Column - 2


NewPriceSht.Unprotect ("WhiteTruckMafia")
'delete columns to left
For i = 1 To FirstCol
    NewPriceSht.Columns(1).Delete
Next i
'delete columns to right
With MasterPriceTbl
    LastCol = .ListColumns(.ListColumns.Count).DataBodyRange.Column + 2
End With
For i = LastCol To NewPriceSht.UsedRange.SpecialCells(xlCellTypeVisible).Columns.Count
    NewPriceSht.Columns(LastCol).Delete
Next i
'delete rows above
For i = 1 To MasterPriceTbl.HeaderRowRange.Row - 5
    NewPriceSht.Rows(1).Delete
Next i

NewPriceSht.UsedRange.Rows.Hidden = False

NewPriceSht.Range("A1").Value = "Prepared For:"
NewPriceSht.Range("B1").Value = Vendor
NewPriceSht.Range("A1").HorizontalAlignment = xlRight
NewPriceSht.Range("B1").HorizontalAlignment = xlLeft
NewPriceSht.Range("A1:B1").Locked = True

MasterPriceTbl.DataBodyRange.Locked = True
MasterPriceTbl.HeaderRowRange.Locked = True

For i = 1 To MasterPriceTbl.ListColumns.Count
    If i <> 1 And i <> 2 And i < 11 Then
        For j = 1 To MasterPriceTbl.ListRows.Count
            Set mCell = MasterPriceTbl.DataBodyRange(j, i)
            If mCell.Value <> "-" And mCell.Value <> 0 Then
                mCell.Locked = False
            Else
                mCell.Locked = True
            End If
        Next j
    End If
Next i

MasterPriceTbl.AutoFilter.ShowAllData
MasterPriceTbl.DataBodyRange.AutoFilter MasterPriceTbl.ListColumns.Count, Vendor

                
NewPriceSht.Protect ("WhiteTruckMafia")
NewExportFile.Close True

Application.ScreenUpdating = True
End Sub

Sub CreateOHDoorPriceExport()

Application.ScreenUpdating = False

Dim SectionalOHDoorPriceTbl As ListObject
Dim ExportSht As Worksheet
Dim NewName As String
Dim NewPriceSht As Worksheet
Dim NewExportFile As Workbook
Dim FirstCol As Integer
Dim LastCol As Integer
Dim FirstRow As Integer
Dim i As Double

Set ExportSht = MasterPriceSht
NewName = "OHDoorPriceSht_" & Format(Date, "mmddyy")

Call SaveAs(NewName, ExportSht)

'If Dir(ThisWorkbook.path & "\" & NewName & ".xlsx") = "" Then
'    Exit Sub
'End If

Set NewExportFile = Workbooks(NewName & ".xlsx")
Set NewPriceSht = NewExportFile.Worksheets(1)
Set SectionalOHDoorPriceTbl = NewPriceSht.ListObjects("SectionalOHDoorPriceTbl")
FirstCol = SectionalOHDoorPriceTbl.ListColumns(1).DataBodyRange.Column - 2


NewPriceSht.Unprotect ("WhiteTruckMafia")
'delete columns to left
For i = 1 To FirstCol
    NewPriceSht.Columns(1).Delete
Next i
'delete columns to right
With SectionalOHDoorPriceTbl
    LastCol = .ListColumns(.ListColumns.Count).DataBodyRange.Column + 2
End With
For i = LastCol To NewPriceSht.UsedRange.SpecialCells(xlCellTypeVisible).Columns.Count
    NewPriceSht.Columns(LastCol).Delete
Next i
'delete rows above
For i = 1 To SectionalOHDoorPriceTbl.HeaderRowRange.Row - 5
    NewPriceSht.Rows(1).Delete
Next i

NewPriceSht.UsedRange.Rows.Hidden = False

SectionalOHDoorPriceTbl.DataBodyRange.Locked = False
NewPriceSht.Protect ("WhiteTruckMafia")
NewExportFile.Close True

Application.ScreenUpdating = True
End Sub

Sub SaveAs(NewName As String, CopySht As Worksheet)
 
    Dim FName           As String
    Dim NewBook         As Workbook
    Application.DisplayAlerts = False
    
    
    
    If Application.OperatingSystem Like "Windows" Then
        FName = ThisWorkbook.path & "\" & NewName & ".xlsx"
        'FName = Replace(FName, "\", ":")
    Else
        FName = ThisWorkbook.path & "/" & NewName & ".xlsx"
    End If
    
    Set NewBook = Workbooks.Add
 
    CopySht.Copy Before:=NewBook.Sheets(1)
 
    If Dir(FName) <> "" Then
        MsgBox "File " & FName & " already exists. To save a new version, delete the existing file and try again."
        NewBook.Close False
        Workbooks.Open (FName)
    Else
        NewBook.SaveAs Filename:=FName
    End If
    
    Application.DisplayAlerts = True
 
End Sub

Sub PrintFloorplan()

Dim DrawSht As Worksheet
Dim LastRow As Integer

Set DrawSht = ThisWorkbook.Worksheets("Wall Drawings")

Application.DisplayAlerts = False

Dim Length As Double
Dim Width As Double
Dim Zoom As Double
Length = (EstSht.Range("Building_Length") + EstSht.Range("e1_GableExtension").Value + EstSht.Range("e3_GableExtension").Value)
Width = (EstSht.Range("Building_Width").Value + EstSht.Range("s2_EaveExtension").Value + EstSht.Range("s4_EaveExtension").Value)

If Length <= 40 And Width <= 80 Then
    Zoom = 50
ElseIf Length <= 70 And Width <= 120 Then
    Zoom = 35
Else
    Zoom = 25
End If
'Drawings (first page is landscape)
Application.PrintCommunication = False
With DrawSht.PageSetup
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .Zoom = Zoom
    .Orientation = xlLandscape
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = True
DrawSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False, From:=1, To:=1
    
Application.DisplayAlerts = True
End Sub

Sub PrintManagerPackage()

Dim CostSht As Worksheet
Dim LastRow As Integer

Call PrintEmployeePackage

Set CostSht = ThisWorkbook.Worksheets("Cost Estimate")

Application.DisplayAlerts = False
'Description
LastRow = CostSht.Cells(Rows.Count, 1).End(xlUp).Row
Application.PrintCommunication = False
With CostSht.PageSetup
    .PrintArea = "A1:F" & LastRow
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = False
CostSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
Application.DisplayAlerts = True

End Sub

Sub PrintDescription()

Dim DescriptionSht As Worksheet
Dim LastRow As Integer

Set DescriptionSht = ThisWorkbook.Worksheets("Project Description")

Application.DisplayAlerts = False
'Description
LastRow = DescriptionSht.Cells(Rows.Count, 1).End(xlUp).Row
Application.PrintCommunication = False
With DescriptionSht.PageSetup
    .PrintArea = "A1:A" & LastRow
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = False
DescriptionSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
Application.DisplayAlerts = True

End Sub

Sub PrintDrawingsOnly()
Dim DescriptionSht As Worksheet
Dim SheetMetalSht As Worksheet
Dim SteelSht As Worksheet
Dim DrawSht As Worksheet
Dim MiscMSht As Worksheet
Dim LastRow As Integer

Set DescriptionSht = ThisWorkbook.Worksheets("Project Description")
Set SheetMetalSht = ThisWorkbook.Worksheets("Employee Materials List")
Set SteelSht = ThisWorkbook.Worksheets("Structural Steel Materials List")
Set MiscMSht = ThisWorkbook.Worksheets("Vendor Misc. Materials")
Set DrawSht = ThisWorkbook.Worksheets("Wall Drawings")

Application.DisplayAlerts = False

Dim Length As Double
Dim Width As Double
Dim Zoom As Double
Length = (EstSht.Range("Building_Length") + EstSht.Range("e1_GableExtension").Value + EstSht.Range("e3_GableExtension").Value)
Width = (EstSht.Range("Building_Width").Value + EstSht.Range("s2_EaveExtension").Value + EstSht.Range("s4_EaveExtension").Value)

If Length <= 50 And Width <= 80 Then
    Zoom = 50
Else
    Zoom = 25
End If
'Drawings (first page is landscape)
Application.PrintCommunication = False
With DrawSht.PageSetup
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .Zoom = Zoom
    .Orientation = xlLandscape
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = True
DrawSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
    
Application.DisplayAlerts = True
End Sub


Sub PrintEmployeePackage()

Dim DescriptionSht As Worksheet
Dim SheetMetalSht As Worksheet
Dim SteelSht As Worksheet
Dim DrawSht As Worksheet
Dim MiscMSht As Worksheet
Dim LastRow As Integer

Set DescriptionSht = ThisWorkbook.Worksheets("Project Description")
Set SheetMetalSht = ThisWorkbook.Worksheets("Employee Materials List")
Set SteelSht = ThisWorkbook.Worksheets("Structural Steel Materials List")
Set MiscMSht = ThisWorkbook.Worksheets("Vendor Misc. Materials")
Set DrawSht = ThisWorkbook.Worksheets("Wall Drawings")

Application.DisplayAlerts = False

Dim Length As Double
Dim Width As Double
Dim Zoom As Double
Length = (EstSht.Range("Building_Length") + EstSht.Range("e1_GableExtension").Value + EstSht.Range("e3_GableExtension").Value)
Width = (EstSht.Range("Building_Width").Value + EstSht.Range("s2_EaveExtension").Value + EstSht.Range("s4_EaveExtension").Value)

If Length <= 50 And Width <= 80 Then
    Zoom = 50
Else
    Zoom = 25
End If
'Drawings (first page is landscape)
Application.PrintCommunication = False
With DrawSht.PageSetup
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .Zoom = Zoom
    .Orientation = xlLandscape
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = True
DrawSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
'Misc
LastRow = MiscMSht.Cells(Rows.Count, 1).End(xlUp).Row
Application.PrintCommunication = False
With MiscMSht.PageSetup
    .PrintArea = "A1:E" & LastRow
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = True
MiscMSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
      
'Steel
LastRow = SteelSht.Cells(Rows.Count, 1).End(xlUp).Row
Application.PrintCommunication = False
With SteelSht.PageSetup
    .PrintArea = "A1:E" & LastRow
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = True
SteelSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False

'Sheet Metal
LastRow = SheetMetalSht.Cells(Rows.Count, 1).End(xlUp).Row
Application.PrintCommunication = False
With SheetMetalSht.PageSetup
    .PrintArea = "A1:E" & LastRow
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = True
SheetMetalSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False

'Description
LastRow = DescriptionSht.Cells(Rows.Count, 1).End(xlUp).Row
Application.PrintCommunication = False
With DescriptionSht.PageSetup
    .PrintArea = "A1:A" & LastRow
    .FitToPagesWide = 1
    .FitToPagesTall = 0
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.25)
    .RightMargin = Application.InchesToPoints(0.25)
    .TopMargin = Application.InchesToPoints(0.25)
    .BottomMargin = Application.InchesToPoints(0.25)
    .HeaderMargin = Application.InchesToPoints(0.2)
    .FooterMargin = Application.InchesToPoints(0.2)
End With
Application.PrintCommunication = False
DescriptionSht.PrintOut Copies:=1, Collate:=True, _
    IgnorePrintAreas:=False
    
Application.DisplayAlerts = True
    
End Sub

Sub SaveAsNewEstimate()

'Export doc as .xlsm

Dim OS As String
OS = Application.OperatingSystem

Dim un As Variant

un = (Environ$("Username"))

Dim wb As Workbook
Set wb = ThisWorkbook

Dim EstName As String
Dim FilePath, FileOnly, PathOnly, NewFilePath, CustomerFolderPath, ClientName As String

FilePath = wb.FullName
FileOnly = wb.Name
PathOnly = Left(FilePath, Len(FilePath) - Len(FileOnly))

With EstSht
    ClientName = .Range("CustomerName").Value
    CustomerFolderPath = PathOnly & ClientName
    EstName = .Range("CustomerName").Value & "_Estimate_" & .Range("Building_Width").Value & "x" & .Range("Building_Length").Value & "x" & .Range("Building_Height").Value & "_" & Month(Now()) & Year(Now())
End With

EstName = ValidWBName(EstName)
NewFilePath = CustomerFolderPath & EstName

If InStr(OS, "Windows") > 0 Then
    MakeNewFolderPC (CustomerFolderPath)
    wb.SaveAs Filename:=NewFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
Else
    MakeNewFolderMAC (ClientName)
    wb.SaveAs Filename:=NewFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
End If

End Sub

Function MakeNewFolderPC(path As String)

Dim fso As New FileSystemObject

'examples for what are the input arguments
'strDir = "Folder"
'strPath = "C:\"

'path = strPath & strDir

If Not fso.FolderExists(path) Then

    ' doesn't exist, so create the folder
    fso.CreateFolder path

End If

End Function

Sub MakeNewFolderMAC(ClientName As String)
    'Note: This macro uses the FileOrFolderExistsOnYourMac function.
    'Note : Use 1 as second argument for File and 2 for Folder
    'Test if the folder with the name TestFolder is on your desktop
    Dim FolderPath As String
    FolderPath = MacScript("return (path to desktop folder) as string") & ClientName

    If Right(FolderPath, 1) = Application.PathSeparator Then
        MsgBox "Remove the / at the end of the FolderPath"
        Exit Sub
    End If

    If FileOrFolderExistsOnYourMac(FolderPath, 2) = True Then
        MsgBox "Folder exists."
    Else
        MkDir MacScript("return POSIX path of (" & Chr(34) & FolderPath & Chr(34) & ")")
        MsgBox "Folder not exists but created ."
    End If
End Sub

Function FileOrFolderExistsOnYourMac(FileOrFolderstr As String, FileOrFolder As Long) As Boolean
    'Ron de Bruin : 13-Dec-2020, for Excel 2016 and higher
    'Function to test if a file or folder exist on your Mac
    'Use 1 as second argument for File and 2 for Folder
    Dim ScriptToCheckFileFolder As String
    Dim FileOrFolderPath As String
    
    If FileOrFolder = 1 Then
        'File test
        On Error Resume Next
        FileOrFolderPath = Dir(FileOrFolderstr & "*")
        On Error GoTo 0
        If Not FileOrFolderPath = vbNullString Then FileOrFolderExistsOnYourMac = True
    Else
        'folder test
        On Error Resume Next
        FileOrFolderPath = Dir(FileOrFolderstr & "*", vbDirectory)
        On Error GoTo 0
        If Not FileOrFolderPath = vbNullString Then FileOrFolderExistsOnYourMac = True
    End If
End Function

Private Function ValidWBName(Arg As String) As String
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .Pattern = "[\\/:\*\?""<>\|]"
        .Global = True
        ValidWBName = .Replace(Arg, "")
    End With
End Function

Sub BayUpdate(BayRange As Range, BLen As Range, ChangeCell As Range)
Dim BayCell As Range
Dim BayValue As Range
Dim BaySum As Integer

'set bay sum
BaySum = HiddenSht.Range("TotalBayLength").Value

'reset any blanks to 0
For Each BayCell In BayRange
    If BayCell.Value = "" Then BayCell.Value = 0
Next BayCell

'check for bay length over total building length
If BaySum > BLen.Value Then
    'error message
    MsgBox "The total bay length cannot exceed the building length! Please correct the data and try again.", vbExclamation, "Excess Bay Length"
    'change cell value back to 0
    ChangeCell.Value = 0
End If
    

End Sub

Sub MaterialsListCaller()
Dim Confirm As Variant
Dim FOCell As Range
Dim ItemCount As Integer
Dim MissingData As Boolean
Dim sqrFootage As Double
Dim BayNum As Double
Dim Bay1Length As Double
Dim LastBayLength As Double
Dim Bay1Overhang As Double
Dim LastBayOverhang As Double
Dim b As clsBuilding

'''Deternine FO trim generation time
With EstSht
    'Personell Doors
    For Each FOCell In Range(.Range("pDoorCell1"), .Range("pDoorCell12"))
        'if cell isn't hidden, door size is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'add FO trim pieces to item count
            ItemCount = ItemCount + 3
        End If
    Next FOCell
    'Overhead Doors
    For Each FOCell In Range(.Range("OHDoorCell1"), .Range("OHDoorCell12"))
        'if cell isn't hidden, door width is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'add FO trim pieces to item count
            ItemCount = ItemCount + (3 * 5)
        End If
    Next FOCell
    'Windows
    For Each FOCell In Range(.Range("WindowCell1"), .Range("WindowCell12"))
        'if cell isn't hidden, door width is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'add FO trim pieces to item count
            ItemCount = ItemCount + 4
        End If
    Next FOCell
    'Misc Fos
    For Each FOCell In Range(.Range("MiscFOCell1"), .Range("MiscFOCell12"))
        'if cell isn't hidden, door width is entered
        If FOCell.EntireRow.Hidden = False And FOCell.offset(0, 1).Value <> "" Then
            'add FO trim pieces to item count
            ItemCount = ItemCount + (4 * 5)
        End If
    Next FOCell
    
    sqrFootage = .Range("Building_Width").Value * .Range("Building_Height").Value * .Range("Building_Length").Value
    ItemCount = ItemCount + (Application.WorksheetFunction.RoundUp(sqrFootage / 50000, 0) * 10)
    
End With




'confirmation message dependent upon item count
If ItemCount > 0 Then
    Confirm = MsgBox("Would you like to generate a materials list using the information entered? This will take approximately " & ItemCount & " seconds to complete.", vbInformation + vbYesNo, "Confirm Materials List Generation")
Else
    Confirm = MsgBox("Would you like to generate a materials list using the information entered?", vbInformation + vbYesNo, "Confirm Materials List Generation")
End If

''' Check for missing information
With EstSht
    Select Case True
    '' Building info
    Case .Range("Roof_Shape").Value = ""
        MsgBox "The building's roof shape must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Roof_Shape")
        MissingData = True
    Case .Range("Building_Width").Value = ""
        MsgBox "The building's width must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Building_Width")
        MissingData = True
    Case .Range("Roof_Pitch").Value = ""
        MsgBox "The building's roof pitch must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Roof_Pitch")
        MissingData = True
    Case .Range("Building_Height").Value = ""
        MsgBox "The building's height must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Building_Height")
        MissingData = True
    Case .Range("Building_Length").Value = ""
        MsgBox "The building's length must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Building_Length")
        MissingData = True
    '' Panel info
    Case .Range("Wall_pShape").Value = ""
        MsgBox "The building's wall panel shape must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Wall_pShape")
        MissingData = True
    Case .Range("Wall_pType").Value = ""
        MsgBox "The building's wall panel type must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Wall_pType")
        MissingData = True
    Case .Range("Wall_Color").Value = ""
        MsgBox "The building's wall color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Wall_Color")
        MissingData = True
    Case .Range("Roof_pShape").Value = ""
        MsgBox "The building's roof panel shape must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Roof_pShape")
        MissingData = True
    Case .Range("Roof_pType").Value = ""
        MsgBox "The building's roof panel type must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Roof_pType")
        MissingData = True
    Case .Range("Roof_Color").Value = ""
        MsgBox "The building's roof color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Roof_Color")
        MissingData = True
    '' Trim Colors
    Case .Range("Rake_tColor").Value = ""
        MsgBox "The building's rake trim color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Rake_tColor")
        MissingData = True
    Case .Range("Eave_tColor").Value = ""
        MsgBox "The building's eave trim color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Eave_tColor")
        MissingData = True
    Case .Range("OutsideCorner_tColor").Value = ""
        MsgBox "The building's corner trim color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("OutsideCorner_tColor")
        MissingData = True
    Case .Range("Base_tColor").Value = ""
        MsgBox "The building's base trim color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("Base_tColor")
        MissingData = True
    'check for wainscot panel type/color
    Case .Range("e1_Wainscot").Value <> "None" And (.Range("e1_Wainscot").offset(0, 1).Value = "" Or .Range("e1_Wainscot").offset(0, 2).Value = "")
        MsgBox "The building's wainscot panel type and panel color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("e1_Wainscot")
        MissingData = True
    Case .Range("s2_Wainscot").Value <> "None" And (.Range("s2_Wainscot").offset(0, 1).Value = "" Or .Range("s2_Wainscot").offset(0, 2).Value = "")
        MsgBox "The building's wainscot panel type and panel color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("s2_Wainscot")
        MissingData = True
    Case .Range("e3_Wainscot").Value <> "None" And (.Range("e3_Wainscot").offset(0, 1).Value = "" Or .Range("e3_Wainscot").offset(0, 2).Value = "")
        MsgBox "The building's wainscot panel type and panel color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("e3_Wainscot")
        MissingData = True
    Case .Range("s4_Wainscot").Value <> "None" And (.Range("s4_Wainscot").offset(0, 1).Value = "" Or .Range("s4_Wainscot").offset(0, 2).Value = "")
        MsgBox "The building's wainscot panel type and panel color must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
        Application.GoTo .Range("s4_Wainscot")
        MissingData = True
    End Select
End With

Dim i As Integer
Dim j As Integer
Dim MissingOHData As Boolean
With EstSht
'Check OHDoor data
For i = 1 To 11
    If .Range("OHDoorCell1").offset(i - 1, 1).Value <> "" And .Range("OHDoorCell1").offset(i - 1, 1).EntireRow.Hidden = False Then
        MissingOHData = False
        For j = 1 To 8
            If .Range("OHDoorCell1").offset(i - 1, j + 1).Value = "" Then
                MsgBox "Overhead door information must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
                MissingOHData = True
                MissingData = True
                Application.GoTo .Range("OHDoorCell1").offset(i - 1, 1)
                Exit For
            End If
        Next j
        If MissingOHData = True Then Exit For
    End If
Next i


'Check PDoor data
For i = 1 To 11
    If .Range("pDoorCell1").offset(i - 1, 1).Value <> "" And .Range("pDoorCell1").offset(i - 1, 1).EntireRow.Hidden = False Then
        MissingOHData = False
        For j = 1 To 7
            If .Range("pDoorCell1").offset(i - 1, j + 1).Value = "" Then
                MsgBox "Personnel door information must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
                MissingOHData = True
                MissingData = True
                Application.GoTo .Range("pDoorCell1").offset(i - 1, 1)
                Exit For
            End If
        Next j
        If MissingOHData = True Then Exit For
    End If
Next i


'Check Window data
For i = 1 To 11
    If .Range("WindowCell1").offset(i - 1, 1).Value <> "" And .Range("WindowCell1").offset(i - 1, 1).EntireRow.Hidden = False Then
        MissingOHData = False
        For j = 1 To 5
            If .Range("WindowCell1").offset(i - 1, j + 1).Value = "" Then
                MsgBox "Window information must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
                MissingOHData = True
                MissingData = True
                Application.GoTo .Range("WindowCell1").offset(i - 1, 1)
                Exit For
            End If
        Next j
        If MissingOHData = True Then Exit For
    End If
Next i


'Check Misc FO data
For i = 1 To 11
    If .Range("MiscFOCell1").offset(i - 1, 1).Value <> "" And .Range("MiscFOCell1").offset(i - 1, 1).EntireRow.Hidden = False Then
        MissingOHData = False
        For j = 1 To 7
            If .Range("MiscFOCell1").offset(i - 1, j + 1).Value = "" Then
                MsgBox "Misc. Framed Opening information must be entered. Please enter this data before generating a materials list.", vbExclamation, "Missing Data"
                MissingOHData = True
                MissingData = True
                Application.GoTo .Range("MiscFOCell1").offset(i - 1, 1)
                Exit For
            End If
        Next j
        If MissingOHData = True Then Exit For
    End If
Next i
End With

'Check that Bay Lenghts + Overhangs don't add up to more than 30'
With EstSht
Dim e1Overhang As Boolean
Dim e1Extension As Boolean
Dim s2Overhang As Boolean
Dim s2Extension As Boolean
Dim e3Overhang As Boolean
Dim e3Extension As Boolean
Dim s4Overhang As Boolean
Dim s4Extension As Boolean
 
If .Range("e1_GableOverhang").Value > 0 Then e1Overhang = True
If .Range("e1_GableExtension").Value > 0 Then e1Extension = True
If .Range("s2_EaveOverhang").Value > 0 Then s2Overhang = True
If .Range("s2_EaveExtension").Value > 0 Then s2Extension = True
If .Range("e3_GableOverhang").Value > 0 Then e3Overhang = True
If .Range("e3_GableExtension").Value > 0 Then e3Extension = True
If .Range("s4_EaveOverhang").Value > 0 Then s4Overhang = True
If .Range("s4_EaveExtension").Value > 0 Then s4Extension = True

 
'Get Bay Num
BayNum = .Range("BayNum").Value
If BayNum = 0 Or BayNum = 1 Then
    If .Range("Building_Length").Value > 30 Then
        MsgBox "Bays cannot be longer than 30'. Please enter values less than 30' before generating a materials list.", vbExclamation, "Bay Length Error"
        Application.GoTo .Range("Building_Length")
        MissingData = True
    End If
End If

'check Extension Overhangs
If e1Extension And e1Overhang Then
    If .Range("e1_GableExtension").Value + .Range("e1_GableOverhang").Value > 30 Then
        MsgBox "The combined extension length and overhang length at endwall 1 is longer than 30'. Please enter values less than 30' before generating a materials list.", vbExclamation, "Extension Length Error"
        Application.GoTo .Range("e1_GableOverhang")
        MissingData = True
    End If
End If
If e3Extension And e3Overhang Then
    If .Range("e3_GableExtension").Value + .Range("e3_GableOverhang").Value > 30 Then
        MsgBox "The combined extension length and overhang length at endwall 3 is longer than 30'. Please enter values less than 30' before generating a materials list.", vbExclamation, "Extension Length Error"
        Application.GoTo .Range("e3_GableOverhang")
        MissingData = True
    End If
End If
'If s2Extension And s2Overhang Then
'    If .Range("s2_EaveExtension").Value + .Range("s2_EaveOverhang").Value > 30 Then
'        MsgBox "The combined extension length and overhang length at sidewall 2 is longer than 30'. Please enter values less than 30' before generating a materials list.", vbExclamation, "Extension Length Error"
'        Application.GoTo .Range("s2_EaveOverhang")
'        MissingData = True
'    End If
'End If
'If s4Extension And s4Overhang Then
'    If .Range("s4_EaveExtension").Value + .Range("s4_EaveOverhang").Value > 30 Then
'        MsgBox "The combined extension length and overhang length at endwall 1 is longer than 30'. Please enter values less than 30' before generating a materials list.", vbExclamation, "Extension Length Error"
'        Application.GoTo .Range("s4_EaveOverhang")
'        MissingData = True
'    End If
'End If




If BayNum > 1 And e1Extension = False And e3Extension = False Then
    'Get Bay 1 Length
    Bay1Length = .Range("Bay1_Length").Value
    
    'Get Last Bay Length
    LastBayLength = .Range("Bay1_Length").offset(BayNum - 1, 0)
    
    'Get Bay 1 Overhang
    Bay1Overhang = .Range("e1_GableOverhang").Value
    
    'Get Last Bay Overhang
    LastBayOverhang = .Range("e3_GableOverhang").Value
    
    If Bay1Length + Bay1Overhang > 30 Then
        MsgBox "The combined bay length (" & Bay1Length & "') and overhang length (" & Bay1Overhang & "') at endwall 1 is longer than 30'. Please enter values less than 30' before generating a materials list.", vbExclamation, "Bay Length Error"
        Application.GoTo .Range("e1_GableOverhang")
        MissingData = True
    End If
    If LastBayLength + LastBayOverhang > 30 Then
        MsgBox "The combined bay length (" & LastBayLength & "') and overhang length (" & LastBayOverhang & "') at endwall 3 is longer than 30'. Please enter values less than 30' before generating a materials list.", vbExclamation, "Bay Length Error"
        Application.GoTo .Range("e3_GableOverhang")
        MissingData = True
    End If
End If
End With
    
    

If Confirm = vbYes And MissingData = False Then
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set b = New clsBuilding
    Set BTItems = New clsBuildertrend
    Call MaterialsListGen.MaterialsListGen(b)
    Call StructuralSteelMaterialsGen.StructuralSteelMaterialsGen(b)
    Call VendorAndPriceLists.VendorMaterialListsGen(b)
    Call VendorAndPriceLists.PriceListGen(b)
    Call VendorAndPriceLists.CostEstimateGen(b)
    Call VendorAndPriceLists.DescriptionGen(b)
    Call BuildertrendGen.GenSheet
    Call ReorderSheets
    'select Description sheet, show completed messaged
    ThisWorkbook.Sheets("Project Description").Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Materials list generation complete!", vbInformation, "Generation Complete"
End If


End Sub

Sub ReorderSheets()

With ThisWorkbook
    .Sheets("Structural Steel Lookup Tables").Move Before:=.Sheets(2)
    .Sheets("Master Price List").Move Before:=.Sheets(2)
    .Sheets("Vendor Sheet Metal Materials").Move Before:=.Sheets(2)
    .Sheets("Vendor Misc. Materials").Move Before:=.Sheets(2)
    .Sheets("Optimized Cut List").Move Before:=.Sheets(2)
    .Sheets("Structural Steel Materials List").Move Before:=.Sheets(2)
    .Sheets("Employee Materials List").Move Before:=.Sheets(2)
    .Sheets("Wall Drawings").Move Before:=.Sheets(2)
    .Sheets("Cost Estimate").Move Before:=.Sheets(2)
    .Sheets("Structural Steel Price List").Move Before:=.Sheets(2)
    .Sheets("Materials Price List").Move Before:=.Sheets(2)
    .Sheets("Project Description").Move Before:=.Sheets(2)
    .Sheets("Project Details").Move Before:=.Sheets(2)
    .Sheets("Buildertrend Estimate").Move Before:=BuildertrendTmp
    BuildertrendTmp.Move After:=.Sheets("Buildertrend Estimate")
End With
    
End Sub

Sub ImportProjectDetails()

Dim Sourcewb As Workbook
Dim Sourcews As Worksheet
Dim Activewb As Workbook
Dim FileToOpen As Variant
Dim i As Integer
Dim j As Integer
Dim Address As String
Dim BayErrorMsg As String
Dim OS As String
Dim FileFound As Boolean
Dim mybook As Workbook
Dim MyPath As String
Dim MyScript As String
Dim MyFiles As String
Dim MySplit As Variant
Dim n As Long
Dim FName As String

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

FileFound = False

OS = Application.OperatingSystem

If InStr(OS, "Windows") > 0 Then
    Application.ScreenUpdating = False
    Set Activewb = ThisWorkbook
    '''''Select File to Import From''''''
    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.xlsm*),*xlsm*")
    'Check Valid File and that it's an Estimating Template
    If FileToOpen <> False Then
        Set Sourcewb = Application.Workbooks.Open(FileToOpen)
        FileFound = True
    End If
Else
    '''' CODE TO OPEN FILE ON MAC
    On Error Resume Next
    MyPath = MacScript("return (path to documents folder) as String")
    'Or use MyPath = "Macintosh HD:Users:Ron:Desktop:TestFolder:"

    ' In the following statement, change true to false in the line "multiple
    ' selections allowed true" if you do not want to be able to select more
    ' than one file. Additionally, if you want to filter for multiple files, change
    ' {""com.microsoft.Excel.xls""} to
    ' {""com.microsoft.excel.xls"",""public.comma-separated-values-text""}
    ' if you want to filter on xls and csv files, for example.
    MyScript = _
    "set applescript's text item delimiters to "","" " & vbNewLine & _
               "set theFiles to (choose file of type " & _
             " {""org.openxmlformats.spreadsheetml.sheet.macroenabled""}" & _
               "with prompt ""Please select a file or files"" default location alias """ & _
               MyPath & """ multiple selections allowed false) as string" & vbNewLine & _
               "set applescript's text item delimiters to """" " & vbNewLine & _
               "return theFiles"

    MyFiles = MacScript(MyScript)
    On Error GoTo 0

    If MyFiles <> "" Then
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
        End With
        
        MyFiles = Replace(MyFiles, ":", "/")
        MyFiles = Replace(MyFiles, "Macintosh HD", "", Count:=1)

        MySplit = Split(MyFiles, ",")
        For n = LBound(MySplit) To UBound(MySplit)

            ' Get the file name only and test to see if it is open.
            FName = Right(MySplit(n), Len(MySplit(n)) - InStrRev(MySplit(n), Application.PathSeparator, , 1))
            If bIsBookOpen(FName) = False Then

                Set mybook = Nothing
                On Error Resume Next
                Set mybook = Workbooks.Open(MySplit(n))
                On Error GoTo 0

                If Not mybook Is Nothing Then
                    Set Sourcewb = mybook
                    FileFound = True
                    GoTo Continue
                End If
            End If
        Next n
    End If
End If

Continue:

If FileFound = False Then
    MsgBox "Please Select Valid File"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    GoTo ExitSub
End If

'Define Import Sheet and check for Project Details using custom function
If sheetExists("Project Details", Sourcewb) Then '''change to named range
    Set Sourcews = Sourcewb.Worksheets("Project Details")
Else
    Sourcewb.Close False
    MsgBox "Please Select Valid File" 'add title, etc.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    GoTo ExitSub
End If

''''''''Import Data; for ranges WITHOUT Change triggers, Turn Events OFF by "FALSE" parameter

Application.EnableEvents = False

CopyNamedRange "Building_Width", Sourcews, False
CopyNamedRange "Roof_Pitch", Sourcews, False
CopyNamedRange "Building_Height", Sourcews, False
CopyNamedRange "Building_Length", Sourcews, False
CopyNamedRange "BayNum", Sourcews, True '''''''''
If Sourcews.Range("BayNum").Value > 0 Then
    Application.Calculation = xlCalculationAutomatic
    For i = 0 To 11
        CopyNamedRange Sourcews.Range("Building_Height").offset(i + 3, 0).Address, Sourcews, True
    Next i
    Application.Calculation = xlCalculationManual
End If
CopyNamedRange "Roof_Shape", Sourcews, False
CopyNamedRange "Wall_pShape", Sourcews, False
CopyNamedRange "Wall_pType", Sourcews, True
CopyNamedRange "Wall_Color", Sourcews, False
CopyNamedRange "LinerPanels", Sourcews, True '''''''''
If Sourcews.Range("LinerPanels").Value = "Yes" Then
    For i = 0 To 4
        For j = 0 To 3
            CopyNamedRange Sourcews.Range("e1_LinerPanels").offset(i, j).Address, Sourcews, True
        Next j
    Next i
End If
CopyNamedRange "Roof_pShape", Sourcews, False
CopyNamedRange "Roof_pType", Sourcews, True
CopyNamedRange "Roof_Color", Sourcews, False
CopyNamedRange "AlterWalls", Sourcews, True '''''''''
If Sourcews.Range("AlterWalls").Value = "Yes" Then
    Sourcews.Unprotect "WhiteTruckMafia"
    EstSht.Unprotect "WhiteTruckMafia"
    CopyNamedRange "e1_WallStatus", Sourcews, True '''''''''
    If Sourcews.Range("e1_WallStatus").Value = "Partial" Then
        CopyNamedRange Sourcews.Range("e1_WallStatus").offset(0, 2).Address, Sourcews, False
    End If
    CopyNamedRange "s2_WallStatus", Sourcews, True '''''''''
    If Sourcews.Range("s2_WallStatus").Value = "Partial" Then
        CopyNamedRange Sourcews.Range("s2_WallStatus").offset(0, 2).Address, Sourcews, False
    End If
    CopyNamedRange "e3_WallStatus", Sourcews, True '''''''''
    If Sourcews.Range("e3_WallStatus").Value = "Partial" Then
        CopyNamedRange Sourcews.Range("e3_WallStatus").offset(0, 2).Address, Sourcews, False
    End If
    CopyNamedRange "s4_WallStatus", Sourcews, True '''''''''
    If Sourcews.Range("s4_WallStatus").Value = "Partial" Then
        CopyNamedRange Sourcews.Range("s4_WallStatus").offset(0, 2).Address, Sourcews, False
    End If
    CopyNamedRange "e1_Expandable", Sourcews, False '''''''''
    'CopyNamedRange Sourcews.Range("e1_Expandable").offset(1, 0).Address, Sourcews, False ''''''''' (not named, use offset)
    CopyNamedRange "e3_Expandable", Sourcews, False '''''''''
    'CopyNamedRange Sourcews.Range("31_Expandable").offset(1, 0).Address, Sourcews, False ''''''''' (not named, use offset)
    Sourcews.Protect "WhiteTruckMafia"
    EstSht.Protect "WhiteTruckMafia"
End If
CopyNamedRange "All_tColors", Sourcews, True '''''''''
CopyNamedRange "FO_tColor", Sourcews, False
CopyNamedRange "Base_tColor", Sourcews, False
CopyNamedRange "Rake_tColor", Sourcews, False
CopyNamedRange "Eave_tColor", Sourcews, False
CopyNamedRange "OutsideCorner_tColor", Sourcews, False
CopyNamedRange "Wainscot", Sourcews, True '''''''''
If Sourcews.Range("Wainscot").Value = "Yes" Then
    For i = 0 To 3
        For j = 0 To 2
            CopyNamedRange Sourcews.Range("e1_Wainscot").offset(i, j).Address, Sourcews, True
        Next j
    Next i
End If
CopyNamedRange "GutterAndDownspouts", Sourcews, True '''''''''
CopyNamedRange "PDoorNum", Sourcews, True '''''''''
If Sourcews.Range("PDoorNum").Value > 0 Then
    For i = 0 To Sourcews.Range("PDoorNum").Value - 1
        For j = 1 To 7
            CopyNamedRange Sourcews.Range("pDoorCell1").offset(i, j).Address, Sourcews, False
        Next j
    Next i
End If
CopyNamedRange "OHDoorNum", Sourcews, True '''''''''
If Sourcews.Range("OHDoorNum").Value > 0 Then
    For i = 0 To Sourcews.Range("OHDoorNum").Value - 1
        For j = 1 To 9
            CopyNamedRange Sourcews.Range("OHDoorCell1").offset(i, j).Address, Sourcews, False
        Next j
    Next i
End If
CopyNamedRange "WindowNum", Sourcews, True '''''''''
If Sourcews.Range("WindowNum").Value > 0 Then
    For i = 0 To Sourcews.Range("WindowNum").Value - 1
        Debug.Print (Sourcews.Range("WindowCell1").offset(i, 1).Value)
        Debug.Print (Sourcews.Range("WindowCell1").offset(i, 1).Address)
        For j = 1 To 6
            If j = 4 Then
                'skip, formula in this cell
            Else
                CopyNamedRange Sourcews.Range("WindowCell1").offset(i, j).Address, Sourcews, False
            End If
        Next j
    Next i
End If
CopyNamedRange "MiscFONum", Sourcews, True '''''''''
If Sourcews.Range("MiscFONum").Value > 0 Then
    For i = 0 To Sourcews.Range("MiscFONum").Value - 1
        For j = 1 To 8
            If j = 6 Then
                'skip, formula in this cell
            Else
                CopyNamedRange Sourcews.Range("MiscFOCell1").offset(i, j).Address, Sourcews, False
            End If
        Next j
    Next i
End If
CopyNamedRange "WallInsulation", Sourcews, False
CopyNamedRange "RoofInsulation", Sourcews, False
CopyNamedRange "RidgeVentQty", Sourcews, False
CopyNamedRange "TranslucentWallPanelQty", Sourcews, False
CopyNamedRange "SkylightQty", Sourcews, False
CopyNamedRange "RidgeVentType", Sourcews, False
CopyNamedRange "TranslucentWallPanelLength", Sourcews, False
CopyNamedRange "SkylightLength", Sourcews, False
CopyNamedRange "e1_GableOverhang", Sourcews, False
CopyNamedRange "s2_EaveOverhang", Sourcews, False
CopyNamedRange "e3_GableOverhang", Sourcews, False
CopyNamedRange "s4_EaveOverhang", Sourcews, False
CopyNamedRange "e1_GableOverhangSoffit", Sourcews, False
CopyNamedRange "s2_EaveOverhangSoffit", Sourcews, False
CopyNamedRange "e3_GableOverhangSoffit", Sourcews, False
CopyNamedRange "s4_EaveOverhangSoffit", Sourcews, False
For i = 0 To 3
    For j = 1 To 4
        Address = EstSht.Range("e1_GableOverhangSoffit").offset(i, j).Address
        CopyNamedRange Address, Sourcews, True
    Next j
Next i
CopyNamedRange "e1_GableExtension", Sourcews, True '''''''''
CopyNamedRange "s2_EaveExtension", Sourcews, True '''''''''
CopyNamedRange "e3_GableExtension", Sourcews, True '''''''''
CopyNamedRange "s4_EaveExtension", Sourcews, True '''''''''
CopyNamedRange "e1_GableExtensionSoffit", Sourcews, True '''''''''
CopyNamedRange "s2_EaveExtensionSoffit", Sourcews, True '''''''''
CopyNamedRange "e3_GableExtensionSoffit", Sourcews, True '''''''''
CopyNamedRange "s4_EaveExtensionSoffit", Sourcews, True '''''''''
For i = 0 To 3
    For j = 1 To 4
        Address = EstSht.Range("e1_GableExtensionSoffit").offset(i, j).Address
        CopyNamedRange Address, Sourcews, True
    Next j
Next i

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

Application.GoTo reference:=EstSht.Range("CustomerName"), scroll:=True

Sourcewb.Close False
ExitSub:

End Sub

Sub CopyNamedRange(Name As String, Sourcews As Worksheet, EnableEvent As Boolean)

If EnableEvent = True Then
    Application.EnableEvents = True
End If

On Error GoTo NextCell
'Ignores cells that are protected and unavailable(which means they would be on the current sheet as well)
'Also avoids errors for old files that are missing named ranges
'Sourcews.Range(Name).Copy
EstSht.Range(Name).Value = Sourcews.Range(Name).Value
If Not EstSht.Range(Name).Validation.Value Then
    EstSht.Range(Name).ClearContents
End If

NextCell:
Application.EnableEvents = False

End Sub

Function sheetExists(SheetName As String, Sourcewb As Workbook) As Boolean
On Error Resume Next
sheetExists = (Sourcewb.Sheets(SheetName).Index > 0)
End Function

Function bIsBookOpen(ByRef szBookName As String) As Boolean
' Contributed by Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function

Public Sub EventsUpdating()
If Application.EnableEvents = True Then
    Application.EnableEvents = False
    Application.ScreenUpdating = False
ElseIf Application.EnableEvents = False Then
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End If
End Sub
