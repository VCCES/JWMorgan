Attribute VB_Name = "BETASelfDestructModule"
Option Explicit

Sub KillMe()
    With ThisWorkbook
        .Saved = True
        .ChangeFileAccess Mode:=xlReadOnly
        Kill .FullName
        .Close False
    End With
End Sub

'to create BETA, add CheckDate() sub to Workbook_Open event
Sub CheckDate()

Dim BetaDate As Date

BetaDate = HiddenSht.Range("BetaDate").Value

If BetaDate < Now Then
    KillMe
End If

End Sub

Sub TestBETA()

CheckDate
End Sub
