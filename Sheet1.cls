VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As String
    Dim r As Range
    Dim c As Range
    
    ' Check intersection with first row and first column
    Set r = Application.Intersect(Target, Rows(1))
    Set c = Application.Intersect(Target, Columns(1))
    
    ' Insert formula when first column is changed
    If Not (c Is Nothing) Then
        If c.Row > 2 Then
            cell = c.Address
            c.Offset(0, 1).Formula = "=IFERROR(VLOOKUP(" & cell & ",Symbols!$A$1:$C$25243,2,FALSE),"""")"
            c.EntireColumn.AutoFit
            c.Offset(0, 1).EntireColumn.AutoFit
        End If
    End If
    
    ' Insert formula when first row is changed
    If Not (r Is Nothing) Then
        If r.Column > 2 Then
            cell = r.Address
            r.Offset(1, 0).Formula = "=IFERROR(VLOOKUP(" & cell & ",Tags!$A$1:$C$88,2,FALSE),"""")"
            r.EntireColumn.AutoFit
        End If
    End If
    
End Sub
