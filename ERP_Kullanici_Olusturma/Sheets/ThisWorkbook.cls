VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public deg As String
Private Sub Workbook_BeforeClose(Cancel As Boolean)
If Val(deg) = 0 Then
MsgBox "Lütfen ÇIKIS Butonunu Kullaniniz.", vbCritical
Cancel = True
End If
End Sub
