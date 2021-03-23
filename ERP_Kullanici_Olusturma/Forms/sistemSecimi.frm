VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sistemSecimi 
   Caption         =   "Sistemler ve Sirketler"
   ClientHeight    =   6672
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12792
   OleObjectBlob   =   "sistemSecimi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sistemSecimi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
Cells(1, 10).Value = ComboBox1.Value
'tamam
End Sub
Private Sub CommandButton2_Click()
For i = 13 To 40 Step 2
    If Cells(1, 10).Value = Cells(2, i).Value Then
        X = Cells(3, i).End(xlDown).Row
        ListBox1.List = Range(Cells(3, i), Cells(X, i)).Value
    End If
Next i
'tamam
End Sub
Private Sub CommandButton3_Click()

If ListBox1.ListIndex = -1 Then
    MsgBox "Lütfen Mevcut Sirketler listesinden seçim yapiniz."
Else:
    ListBox2.AddItem ListBox1.Text
    ListBox1.RemoveItem ListBox1.ListIndex
End If

'tamam
End Sub
Private Sub CommandButton4_Click()

If ListBox2.ListIndex = -1 Then
    MsgBox "Lütfen Güncel Sirketler listesinden seçim yapiniz."
Else:
    ListBox1.AddItem ListBox2.Text
    ListBox2.RemoveItem ListBox2.ListIndex
End If
'tamam
End Sub

Private Sub CommandButton5_Click()
Range("G5:G8").Clear

Range("G10:G14").Clear

Range("A3:A300").Clear
Range("B4:B300").Clear
Range("C4:B300").Clear
Range("D4:B300").Clear
Range("E4:B300").Clear
Range("J1:J1").Clear

Cells(1, 3) = "devam"
ThisWorkbook.deg = 1
ActiveWorkbook.Close
Excel.Application.Quit
End Sub

Private Sub onay_Click()
     
t = Cells(3, 1).End(xlDown).Row
Range(Cells(3, 1), Cells(t, 1)).Clear

If ListBox2.ListCount = 0 Then
    MsgBox "Güncel sirketler listesine sirket ekleyiniz."

Else:
Z = ListBox2.ListCount
Range("A3:A" & Z + 2).Value = ListBox2.List
f = Cells(3, 1).End(xlDown).Row
    If Cells(4, 1) = "" Then
        MsgBox "Güncel Sirketler listesinde en az iki sirket olmalidir."
    Else:
        Cells(f + 1, 1) = "END"
    End If
End If

Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.TextToColumns Destination:=Range("A2"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
    :=Array(1, 2), TrailingMinusNumbers:=True

 
r = Cells(3, 1).End(xlDown).Row
For i = 3 To r - 1
    If Len(Cells(i, 1)) = 1 Then
        Cells(i, 1).Value = "000" & Cells(i, 1).Value
    ElseIf Len(Cells(i, 1)) = 2 Then
        Cells(i, 1).Value = "00" & Cells(i, 1).Value
    ElseIf Len(Cells(i, 1)) = 3 Then
        Cells(i, 1).Value = "0" & Cells(i, 1).Value
    End If
Next i
End Sub

Private Sub ToggleButton1_Click()
If ListBox2.Value = "" And Cells(3, 1) = "" And Cells(4, 1) = "" Then
    MsgBox "Lütfen sirket veya sirketleri seçiniz. Seçim yaptiysaniz onaylayiniz."
Else:
    kullaniciBilgileri.Show
End If
End Sub

Private Sub ToggleButton2_Click()
If ListBox2.Value = "" And Cells(3, 1) = "" And Cells(4, 1) = "" Then
    MsgBox "Lütfen sirket veya sirketleri seçiniz. Seçim yaptiysaniz onaylayiniz."
Else:
    yeniOlusturma.Show
End If
End Sub

Private Sub ToggleButton3_Click()
sonuc.Show
End Sub

Private Sub topluEkle_Click()
If ListBox2.ListCount <> 0 Then
    MsgBox "Güncel sirketler listesine ekli olan sirketleri Mevcut sirketler listesine ekleyiniz. Bunun için sirket kodunun üstüne tiklayip yön tusuna basiniz."
ElseIf ListBox2.ListCount = 0 Then
    ListBox2.List = ListBox1.List
    ListBox1.Clear
End If
End Sub

Private Sub topluKaldir_Click()
If ListBox1.ListCount <> 0 Then
    MsgBox "Mevcut sirketler listesine ekli olan sirketleri Güncel sirketler listesine ekleyiniz. Bunun için sirket kodunun üstüne tiklayip yön tusuna basiniz."
ElseIf ListBox1.ListCount = 0 Then
    ListBox1.List = ListBox2.List
    ListBox2.Clear
    f = Cells(3, 1).End(xlDown).Row
    Range(Cells(3, 1), Cells(f, 1)).Clear
End If
End Sub

