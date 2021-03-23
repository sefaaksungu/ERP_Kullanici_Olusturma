VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sonuc 
   ClientHeight    =   6672
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12792
   OleObjectBlob   =   "sonuc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sonuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
p = Cells(4, 4).End(xlDown).Row
basarisiz2.List = Range(Cells(4, 4), Cells(p, 4)).Value
r = Cells(4, 3).End(xlDown).Row
basarisiz.List = Range(Cells(4, 3), Cells(r, 3)).Value
u = Cells(4, 2).End(xlDown).Row
basarili.List = Range(Cells(4, 2), Cells(u, 2)).Value
n = Cells(4, 5).End(xlDown).Row
kullaniciYok.List = Range(Cells(4, 5), Cells(n, 5)).Value
End Sub


