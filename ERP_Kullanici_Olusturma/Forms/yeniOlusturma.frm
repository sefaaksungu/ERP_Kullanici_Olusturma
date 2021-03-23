VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} yeniOlusturma 
   Caption         =   "Yeni Kullanici Olusturma"
   ClientHeight    =   6672
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12792
   OleObjectBlob   =   "yeniOlusturma.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "yeniOlusturma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'Sub kullanici_sifre_degisimi()
If mevcutKullanici.Value = "" And msifre = "" And ysifre = "" And yeniKullaniciKodu = "" And yeniKullaniciIsim = "" And yeniKullaniciSoyisim = "" And yeniKullaniciMail = "" And yetkiSablonu = "" Then
    MsgBox "Lütfen seçimleri doldurunuz ve sifreleri en az 8 karakter ve bir büyük harf giriniz."
    
'ElseIf Len(msifre.Value) >= 8 And Len(ysifre.Value) >= 8 Then
Else:
    Dim baglan As New Selenium.WebDriver, elementler As WebElements, uyari As List
    baglan.Start "chrome"
    For i = 13 To 40 Step 2
        If Cells(1, 10).Value = Cells(2, i) Then
            baglan.Get Cells(2, i + 1).Value & "*****"
        End If
    Next i
    
    baglan.Window.Maximize
    
    '
    h = 4
    k = 4
    m = 4
    For i = 13 To 40 Step 2
        If Cells(1, 10).Value = Cells(2, i) Then
            X = Cells(3, 1).End(xlDown).Row
            For j = 3 To X
                baglan.FindElementById("*****").SendKeys mevcutKullanici.Value & Cells(j, 1).Value
                baglan.FindElementById("*****").SendKeys msifre.Value
                baglan.Wait 3000
                
                baglan.FindElementByXPath("(//*[contains(.,'*****')])[last()]").Click
                baglan.Wait 3000
                    
                Set elementler = baglan.FindElementsByXPath("/html/body")
                For Each element In elementler
                metin = element.Text
                If metin Like "*Hata*" Then
                baglan.Wait 3000
                Range("C1") = metin
                End If
                Next element
                
                
                If Cells(1, 3) Like "*Bilgileri Hata*" Then
                    Cells(k, 3).Value = Cells(j, 1).Value
                    k = k + 1
                    GoTo atla
                Else:
                    baglan.Get Cells(2, i + 1).Value & "*****"
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 4000
                    baglan.FindElementById("*****").SendKeys yeniKullaniciKodu.Value & Cells(j, 1).Value 'kullanici kodu girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 2000
                    
                    baglan.ExecuteScript ("$('#*****').val(" & Cells(11, 8) & ").change()") 'kullanici adi girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 2000
                    
                    baglan.ExecuteScript ("$('#*****').val(" & Cells(12, 8) & ").change()") 'kullanici soyadi girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 2000
                    
                    baglan.ExecuteScript ("$('#*****').val(" & Cells(13, 8) & ").change()") 'kullanici mail girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 2000
                    
                    baglan.ExecuteScript ("$('#*****').val(" & Cells(14, 8) & ").change()") 'kullanici sifre girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 2000
                    
                    baglan.ExecuteScript ("$('#*****').val(" & Cells(14, 8) & ").change()") 'kullanici yeniden sifre girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 2000
                    
                    baglan.ExecuteScript ("$('#*****').val(" & Cells(15, 8) & ").change()") 'kullanici bekleme süresi girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 3000
                    
                    baglan.FindElementByXPath("//label[contains(.,'" & Cells(16, 7).Value & "')]").Click
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 2000
                    
                    baglan.FindElementByXPath("(//*[contains(.,'*****')])[last()]").Click 'kaydet
                    
                    Set elementler = baglan.FindElementsByXPath("/html/body")
                    For Each element In elementler
                    metin = element.Text
                    If metin Like "*son 5*" Then
                    baglan.Wait 3000
                    Range("C1") = metin
                    Cells(m, 4).Value = Cells(j, 1).Value
                    m = m + 1
                    baglan.Get Cells(2, i + 1).Value & "*****"
                    GoTo atla
                    End If
                    Next element
                    
                    
                    Cells(h, 2).Value = Cells(j, 1).Value
                    h = h + 1
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 3000
                    baglan.Get Cells(2, i + 1).Value & "*****"
atla:
                End If
                Cells(1, 3).Value = "devam"
                If Cells(j + 1, 1).Value = "END" Then
                    sonuc.Show
                End If
            Next j
        End If
    Next i
End If
End Sub

Private Sub ToggleButton1_Click()
bilgi.Show
End Sub

Private Sub yeniKullaniciIsim_Change()
Cells(11, 7).Value = yeniKullaniciIsim.Value
'tamam
End Sub

Private Sub yeniKullaniciKodu_Change()
Cells(10, 7).Value = yeniKullaniciKodu.Value
'tamam
End Sub

Private Sub yeniKullaniciMail_Change()
Cells(13, 7).Value = yeniKullaniciMail.Value
'tamam
End Sub

Private Sub yeniKullaniciSoyisim_Change()
Cells(12, 7).Value = yeniKullaniciSoyisim.Value
'tamam
End Sub

Private Sub yetkiSablonu_Change()
Cells(16, 7).Value = yetkiSablonu.Value
'tamam
End Sub

Private Sub ysifre_Change()
Cells(14, 7).Value = ysifre.Value
'tamam
'If Len(msifre.Value) < 8 Then
'MsgBox "En az 8 karakter ve bir büyük harf giriniz."
End Sub

