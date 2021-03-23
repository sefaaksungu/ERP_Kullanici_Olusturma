VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} kullaniciBilgileri 
   Caption         =   "Kullanici Bilgileri"
   ClientHeight    =   6672
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12792
   OleObjectBlob   =   "kullaniciBilgileri.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "kullaniciBilgileri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub baslat_Click()
'Sub kullanici_sifre_degisimi()
If kullanici.Value = "" And sifre = "" And yeniSifre = "" And kullanici1 = "" Then
    MsgBox "Lütfen Sistemlerden birini seçiniz."
Else:
    Dim baglan As New Selenium.WebDriver, elementler As WebElements, uyari As List, element As WebElement
    baglan.Start "chrome"
    For i = 13 To 40 Step 2
        If Cells(1, 10).Value = Cells(2, i) Then
            baglan.Get Cells(2, i + 1).Value & "login"
        End If
    Next i
    
    baglan.Window.Maximize
    h = 4
    k = 4
    m = 4
    u = 4
    For i = 13 To 40 Step 2
        If Cells(1, 10).Value = Cells(2, i) Then
            X = Cells(3, 1).End(xlDown).Row
            For j = 3 To X
                baglan.FindElementById("*****").SendKeys kullanici.Value & Cells(j, 1).Value
                baglan.FindElementById("*****").SendKeys sifre.Value
                baglan.Wait 3000
                    
                baglan.FindElementByXPath("(//*[contains(.,'Giri')])[last()]").Click
                baglan.Wait 3000
                    
                Set elementler = baglan.FindElementsByXPath("/html/body")
                For Each element In elementler
                metin = element.Text
                If metin Like "*Hata*" Then
                baglan.Wait 3000
                Range("C1") = metin
                End If
                Next element
                    
                
                If Cells(1, 3) Like "*Hata*" Then
                    Cells(k, 3).Value = Cells(j, 1).Value
                    k = k + 1
                    baglan.Wait 2000
                    GoTo atla
                        
                Else:
                    baglan.Get Cells(2, i + 1).Value & "users"
                    baglan.WaitForScript "!jQuery.active"
                    baglan.ExecuteScript ("$('#*****').val(" & Cells(7, 8) & ").change()") 'kullanici adi girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 3000
                    baglan.FindElementByXPath("(//*[contains(.,'*****')])[last()]").Click 'listeleme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 3000
                    
                    Set elementler = baglan.FindElementsByXPath("/html/body")
                    For Each element In elementler
                    metin = element.Text
                    If metin Like "*kriterlere uyan bir*" Then
                    baglan.Wait 3000
                    Range("C1") = metin
                    End If
                    Next element
                
                    If Cells(1, 3).Value Like "*kriterlere uyan bir*" Then
                    Cells(u, 5).Value = Cells(j, 1).Value
                    u = u + 1
                    baglan.Get Cells(2, i + 1).Value & "*****"
                    GoTo atla
                    End If
                
                    baglan.FindElementByXPath("//*[@id='*****']/tbody/tr/td[1]/a").Click 'kullanciya girme
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 3000
                    baglan.FindElementByXPath("(//*[contains(.,'*****')])[last()]").Click 'düzenleme tiklama
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 3000
                    baglan.ExecuteScript ("$('*****').val(" & Cells(8, 8) & ").change()") 'ilk sifre girisi
                    baglan.WaitForScript "!jQuery.active"
                    baglan.ExecuteScript ("$('*****').val(" & Cells(8, 8) & ").change()") 'ikinci sifre girisi
                    baglan.WaitForScript "!jQuery.active"
                    baglan.Wait 4000
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

Private Sub kullanici_Change()
Cells(5, 7).Value = kullanici.Value
End Sub

Private Sub kullanici1_Change()
Cells(7, 7).Value = kullanici1.Value
End Sub

Private Sub sifre_Change()
Cells(6, 8).Value = sifre.Value
End Sub

Private Sub yeniSifre_Change()
Cells(8, 7).Value = yeniSifre.Value
End Sub
