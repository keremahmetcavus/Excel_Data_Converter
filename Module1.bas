Attribute VB_Name = "Module1"
Sub CC_U()
    
    
    On Error GoTo 1:
    
    Sheets("DATA").Select ' SHEETÝ SEÇEN KOD
    Set BARKOD_F = Sheets("DATA").Rows("1").Find(What:="BARCODE") 'BARKOD YAZISININ BULAN KOD OLDUÐU HÜCREYÝ TANIMLAR.
    BARKOD_F_C = BARKOD_F.Column 'BARKOD YAZISININ BULUNDUÐU HÜCRENÝN COLUMN NUMARASINI DEÐERE ATAR.
    
        If BARKOD_F_C = 1 Then
        Else
            Set BARKOD_F_L = BARKOD_F.Offset(0, -1) 'BARKODUN SOLUNDAKÝ HÜCREYÝ SEÇEN KOD.
            Set BARKOD_F_L_ALL = BARKOD_F.End(xlToLeft)  'BARKODUN EN SOLUNDAKÝ HÜCREYÝ SEÇEN KOD A1 OLARAK _
            DEÐÝÞTÝRÝLECEK.
            Set BARKOD_F_L_ALLD = BARKOD_F_L_ALL.End(xlDown) 'BARKODUN EN ALTINDAKÝ HÜCREYE KADAR SÝLER.
            ActiveSheet.Range(BARKOD_F_L, BARKOD_F_L_ALLD).Delete Shift:=xlToLeft ' YUKARIDA SEÇÝLEN TÜM HÜCRE _
            ARALIÐINI SÝLEN VE YANINDAKÝ TÜM HÜCRELERÝ SOLA KAYDIRAN KOD.
        End If
    
        ThisWorkbook.Sheets.Add.Name = "Calculating..." 'BELÝRTÝLEN ÝSÝMDE BÝR SHEET AÇAR
        Sheets("Calculating...").Select 'BELÝRTÝLEN ÝSÝMDEKÝ SHEETÝ SEÇER
        Range("A1").Select 'BELÝRTÝLEN HÜCREYÝ SEÇER
    
        Selection.Consolidate Sources:= _
            "'DATA'!R1C1:R1000000C10000" _
            , Function:=xlMax, TopRow:=True, LeftColumn:=True, CreateLinks:=False
    
        Sheets("DATA").Select
        Rows("1:1").Select
        Selection.Copy
        Sheets("Calculating...").Select
        Rows("1:1").Select
        ActiveSheet.Paste
            
    'calculatingten Sonuc'a atan kodlarýn hepsi aþaðýdadýr
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="BARCODE")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(1).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="RFV")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
    Selection.Copy
    Sheets("SONUC").Select
    Columns(8).Select
    ActiveSheet.Paste
    Sheets("Calculating...").Select
    
    Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="H1RFV")
    UNFORM_FOUND_C = UNFORM_FOUND.Column
    Columns(UNFORM_FOUND_C).Select
    Selection.Copy
    Sheets("SONUC").Select
    Columns(9).Select
    ActiveSheet.Paste
    Sheets("Calculating...").Select
    
    Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="LFV")
    UNFORM_FOUND_C = UNFORM_FOUND.Column
    Columns(UNFORM_FOUND_C).Select
    Selection.Copy
    Sheets("SONUC").Select
    Columns(10).Select
    ActiveSheet.Paste
    Sheets("Calculating...").Select
    
    Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="CONICITY")
    UNFORM_FOUND_C = UNFORM_FOUND.Column
    Columns(UNFORM_FOUND_C).Select
    Selection.Copy
    Sheets("SONUC").Select
    Columns(11).Select
    ActiveSheet.Paste
    Sheets("Calculating...").Select
    
    Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="PLY")
    UNFORM_FOUND_C = UNFORM_FOUND.Column
    Columns(UNFORM_FOUND_C).Select
    Selection.Copy
    Sheets("SONUC").Select
    Columns(12).Select
     ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="RRO")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(2).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="LRO")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(3).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="DEPRES")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(4).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="WOBBLE")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(5).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="UNB_DY_LO")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(6).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="UNB_DY_UP")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(7).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        'calculatingten Sonuc'a atan kodlarýn sonu
        
        Application.DisplayAlerts = False 'SÝLERKEN TÜM HATA UYARILARINI VERMESÝNÝ KAPATIR
        Sheets("Calculating...").Delete 'SHEETÝ SÝLER.
        Application.DisplayAlerts = True 'SÝLERKEN TÜM HATA UYARINI VERMESÝNÝ AÇAR.
        'YUKARDAKÝ KODLAR PROGRAM BÝTÝNCE ÇALIÞTIRILACAK.
        Sheets("Uniformita Bulucu").Select
        
        'CONICTY'YÝ HESAPLAYAN KODUN BAÞLANGICI
        '----------------------------------
        '----------------------------------
        '----------------------------------
        '----------------------------------
        '----------------------------------
        '----------------------------------
        
        Sheets("DATA").Select ' SHEETÝ SEÇEN KOD
        Set BARKOD_F = Sheets("DATA").Rows("1").Find(What:="BARCODE") 'BARKOD YAZISININ BULAN KOD OLDUÐU HÜCREYÝ TANIMLAR.
        BARKOD_F_C = BARKOD_F.Column 'BARKOD YAZISININ BULUNDUÐU HÜCRENÝN COLUMN NUMARASINI DEÐERE ATAR.
    
        If BARKOD_F_C = 1 Then
        Else
            Set BARKOD_F_L = BARKOD_F.Offset(0, -1) 'BARKODUN SOLUNDAKÝ HÜCREYÝ SEÇEN KOD.
            Set BARKOD_F_L_ALL = BARKOD_F.End(xlToLeft)  'BARKODUN EN SOLUNDAKÝ HÜCREYÝ SEÇEN KOD A1 OLARAK _
            DEÐÝÞTÝRÝLECEK.
            Set BARKOD_F_L_ALLD = BARKOD_F_L_ALL.End(xlDown) 'BARKODUN EN ALTINDAKÝ HÜCREYE KADAR SÝLER.
            ActiveSheet.Range(BARKOD_F_L, BARKOD_F_L_ALLD).Delete Shift:=xlToLeft ' YUKARIDA SEÇÝLEN TÜM HÜCRE _
            ARALIÐINI SÝLEN VE YANINDAKÝ TÜM HÜCRELERÝ SOLA KAYDIRAN KOD.
        End If
    
        ThisWorkbook.Sheets.Add.Name = "Calculating..." 'BELÝRTÝLEN ÝSÝMDE BÝR SHEET AÇAR
        Sheets("Calculating...").Select 'BELÝRTÝLEN ÝSÝMDEKÝ SHEETÝ SEÇER
        Range("A1").Select 'BELÝRTÝLEN HÜCREYÝ SEÇER
    
        Selection.Consolidate Sources:= _
            "'DATA'!R1C1:R1000000C10000" _
            , Function:=xlSum, TopRow:=True, LeftColumn:=True, CreateLinks:=False
    
        Sheets("DATA").Select
        Rows("1:1").Select
        Selection.Copy
        Sheets("Calculating...").Select
        Rows("1:1").Select
        ActiveSheet.Paste
            
        'calculatingten Sonuc'a atan kodlarýn hepsi aþaðýdadýr
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="CONICITY")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(11).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        
        Application.DisplayAlerts = False 'SÝLERKEN TÜM HATA UYARILARINI VERMESÝNÝ KAPATIR
        Sheets("Calculating...").Delete 'SHEETÝ SÝLER.
        Application.DisplayAlerts = True 'SÝLERKEN TÜM HATA UYARINI VERMESÝNÝ AÇAR.
        'YUKARDAKÝ KODLAR PROGRAM BÝTÝNCE ÇALIÞTIRILACAK.
        Sheets("Uniformita Bulucu").Select
1:
    
End Sub

    'Set RRO_F = Sheets("DATA").Rows("1").Find(What:="RRO") 'RRO YAZISININ BULAN KOD OLDUÐU HÜCREYÝ TANIMLAR
    'Set LRO_F = Sheets("DATA").Rows("1").Find(What:="LRO") 'LRO YAZISININ BULAN KOD OLDUÐU HÜCREYÝ TANIMLAR _
        Set BARKOD_F_D = BARKOD_F.End(xlDown) ' BARKODUN EN ALTINDAKÝ HÜCRENÝN KORDÝNATINI ATAR _
        Set BARKOD_F_R = BARKOD_F.End(xlToRight) ' BARKODUN EN SAÐINDAKÝ HÜCRENÝN KORDÝNATINI ATAR _
            BARKOD_F_DC = BARKOD_F_D.Column _
            BARKOD_F_DR = BARKOD_F_D.Row 'ÞÝMDÝLÝK LAZIM OLMAYAN BÝR KOD _
            BARKOD_F_RC = BARKOD_F_R.Column 'ÞÝMDÝLÝK LAZIM OLMAYAN BÝR KOD _
            BARKOD_F_RR = BARKOD_F_R.Row 'ÞÝMDÝLÝK LAZIM OLMAYAN BÝR KOD
    'NOT: YUKARIDAKÝ KODLARI OLMASI GEREKEN YERE KONUMLANDIR
    'ActiveSheet.Range(BARKOD_F_D, BARKOD_F_R).Select '2 HÜCRENÝN ARASINDAKÝ TÜM HERÞEYÝ SEÇEN KOD
    
    
Sub sýfýrla()

    Sheets("DATA").Select
    Range("A1:ZZ100000").Select
    Selection.ClearContents
    Sheets("SONUC").Select
    Range("A2:ZZ100000").Select
    Selection.ClearContents
    Sheets("Uniformita Bulucu").Select
    'ActiveSheet.Unprotect
    Range("A2:A201").Select
    Selection.ClearContents
    Sheets("Kontrol Paneli").Select
    
End Sub


