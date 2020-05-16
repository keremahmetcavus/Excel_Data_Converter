Attribute VB_Name = "Module1"
Sub CC_U()
    
    
    On Error GoTo 1:
    
    Sheets("DATA").Select ' SHEET� SE�EN KOD
    Set BARKOD_F = Sheets("DATA").Rows("1").Find(What:="BARCODE") 'BARKOD YAZISININ BULAN KOD OLDU�U H�CREY� TANIMLAR.
    BARKOD_F_C = BARKOD_F.Column 'BARKOD YAZISININ BULUNDU�U H�CREN�N COLUMN NUMARASINI DE�ERE ATAR.
    
        If BARKOD_F_C = 1 Then
        Else
            Set BARKOD_F_L = BARKOD_F.Offset(0, -1) 'BARKODUN SOLUNDAK� H�CREY� SE�EN KOD.
            Set BARKOD_F_L_ALL = BARKOD_F.End(xlToLeft)  'BARKODUN EN SOLUNDAK� H�CREY� SE�EN KOD A1 OLARAK _
            DE���T�R�LECEK.
            Set BARKOD_F_L_ALLD = BARKOD_F_L_ALL.End(xlDown) 'BARKODUN EN ALTINDAK� H�CREYE KADAR S�LER.
            ActiveSheet.Range(BARKOD_F_L, BARKOD_F_L_ALLD).Delete Shift:=xlToLeft ' YUKARIDA SE��LEN T�M H�CRE _
            ARALI�INI S�LEN VE YANINDAK� T�M H�CRELER� SOLA KAYDIRAN KOD.
        End If
    
        ThisWorkbook.Sheets.Add.Name = "Calculating..." 'BEL�RT�LEN �S�MDE B�R SHEET A�AR
        Sheets("Calculating...").Select 'BEL�RT�LEN �S�MDEK� SHEET� SE�ER
        Range("A1").Select 'BEL�RT�LEN H�CREY� SE�ER
    
        Selection.Consolidate Sources:= _
            "'DATA'!R1C1:R1000000C10000" _
            , Function:=xlMax, TopRow:=True, LeftColumn:=True, CreateLinks:=False
    
        Sheets("DATA").Select
        Rows("1:1").Select
        Selection.Copy
        Sheets("Calculating...").Select
        Rows("1:1").Select
        ActiveSheet.Paste
            
    'calculatingten Sonuc'a atan kodlar�n hepsi a�a��dad�r
    
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
    
        'calculatingten Sonuc'a atan kodlar�n sonu
        
        Application.DisplayAlerts = False 'S�LERKEN T�M HATA UYARILARINI VERMES�N� KAPATIR
        Sheets("Calculating...").Delete 'SHEET� S�LER.
        Application.DisplayAlerts = True 'S�LERKEN T�M HATA UYARINI VERMES�N� A�AR.
        'YUKARDAK� KODLAR PROGRAM B�T�NCE �ALI�TIRILACAK.
        Sheets("Uniformita Bulucu").Select
        
        'CONICTY'Y� HESAPLAYAN KODUN BA�LANGICI
        '----------------------------------
        '----------------------------------
        '----------------------------------
        '----------------------------------
        '----------------------------------
        '----------------------------------
        
        Sheets("DATA").Select ' SHEET� SE�EN KOD
        Set BARKOD_F = Sheets("DATA").Rows("1").Find(What:="BARCODE") 'BARKOD YAZISININ BULAN KOD OLDU�U H�CREY� TANIMLAR.
        BARKOD_F_C = BARKOD_F.Column 'BARKOD YAZISININ BULUNDU�U H�CREN�N COLUMN NUMARASINI DE�ERE ATAR.
    
        If BARKOD_F_C = 1 Then
        Else
            Set BARKOD_F_L = BARKOD_F.Offset(0, -1) 'BARKODUN SOLUNDAK� H�CREY� SE�EN KOD.
            Set BARKOD_F_L_ALL = BARKOD_F.End(xlToLeft)  'BARKODUN EN SOLUNDAK� H�CREY� SE�EN KOD A1 OLARAK _
            DE���T�R�LECEK.
            Set BARKOD_F_L_ALLD = BARKOD_F_L_ALL.End(xlDown) 'BARKODUN EN ALTINDAK� H�CREYE KADAR S�LER.
            ActiveSheet.Range(BARKOD_F_L, BARKOD_F_L_ALLD).Delete Shift:=xlToLeft ' YUKARIDA SE��LEN T�M H�CRE _
            ARALI�INI S�LEN VE YANINDAK� T�M H�CRELER� SOLA KAYDIRAN KOD.
        End If
    
        ThisWorkbook.Sheets.Add.Name = "Calculating..." 'BEL�RT�LEN �S�MDE B�R SHEET A�AR
        Sheets("Calculating...").Select 'BEL�RT�LEN �S�MDEK� SHEET� SE�ER
        Range("A1").Select 'BEL�RT�LEN H�CREY� SE�ER
    
        Selection.Consolidate Sources:= _
            "'DATA'!R1C1:R1000000C10000" _
            , Function:=xlSum, TopRow:=True, LeftColumn:=True, CreateLinks:=False
    
        Sheets("DATA").Select
        Rows("1:1").Select
        Selection.Copy
        Sheets("Calculating...").Select
        Rows("1:1").Select
        ActiveSheet.Paste
            
        'calculatingten Sonuc'a atan kodlar�n hepsi a�a��dad�r
    
        Set UNFORM_FOUND = Sheets("Calculating...").Rows("1").Find(What:="CONICITY")
        UNFORM_FOUND_C = UNFORM_FOUND.Column
        Columns(UNFORM_FOUND_C).Select
        Selection.Copy
        Sheets("SONUC").Select
        Columns(11).Select
        ActiveSheet.Paste
        Sheets("Calculating...").Select
    
        
        Application.DisplayAlerts = False 'S�LERKEN T�M HATA UYARILARINI VERMES�N� KAPATIR
        Sheets("Calculating...").Delete 'SHEET� S�LER.
        Application.DisplayAlerts = True 'S�LERKEN T�M HATA UYARINI VERMES�N� A�AR.
        'YUKARDAK� KODLAR PROGRAM B�T�NCE �ALI�TIRILACAK.
        Sheets("Uniformita Bulucu").Select
1:
    
End Sub

    'Set RRO_F = Sheets("DATA").Rows("1").Find(What:="RRO") 'RRO YAZISININ BULAN KOD OLDU�U H�CREY� TANIMLAR
    'Set LRO_F = Sheets("DATA").Rows("1").Find(What:="LRO") 'LRO YAZISININ BULAN KOD OLDU�U H�CREY� TANIMLAR _
        Set BARKOD_F_D = BARKOD_F.End(xlDown) ' BARKODUN EN ALTINDAK� H�CREN�N KORD�NATINI ATAR _
        Set BARKOD_F_R = BARKOD_F.End(xlToRight) ' BARKODUN EN SA�INDAK� H�CREN�N KORD�NATINI ATAR _
            BARKOD_F_DC = BARKOD_F_D.Column _
            BARKOD_F_DR = BARKOD_F_D.Row '��MD�L�K LAZIM OLMAYAN B�R KOD _
            BARKOD_F_RC = BARKOD_F_R.Column '��MD�L�K LAZIM OLMAYAN B�R KOD _
            BARKOD_F_RR = BARKOD_F_R.Row '��MD�L�K LAZIM OLMAYAN B�R KOD
    'NOT: YUKARIDAK� KODLARI OLMASI GEREKEN YERE KONUMLANDIR
    'ActiveSheet.Range(BARKOD_F_D, BARKOD_F_R).Select '2 H�CREN�N ARASINDAK� T�M HER�EY� SE�EN KOD
    
    
Sub s�f�rla()

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


