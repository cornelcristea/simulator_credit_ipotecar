Attribute VB_Name = "salveazaPDF"
Sub SalveazaPDF_Buton()
    Dim numeFisier As String
    Dim dataCurenta As Date
    Dim dataCod As String
    Dim caleFisier As Variant
    Dim raspuns As VbMsgBoxResult
    Dim oraCurenta As Date
    Dim salvareSheet As Worksheet
    
    oraCurenta = Time
    dataCurenta = Date
    dataCod = Format(dataCurenta, "ddMMyy")
    oraCod = Format(oraCurenta, "hhmm")
    numeFisier = "scadentar_" & dataCod & oraCod & ".pdf"
    
    Set salvareSheet = ActiveSheet
    caleFisier = Application.GetSaveAsFilename(numeFisier, "PDF Files (*.pdf), *.pdf")
    
    If caleFisier <> "False" Then ' verifica daca salvarea nu a fost oprita
        salvareSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caleFisier, Quality:=xlQualityStandard
        raspuns = MsgBox("Fisierul PDF a fost salvat cu succes!" & vbNewLine & "Doriti sa deschideti documentul?", vbYesNo + vbQuestion, "Confirmare")
        If raspuns = vbYes Then
            ThisWorkbook.FollowHyperlink caleFisier
        End If
    End If
End Sub



