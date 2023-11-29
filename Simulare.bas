Attribute VB_Name = "Simulare"
Sub Simulare_Buton()

    '========================
    ' variabile de lucru
    '========================
    Dim valoareImprumut As Double
    Dim perioadaImprumut As Double
    Dim perioadaRataFix As Integer
    Dim perioadaRataVar As Integer
    Dim dobandaFix As Double
    Dim dobandaVar As Double
    Dim rataLunarFix As Double
    Dim rataLunarVar As Double
    Dim rataLunarDAE As Double
    Dim soldInitial As Double
    Dim soldRamas As Double
    Dim principal As Double
    Dim rata As Double
    Dim dobanda As Double
    Dim tipRata As String
    Dim dataCurenta As Date
    Dim plataAnticipata As Double
    Dim data As Date
    Dim dobandaTotala As Double
    Dim iLuna As Integer
    Dim valoareImobil As Double
    Dim valoareAvans As Double
    Dim venitLunar As Double
    Dim gradIndatorare As Double
    Dim asigViata As Double
    Dim salariuBanca As Double
    Dim imobilEco As Double
    Dim procentAvans As Double
    Dim perioadaAnticipata As Integer
    
    
    ' date de intrare de la utilizator
    venitLunar = range("C3").Value
    valoareImobil = range("C4").Value
    valoareAvans = range("C5").Value
    perioadaImprumut = range("B6").Value
    asigViata = range("C7").Value
    salariuBanca = range("C8").Value
    imobilEco = range("C9").Value
    tipRata = range("B10").Value
    perioadaRataFix = range("B11").Value
    perioadaRataVar = range("B12").Value
    dobandaFix = range("C11").Value
    dobandaVar = range("C12").Value
    dobandaDAE = range("B13").Value
    
    
    ' date calculate
    procentAvans = valoareAvans / valoareImobil
    dataCurenta = Date 'range("F2").Value
    valoareImprumut = valoareImobil - valoareAvans 'range("G3").Value
    rataLunarFix = Pmt(dobandaFix / 12, perioadaImprumut * 12, -valoareImprumut) 'range("G4").Value
    rataLunarVar = Pmt(dobandaVar / 12, perioadaImprumut * 12, -valoareImprumut) 'range("G5").Value
    rataLunarDAE = Pmt(dobandaDAE / 12, perioadaImprumut * 12, -valoareImprumut) 'range("G6").Value
    gradIndatorare = rataLunarDAE / venitLunar 'range("F9").Value
    dobandaTotala = 0
    perioadaAnticipata = 0
    
    '========================
    ' generare scadentar
    '========================
    formatareTabel
    
    ' verificare parametrii
    'verificareParametru perioadaImprumut, 30, 2, "Perioada unui credit ipotecar trebuie sa fie maxim 30 de ani", "A6:B6:C6"
    verificareParametru procentAvans, 0.15, 1, "Avansul dumneavoastra este insuficient." & vbNewLine & "Acesta trebuie sa fie minim 15% din valoarea imobilului.", "A5:B5:C5"
    verificareParametru gradIndatorare, 0.25, 2, "Gradul de indatorare este prea mare." & vbNewLine & "Rata lunara trebuie sa fie maxim 25% din venitul lunar.", "E9:F9:G9"
           
    ' completare date pentru fiecare luna
    For iLuna = 0 To (perioadaImprumut * 12)
        If iLuna <> 0 Then
            data = DateSerial(Year(dataCurenta), Month(datCurenta) + iLuna, Day(dataCurenta))
            plataAnticipata = range("G" & 16 + iLuna - 1).Value

            ' verificare plata anticipata
            If plataAnticipata <> 0 Then
                Dim jLuna As Integer
                Dim perioadaAvans As Integer

                soldRamas = soldRamas - plataAnticipata
                perioadaAvans = Abs(plataAnticipata / principal)
                perioadaAnticipata = perioadaAnticipata + perioadaAvans

                For jLuna = 1 To perioadaAvans
                    data = DateSerial(Year(dataCurenta), Month(datCurenta) + iLuna, Day(dataCurenta))
                    completareDate data, iLuna, 0, 0, 0, 0
                    Rows(16 + iLuna).EntireRow.Hidden = True

                    If perioadaAvans = 1 Or jLuna = perioadaAvans Then
                        Exit For
                    Else
                        iLuna = iLuna + 1
                    End If
                Next jLuna
            ' verificare tip de rata
            Else
                Select Case tipRata
                    Case "Rate egale" ' principal crescator si dobanda descrescatoare
                        If iLuna <= (perioadaRataFix * 12) Then ' perioada rata fixa
                            rata = rataLunarFix
                            dobanda = (dobandaFix / 12) * soldRamas
                        Else ' perioada rata variabila
                            rata = rataLunarVar
                            dobanda = (dobandaVar / 12) * soldRamas
                        End If
                        principal = rata - dobanda

                    Case "Rate descrescatoare" ' principal constant si dobanda descrescatoare (cel mai avantajos)
                        principal = valoareImprumut / (perioadaImprumut * 12)
                        If iLuna <= (perioadaRataFix * 12) Then ' perioada rata fixa
                            dobanda = (dobandaFix / 12) * soldRamas
                        Else ' perioada rata variabila
                            dobanda = (dobandaVar / 12) * soldRamas
                        End If
                        rata = principal + dobanda
                End Select

                soldRamas = soldRamas - principal

                ' calcul dobanda totala
                dobandaTotala = dobandaTotala + dobanda

                ' completare date in tabel
                If soldRamas >= 0 Then
                    completareDate data, iLuna, principal, dobanda, rata, soldRamas
                Else
                    completareDate data, iLuna, principal, dobanda, rata, 0
                    Exit For
                End If
            End If
        Else
            soldRamas = valoareImprumut
            completareDate dataCurenta, iLuna, 0, 0, 0, soldRamas
        End If
    Next iLuna
    
    ' perioada finala cu plata anticipata
    If perioadaAnticipata <> 0 Then
        Dim perioadaFinala As Integer
        perioadaFinala = (perioadaImprumut * 12) - perioadaAnticipata
        range("F10").Value = perioadaFinala / 12
        range("G10").Value = perioadaFinala
    Else
        range("F10").Value = perioadaImprumut
        range("G10").Value = perioadaImprumut * 12
    End If
    
    range("G7").Value = Round(dobandaTotala, 0)
    range("G8").Value = valoareImprumut + dobandaTotala ' total de rambursat


End Sub


Function formatareTabel()
    'Dim raspuns As VbMsgBoxResult
    
    'raspuns = MsgBox("Resetati platile anticipate?", vbYesNo + vbQuestion, "Confimare")
    'If raspuns = vbYes Then
        'range("G18:G" & 18 + 360).ClearContents
    'End If
    
    ' stergere date vechi
    range("A16:F" & 16 + 360).ClearContents
    
    ' afisare randuri ascunse
    ActiveSheet.Rows.Hidden = False
    Rows(16 + 360).RowHeight = 11.25
 
    ' formatare text
    range("A16:A" & 18 + 360).NumberFormat = "dd/mm/yyyy"
    range("C16:G" & 18 + 360).NumberFormat = "#,##0.00 [$RON]"

    ' centrare text
    range("A15:G" & 15 + 360).HorizontalAlignment = xlCenter
    
    ' adaugare borduri
    'Range("A17:G" & 17 + number * 12).BorderAround (1)

    ' cap de tabel
    range("A15").Value = "Data platii"
    range("B15").Value = "Luna"
    range("C15").Value = "Principal"
    range("D15").Value = "Dobanda"
    range("E15").Value = "Rata lunara"
    range("F15").Value = "Sold ramas"
    range("G15").Value = "Plata anticipata"

End Function


Function completareDate(data As Date, luna As Integer, principal As Double, dobanda As Double, rata As Double, soldRamas As Double)
    range("A" & 16 + luna).Value = data
    range("B" & 16 + luna).Value = luna
    range("C" & 16 + luna).Value = principal
    range("D" & 16 + luna).Value = dobanda
    range("E" & 16 + luna).Value = rata
    range("F" & 16 + luna).Value = soldRamas
End Function


Function verificareParametru(param As Double, ref As Double, operator As Integer, mesaj As String, celule As String)
    Dim eroarePrezenta As Boolean
    eroare = False
    
    ' operator: 1 este mai mic, 2 este mai mare, 3 este egal
    Select Case operator
        Case 1
            If param < ref Then
                eroare = True
            End If
        Case 2
            If param > ref Then
                eroare = True
            End If
        Case 3
            If param = ref Then
                eroare = True
            End If
    End Select
    
    If eroare Then
        range(celule).Font.Color = RGB(190, 0, 0)
        MsgBox mesaj, vbOKOnly, "Avertisment"
    Else
        range(celule).Font.Color = RGB(0, 0, 0)
    End If
End Function





