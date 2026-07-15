Option Explicit

Sub ColocarValoresComissao(Nm, Dc, Tx, Pr)
    'On Error GoTo Erro
    Dim Tx1, Tx2, Parc, Parc1, Parc2            As Double
    Dim Np1, Dp1                                As String
    Dim Enc                                     As Boolean
    Dim ProbTx, ProbParc                        As Variant
    Dim TxBusca                                 As String
    Laux = 2
    'Encontrar os Dados para Comissão
    Tx = CDbl(Replace(Tx, ".", ","))
    Parc = Pr
    Laux = 2
    Enc = False
    Vlenc(12) = Empty
    PontosParaCorrecao
    While Enc = False
        Np1 = P01.Cells(Laux, "A")
        Dp1 = P01.Cells(Laux, "B")
        Tx1 = Round(P01.Cells(Laux, "C") * 100, 4)
        Tx2 = Round(P01.Cells(Laux, "D") * 100, 4)
        Parc1 = P01.Cells(Laux, "E")
        Parc2 = P01.Cells(Laux, "F")
        If Verificar = True Then
            Stop
            If Laux = 10 Then Stop
        End If
        If Nm = Np1 And Dc = Dp1 And Tx >= Tx1 And Tx <= Tx2 And Parc >= Parc1 And Parc <= Parc2 Then
            If Trim(UCase(P01.Cells(1, "G"))) = UCase(Vlenc(6)) Then
                Vlenc(5) = P01.Cells(Laux, "G")
              ElseIf Trim(UCase(P01.Cells(1, "H"))) = UCase(Vlenc(6)) Then
                Vlenc(5) = P01.Cells(Laux, "H")
              ElseIf Trim(UCase(P01.Cells(1, "I"))) = UCase(Vlenc(6)) Then
                Vlenc(5) = P01.Cells(Laux, "I")
              ElseIf Trim(UCase(P01.Cells(1, "J"))) = UCase(Vlenc(6)) Then
                Vlenc(5) = P01.Cells(Laux, "J")
            End If
            Enc = True
          Else
            Laux = Laux + 1
        End If
        If P01.Cells(Laux, "I") = "" Then
            Dim Averif  As Long
            Dim QtVerif As Long
            Dim Lverif  As Long
            
            ProbTx = 0
            ProbParc = 0
            Lverif = P01.Cells(P01.Rows.Count, "A").End(xlUp).Row
            For Averif = 2 To Lverif
                If P01.Cells(Averif, "A") = Nm And P01.Cells(Averif, "B") = Dc And P01.Cells(Averif, "C") <= Tx / 100 Then
                    ProbTx = ProbTx + 1
                End If
            Next Averif
            
            For Averif = 2 To Lverif
                If P01.Cells(Averif, "A") = Nm And P01.Cells(Averif, "B") = Dc And P01.Cells(Averif, "E") <= Parc Then
                    ProbParc = ProbParc + 1
                End If
            Next Averif
            
            If ProbTx = 0 Then
                Verificar = True
                Vlenc(12) = "Problema na Taxa"
            End If
            If ProbParc = 0 Then
                Verificar = True
                Vlenc(12) = "Problema na Parcela"
            End If
            Vlenc(5) = 0
            Enc = True
        End If
    Wend
    Exit Sub
Erro:
    Call Notificar("Houve um erro no Fechamento Mensal", "R")
End Sub