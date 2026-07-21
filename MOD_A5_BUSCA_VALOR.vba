Option Explicit

Sub ColocarValoresComissao(nm As String, Dc As String, Tx As Double, Pr As Integer)
    'On Error GoTo Erro
    Dim Tx1         As Double, Tx2          As Double, Parc As Double, Parc1    As Double, Parc2    As Double
    Dim Np1         As String, Dp1          As String, Laux As Long, L          As Long, C          As Long
    Dim Enc         As Boolean, TxBusca     As String
    Dim ProbTx      As Integer, ProbParc    As Integer
    
    Laux = 2
    'Encontrar os Dados para Comissão
    Tx = CDbl(Replace(Tx, ".", ","))
    Parc = Pr
    Call PontosParaCorrecao(nm, Dc, Tx, Pr)
    For L = 2 To UBound(tb_Comissao, 1)
        Np1 = tb_Comissao(L, 1)
        Dp1 = tb_Comissao(L, 2)
        Tx1 = Round(tb_Comissao(L, 3) * 100, 4)
        Tx2 = Round(tb_Comissao(L, 4) * 100, 4)
        Parc1 = tb_Comissao(L, 5)
        Parc2 = tb_Comissao(L, 6)

        Vlenc(12) = Empty
        
        Np1 = tb_Comissao(L, 1)
        Dp1 = tb_Comissao(L, 2)
        Tx1 = Round(tb_Comissao(L, 3) * 100, 4)
        Tx2 = Round(tb_Comissao(L, 4) * 100, 4)
        Parc1 = tb_Comissao(L, 5)
        Parc2 = tb_Comissao(L, 6)
        If Verificar = True Then
            If L = LinhaVerificar Then
                Stop
                Verificar = False
            End If
        End If
        If nm = Np1 And Dc = Dp1 And Tx >= Tx1 And Tx <= Tx2 And Parc >= Parc1 And Parc <= Parc2 Then
            If Trim(UCase(tb_Comissao(1, 7))) = UCase(Vlenc(6)) Then
                Vlenc(5) = tb_Comissao(L, 7)
                Exit For
              ElseIf Trim(UCase(tb_Comissao(1, 8))) = UCase(Vlenc(6)) Then
                Vlenc(5) = tb_Comissao(L, 8)
                Exit For
              ElseIf Trim(UCase(tb_Comissao(1, 9))) = UCase(Vlenc(6)) Then
                Vlenc(5) = tb_Comissao(L, 9)
                Exit For
              ElseIf Trim(UCase(tb_Comissao(1, 10))) = UCase(Vlenc(6)) Then
                Vlenc(5) = tb_Comissao(L, 10)
                Exit For
            End If
            If tb_Comissao(L, 1) = nm And tb_Comissao(L, 2) = Dc And tb_Comissao(L, 3) <= Tx / 100 Then ProbTx = ProbTx + 1
            If tb_Comissao(L, 1) = nm And tb_Comissao(L, 2) = Dc And tb_Comissao(L, 5) <= Parc Then ProbParc = ProbParc + 1
            Exit For
          ElseIf L = UBound(tb_Comissao, 1) Then
            If ProbTx = 0 Then
                Verificar = True
                Vlenc(12) = "Abaixo da Taxa Minima"
            End If
            If ProbParc = 0 Then
                Verificar = True
                Vlenc(12) = "Fora do Prazo"
            End If
            Vlenc(5) = 0
          Else
            If tb_Comissao(L, 1) = nm And tb_Comissao(L, 2) = Dc And tb_Comissao(L, 3) <= Tx / 100 Then ProbTx = ProbTx + 1
            If tb_Comissao(L, 1) = nm And tb_Comissao(L, 2) = Dc And tb_Comissao(L, 5) <= Parc Then ProbParc = ProbParc + 1
        End If
    Next L
    
    Exit Sub
Erro:
    Call Notificar("Houve um erro no Fechamento Mensal", "R")
End Sub

