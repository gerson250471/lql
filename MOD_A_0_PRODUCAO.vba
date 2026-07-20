Option Explicit
Public Contrato             As String
Public Verificar            As Boolean
Public Vlenc(15)            As Variant
Public bdProd()             As Variant
Public Prod                 As Integer
Public LimiteBd             As Long
Public LinhaVerificar       As Long
'------------------------------------
'Configuração do Banco de Dados:
'  COL      DESCRIÇÃO           N°
'( A  )     DATA MOVIMENTO      00
'( B  )     CPF                 01
'( C  )     BANCO               02
'( D  )     CONVENIO            03
'( E  )     CONTRATO            04
'( F  )     DATA CONTRATO       05
'( G  )     TAXA                06
'( H  )     PARCELA             07
'( I  )     CHAVE J             08
'( J  )     COMISSAO_PF         09
'( K  )     RESTRICAO_RCC       10
'( L  )     ANO                 11
'( M  )     MÊS                 12
'( N  )     PROMOTOR            13
'( O  )     PRODUTO             14
'( P  )     COMISSÃO            15
'( Q  )     PERFIL              16
'( R  )     VALOR               17
'( S  )     DESCRIÇÃO           18
'( T  )     VALOR BRUTO         19
'( U  )     VALOR LIQUIDO       20
'( V  )     VALOR CONSIDERADO   21
'( W  )     AGENCIA             22
'( X  )     EMPRESA             23
'( Y  )     DESC. CONVENIO      24
'( Z  )     OBSERVAÇÃO          25
'( AA  )    PAGO EM             26
'------------------------------------

Sub ArquivarProducaoMes()
    'On Error GoTo Erro
    Dim Home    As String, Job          As String, Descricao    As String, JobTrab      As String
    Dim L       As Long, Laux           As Long, Lcli           As Long, NaoEncontrou   As Integer
    Dim I       As Long, A              As Long, QtchaveJ       As Long, Perfil         As String
    Dim Qt      As Integer, Conv        As Double, nm           As String, Promotor     As String
    Dim ChaveJ  As String, ChaveProc    As String, QtBd         As Integer, Vl          As Double
    
    'Identificando o Arquivo Home
    Home = ActiveWorkbook.Name
    
    'Obtendo Arquivo para trabalhar
    Job = Application.GetOpenFilename("Excel,*.xlsx", , "Favor informar o Arquivo a ser trabalhado")
    
Voltar:
    
     'Limpeza Inicial
    P_13.Range("A1:CF20000").ClearContents
    P_13.Cells(1, "CA") = "Chave J"
    P_13.Cells(1, "CB") = "Cadastro"
    P_13.Cells(1, "CC") = "Quant"
    P_13.Cells(1, "CD") = "Situação"
    P_13.Cells(1, "CD") = "Valido"
    P_05.AutoFilterMode = False
    
    TabelaEmMemoria
    
    'Enquanto estou codificando
'    P_01.Range("A2:Z1000").EntireRow.Delete
    
    '-trabalho
    Application.DisplayAlerts = False
    Workbooks.Open Job
    JobTrab = ActiveWorkbook.Name
    L = ActiveWorkbook.Sheets(1).Cells(Cells.Rows.Count, "A").End(xlUp).Row
    ActiveWorkbook.Sheets(1).Range("A1:AK" & L).Copy
    Windows(Home).Activate
    P_13.Cells(1, 1).PasteSpecial xlValues
    
    tb_Producao = P_13.Range("A1:AK" & L)
    
    Windows(JobTrab).Activate
    Application.CutCopyMode = False
    ActiveWorkbook.Close
    Windows(Home).Activate
    Application.DisplayAlerts = True
    
    'Obtendo Chaves J
    L = P_13.Cells(P_13.Cells.Rows.Count, "A").End(xlUp).Row
    Laux = 2
    Prod = 0
    For I = 2 To L
        Qt = WorksheetFunction.CountIf(Range("CA:CA"), P_13.Cells(I, "D"))
        If Qt = 0 Then
            P_13.Cells(Laux, "CA") = P_13.Cells(I, "D")
            P_13.Cells(Laux, "CB").FormulaR1C1 = "=COUNTIF(PROMOTOR!C[-79],RC[-1])"
            P_13.Cells(Laux, "CC").FormulaR1C1 = "=COUNTIF(C[-77],RC[-2])"
            Prod = Prod + 1
            Laux = Laux + 1
          Else
            Prod = Prod + 1
        End If
    Next I
    
    P_13.Select
    P_13.Cells(Laux, "CE").Select
    P_13.Cells(Laux, "CE") = Prod
    P_13.Cells(1, "CF") = WorksheetFunction.Sum(P_13.Range("CC:CC"))
    If P_13.Cells(Laux, "CE") = P_13.Cells(1, "CF") Then
        
      Else
        GoTo Voltar
    End If
    
    ReDim bdProd(Prod, 26)
    QtchaveJ = Laux - 1
    'Inicio dos trabalhos
    LimiteBd = Prod
    Prod = 0
    NaoEncontrou = 0
    
    For I = 2 To QtchaveJ
        ChaveJ = P_13.Cells(I, "CA")
        'Encontrar Nome Promotor e Perfil
        
        For A = 1 To UBound(tb_Promotor, 1)
            If tb_Promotor(A, 1) = ChaveJ Then
                Promotor = tb_Promotor(A, 2)
                Perfil = tb_Promotor(A, 3)
                Vlenc(6) = Perfil
                Exit For
              ElseIf A = UBound(tb_Promotor, 1) Then
                Promotor = "Cadastrar Promotor " & NaoEncontrou
                Perfil = "BLACK"
                Vlenc(6) = Perfil
                NaoEncontrou = NaoEncontrou + 1
            End If
        Next A
        'Preparar dados obtidos para arquivamento
        Call PreparaDadosArquivamento(Promotor, Perfil, ChaveJ)
    
        P_13.Select
        P_13.Cells(I, "CD").Select
        P_13.Cells(I, "CD") = "Lançado"
        P_13.Cells(I, "CE") = Prod - WorksheetFunction.Sum(P_13.Range(P_13.Cells(1, "CE"), P_13.Cells(I - 1, "CE")))
        
    Next I

    'Arquivar dados
    L = P_01.Cells(P_01.Rows.Count, "A").End(xlUp).Row
    LimiteBd = Prod - 1
    Stop
    For A = 0 To LimiteBd
        QtBd = WorksheetFunction.CountIf(P_01.Range("E:E"), bdProd(A, 4))
        
        If QtBd = 0 Then
            L = L + 1
            '( A  )     DATA MOVIMENTO      00
            P_01.Cells(L, "A") = bdProd(A, 0)
            '( B  )     CPF                 01
            P_01.Cells(L, "B") = bdProd(A, 1)
            '( C  )     BANCO               02
            P_01.Cells(L, "C") = bdProd(A, 2)
            '( D  )     CONVENIO            03
            P_01.Cells(L, "D") = bdProd(A, 3)
            '( E  )     CONTRATO            04
            P_01.Cells(L, "E") = bdProd(A, 4)
            '( F  )     DATA CONTRATO       05
            P_01.Cells(L, "F") = bdProd(A, 5)
            '( G  )     TAXA                06
            P_01.Cells(L, "G") = bdProd(A, 6)
            '( H  )     PARCELA             07
            P_01.Cells(L, "H") = bdProd(A, 7)
            '( I  )     CHAVE J             08
            P_01.Cells(L, "I") = bdProd(A, 8)
            '( J  )     COMISSAO_PF         09
            P_01.Cells(L, "J") = bdProd(A, 9)
            '( K  )     RESTRICAO_RCC       10
            P_01.Cells(L, "K") = bdProd(A, 10)
            '( L  )     ANO                 11
            P_01.Cells(L, "L") = bdProd(A, 11)
            '( M  )     MÊS                 12
            P_01.Cells(L, "M") = bdProd(A, 12)
            '( N  )     PROMOTOR            13
            P_01.Cells(L, "N") = bdProd(A, 13)
            '( O  )     PRODUTO             14
            P_01.Cells(L, "O") = bdProd(A, 14)
            '( P  )     COMISSÃO            15
            P_01.Cells(L, "P") = bdProd(A, 15)
            '( Q  )     PERFIL              16
            P_01.Cells(L, "Q") = bdProd(A, 16)
            '( R  )     VALOR               17
            P_01.Cells(L, "R") = bdProd(A, 17)
            '( S  )     DESCRIÇÃO           18
            P_01.Cells(L, "S") = bdProd(A, 18)
            '( T  )     VALOR BRUTO         19
            P_01.Cells(L, "T") = bdProd(A, 19)
            '( U  )     VALOR LIQUIDO       20
            P_01.Cells(L, "U") = bdProd(A, 20)
            '( V  )     VALOR CONSIDERADO   21
            P_01.Cells(L, "V") = bdProd(A, 21)
            '( W )     AGENCIA             26
            P_01.Cells(L, "AA") = bdProd(A, 22)
            '( X  )     EMPRESA             22
            P_01.Cells(L, "X") = bdProd(A, 23)
            '( Y  )     DESC. CONVENIO      23
            P_01.Cells(L, "Y") = bdProd(A, 24)
            '( Z  )     OBSERVAÇÃO          24
            P_01.Cells(L, "Z") = bdProd(A, 25)
            '( AA  )     PAGO EM             25
            P_01.Cells(L, "AA") = bdProd(A, 26)
        End If
    Next A
    
    Call Notificar("Arquivado com Sucesso", "A")
    ActiveWorkbook.Save
    Exit Sub
Erro:
    Call Notificar("Houve um erro durante o Arquivamento do mês", "R")
End Sub

Sub PontosParaCorrecao(Produto As String, Dc_Prod As String, taxa As Double, Parcela As Integer)
    Dim Lcorrecao   As Long
    Dim Lbusca      As Long
    Verificar = False
    
    For Lcorrecao = 2 To UBound(tb_Verificacao, 1)
        If tb_Verificacao(Lcorrecao, 1) = CDbl(Contrato) Then
            Verificar = True
            
            For Lbusca = 2 To UBound(tb_Comissao, 1)
                If Lbusca = 110 Then Stop
                If tb_Comissao(Lbusca, 1) = Produto And tb_Comissao(Lbusca, 2) = Dc_Prod And tb_Comissao(Lbusca, 4) > taxa / 100 And tb_Comissao(Lbusca, 6) > Parcela Then
                    LinhaVerificar = Lbusca
                    Exit For
                End If
            Next Lbusca
            Exit For
        End If
    Next Lcorrecao
    
End Sub

Private Sub PreparaDadosArquivamento(Promotor, Perfil, ChaveJ)
    Dim L   As Long, Qt             As Integer, Laux As Long
    Dim nm  As String, Descricao    As String, Vl   As Double, Lcli As Long
    
    For A = 2 To UBound(tb_Producao, 1)
        If tb_Producao(A, 4) = ChaveJ Then
            '( A  )     DATA MOVIMENTO      00
            bdProd(Prod, 0) = CDate(tb_Producao(A, 2))
            '( B  )     CPF                 01
            bdProd(Prod, 1) = tb_Producao(A, 23)
            '( C  )     BANCO               02
            bdProd(Prod, 2) = "BANCO DO BRASIL"
            '( D  )     CONVENIO            03
            bdProd(Prod, 3) = tb_Producao(A, 7)
            '( E  )     CONTRATO            04
            bdProd(Prod, 4) = tb_Producao(A, 9)
            '( F  )     DATA CONTRATO       05
            bdProd(Prod, 5) = CDate(tb_Producao(A, 8))
            '( G  )     TAXA                06
            bdProd(Prod, 6) = CDbl(Replace(tb_Producao(A, 17), ".", ","))
            '( H  )     PARCELA             07
            bdProd(Prod, 7) = CInt(tb_Producao(A, 10))
            '( I  )     CHAVE J             08
            bdProd(Prod, 8) = ChaveJ
            '( J  )     COMISSAO_PF         09
            bdProd(Prod, 9) = Empty
            '( K  )     RESTRICAO_RCC       10
            bdProd(Prod, 10) = tb_Producao(A, 28)
            '( L  )     ANO                 11
            bdProd(Prod, 11) = Year(CDate(tb_Producao(A, 2)))
            '( M  )     MÊS                 12
            bdProd(Prod, 12) = Month(CDate(tb_Producao(A, 2)))
            '( N  )     PROMOTOR            13
            bdProd(Prod, 13) = Promotor
            '( O  )     PRODUTO             14
            bdProd(Prod, 14) = CDbl(tb_Producao(A, 5))
            'Montar Produto
            For L = 2 To UBound(tb_Produto, 1)
                If CDbl(tb_Producao(A, 5)) = tb_Produto(L, 1) Then
                    nm = tb_Produto(L, 2)
                    Descricao = tb_Produto(L, 3)
                    Exit For
                End If
            Next L
            
            'OBTER CONVÊNIO
            If nm = "CONSIGNADO" Then
                'VERIFICAR CONVÊNIO
                Qt = WorksheetFunction.CountIf(P_06.Range("A:A"), CDbl(P_13.Cells(A, "G")))
                If Qt > 0 Then
                    Laux = 2
                    While P_06.Cells(Laux, "A") <> CDbl(P_13.Cells(A, "G"))
                        Laux = Laux + 1
                    Wend
                    If P_06.Cells(Laux, "B") <> "INSS" Then
                        nm = nm & " " & P_06.Cells(Laux, "B")
                        Vlenc(4) = nm
                        Contrato = CDbl(P_13.Cells(A, "I"))
                        Call ColocarValoresComissao(nm, Descricao, Replace(P_13.Cells(A, "Q"), ".", ","), CDbl(P_13.Cells(A, "J")))
                    Else
                        nm = nm & " " & P_06.Cells(Laux, "B")
                        Vlenc(4) = nm
                        Contrato = CDbl(P_13.Cells(A, "I"))
                        Descricao = Replace(Descricao, "CONSIGNADO", "CONSIGNADO INSS")
                        Call ColocarValoresComissao(nm, Descricao, Replace(P_13.Cells(A, "Q"), ".", ","), CDbl(P_13.Cells(A, "J")))
                    End If
                Else
                    If P_13.Cells(A, "T") = "1" Then
                        nm = nm & " " & "PÚBLICO"
                        Vlenc(4) = nm
                        Contrato = CDbl(P_13.Cells(A, "I"))
                        Call ColocarValoresComissao(nm, Descricao, Replace(P_13.Cells(A, "Q"), ".", ","), CDbl(P_13.Cells(A, "J")))
                    Else
                        nm = nm & " " & "PRIVADO"
                        Vlenc(4) = nm
                        Contrato = CDbl(P_13.Cells(A, "I"))
                        Call ColocarValoresComissao(nm, Descricao, Replace(P_13.Cells(A, "Q"), ".", ","), CDbl(P_13.Cells(A, "J")))
                    End If
                End If
              ElseIf nm = "NÃO CONSIGNADO" Then
                    Vlenc(4) = nm
                    Contrato = CDbl(P_13.Cells(A, "I"))
                    Call ColocarValoresComissao(nm, Descricao, Replace(P_13.Cells(A, "Q"), ".", ","), CDbl(P_13.Cells(A, "J")))
              ElseIf nm = "CRÉDITO ADIANTAMENTO" Then
                    Vlenc(4) = nm
                    Contrato = CDbl(P_13.Cells(A, "I"))
                    Call ColocarValoresComissao(nm, Descricao, Replace(P_13.Cells(A, "Q"), ".", ","), CDbl(P_13.Cells(A, "J")))
              ElseIf nm = "CDC FGTS" Then
                    Stop
              ElseIf nm = "PORTABILIDADE" Then
                    Stop
              End If
            '-----------------
            '( P  )     COMISSÃO            15
            bdProd(Prod, 15) = Vlenc(5)
            '( Q  )     PERFIL              16
            bdProd(Prod, 16) = Perfil
            '( R  )     VALOR               17
            Vl = CDbl(Replace(tb_Producao(A, 12), ".", ","))
            bdProd(Prod, 17) = Vl * Vlenc(5)
            '( S  )     DESCRIÇÃO           18
            bdProd(Prod, 18) = Descricao
            '( T  )     VALOR BRUTO         19
            bdProd(Prod, 19) = CDbl(Replace(tb_Producao(A, 11), ".", ","))
            '( U  )     VALOR LIQUIDO       20
            bdProd(Prod, 20) = CDbl(Replace(tb_Producao(A, 12), ".", ","))
            '( V  )     VALOR CONSIDERADO   21
            bdProd(Prod, 21) = Vl
            '( W  )AGENCIA
            bdProd(Prod, 22) = Empty
            '( X  )     EMPRESA             22
            bdProd(Prod, 23) = Empty
            '( Y  )     DESC. CONVENIO      23
            bdProd(Prod, 24) = Vlenc(4)
            '( Z  )     OBSERVAÇÃO          24
            bdProd(Prod, 25) = Vlenc(12)
            '( AA  )     PAGO EM             25
            bdProd(Prod, 26) = Empty
            '++++++++++++++++++++++++++++++++++++++++++
            'Verificar se CPF consta no bd_Cliente para inclusão
            Qt = WorksheetFunction.CountIf(P_00.Range("A:A"), tb_Producao(A, 23))
            If Qt = 0 Then
                Lcli = P_00.Cells(P_00.Rows.Count, "A").End(xlUp).Row + 1
                'CPF
                P_00.Cells(Lcli, "A") = P_13.Cells(A, "W")
                'Cliente
                P_00.Cells(Lcli, "B") = P_13.Cells(A, "X")
                'Banco
                P_00.Cells(Lcli, "C") = "Banco do Brasil"
                'Data de inclusão
                P_00.Cells(Lcli, "D") = Now
            End If
            Prod = Prod + 1
        End If
    Next A
    
End Sub
