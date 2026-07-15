Option Explicit
Public Contrato     As String
Public Verificar    As Boolean
Public Vlenc(15)    As Variant
Public bdProd()     As Variant
Public Prod         As Integer
Public LimiteBd     As Long
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
'( Y  )     VALOR BRUTO         19
'( Z  )     VALOR LIQUIDO       20
'( T  )     VALOR CONSIDERADO   21
'( U  )     EMPRESA             22
'( V  )     DESC. CONVENIO      23
'( W  )     OBSERVAÇÃO          24
'( X  )     PAGO EM             25
'( AA )     AGENCIA             26
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
    
     'Limpeza Inicial
    P03.Range("A1:CE20000").ClearContents
    P03.Cells(1, "CA") = "Chave J"
    P03.Cells(1, "CB") = "Cadastro"
    P03.Cells(1, "CC") = "Quant"
    P03.Cells(1, "CD") = "Situação"
    P01.AutoFilterMode = False
    
    'Enquanto estou codificando
    'P_01.Range("A2:Z1000").EntireRow.Delete
    
    'Abertura do Arquivo e obtenção dos dados
    Workbooks.Open Job
    JobTrab = ActiveWorkbook.Name
    L = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    Range(Cells(1, "A"), Cells(L, "AK")).Copy
    
    'Colocando no arquivo temporário para trabalho
    Windows(Home).Activate
    P03.Cells(1, 1).PasteSpecial xlValues
    Application.CutCopyMode = False
    P03.Select
    P03.Cells(1, "AE").Select
    
    'Fechando Job
    Application.DisplayAlerts = False
    Windows(JobTrab).Activate
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Windows(Home).Activate
    
    'Obtendo Chaves J
    L = P03.Cells(P03.Cells.Rows.Count, "A").End(xlUp).Row
    Laux = 2
    Prod = 0
    For I = 2 To L
        Qt = WorksheetFunction.CountIf(Range("CA:CA"), P03.Cells(I, "D"))
        If Qt = 0 Then
            P03.Cells(Laux, "CA") = P03.Cells(I, "D")
            P03.Cells(Laux, "CB").FormulaR1C1 = "=COUNTIF(P04!C[-79],RC[-1])"
            P03.Cells(Laux, "CC").FormulaR1C1 = "=COUNTIF(C[-77],RC[-2])"
            Prod = Prod + 1
            Laux = Laux + 1
          Else
            Prod = Prod + 1
        End If
    Next I
    
    P03.Cells(Laux, "CC") = Prod
    ReDim bdProd(Prod, 26)
    QtchaveJ = Laux - 1
    'Inicio dos trabalhos
    LimiteBd = Prod
    Prod = 0
    NaoEncontrou = 0
    
    For I = 2 To QtchaveJ
        ChaveJ = P03.Cells(I, "CA")
        'Encontrar Nome Promotor e Perfil
        If P03.Cells(I, "CB") > 0 Then
            L = P_05.Cells(P_05.Rows.Count, "A").End(xlUp).Row
            For A = 2 To L
                ChaveProc = P_05.Cells(A, "A")
                If ChaveProc = ChaveJ Then
                    Promotor = P_05.Cells(A, "B")
                    Perfil = P_05.Cells(A, "C")
                    Vlenc(6) = Perfil
                    Exit For
                End If
            Next A
            'Preparar dados obtidos para arquivamento
            Call PreparaDadosArquivamento(Promotor, Perfil, ChaveJ)
          Else
            'Vou precisar ajustar esse ponto
            Promotor = "Cadastrar Promotor " & NaoEncontrou
            Perfil = "BLACK"
            Vlenc(6) = Perfil
            NaoEncontrou = NaoEncontrou + 1
            'Preparar dados obtidos para arquivamento
            Call PreparaDadosArquivamento(Promotor, Perfil, ChaveJ)
        End If
        P03.Select
        P03.Cells(I, "CD").Select
        P03.Cells(I, "CD") = "Lançado"
        P03.Cells(I, "CE") = Prod - WorksheetFunction.Sum(P03.Range(P03.Cells(1, "CE"), P03.Cells(I - 1, "CE")))
        
    Next I
    
    'Arquivar dados
    L = P_01.Cells(P_01.Rows.Count, "A").End(xlUp).Row
    LimiteBd = Prod - 1
    For A = 0 To LimiteBd
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
        '( Y  )     VALOR BRUTO         19
        P_01.Cells(L, "T") = bdProd(A, 19)
        '( Z  )     VALOR LIQUIDO       20
        P_01.Cells(L, "U") = bdProd(A, 20)
        '( T  )     VALOR CONSIDERADO   21
        P_01.Cells(L, "V") = bdProd(A, 21)
        '( U  )     EMPRESA             22
        P_01.Cells(L, "W") = bdProd(A, 22)
        '( V  )     DESC. CONVENIO      23
        P_01.Cells(L, "X") = bdProd(A, 23)
        '( W  )     OBSERVAÇÃO          24
        P_01.Cells(L, "Y") = bdProd(A, 24)
        '( X  )     PAGO EM             25
        P_01.Cells(L, "Z") = bdProd(A, 25)
        '( AA )     AGENCIA             26
        P_01.Cells(L, "AA") = bdProd(A, 26)
    Next A
    
    Call Notificar("Arquivado com Sucesso", "A")
    ActiveWorkbook.Save
    Exit Sub
Erro:
    Call Notificar("Houve um erro durante o Arquivamento do mês", "R")
End Sub

Sub PontosParaCorrecao()
    Verificar = False
    'If Contrato = 211118906 Then Verificar = True
End Sub

Private Sub PreparaDadosArquivamento(Promotor, Perfil, ChaveJ)
    Dim L   As Long, Qt             As Integer, Laux As Long
    Dim nm  As String, Descricao    As String, Vl   As Double, Lcli As Long
    
    L = P03.Cells(Cells.Rows.Count, "A").End(xlUp).Row
    For A = 2 To L
        If P03.Cells(A, "O") <> "" Then
            
            ElseIf P03.Cells(A, "D") = ChaveJ Then

            '( A  )     DATA MOVIMENTO      00
            bdProd(Prod, 0) = CDate(P03.Cells(A, "B"))
            '( B  )     CPF                 01
            bdProd(Prod, 1) = P03.Cells(A, "W")
            '( C  )     BANCO               02
            bdProd(Prod, 2) = "BANCO DO BRASIL"
            '( D  )     CONVENIO            03
            bdProd(Prod, 3) = P03.Cells(A, "G")
            '( E  )     CONTRATO            04
            bdProd(Prod, 4) = P03.Cells(A, "I")
            '( F  )     DATA CONTRATO       05
            bdProd(Prod, 5) = CDate(P03.Cells(A, "H"))
            '( G  )     TAXA                06
            bdProd(Prod, 6) = CDbl(Replace(P03.Cells(A, "Q"), ".", ","))
            '( H  )     PARCELA             07
            bdProd(Prod, 7) = CInt(P03.Cells(A, "J"))
            '( I  )     CHAVE J             08
            bdProd(Prod, 8) = ChaveJ
            '( J  )     COMISSAO_PF         09
            bdProd(Prod, 9) = Empty
            '( K  )     RESTRICAO_RCC       10
            bdProd(Prod, 10) = P03.Cells(A, "AB")
            '( L  )     ANO                 11
            bdProd(Prod, 11) = Year(CDate(P03.Cells(A, "B")))
            '( M  )     MÊS                 12
            bdProd(Prod, 12) = Month(CDate(P03.Cells(A, "B")))
            '( N  )     PROMOTOR            13
            bdProd(Prod, 13) = Promotor
            '( O  )     PRODUTO             14
            bdProd(Prod, 14) = CDbl(P03.Cells(A, "E"))
            'Montar Produto
            Qt = WorksheetFunction.CountIf(P_07.Range("A:A"), P03.Cells(A, "E"))
            If Qt > 0 Then
                Laux = 2
                While P_07.Cells(Laux, "A") <> CDbl(P03.Cells(A, "E"))
                    Laux = Laux + 1
                Wend
                nm = P_07.Cells(Laux, "B")
                Descricao = P_07.Cells(Laux, "C")

                If nm = "CONSIGNADO" Then
                    'VERIFICAR CONVÊNIO
                    Qt = WorksheetFunction.CountIf(P_06.Range("A:A"), CDbl(P03.Cells(A, "G")))
                    If Qt > 0 Then
                        Laux = 2
                        While P_06.Cells(Laux, "A") <> CDbl(P03.Cells(A, "G"))
                            Laux = Laux + 1
                        Wend
                        If P_06.Cells(Laux, "B") <> "INSS" Then
                            nm = nm & " " & P_06.Cells(Laux, "B")
                            Vlenc(4) = nm
                            Contrato = CDbl(P03.Cells(A, "I"))
                            Call ColocarValoresComissao(nm, Descricao, P03.Cells(A, "Q"), CDbl(P03.Cells(A, "J")))
                        Else
                            nm = nm & " " & P_06.Cells(Laux, "B")
                            Vlenc(4) = nm
                            Contrato = CDbl(P03.Cells(A, "I"))
                            Descricao = Replace(Descricao, "CONSIGNADO", "CONSIGNADO INSS")
                            Call ColocarValoresComissao(nm, Descricao, P03.Cells(A, "Q"), CDbl(P03.Cells(A, "J")))
                        End If
                    Else
                        If P03.Cells(A, "T") = "1" Then
                            nm = nm & " " & "PÚBLICO"
                            Vlenc(4) = nm
                            Contrato = CDbl(P03.Cells(A, "I"))
                            Call ColocarValoresComissao(nm, Descricao, P03.Cells(A, "Q"), CDbl(P03.Cells(A, "J")))
                        Else
                            nm = nm & " " & "PRIVADO"
                            Vlenc(4) = nm
                            Contrato = CDbl(P03.Cells(A, "I"))
                            Call ColocarValoresComissao(nm, Descricao, P03.Cells(A, "Q"), CDbl(P03.Cells(A, "J")))
                        End If
                    End If
                ElseIf nm = "NÃO CONSIGNADO" Then
                    Vlenc(4) = nm
                    Contrato = CDbl(P03.Cells(A, "I"))
                    Call ColocarValoresComissao(nm, Descricao, P03.Cells(A, "Q"), CDbl(P03.Cells(A, "J")))
                ElseIf nm = "CRÉDITO ADIANTAMENTO" Then
                    Vlenc(4) = nm
                    Contrato = CDbl(P03.Cells(A, "I"))
                    Call ColocarValoresComissao(nm, Descricao, P03.Cells(A, "Q"), CDbl(P03.Cells(A, "J")))
                Else
                    Stop
                End If
            End If
            '( P  )     COMISSÃO            15
            bdProd(Prod, 15) = Vlenc(5)
            '( Q  )     PERFIL              16
            bdProd(Prod, 16) = Perfil
            '( R  )     VALOR               17
            If P03.Cells(A, "E") = "2887" Then
                Vl = CDbl(Replace(P03.Cells(A, "J"), ".", ","))
                bdProd(Prod, 17) = Vl * Vlenc(5)
            Else
                Vl = CDbl(Replace(P03.Cells(A, "K"), ".", ","))
                bdProd(Prod, 17) = Vl * Vlenc(5)
            End If
            '( S  )     DESCRIÇÃO           18
            bdProd(Prod, 18) = Descricao
            '( Y  )     VALOR BRUTO         19
            bdProd(Prod, 19) = CDbl(Replace(P03.Cells(A, "K"), ".", ","))
            '( Z  )     VALOR LIQUIDO       20
            bdProd(Prod, 20) = CDbl(Replace(P03.Cells(A, "L"), ".", ","))
            '( T  )     VALOR CONSIDERADO   21
            bdProd(Prod, 21) = Vl
            '( U  )     EMPRESA             22
            bdProd(Prod, 22) = Empty
            '( V  )     DESC. CONVENIO      23
            bdProd(Prod, 23) = Vlenc(4)
            '( W  )     OBSERVAÇÃO          24
            bdProd(Prod, 24) = Vlenc(12)
            '( X  )     PAGO EM             25
            bdProd(Prod, 25) = Empty
            '( AA )     AGENCIA             26
            bdProd(Prod, 26) = Empty
            '++++++++++++++++++++++++++++++++++++++++++
            'Verificar se CPF consta no bd_Cliente para inclusão
            Qt = WorksheetFunction.CountIf(P_00.Range("A:A"), P03.Cells(A, "W"))
            If Qt = 0 Then
                Lcli = P_00.Cells(P_00.Rows.Count, "A").End(xlUp).Row + 1
                'CPF
                P_00.Cells(Lcli, "A") = P03.Cells(A, "W")
                'Cliente
                P_00.Cells(Lcli, "B") = P03.Cells(A, "X")
                'Banco
                P_00.Cells(Lcli, "C") = "Banco do Brasil"
            End If
            Prod = Prod + 1
        End If
    Next A
    
End Sub
