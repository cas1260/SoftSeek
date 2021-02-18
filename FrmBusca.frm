VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmBusca 
   Caption         =   "SoftSeek 1.0"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   ForeColor       =   &H00800000&
   Icon            =   "FrmBusca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   12660
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar Barra 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   7635
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      ForeColor       =   &H00800000&
      Height          =   6315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2835
      Begin VB.CommandButton CmdInic 
         Caption         =   "Iniciar Pesquisa"
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   6000
         Width           =   2685
      End
      Begin VB.ListBox LstUrl 
         Appearance      =   0  'Flat
         Columns         =   3
         Height          =   5430
         ItemData        =   "FrmBusca.frx":08CA
         Left            =   60
         List            =   "FrmBusca.frx":08E0
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   540
         Width           =   2685
      End
      Begin VB.TextBox TxtBusca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   60
         TabIndex        =   0
         Top             =   210
         Width           =   2685
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Buscar"
      ForeColor       =   &H00800000&
      Height          =   6315
      Left            =   2850
      TabIndex        =   3
      Top             =   0
      Width           =   9795
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   6045
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   10663
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FormatString    =   $"FrmBusca.frx":0964
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin VB.Label LblStatus 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   30
      TabIndex        =   8
      Top             =   6360
      Width           =   12555
   End
   Begin VB.Label LblUrl 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   30
      TabIndex        =   7
      Top             =   6960
      Width           =   12585
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Url Atual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   6
      Top             =   6720
      Width           =   12555
   End
End
Attribute VB_Name = "FrmBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UrlLida() As String
Dim TipoBanco As String
Dim TotalPagina As Long
Dim TotalPaginaValida As Long
Dim TotalPalavra As Long
Dim UrlOk()  As String
Dim Nivel As Long
Dim CancelaAcao As Boolean

Private Sub InicioBusca()
    
    Dim UrlBase() As String
    Dim TBanco() As String
    Dim Texto As String
    
    ReDim UrlBase(10) As String
    ReDim TBanco(10) As String
    
        
    'Grid.Clear
    'Grid.Rows = 2
    'Grid.FormatString = "Codigo               |Url de Acesso                                                                                                                                                                             |Data / Hora                         "
    
    Barra.Min = 0
    Barra.Value = 0
    Barra.Max = 1
    
    CancelaAcao = False
    
    TxtBusca.Enabled = False
    
'    ReDim UrlOk(0) As String
    
    LblStatus.Caption = "Inicio de Busca, Aguardade um momento..."
    
    DoEvents
    
    'UrlBase(0) = "https://www.licitacoes-e.com.br/aop/lct/licitacao/consulta/ListarQtdLctcCprd.jsp?situacao-licitacao=4"
    UrlBase(0) = "https://www.licitacoes-e.com.br/aop/lct/licitacao/consulta/ListarQtdLctcCprd.jsp?situacaoLicitacao=4"
    TBanco(0) = "BB"
    
    UrlBase(1) = "https://wwws.sistemas.mg.gov.br/cotacao/do/getProcessoLista?opcaoConsulta=3&ano=2008"
                 'https://wwws.sistemas.mg.gov.br/cotacao/do/getProcessoLista
    TBanco(1) = "CMG" 'compras.mg.gov.br
    
    UrlBase(2) = "https://wwws.licitanet.mg.gov.br/licita_futuras.asp?ano=2007&descri=&OK=OK&orgao=0"
    TBanco(2) = "LCF"
    
    UrlBase(3) = "https://wwws.licitanet.mg.gov.br/licita_acolhe.asp?ano=2007&descri=&OK=OK&orgao=0"
    TBanco(3) = "LCA"
    
    UrlBase(4) = "http://www.bec.sp.gov.br/publico/aspx/Home.aspx"
    TBanco(4) = "BEC"
    
    UrlBase(5) = "http://www.conlicitacao.com.br/bd/pagina3.php?chave=&nome=&editaltipo=dsntrl&edital=&endereco=&cidade=&agrup%5B%5D=dsntrl&objeto=&estado=dsntrl&vigente=S&datahora_val_de=dd%2Fmm%2Faaaa&datahora_val_ate=dd%2Fmm%2Faaaa&datahora_inc_de=dd%2Fmm%2Faaaa&datahora_inc_ate=dd%2Fmm%2Faaaa"
    TBanco(5) = "conlicitacao"
    
    
    DoEvents
    
    TotalPagina = 0
    TotalPaginaValida = 0
    
    For x = 0 To LstUrl.ListCount - 1
        
        If LstUrl.Selected(x) = True Then
            TipoBanco = TBanco(x)
            If TipoBanco = "CMG" Then
                Call ReadUrl("https://wwws.sistemas.mg.gov.br/cotacao/do/setUpProcessoConsultaForm", "POST")
                             'https://wwws.sistemas.mg.gov.br/cotacao/do/setUpProcessoConsultaForm
                            'https://wwws.sistemas.mg.gov.br/cotacao/do/setUpProcessoConsultaForm
            ElseIf TipoBanco = "conlicitacao" Then
                Texto = ReadUrl("http://www.conlicitacao.com.br/bd/login_novo.php?login=grupomitra&senha=grupomitra", "POST")
            End If
            Texto = ReadUrl(UrlBase(x), "POST")
            BuscaSubPagina Texto, UrlBase(x)
        End If
        DoEvents
    Next
    
    If CancelaAcao = True Then
        MsgBox "Operação cancelada!", vbCritical, "Atenção"
    Else
        MsgBox "Busca Complenta", vbInformation, "OK"
    End If

    TotalPagina = 0
    TotalPaginaValida = 0
    TotalPalavra = 0

    TxtBusca.Enabled = True
End Sub

Private Sub BuscaSubPagina(Texto As String, UrlP As String)
    Dim Lngx As Long, lngY                        As Long
    Dim xlngX As Long, xlngY                      As Long
    Dim TextoAux As String, UrlTemp               As String
    Dim xControle As Long, NovoVetor              As Boolean
    Dim TextoPesq                                 As String
    Dim TextoASerPesq()                           As String
    Dim y                                         As Long
    Dim Codigo As String, Sql                     As String
    Dim lngPesq As Long, lngPesq1 As Long, LngXXX As Long
    Dim strDetalhe                                As String

    If CancelaAcao = True Then Exit Sub

    TextoAux = LCase(Texto)

    TextoASerPesq = Split(TxtBusca.Text, ",")

    LblUrl.Caption = UrlP

    For x = 0 To UBound(TextoASerPesq)
        If InStr(x + 1, TextoAux, LCase(TextoASerPesq(x))) > 0 Then
            NovoVetor = True
            For y = 0 To UBound(UrlOk)
                If Trim(LCase(UrlOk(xControle))) = Trim(LCase(UrlP)) Then
                    NovoVetor = False
                    Exit For
                End If
                DoEvents
            Next
            If NovoVetor = True Then
                If TipoBanco = "conlicitacao" Then

                    MsgBox "teste"
                ElseIf TipoBanco = "BB" Then
                    If InStr(1, LCase(UrlP), LCase("ExibirLicitacao")) > 0 Or InStr(1, LCase(UrlP), LCase("consultar-detalhes-licitacao")) > 0 Then

                        lngPesq = InStr(1, UrlP, "id-licitacao")
                        If lngPesq = 0 Then
                            lngPesq = InStr(1, UrlP, "numeroLicitacao")

                            lngPesq1 = InStr(lngPesq, UrlP, "&")

                            If lngPesq1 > 0 Then
                                Codigo = Mid(UrlP, lngPesq + 16, lngPesq1 - lngPesq - 16)
                            Else
                                Codigo = Mid(UrlP, lngPesq + 13, Len(UrlP) - lngPesq)
                            End If
                        Else
                            lngPesq1 = InStr(lngPesq, UrlP, "&")

                            If lngPesq1 > 0 Then
                                Codigo = Mid(UrlP, lngPesq, lngPesq1 - lngPesq - 2)
                            Else
                                Codigo = Mid(UrlP, lngPesq + 13, Len(UrlP) - lngPesq)
                            End If
                        End If

                        For LngXXX = 0 To Grid.Rows - 1
                            If Trim(Grid.TextMatrix(LngXXX, 0)) = Codigo Then
                                GoTo PulaReg
                            End If
                        Next

                        ReDim Preserve UrlOk(UBound(UrlOk) + 1)
                        UrlOk(UBound(UrlOk)) = UrlP
                        If Grid.TextMatrix(1, 0) <> "" Then
                            Grid.Rows = Grid.Rows + 1
                        End If

                        Grid.TextMatrix(Grid.Rows - 1, 0) = Codigo
                        Grid.TextMatrix(Grid.Rows - 1, 1) = UrlP
                        Grid.TextMatrix(Grid.Rows - 1, 2) = Now()

                        Sql = "insert into tbllogs (Url, Codigo, Origem, DataHora) Values ('" & UrlP & "', '" & Codigo & "', '" & TipoBanco & "', '" & Now() & "')"
                        Cn.Execute Sql

PulaReg:
                    End If
                    TotalPalavra = TotalPalavra + 1
                ElseIf TipoBanco = "CMG" Then
                    'GoTo PulaReg2

                    If InStr(1, UrlP, "processo=") > 0 Then

                        lngPesq = InStr(1, UrlP, "processo=")
                        lngPesq1 = InStr(lngPesq, UrlP, "&")

                        If lngPesq1 > 0 Then
                            Codigo = Mid(UrlP, lngPesq + 9, lngPesq1 - lngPesq - 9)
                        Else
                            Codigo = Mid(UrlP, lngPesq + 9, Len(UrlP) - lngPesq)
                        End If

                        For LngXXX = 0 To Grid.Rows - 1
                            If Trim(Grid.TextMatrix(LngXXX, 0)) = Codigo Then
                                GoTo PulaReg2
                            End If
                        Next

                        ReDim Preserve UrlOk(UBound(UrlOk) + 1)
                        UrlOk(UBound(UrlOk)) = UrlP
                        If Grid.TextMatrix(1, 0) <> "" Then
                            Grid.Rows = Grid.Rows + 1
                        End If

                        Grid.TextMatrix(Grid.Rows - 1, 0) = Codigo
                        Grid.TextMatrix(Grid.Rows - 1, 1) = UrlP
                        Grid.TextMatrix(Grid.Rows - 1, 2) = Now()

                        Sql = "insert into tbllogs (Url, Codigo, Origem, DataHora) Values ('" & UrlP & "', '" & Codigo & "', '" & TipoBanco & "', '" & Now() & "')"
                        Cn.Execute Sql
                    End If
PulaReg2:
                    TotalPalavra = TotalPalavra + 1


                ElseIf TipoBanco = "LCF" Or TipoBanco = "LCA" Then
                    If InStr(1, UrlP, "idLicitacao=") > 0 Then

                        lngPesq = InStr(1, UrlP, "idLicitacao=")
                        lngPesq1 = InStr(lngPesq, UrlP, "&")

                        If lngPesq1 > 0 Then
                            Codigo = Mid(UrlP, lngPesq + 12, lngPesq1 - lngPesq - 12)
                        Else
                            Codigo = Mid(UrlP, lngPesq + 9, Len(UrlP) - lngPesq)
                        End If

                        For LngXXX = 0 To Grid.Rows - 1
                            If Trim(Grid.TextMatrix(LngXXX, 0)) = Codigo Then
                                GoTo PulaReg3
                            End If
                        Next

                        ReDim Preserve UrlOk(UBound(UrlOk) + 1)
                        UrlOk(UBound(UrlOk)) = UrlP
                        If Grid.TextMatrix(1, 0) <> "" Then
                            Grid.Rows = Grid.Rows + 1
                        End If

                        Grid.TextMatrix(Grid.Rows - 1, 0) = Codigo
                        Grid.TextMatrix(Grid.Rows - 1, 1) = UrlP
                        Grid.TextMatrix(Grid.Rows - 1, 2) = Now()

                        Sql = "insert into tbllogs (Url, Codigo, Origem, DataHora) Values ('" & UrlP & "', '" & Codigo & "', '" & TipoBanco & "', '" & Now() & "')"
                        Cn.Execute Sql
                    End If
                    
                ElseIf TipoBanco = "BEC" Then
                    If InStr(1, UrlP, "nroOC=") > 0 Then

                        lngPesq = InStr(1, UrlP, "nroOC=")
                        lngPesq1 = InStr(lngPesq, UrlP, "&")

                        If lngPesq1 > 0 Then
                            Codigo = Mid(UrlP, lngPesq + 6, lngPesq1 - lngPesq - 12)
                        Else
                            Codigo = Mid(UrlP, lngPesq + 6, Len(UrlP) - lngPesq)
                        End If

                        For LngXXX = 0 To Grid.Rows - 1
                            If Trim(Grid.TextMatrix(LngXXX, 0)) = Codigo Then
                                GoTo PulaReg3
                            End If
                        Next

                        ReDim Preserve UrlOk(UBound(UrlOk) + 1)
                        UrlOk(UBound(UrlOk)) = UrlP
                        If Grid.TextMatrix(1, 0) <> "" Then
                            Grid.Rows = Grid.Rows + 1
                        End If

                        Grid.TextMatrix(Grid.Rows - 1, 0) = Codigo
                        Grid.TextMatrix(Grid.Rows - 1, 1) = UrlP
                        Grid.TextMatrix(Grid.Rows - 1, 2) = Now()

                        Sql = "insert into tbllogs (Url, Codigo, Origem, DataHora) Values ('" & UrlP & "', '" & Codigo & "', '" & TipoBanco & "', '" & Now() & "')"
                        Cn.Execute Sql
                    End If
PulaReg3:
                    TotalPalavra = TotalPalavra + 1

                End If
            End If
        End If
    Next

    For x = 0 To Len(Texto)

        DoEvents


        LblStatus.Caption = "Total de Pagina :" & TotalPagina & " Pagina validas " & TotalPaginaValida & " Palavra encontrada: " & TotalPalavra & " Nivel : " & Nivel
        
        If Nivel > Barra.Max Then
            Barra.Max = Nivel
        End If
        
        Barra.Value = Nivel
        
        DoEvents

        Lngx = InStr(x + 1, TextoAux, "<a href")
        If Lngx = 0 Then
            Exit Sub
        End If

        lngY = InStr(Lngx + 1, TextoAux, ">")

        UrlTemp = Mid(Texto, Lngx + 7, lngY - Lngx - 8)

        If InStr(1, UrlTemp, "detalhar") > 0 Then
            'Stop
        End If

        If Left(UrlTemp, 1) = Chr(34) Then UrlTemp = Right(UrlTemp, Len(UrlTemp) - 1)
        If Left(UrlTemp, 1) = "=" Then UrlTemp = Right(UrlTemp, Len(UrlTemp) - 1)
        If Left(UrlTemp, 1) = Chr(34) Then UrlTemp = Right(UrlTemp, Len(UrlTemp) - 1)

        DoEvents

        xlngX = InStr(1, UrlTemp, Chr(34))

        If xlngX = 0 Then
            xlngX = InStr(1, UrlTemp, "'")
        End If

        If xlngX > 0 Then

            xlngY = InStr(xlngX + 1, UrlTemp, Chr(34))

            If xlngY = 0 Then
                xlngY = InStr(xlngX + 1, UrlTemp, "'")
            End If

            If xlngY = 0 Then
                UrlTemp = Right(UrlTemp, Len(UrlTemp) - xlngX)
            Else
                If ContaChr(UrlTemp, Chr(34)) > 1 Then
                    UrlTemp = Left(UrlTemp, xlngX - 1)
                Else
                    UrlTemp = Mid(UrlTemp, xlngX + 1, xlngY - xlngX - 1)
                End If
                ''Mid(UrlTemp, lngX, Len(UrlTemp))
            End If

        End If

        UrlTemp = Replace(UrlTemp, "&amp;", "&")

        TotalPagina = TotalPagina + 1

        If TipoBanco = "BB" Then

            If (InStr(1, LCase(UrlTemp), "licitacoes") > 0 And InStr(1, LCase(UrlTemp), "consulta") > 0) Or (InStr(1, LCase(UrlTemp), "detalhar(") > 0) Then
                'detalhar(


                If InStr(1, LCase(UrlTemp), LCase("PesquisaAvancada")) = 0 Then
                    If InStr(1, LCase(UrlTemp), LCase("LicitacoesAcompanhaveis")) > 0 Or _
                       InStr(1, LCase(UrlTemp), LCase("ExibirLicitacao")) > 0 Or _
                       InStr(1, LCase(UrlTemp), "detalhar(") > 0 Then

                        If InStr(1, LCase(UrlTemp), "detalhar(") > 0 Then
                            strDetalhe = Right(UrlTemp, Len(UrlTemp) - 20)
                            strDetalhe = Left(strDetalhe, Len(strDetalhe) - 1)
                            'function detalhar(licitacao) { document.detalhar.numeroLicitacao.value = licitacao; document.detalhar.submit(); }
                            UrlTemp = "https://www.licitacoes-e.com.br/aop/consultar-detalhes-licitacao.aop?numeroLicitacao=" & strDetalhe & "&opcao=consultarDetalheLicitacao"
                        End If


                        DoEvents
                        NovoVetor = True
                        For xControle = 0 To UBound(UrlLida)
                            If LCase(UrlLida(xControle)) = LCase(UrlTemp) Then
                                NovoVetor = False
                                Exit For
                            End If
                            DoEvents
                            If CancelaAcao = True Then Exit Sub
                        Next

                        If CancelaAcao = True Then Exit Sub

                        If NovoVetor = True Then
                            If InStr(1, UrlTemp, "#") = 0 Then
                                ReDim Preserve UrlLida(UBound(UrlLida) + 1)
                                UrlLida(UBound(UrlLida)) = UrlTemp
                                DoEvents
                                TextoPesq = ReadUrl(UrlTemp, "POST")
                                Nivel = Nivel + 1
                                BuscaSubPagina TextoPesq, UrlTemp
                                Nivel = Nivel - 1
                            End If
                            DoEvents
                        End If
                        DoEvents
                        TotalPaginaValida = TotalPaginaValida + 1

                        If CancelaAcao = True Then Exit Sub

                    End If
                End If
            End If
        ElseIf TipoBanco = "CMG" Then
            UrlTemp = Replace(UrlTemp, "&amp;", "&")
            If LCase(Left(UrlTemp, 7)) <> "https://" Then
                UrlTemp = "https://wwws.sistemas.mg.gov.br" & UrlTemp
            End If
            DoEvents
            DoEvents
            NovoVetor = True
            For xControle = 0 To UBound(UrlLida)
                If LCase(UrlLida(xControle)) = LCase(UrlTemp) Then
                    NovoVetor = False
                    Exit For
                End If
                DoEvents
                If CancelaAcao = True Then Exit Sub
            Next

            If CancelaAcao = True Then Exit Sub

            If NovoVetor = True Then
                If InStr(1, UrlTemp, "#") = 0 Then
                    ReDim Preserve UrlLida(UBound(UrlLida) + 1)
                    UrlLida(UBound(UrlLida)) = UrlTemp
                    DoEvents
                    TextoPesq = ReadUrl(UrlTemp, "POST")
                    Nivel = Nivel + 1
                    BuscaSubPagina TextoPesq, UrlTemp
                    Nivel = Nivel - 1
                End If
                DoEvents
            End If
            DoEvents
            TotalPaginaValida = TotalPaginaValida + 1

            If CancelaAcao = True Then Exit Sub

        ElseIf TipoBanco = "LCF" Or TipoBanco = "LCA" Then
            UrlTemp = Replace(UrlTemp, "&amp;", "&")
            If LCase(Left(UrlTemp, 7)) <> "https://" Then
                UrlTemp = "https://wwws.licitanet.mg.gov.br/" & UrlTemp
            End If
            DoEvents
            DoEvents
            NovoVetor = True
            For xControle = 0 To UBound(UrlLida)
                If LCase(UrlLida(xControle)) = LCase(UrlTemp) Then
                    NovoVetor = False
                    Exit For
                End If
                DoEvents
                If CancelaAcao = True Then Exit Sub
            Next

            If CancelaAcao = True Then Exit Sub

            If NovoVetor = True Then
                If InStr(1, UrlTemp, "#") = 0 Then

                    UrlTemp = Replace(UrlTemp, Chr(13), "")
                    UrlTemp = Replace(UrlTemp, Chr(10), "")
                    UrlTemp = Replace(UrlTemp, Chr(9), "")

                    ReDim Preserve UrlLida(UBound(UrlLida) + 1)
                    UrlLida(UBound(UrlLida)) = UrlTemp
                    DoEvents
                    TextoPesq = ReadUrl(UrlTemp, "POST")
                    Nivel = Nivel + 1
                    BuscaSubPagina TextoPesq, UrlTemp
                    Nivel = Nivel - 1
                End If
                DoEvents
            End If
            DoEvents
            TotalPaginaValida = TotalPaginaValida + 1

            If CancelaAcao = True Then Exit Sub

        ElseIf TipoBanco = "BEC" Then
            'Debug.Print UrlTemp
            UrlTemp = Replace(UrlTemp, "&amp;", "&")
            Debug.Print UrlTemp
            If InStr(1, UrlTemp, "dtgNatDespOc") > 0 Then
'                MsgBox "Javascript"
            End If

            If InStr(1, LCase(UrlTemp), "resumonatureza") > 0 Or InStr(1, LCase(UrlTemp), "dadosoc") > 0 Or InStr(1, LCase(UrlTemp), "editaloc") > 0 Or InStr(1, UrlTemp, "dtgNatDespOc") > 0 Then
                If InStr(1, UrlTemp, "dtgNatDespOc") > 0 Then
                    UrlTemp = BuscaValor("action", Texto) & "&__EVENTTARGET=" & Replace(UrlTemp, "$", ":")
                    UrlTemp = UrlTemp & "&__EVENTARGUMENT=&__VIEWSTATE=" & BuscaValor("__VIEWSTATE" & Chr(34), Texto, 14)
                    UrlTemp = "http://www.bec.sp.gov.br/publico/aspx/" & UrlTemp
                Else
                    UrlTemp = "http://www.bec.sp.gov.br/publico/aspx/" & UrlTemp
                End If

                DoEvents
                DoEvents
                NovoVetor = True
                For xControle = 0 To UBound(UrlLida)
                    If LCase(UrlLida(xControle)) = LCase(UrlTemp) Then
                        NovoVetor = False
                        Exit For
                    End If
                    DoEvents
                    If CancelaAcao = True Then Exit Sub
                Next

                If CancelaAcao = True Then Exit Sub

                If NovoVetor = True Then
                    If InStr(1, UrlTemp, "#") = 0 Then

                        UrlTemp = Replace(UrlTemp, Chr(13), "")
                        UrlTemp = Replace(UrlTemp, Chr(10), "")
                        UrlTemp = Replace(UrlTemp, Chr(9), "")

                        ReDim Preserve UrlLida(UBound(UrlLida) + 1)
                        UrlLida(UBound(UrlLida)) = UrlTemp
                        DoEvents
                        TextoPesq = ReadUrl(UrlTemp, "GET")
                        Nivel = Nivel + 1
                        BuscaSubPagina TextoPesq, UrlTemp
                        Nivel = Nivel - 1
                    End If
                    DoEvents
                End If
                DoEvents
                TotalPaginaValida = TotalPaginaValida + 1

                If CancelaAcao = True Then Exit Sub
            End If
        End If
        x = lngY
        'LblStatus.Caption = "Total de Pagina :" & TotalPagina & " Pagina validas " & TotalPaginaValida & " Palavra encontrada: " & TotalPalavra & " Nivel : " & Nivel
        'DoEvents

        If Lngx > Len(Texto) Then Exit Sub
        DoEvents

        If CancelaAcao = True Then Exit Sub

    Next
End Sub

Private Sub CmdInic_Click()
    If CDate("03/03/2008") <= Date Then
        Err.Raise 6
        Exit Sub
    End If

    If Left(CmdInic.Caption, 1) = "C" Then
        CmdInic.Caption = "Iniciar Pesquisa"
        CancelaAcao = True
    Else
        CmdInic.Caption = "Cancelar"
        InicioBusca
        CmdInic.Caption = "Iniciar Pesquisa"

    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Form_Load_Error
    
    ReDim UrlLida(1) As String
    Dim Rs As ADODB.Recordset
    Dim LinhaAtual As Long
    ReDim UrlOk(0) As String
    Set Cn = New ADODB.Connection
    
    Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\softseek.mdb" & ";Persist Security Info=False"

    Set Rs = New ADODB.Recordset

    Grid.Clear
    Grid.Rows = 2
    Grid.FormatString = "Codigo    |Url de Acesso                                                                                                                       |Data / Hora                         "
    '"Codigo               |Url de Acesso                                                                                                                                                                             |Data / Hora                         "

    Rs.Open "Select * From tblLogs", Cn
    Do While Not Rs.EOF
        
        LinhaAtual = Grid.Rows - 1
        
        If Grid.TextMatrix(LinhaAtual, 0) <> "" Then
            Grid.AddItem ""
        End If
        LinhaAtual = Grid.Rows - 1
        
        Grid.TextMatrix(LinhaAtual, 0) = Rs.Fields("Codigo").Value
        Grid.TextMatrix(LinhaAtual, 1) = Rs.Fields("Url").Value
        Grid.TextMatrix(LinhaAtual, 2) = Rs.Fields("DataHora").Value
        ReDim Preserve UrlOk(UBound(UrlOk) + 1)
        UrlOk(UBound(UrlOk)) = Rs.Fields("Url").Value
        Rs.MoveNext
    Loop
    Rs.Close
    


    On Error GoTo 0
    Exit Sub

Form_Load_Error:

    MsgBox "Form_Load Of Formulário FrmBusca" & Chr(13) & Err.Description & " - " & Err.Number
    End

End Sub

Private Sub Grid_DblClick()
    If Trim(Grid.TextMatrix(Grid.Row, Grid.Col)) <> "" Then
        Call ShellExecute(hwnd, "open", Grid.TextMatrix(Grid.Row, Grid.Col), vbNullString, vbNullString, conSwNormal)
    End If
End Sub

Private Sub TxtBusca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then InicioBusca
End Sub
Private Function ContaChr(strTexto As String, strPalavra As String) As Long
    Dim strNovapalavra As String
    strNovapalavra = Replace(strTexto, strPalavra, "")
    ContaChr = Len(strTexto) - Len(strNovapalavra)
End Function
Private Function BuscaValor(strPalavra As String, strTexto As String, Optional lngADD As Long = 0) As String
    Dim Lngx0                                     As Long
    Dim lngX1                                     As Long
    Dim lngX2                                     As Long
    Dim strTexto2                                  As String
    
    Lngx0 = InStr(1, LCase(strTexto), LCase(strPalavra))
    lngX1 = InStr(Lngx0 + 1 + lngADD, LCase(strTexto), Chr(34))
    lngX2 = InStr(lngX1 + 1, LCase(strTexto), Chr(34))
    
    strTexto2 = Mid(strTexto, lngX1 + 1, lngX2 - lngX1 - 1)
    
    Lngx0 = InStr(1, LCase(strTexto2), "&")
    
    If Lngx0 > 0 Then
        strTexto2 = Left(strTexto2, Lngx0 - 1)
    End If
    
    BuscaValor = strTexto2
    
End Function

