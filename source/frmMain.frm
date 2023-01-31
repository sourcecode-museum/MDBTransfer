VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000014&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtErros 
      Height          =   5805
      Left            =   5625
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      Top             =   570
      Width           =   5040
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   10695
      TabIndex        =   13
      Top             =   0
      Width           =   10695
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H80000010&
         Caption         =   "Criar Backup"
         Height          =   195
         Index           =   0
         Left            =   9375
         TabIndex        =   19
         Top             =   315
         Width           =   1245
      End
      Begin VB.CheckBox chkOptions 
         BackColor       =   &H80000010&
         Caption         =   "Deletar Registros Transferidos"
         Height          =   195
         Index           =   1
         Left            =   6435
         TabIndex        =   18
         Top             =   315
         Width           =   2490
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Importar"
         Height          =   495
         Index           =   2
         Left            =   2700
         Picture         =   "frmMain.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Width           =   1335
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   495
         Index           =   1
         Left            =   1365
         Picture         =   "frmMain.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Width           =   1335
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Limpar"
         Height          =   495
         Index           =   0
         Left            =   30
         Picture         =   "frmMain.frx":0F9E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   210
      Left            =   2955
      TabIndex        =   12
      Top             =   6420
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Top             =   6375
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Key             =   "Transferidos"
            Object.ToolTipText     =   "Registros Transfêridos"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Key             =   "Erros"
            Object.ToolTipText     =   "Erros de Transferência"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13679
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame 
      BackColor       =   &H80000014&
      Caption         =   "SQL Where (Tabela Origem)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   0
      Left            =   45
      TabIndex        =   9
      Top             =   5520
      Width           =   5550
      Begin VB.TextBox txtSQLWhere 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   45
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   210
         Width           =   5445
      End
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   1920
      Left            =   30
      TabIndex        =   8
      Top             =   3570
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   3387
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483626
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Campo de Origem"
         Object.Width           =   6165
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Campos de Destino"
         Object.Width           =   6165
      EndProperty
   End
   Begin VB.Frame Frame 
      BackColor       =   &H80000014&
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Index           =   2
      Left            =   2850
      TabIndex        =   1
      Top             =   510
      Width           =   2745
      Begin VB.ListBox LstCampos 
         BackColor       =   &H80000016&
         Height          =   2400
         Index           =   2
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   2625
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   15
         TabIndex        =   7
         Top             =   2610
         Width           =   2700
      End
      Begin VB.Label lblTamanho 
         Caption         =   "Tamanho:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   15
         TabIndex        =   6
         Top             =   2790
         Width           =   2700
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H80000014&
      Caption         =   "Origem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3030
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   510
      Width           =   2745
      Begin MSComDlg.CommonDialog cDialogo 
         Left            =   1395
         Top             =   2460
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Abrir"
      End
      Begin VB.ListBox LstCampos 
         BackColor       =   &H80000016&
         Height          =   2400
         Index           =   1
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   195
         Width           =   2625
      End
      Begin VB.Label lblTamanho 
         Caption         =   "Tamanho:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   15
         TabIndex        =   5
         Top             =   2790
         Width           =   2700
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   15
         TabIndex        =   4
         Top             =   2595
         Width           =   2700
      End
   End
   Begin VB.Menu menuAbrir 
      Caption         =   "&Abrir Tabela"
      Begin VB.Menu mnuAbrir 
         Caption         =   "Tabela &Origem"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "Tabela &Destino"
         Index           =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuAbrir 
         Caption         =   "&Sair"
         Index           =   4
      End
   End
   Begin VB.Menu menuConfigurar 
      Caption         =   "&Manutenção"
      Begin VB.Menu mnuConfigurar 
         Caption         =   "&Alterar"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuConfigurar 
         Caption         =   "&Excluir"
         Index           =   2
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moConn          As ADODB.Connection
Private moCatalogo      As ADOX.Catalog
Private moTBLOrigem     As ADOX.Table
Private moTBLDestino    As ADOX.Table

Private msConexão       As String
Private masListaArray() As String
Private mbAlterando     As Boolean
Private msTabAtual      As String
Private mlCfgID         As Long

Private Enum eTipo
  Origem = 1
  Destino = 2
End Enum

'Variaveis para Gravar no Banco
Private msBancoOrigem   As String
Private msBancoDestino  As String

Private Sub cmdButton_Click(Index As Integer)
  Select Case Index
    Case Is = 0 '"Limpar"
      Call LimparTodos
    
    Case Is = 1 '"Salvar":
     Call SalvarConfig
     cmdButton(2).Enabled = True
    Case Is = 2 '"Transferir"
      cmdButton(Index).Enabled = False
      Call SalvarConfig
      modTransfer.Transferir mlCfgID
  End Select
End Sub

Private Sub Form_Load()
   ListView.ColumnHeaders(1).Width = ListView.Width / 2 - 50
   ListView.ColumnHeaders(2).Width = ListView.ColumnHeaders(1).Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set moCatalogo = Nothing
  Set moTBLDestino = Nothing
  Set moTBLOrigem = Nothing
  Set frmMain = Nothing
  End
End Sub

Private Sub mnuAbrir_Click(Index As Integer)
   Dim sTitle As String
   Dim bOK    As Boolean
      
   If Index = Origem Then
      sTitle = "Origem"
   ElseIf Index = Destino Then
      sTitle = "Destino"
   Else
      Unload Me
      End
   End If
   
   With cDialogo
      .DialogTitle = "Abrir Arquivo de " & sTitle
      .Filter = " Access (*.MDB) |*.MDB| DBASE (*.DBF) | *.DBF| Todos Arquivos |*.DBF; *.MDB"
    '  .Filter = " Microsoft Access (*.MDB)|*.mdb"
      .FilterIndex = 3
      .Flags = &H4
      .InitDir = GetSetting(App.EXEName, "Paths", sTitle, App.Path)
      .CancelError = True
      
      On Error Resume Next
      .ShowOpen
      If Err.Number = 0 Then
        SaveSetting App.EXEName, "Paths", sTitle, .FileName
         bOK = TestarArquivo(Index, .FileTitle, .FileName)
      End If
      On Error GoTo 0
   
      If bOK Then
         Call CarregarListaCampos(Index)
      End If
   End With
End Sub

Private Function TestarArquivo(ByVal peTipo As eTipo, _
                               ByVal psArquivo As String, _
                               ByVal psPath As String) As Boolean
  Dim sMSG  As String
  Static bTestarArq As Boolean
  
  If Dir(psPath) = "" Then
    If peTipo = Origem Then
       sMSG = "Arquivo de Origem inválido!!!"
    Else
       sMSG = "Arquivo de Destino inválido!!!"
    End If
    MsgBox sMSG, vbCritical, "Erro..."
    mnuAbrir_Click CInt(peTipo)
    Exit Function
  End If
    
  msConexão = cProvider & psPath
   
  If peTipo = Origem Then
    msBancoOrigem = psArquivo
    gsPathOrigem = psPath
  Else
    msBancoDestino = psArquivo
    gsPathDestino = psPath
  End If
  
  TestarArquivo = True
End Function

Private Sub CarregarListaCampos(ByVal peTipo As eTipo)
   Dim i        As Integer
   Dim Index    As Integer
   Dim oTBLTemp As ADOX.Table
   Dim oConn    As ADODB.Connection
   
   Set oConn = New ADODB.Connection
   Set moCatalogo = New ADOX.Catalog
   
   oConn.Open msConexão
   Set moCatalogo.ActiveConnection = oConn
   
   If Not mbAlterando Then
      Call CarregarListaTabelas(moCatalogo)
      msTabAtual = frmSelect.CarregarLista(masListaArray)
   End If
   
   If msTabAtual = "" Then Exit Sub
   
   Index = CInt(peTipo)

   Set oTBLTemp = moCatalogo.Tables(msTabAtual)
   On Error GoTo 0
   If oTBLTemp.Columns.Count <= 0 Then
      MsgBox "Banco de Dados vazio ou Danificado!", vbCritical, "Erro..."
      Exit Sub
   End If
   
   Call LimparListaConfig
   LstCampos(Index).Clear
   For i = 0 To oTBLTemp.Columns.Count - 1
     LstCampos(Index).AddItem oTBLTemp.Columns(i).Name
   Next
   
   If Index = Destino Then
      Frame(Index).Caption = "Destino: " & msTabAtual
      gsTBLDestino = msTabAtual
      Set moTBLDestino = oTBLTemp
   Else
      Frame(Index).Caption = "Origem: " & msTabAtual
      gsTBLOrigem = msTabAtual
      Set moTBLOrigem = oTBLTemp
   End If
   LstCampos(Index).ListIndex = 0

   msTabAtual = ""
   oConn.Close
   Set oConn = Nothing
   Set moCatalogo = Nothing
End Sub

Private Sub LimparListaConfig()
   While ListView.ListItems.Count > 0
      ListView.SetFocus
      ListView_DblClick
   Wend
End Sub

Private Sub CarregarListaTabelas(pCatalogo As Catalog)
   Dim i As Integer
   Dim nCount As Integer
   Dim sNome As String

   For i = 0 To pCatalogo.Tables.Count - 1
      With pCatalogo.Tables(i)
         If .Type = "TABLE" Then
            sNome = pCatalogo.Tables(i).Name
            If Left$(sNome, 4) <> "MSys" Then
               ReDim Preserve masListaArray(nCount)
               masListaArray(nCount) = sNome
               nCount = nCount + 1
            End If
         End If
      End With
   Next
End Sub

Private Sub ListView_DblClick()
  Dim List As ListItem
  
  On Error Resume Next
  If ListView.ListItems.Count = 0 Then Exit Sub
  
  Set List = ListView.SelectedItem
  LstCampos(Origem).AddItem List.Text
  LstCampos(Destino).AddItem List.SubItems(1)
  
  ListView.ListItems.Remove List.Text
  If ListView.ListItems.Count = 0 Then
    cmdButton(1).Enabled = False 'Botao Salvar
  End If
End Sub

Private Sub LstCampos_Click(Index As Integer)
   Dim oTBLTemp As ADOX.Table
   Dim sTipo    As String
   
   If Index = Origem Then
      Set oTBLTemp = moTBLOrigem
   Else
      Set oTBLTemp = moTBLDestino
   End If
     
   On Error Resume Next
   sTipo = NomeTipo(oTBLTemp.Columns(LstCampos(Index).Text).Type)
   lblTipo(Index).Caption = "Tipo: " & sTipo
   lblTamanho(Index).Caption = "Tamanho: " & oTBLTemp.Columns(LstCampos(Index).Text).DefinedSize
   On Error GoTo 0
   
   Set oTBLTemp = Nothing
End Sub

Private Function NomeTipo(nTipo As Integer) As String
   Select Case nTipo
      Case 3
         NomeTipo = "Inteiro"
      Case 6
         NomeTipo = "Moeda"
      Case 7
         NomeTipo = "DataHora"
      Case 11
         NomeTipo = "Lógico"
      Case 129
         NomeTipo = "Caracter"
      Case Else
         NomeTipo = nTipo
   End Select
End Function

Private Sub LstCampos_DblClick(Index As Integer)
   Dim List As ListItem
   Dim sOrigem As String
   Dim sDestino As String
   
   If LstCampos(Origem).ListCount = 0 Or _
      LstCampos(Destino).ListCount = 0 Then
      Exit Sub
   End If
   cmdButton(1).Enabled = True 'TBar.ButtonEnabled("Salvar") = True

   sOrigem = LstCampos(Origem).Text
   sDestino = LstCampos(Destino).Text
   Set List = ListView.ListItems.Add(, sOrigem, sOrigem)
   List.SubItems(1) = sDestino
   
   On Error Resume Next
   LstCampos(Origem).RemoveItem LstCampos(Origem).ListIndex
   LstCampos(Destino).RemoveItem LstCampos(Destino).ListIndex
   LstCampos(Origem).ListIndex = 0
   LstCampos(Origem).SetFocus
   On Error GoTo 0
End Sub

Private Sub mnuConfigurar_Click(Index As Integer)
   Select Case Index
      Case 1
         Call Alterar
      Case 2
         Call Excluir
   End Select
End Sub

Private Sub SalvarConfig()
  Dim i         As Integer
  Dim RS        As ADODB.Recordset
  Dim sSQL      As String
  Dim sConn     As String
  Dim sCampos   As String
  Dim sValores  As String
      
  
'  cArqINI.PathFile = PathINI
'  sConn = cArqINI.Ler("CONEXAO", "DATABASE", "")
'  Set cArqINI = Nothing

  sSQL = "SELECT * FROM CONFIGURACAO"
   
  If mlCfgID > 0 Then
    sSQL = sSQL & " WHERE ID = " & mlCfgID
  End If
  
  Set RS = New ADODB.Recordset
  
  RS.Open sSQL, ConnectionString, adOpenStatic, adLockOptimistic
  
  If Not mbAlterando Then
    If RS.RecordCount > 0 Then
      On Error Resume Next
      RS.MoveLast
      On Error GoTo 0
    End If
    RS.AddNew
  End If
  
  RS!PATHORIGEM = gsPathOrigem
  RS!TBLORIGEM = gsTBLOrigem
  RS!BANCOORIGEM = msBancoOrigem
  
  RS!PATHDESTINO = gsPathDestino
  RS!TBLDESTINO = gsTBLDestino
  RS!BANCODESTINO = msBancoDestino
  gsWhere = txtSQLWhere.Text & ""
  RS!SQLWHERE = gsWhere
  RS.Update
  mlCfgID = RS!ID
  RS.Close
   
  sSQL = "SELECT * FROM CAMPOSTRANSF WHERE CFG_ID = " & mlCfgID
  RS.Open sSQL, ConnectionString, adOpenKeyset, adLockOptimistic
  For i = 1 To ListView.ListItems.Count
    RS.Filter = "Posicao = " & i
    If RS.EOF Then
      RS.AddNew
    End If
    RS!CFG_ID = mlCfgID
    RS!POSICAO = i
    RS!CMPORIGEM = ListView.ListItems(i).Text
    RS!CMPDESTINO = ListView.ListItems(i).SubItems(1)
    RS.Update
  Next
  
  RS.Close
  Set RS = Nothing
  
  MsgBox "Gravação Completada.", vbInformation, "MDBTranfer v" & App.Major & "." & App.Minor
End Sub

Private Sub LimparTodos()
  Call ModoConfiguração(False)

  LstCampos(Origem).Clear
  LstCampos(Destino).Clear
  ListView.ListItems.Clear
  
  mlCfgID = 0

  Frame(1).Caption = "Tabela Origem: "
  Frame(2).Caption = "Tabela Destino: "
  
  ReDim sListaArray(0)
  Set moCatalogo = Nothing
  Set moTBLDestino = Nothing
  Set moTBLOrigem = Nothing
End Sub

Private Sub LstCampos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   On Error GoTo Sair:
   If KeyCode = vbKeyDelete Then
      LstCampos(Index).RemoveItem LstCampos(Index).ListIndex
   End If
   Exit Sub
Sair:
End Sub

Private Sub Alterar()
  Dim RsConfig  As ADODB.Recordset
  Dim RsCampos  As ADODB.Recordset
  Dim lCfgID    As Long
  
  Set RsConfig = New ADODB.Recordset
  With RsConfig
    .CursorLocation = adUseClient
    .Open "SELECT * FROM CONFIGURACAO", modTransfer.ConnectionString

    If .RecordCount = 0 Then
      MsgBox "Não há configuração existente!", vbInformation, "Aviso..."
    Else
      lCfgID = frmSelect.CarregarListaConfig(RsConfig)
      If lCfgID > 0 Then
        .Filter = "ID = " & lCfgID
  
        Call ModoConfiguração(True)
        msTabAtual = !TBLORIGEM
        If TestarArquivo(Origem, msTabAtual, !PATHORIGEM) Then
          Call CarregarListaCampos(Origem)
  
          msTabAtual = !TBLDESTINO
          If TestarArquivo(Destino, msTabAtual, !PATHDESTINO) Then
            Call CarregarListaCampos(Destino)
            txtSQLWhere.Text = !SQLWHERE & ""
            Set RsCampos = New ADODB.Recordset
            With RsCampos
              .CursorLocation = adUseClient
              .Open "SELECT * FROM CAMPOSTRANSF WHERE CFG_ID = " & lCfgID, modTransfer.ConnectionString
            
              While Not .EOF
                LstCampos(Origem).Text = !CMPORIGEM
                LstCampos(Destino).Text = !CMPDESTINO
                LstCampos_DblClick 1
                .MoveNext
              Wend
            End With
          End If
          mlCfgID = lCfgID
        End If
      End If
    End If
  End With
TrataErro:
  Set RsConfig = Nothing
  Set RsCampos = Nothing
End Sub

Private Sub Excluir()
  Dim CN         As ADODB.Connection
  Dim sSQLDelete As String
  
  If mlCfgID = 0 Then Exit Sub
  
  Set CN = New ADODB.Connection
  
  CN.Open modTransfer.ConnectionString
  
  If MsgBox("Você deseja excluir esta configuração? ", _
            vbQuestion + vbYesNo, "Confirmação...") = vbYes Then

    'Deleta os Registros da Tabela de Configuração
    sSQLDelete = "DELETE * FROM CONFIGURACAO WHERE ID = " & mlCfgID
    CN.Execute sSQLDelete

    'Deleta os Registros da Tabela de Campos da Configuração
    sSQLDelete = "DELETE * FROM CAMPOSTRANSF WHERE CFG_ID = " & mlCfgID
    CN.Execute sSQLDelete
  End If
    
  CN.Close
  Set CN = Nothing
  Call LimparTodos
End Sub

Private Sub ModoConfiguração(bAlterando As Boolean)
   mbAlterando = bAlterando
   menuAbrir.Enabled = Not bAlterando
   mnuConfigurar(1).Enabled = Not bAlterando
   cmdButton(1).Enabled = bAlterando ' TBar.ButtonEnabled("Salvar") = bAlterando
   cmdButton(2).Enabled = bAlterando 'TBar.ButtonEnabled("Transferir") = bAlterando
End Sub

Public Sub Progresso(ByVal lProgresso As Long)
  On Error Resume Next
  Progress.Value = lProgresso
End Sub
