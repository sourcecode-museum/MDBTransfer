VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formulário de Seleção"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   2790
      TabIndex        =   1
      Top             =   0
      Width           =   2850
      Begin VB.CommandButton cmdButton 
         Caption         =   "&Cancelar"
         Height          =   495
         Index           =   1
         Left            =   1400
         Picture         =   "frmSelect.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1400
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Con&firmar"
         Height          =   495
         Index           =   0
         Left            =   0
         Picture         =   "frmSelect.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1400
      End
   End
   Begin MSComctlLib.ListView List 
      Height          =   2670
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   4710
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvRetorno    As Variant

Const BtnConfirmar   As Integer = 0
Const BtnCancelar    As Integer = 1

Const LblErros       As Integer = 1
Const LblTransf      As Integer = 2
Const LblProces      As Integer = 3
Const LblTotal       As Integer = 4

Public Function CarregarLista(ByRef psListaArray() As String) As String
  Dim oLST    As ListItem
  Dim iFor    As Integer
  
  List.ColumnHeaders.Add , , "Tabela", List.Width - 350
  
  For iFor = 0 To UBound(psListaArray)
    List.ListItems.Add , , psListaArray(iFor)
  Next
  
  Show vbModal
  CarregarLista = mvRetorno
End Function

Public Function CarregarListaConfig(ByRef pRSConfig As ADODB.Recordset) As Long
   Dim i As Integer, nWidth As Integer
   Dim lst As ListItem, nErro As Integer
      
   List.ColumnHeaders.Add , , "Codigo", 0
   List.ColumnHeaders.Add , , "Arquivo Origem", 1300
   List.ColumnHeaders.Add , , "Endereço Origem", 2500
   List.ColumnHeaders.Add , , "Arquivo Destino", 1300
   List.ColumnHeaders.Add , , "Endereço Destino", 2500
   
   Me.Width = 7800
   List.Width = 7800 - (List.Left * 2)
   
  With pRSConfig
    .MoveFirst
    While Not .EOF
      Set lst = List.ListItems.Add(, , !ID)
      lst.SubItems(1) = !TBLORIGEM
      lst.SubItems(2) = !PATHORIGEM
      lst.SubItems(3) = !TBLDESTINO
      lst.SubItems(4) = !PATHDESTINO
      .MoveNext
    Wend
  End With
   
  Show vbModal
  CarregarListaConfig = Val(mvRetorno)
End Function

Private Sub Terminar()
   Unload Me
   Set frmSelect = Nothing
End Sub

Private Sub cmdButton_Click(Index As Integer)
  Select Case Index
    Case Is = 0 '"Confirmar"
      mvRetorno = List.SelectedItem.Text
      
    Case Is = 1 '"Cancelar"
      mvRetorno = ""
  
  End Select
  
  Call Terminar
End Sub

Private Sub List_BeforeLabelEdit(Cancel As Integer)
  Cancel = True
End Sub

Private Sub List_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Dim nChave As Integer
   
   nChave = CInt(ColumnHeader.Position)
   With List
      .SortKey = .ColumnHeaders.Item(nChave).Position - 1
      If .SortOrder = lvwAscending Then
         .SortOrder = lvwDescending
      Else
         .SortOrder = lvwAscending
      End If
   End With
End Sub

Private Sub List_DblClick()
  cmdButton_Click 0
End Sub

Private Sub List_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then List_DblClick
End Sub
