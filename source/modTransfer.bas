Attribute VB_Name = "modTransfer"
Option Explicit

Global Const cProvider As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source="

Private msArraySQLInsert()          As String

Private msTabelaDestino As String
Private msTabelaOrigem  As String
Private mbCancelar      As Boolean

Global gsPathOrigem     As String
Global gsTBLOrigem      As String

Global gsPathDestino    As String
Global gsTBLDestino     As String

Global gsWhere          As String

Public Property Get ConnectionString() As String
  ConnectionString = cProvider & App.Path & "\MDBTransf.mde"
End Property

Public Sub Transferir(ByVal pCfgID As Long)
  Dim RS              As ADODB.Recordset
  Dim sConn           As String
  Dim i               As Integer
  Dim saCMPOrigem()   As String
  Dim saCMPDestino()  As String
  
  sConn = ConnectionString
  Set RS = New ADODB.Recordset
  With RS
    .CursorLocation = adUseServer
    .Open "SELECT * FROM CAMPOSTRANSF WHERE CFG_ID = " & pCfgID, sConn
     
    While Not .EOF
      ReDim Preserve saCMPOrigem(i)
      ReDim Preserve saCMPDestino(i)
      saCMPOrigem(i) = !CMPORIGEM
      saCMPDestino(i) = !CMPDESTINO
      i = i + 1
      .MoveNext
    Wend
    .Close
  End With
  Set RS = Nothing
   
  If CriarSQLInsert(saCMPOrigem, saCMPDestino) Then
    Call ExecutarTransfer
  End If
End Sub

Private Function CriarSQLInsert(ByRef paCMPOrigem() As String, _
                                ByRef paCMPDestino() As String) As Boolean
  Dim i As Integer, nI As Integer
  Dim sSQL As String
  Dim sInsert As String, sInsCampos As String, sInsValores As String
  Dim RsOrigem   As ADODB.Recordset
  
  Set RsOrigem = CriarRSOrigem(paCMPOrigem)
  If RsOrigem Is Nothing Then Exit Function
  
  CriarSQLInsert = Not RsOrigem.EOF
  
  While Not RsOrigem.EOF
    sInsert = "Insert Into " & gsTBLDestino
    sInsCampos = "("
    sInsValores = "("
  
    For i = 0 To UBound(paCMPDestino)
      If Trim(RsOrigem(i).Value) <> "" Then
        sInsCampos = sInsCampos & paCMPDestino(i) & ", "
        
        Select Case RsOrigem(i).Type
          Case Is = adDBDate, adDate, adDBTime
            sInsValores = sInsValores & "#" & Trim(RsOrigem(i).Value) & "#, "
          
          Case Is = adCurrency, adVarChar, adDouble, adChapter, adChar, adWChar, adVarWChar, adBSTR
            sInsValores = sInsValores & """" & Trim(RsOrigem(i).Value) & """, "
          
          Case Is = adBoolean, adInteger
            sInsValores = sInsValores & Trim(RsOrigem(i).Value) & ", "
          
          Case Else
            sInsValores = sInsValores & """" & Trim(RsOrigem(i).Value) & """, "

        End Select
      End If
    Next
    
    ReDim Preserve msArraySQLInsert(nI)
    sInsCampos = Left(sInsCampos, Len(sInsCampos) - 2) & ")"
    sInsValores = Left(sInsValores, Len(sInsValores) - 2) & ")"
    sInsert = sInsert & sInsCampos & "values" & sInsValores
  
    msArraySQLInsert(nI) = sInsert
    nI = nI + 1
    DoEvents
    RsOrigem.MoveNext
  Wend
  
  RsOrigem.Close
  Set RsOrigem = Nothing
End Function

Private Function CriarRSOrigem(ByRef paCMPOrigem() As String) As ADODB.Recordset
   Dim CN     As ADODB.Connection
   Dim sSQL   As String
   Dim i      As Integer
   
   On Error GoTo Sair:
   
   Set CN = New ADODB.Connection
   
   CN.CursorLocation = adUseServer
   
   CN.Open cProvider & gsPathOrigem

   
   sSQL = "SELECT "
   For i = 0 To UBound(paCMPOrigem)
      sSQL = sSQL & paCMPOrigem(i) & ","
   Next
   sSQL = Left(sSQL, Len(sSQL) - 1) & " from " & gsTBLOrigem
   
   If gsWhere <> "" Then
   sSQL = sSQL & " WHERE " & gsWhere
   End If
   
   Set CriarRSOrigem = CN.Execute(sSQL)
      
   Set CN = Nothing
   Exit Function
Sair:
   MsgBox "Arquivo de Origem danificado ou não encontrado!", vbCritical, "Erro..."
End Function

Private Sub ExecutarTransfer()
  Dim CN   As ADODB.Connection
  Dim nI As Long
  Dim nCount As Long, nCountTemp As Long
  Dim nErros As Long, nTransf As Long
    
  Set CN = New ADODB.Connection
  CN.Open cProvider & gsPathDestino
  
  nCount = UBound(msArraySQLInsert)
  frmMain.StatusBar.Panels("Transferidos").Text = "0 de " & nCount + 1
  
  For nI = 0 To nCount
    If Not mbCancelar Then
      On Error Resume Next
      CN.Execute msArraySQLInsert(nI)
      nCountTemp = nCountTemp + 1
      
      If Err.Number <> 0 Then
        frmMain.txtErros.Text = frmMain.txtErros.Text & nCountTemp & " - " & msArraySQLInsert(nI) & vbCrLf
        Debug.Print nCountTemp & " - " & msArraySQLInsert(nI)
        nErros = nErros + 1
        frmMain.StatusBar.Panels("Erros").Text = "Erros: " & nErros
      Else
        nTransf = nTransf + 1
        frmMain.StatusBar.Panels("Transferidos").Text = nTransf & " de " & nCount + 1
      End If
      DoEvents
      frmMain.Progresso (Int(((nI + 1) * 100) \ (nCount + 1)))
    End If
  Next

  If nErros > 0 Then
     MsgBox "Erro na Transferêcia de " & nErros & " dos " & nCountTemp & _
            " Registros do Arquivo de Origem!", vbInformation, "Erro..."
  End If
  
  CN.Close
  Set CN = Nothing
End Sub

'Cancela a Transferência
Public Sub Cancelar()
   mbCancelar = True
End Sub


