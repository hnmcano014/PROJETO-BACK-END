VERSION 5.00
Begin VB.Form TELA_CADASTRO_TELA_EXTRACAO_DE_INFORMACAO 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox TELA_CADASTRO_FILTROSITUACAOCLIENTE 
      Height          =   2205
      Left            =   1920
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton TELA_CADASTRO_BOTAOEXCEL 
      Caption         =   "Extrair relatório para um excel"
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro Situação do cliente"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "TELA_CADASTRO_TELA_EXTRACAO_DE_INFORMACAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const adStateOpen As Integer = 1
Dim conexaoBD As Object
Dim SITUACAOCLIENTE_FILTRO As String

Private Sub Form_Load()
    ' ########## INICIANDO CONEXAO COM O BANCO DE DADOS --//
    Set conexaoBD = CreateObject("ADODB.Connection")
    conexaoBD.ConnectionString = "Driver={PostgreSQL ODBC Driver(UNICODE)};Server=localhost;Port=5433;Database=postgres;UID=postgres;PWD=hnmcano;"
    conexaoBD.Open

    ' ### VALIDACAO DE CONEXAO -->>
    If conexaoBD.State <> adStateOpen Then
        MsgBox "Erro ao conectar ao banco de dados: " & conexaoBD.Errors(0).Description
        Unload Me
    End If

    ' ########## CARREGANDO FUNCAO DE LISTA DE FILTRO --//
    LISTA_FILTRO_SITUACAOCLIENTE
End Sub

Private Sub LISTA_FILTRO_SITUACAOCLIENTE()
    ' ########## CONTRUINDO A LISTA DE SITUACAO_CLIENTE --//
    TELA_CADASTRO_FILTROSITUACAOCLIENTE.Clear
    Dim strSQL As String
    strSQL = "SELECT DISTINCT SITUACAO_CLIENTE FROM bdados_processo;"
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conexaoBD

    ' ########## PREENCHENDO O FILTRO DE LISTA --//
    Do While Not rs.EOF
        TELA_CADASTRO_FILTROSITUACAOCLIENTE.AddItem rs("SITUACAO_CLIENTE").Value
        rs.MoveNext
    Loop

    rs.Close
End Sub

Private Sub TELA_CADASTRO_BOTAOEXCEL_Click()
    ' ########## CONSULTA COM OS ITENS SELECIONADOS --//
    Dim strSQL As String
    strSQL = "SELECT * FROM bdados_processo"

    ' ########## CONSULTA COM ITENS SELECIONADOS --//
    If TELA_CADASTRO_FILTROSITUACAOCLIENTE.ListCount > 0 Then
        Dim i As Integer
        Dim selectedItems As String

        ' ########## VALIDACAO COM ITENS DO FILTROS --//
        For i = 0 To TELA_CADASTRO_FILTROSITUACAOCLIENTE.ListCount - 1
            If TELA_CADASTRO_FILTROSITUACAOCLIENTE.Selected(i) Then
                If Len(selectedItems) > 0 Then
                    selectedItems = selectedItems & ","
                End If
                selectedItems = selectedItems & "'" & TELA_CADASTRO_FILTROSITUACAOCLIENTE.List(i) & "'"
            End If
        Next i

        ' ########## CONSULTA COM FILTRO --//
        If Len(selectedItems) > 0 Then
            strSQL = strSQL & " AND SITUACAO_CLIENTE IN (" & selectedItems & ")"
        End If
    End If

    ' ########## ARMAZENANDO RESULTADO
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conexaoBD

    If Not rs.EOF Then
        Dim xlApp As Object
        Set xlApp = CreateObject("Excel.Application")

        Dim xlWorkbook As Object
        Set xlWorkbook = xlApp.Workbooks.Add

        Dim xlWorksheet As Object
        Set xlWorksheet = xlWorkbook.Worksheets(1)

        For i = 1 To rs.Fields.Count
            xlWorksheet.Cells(1, i).Value = rs.Fields(i - 1).Name
        Next i

        xlWorksheet.Range("A2").CopyFromRecordset rs

        ' ########## SALVANDO RELATÓRIO --//
        Dim fileName As String
        fileName = Environ("USERPROFILE") & "\Downloads\Consulta_Resultado.xlsx"
        xlWorkbook.SaveAs fileName

        xlApp.Quit

        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        Set xlApp = Nothing

        ' ########## ARQUIVO SALVO --//
        MsgBox "Consulta salva em: " & fileName, vbInformation
    Else
        MsgBox "Nenhum resultado encontrado.", vbInformation
    End If

    rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not conexaoBD Is Nothing Then
        If conexaoBD.State = adStateOpen Then
            conexaoBD.Close
        End If
        Set conexaoBD = Nothing
    End If
End Sub

