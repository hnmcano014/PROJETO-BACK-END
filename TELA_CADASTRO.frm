VERSION 5.00
Begin VB.Form TELA_CADASTRO 
   Caption         =   "Tela de cadastro cliente"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton TELA_CADASTRO_BOTAOBUSCAR 
      Caption         =   "BUSCAR"
      Height          =   495
      Left            =   11640
      TabIndex        =   29
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox TELA_CADASTRO_TXTORIGEM 
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Top             =   6600
      Width           =   4695
   End
   Begin VB.TextBox TELA_CADASTRO_TXTTIPOCLIENTE 
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   5640
      Width           =   4695
   End
   Begin VB.TextBox TELA_CADASTRO_TXTSITUACAOCLIENTE 
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox TELA_CADASTRO_TXTCOMPLEMENTO 
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   6600
      Width           =   4695
   End
   Begin VB.TextBox TELA_CADASTRO_TXTENDERECO 
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   5640
      Width           =   4695
   End
   Begin VB.TextBox TELA_CADASTRO_TXTCIDADE 
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   4680
      Width           =   4695
   End
   Begin VB.TextBox TELA_CADASTRO_TXTEMAIL 
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox TELA_CADASTRO_TXTTELEFONE 
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox TELA_CADASTRO_TXTRG 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox TELA_CADASTRO_TXTID 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton TELA_CADASTRO_BOTAONOVO 
      Caption         =   "Criar novo cadastro"
      Height          =   855
      Left            =   10320
      TabIndex        =   6
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CommandButton TELA_CADASTRO_BOTAOEXTRAIR 
      Caption         =   "Extrair informação de cadastro"
      Height          =   855
      Left            =   10320
      TabIndex        =   5
      Top             =   6000
      Width           =   2655
   End
   Begin VB.ComboBox TELA_CADASTRO_FILTRO 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "NOME - RG"
      Top             =   360
      Width           =   11295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Informações Adicionais"
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   9855
      Begin VB.Label Label11 
         Caption         =   "Origem"
         Height          =   255
         Left            =   5040
         TabIndex        =   28
         Top             =   3000
         Width           =   4695
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo de cliente"
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   2040
         Width           =   4575
      End
      Begin VB.Label Label9 
         Caption         =   "Situação cliente"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label8 
         Caption         =   "Complemento"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   4695
      End
      Begin VB.Label Label7 
         Caption         =   "Endereço"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   4695
      End
      Begin VB.Label Label6 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Menu de opções"
      Height          =   4455
      Left            =   10200
      TabIndex        =   1
      Top             =   3360
      Width           =   2895
      Begin VB.CommandButton TELA_CADASTRO_BOTAOEXCLUIR 
         BackColor       =   &H80000000&
         Caption         =   "Excluir cadastro"
         Height          =   435
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   31
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton TELA_CADASTRO_BOTAOATUALIZAR 
         Caption         =   "Atualizar cadastro"
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Informações do Cliente"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   12975
      Begin VB.TextBox TELA_CADASTRO_TXTNOME 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   11175
      End
      Begin VB.Label Label5 
         Caption         =   "EMAIL"
         Height          =   255
         Left            =   8760
         TabIndex        =   16
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "TELEFONE"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label3 
         Caption         =   "RG"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label2 
         Caption         =   "Nome"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   600
         Width           =   11175
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label12 
      Caption         =   "FILTRO: NOME - RG"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "TELA_CADASTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const adStateOpen As Integer = 1
Dim conexaoBD As Object
Dim NOME_FILTRO As String
Dim RG_FILTRO As Long
Dim ID As Double

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
    LISTA_FILTRO
End Sub

Private Sub TELA_CADASTRO_BOTAOBUSCAR_Click()
    If Len(NOME_FILTRO) > 0 And RG_FILTRO > 0 Then
        PreencherTextBoxes
    Else
        MsgBox "Por favor, selecione um item no filtro antes de atualizar as informações."
    End If
End Sub

Private Sub LISTA_FILTRO()
    TELA_CADASTRO_FILTRO.Clear
    ' ########### Consultando NOME - RG distintos --//
    Dim strSQL As String
    strSQL = "SELECT DISTINCT NOME, RG FROM bdados_processo;"
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conexaoBD
    ' ########## PREENCHENDO O FILTRO DE LISTA --//
    Do While Not rs.EOF
        TELA_CADASTRO_FILTRO.AddItem rs("NOME").Value & " - " & rs("RG").Value
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub PreencherTextBoxes()
    ' ########## REALIZANDO CONSULTA PARA ALOCAR AS INFORMACOES DO FILTRO --//
    Dim strSQL As String
    strSQL = "SELECT ID,NOME, RG, TELEFONE, ENDERECO, EMAIL, ID_CLIENTE, CIDADE, SITUACAO_CLIENTE,ORIGEM,TIPO_CLIENTE,COMPLEMENTO  FROM bdados_processo WHERE NOME = '" & NOME_FILTRO & "' AND RG = " & RG_FILTRO & ";"
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open strSQL, conexaoBD

    ' ########## PREECNHENDO AS INFORMACOES --//
    If Not rs.EOF Then
        TELA_CADASTRO_TXTNOME.Text = rs("NOME").Value
        TELA_CADASTRO_TXTID.Text = rs("ID_CLIENTE").Value
        TELA_CADASTRO_TXTRG.Text = rs("RG").Value
        TELA_CADASTRO_TXTSITUACAOCLIENTE.Text = rs("SITUACAO_CLIENTE").Value
        TELA_CADASTRO_TXTORIGEM.Text = rs("ORIGEM").Value
        TELA_CADASTRO_TXTTIPOCLIENTE.Text = rs("TIPO_CLIENTE").Value
        TELA_CADASTRO_TXTCOMPLEMENTO.Text = rs("COMPLEMENTO").Value
        TELA_CADASTRO_TXTTELEFONE.Text = rs("TELEFONE").Value
        TELA_CADASTRO_TXTEMAIL.Text = rs("EMAIL").Value
        TELA_CADASTRO_TXTCIDADE.Text = rs("CIDADE").Value
        TELA_CADASTRO_TXTENDERECO.Text = rs("ENDERECO").Value
        
        ID = rs("ID").Value
    ' ### VALIDANDO O PREENCHIMENTO -->>
    Else
        MsgBox "Não há registros correspondentes aos filtros."
    End If
    
    rs.Close
End Sub

Private Sub TELA_CADASTRO_BOTAONOVO_Click()
    ' ########## CLIENTANDO VARIAVEIS --//
    Dim NOME As String
    Dim RG As Double
    Dim ENDERECO As String
    Dim TELEFONE As Double
    Dim COMPLEMENTO As String
    Dim CIDADE As String
    Dim SITUACAO_CLIENTE As String
    Dim ORIGEM As String
    Dim TIPO_CLIENTE As String
    Dim ID_CLIENTE As Double
    Dim EMAIL As String
    Dim strSQL As String
    ' ###
    If TELA_CADASTRO_TXTNOME.Text = "" Or _
       TELA_CADASTRO_TXTRG.Text = "" Or _
       TELA_CADASTRO_TXTENDERECO.Text = "" Or _
       TELA_CADASTRO_TXTTELEFONE.Text = "" Or _
       TELA_CADASTRO_TXTEMAIL.Text = "" Or _
       TELA_CADASTRO_TXTCOMPLEMENTO.Text = "" Or _
       TELA_CADASTRO_TXTCIDADE.Text = "" Or _
       TELA_CADASTRO_TXTSITUACAOCLIENTE.Text = "" Or _
       TELA_CADASTRO_TXTORIGEM.Text = "" Or _
       TELA_CADASTRO_TXTTIPOCLIENTE.Text = "" Or _
       TELA_CADASTRO_TXTID.Text = "" Then
       
        MsgBox "Por favor, preencha todas as informações corretamente antes de continuar."
        Exit Sub
    End If
    
    NOME = TELA_CADASTRO_TXTNOME.Text
    
    ' ### VALIDACAO DO CLIENTE -->>
    If Not IsNumeric(TELA_CADASTRO_TXTRG.Text) Then
        MsgBox "O campo RG deve ser um valor numérico."
        Exit Sub
    End If
    RG = CDbl(TELA_CADASTRO_TXTRG.Text)
    
    ENDERECO = TELA_CADASTRO_TXTENDERECO.Text
    
    ' ### VALIDACAO DO TELEFONE CLIENTE -->>
    If Not IsNumeric(TELA_CADASTRO_TXTTELEFONE.Text) Then
        MsgBox "O campo TELEFONE deve ser um valor numérico."
        Exit Sub
    End If
    TELEFONE = CDbl(TELA_CADASTRO_TXTTELEFONE.Text)
    
    COMPLEMENTO = TELA_CADASTRO_TXTCOMPLEMENTO.Text
    CIDADE = TELA_CADASTRO_TXTCIDADE.Text
    SITUACAO_CLIENTE = TELA_CADASTRO_TXTSITUACAOCLIENTE.Text
    ORIGEM = TELA_CADASTRO_TXTORIGEM.Text
    TIPO_CLIENTE = TELA_CADASTRO_TXTTIPOCLIENTE.Text
    
    ' ### VALIDACAO DO ID CLIENTE -->>
    If Not IsNumeric(TELA_CADASTRO_TXTID.Text) Then
        MsgBox "O campo ID_CLIENTE deve ser um valor numérico."
        Exit Sub
    End If
    ID_CLIENTE = CDbl(TELA_CADASTRO_TXTID.Text)
    
    EMAIL = TELA_CADASTRO_TXTEMAIL.Text
    
    ' ########## INSERINDO NOVAS INFORMACOES NO BANCO DE DADOS --//
    strSQL = "INSERT INTO bdados_processo (NOME, RG, ENDERECO, TELEFONE, EMAIL, COMPLEMENTO, CIDADE, SITUACAO_CLIENTE, ORIGEM, TIPO_CLIENTE, ID_CLIENTE) VALUES ('" & NOME & "', " & RG & ", '" & ENDERECO & "', " & TELEFONE & ", '" & EMAIL & "', '" & COMPLEMENTO & "', '" & CIDADE & "', '" & SITUACAO_CLIENTE & "', '" & ORIGEM & "', '" & TIPO_CLIENTE & "', '" & ID_CLIENTE & "');"
    
    On Error Resume Next
    conexaoBD.Execute strSQL
    If Err.Number = 0 Then
        MsgBox "Dados inseridos com sucesso!"
    Else
        MsgBox "Erro ao inserir dados: " & Err.Description
    End If
    LISTA_FILTRO
    On Error GoTo 0
End Sub

Private Sub AtualizarCadastro()
    Dim strSQL As String
    
    ' ### VALIDACAO DO RG -->>
    If Not IsNumeric(TELA_CADASTRO_TXTRG.Text) Then
        MsgBox "O campo RG deve ser um valor numérico."
        Exit Sub
    End If
    
    ' ### VALIDACAO DO TELEFONE -->>
    If Not IsNumeric(TELA_CADASTRO_TXTTELEFONE.Text) Then
        MsgBox "O campo TELEFONE deve ser um valor numérico."
        Exit Sub
    End If
    
    ' ### VALIDACAO DO ID CLIENTE -->>
    If Not IsNumeric(TELA_CADASTRO_TXTID.Text) Then
        MsgBox "O campo ID_CLIENTE deve ser um valor numérico."
        Exit Sub
    End If
    
    ' ########## ALTERACAO DO CADASTRO PELO ID DO BANCO DE DADOS --//
    strSQL = "UPDATE bdados_processo SET NOME = '" & TELA_CADASTRO_TXTNOME.Text & "', RG = '" & TELA_CADASTRO_TXTRG.Text & "', ENDERECO = '" & TELA_CADASTRO_TXTENDERECO.Text & "', TELEFONE = '" & TELA_CADASTRO_TXTTELEFONE.Text & "', EMAIL = '" & TELA_CADASTRO_TXTEMAIL.Text & "', COMPLEMENTO = '" & TELA_CADASTRO_TXTCOMPLEMENTO.Text & "', SITUACAO_CLIENTE = '" & TELA_CADASTRO_TXTSITUACAOCLIENTE.Text & "', TIPO_CLIENTE = '" & TELA_CADASTRO_TXTTIPOCLIENTE.Text & "', ORIGEM = '" & TELA_CADASTRO_TXTORIGEM.Text & "', CIDADE = '" & TELA_CADASTRO_TXTCIDADE.Text & "', ID_CLIENTE = '" & TELA_CADASTRO_TXTID.Text & "' WHERE ID = " & ID & ";"
    
    On Error Resume Next
    conexaoBD.Execute strSQL, , 1 ' 1 representa adCmdText
    
    ' ### VALIDACAO -->>
    If Err.Number = 0 Then
        MsgBox "Cadastro atualizado com sucesso!"
    Else
        MsgBox "Erro ao atualizar dados: " & Err.Description
    End If
    
    On Error GoTo 0
End Sub

Private Sub TELA_CADASTRO_BOTAOATUALIZAR_Click()
    AtualizarCadastro
    LISTA_FILTRO
End Sub

Private Sub ExcluirCadastro()
    Dim strSQL As String
    ' ########## ALTERACAO DO CADASTRO PELO ID DO BANCO DE DADOS --//
    strSQL = "DELETE FROM bdados_processo WHERE ID = " & ID & ";"
    On Error Resume Next
    conexaoBD.Execute strSQL, , 1
    ' ### VALIDACAO -->>
    If Err.Number = 0 Then
        MsgBox "Cadastro Excluido com sucesso!"
    Else
        MsgBox "Nenhum cadastro na tela."
    End If
    On Error GoTo 0
End Sub

Private Sub TELA_CADASTRO_BOTAOEXCLUIR_Click()
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja seguir com a exclusão?", vbQuestion + vbYesNo, "Confirmação")

    If resposta = vbYes Then
        ExcluirCadastro
    Else
        Exit Sub
    End If
    LISTA_FILTRO
End Sub

Private Sub TELA_CADASTRO_BOTAOEXTRAIR_Click()
TELA_CADASTRO_TELA_EXTRACAO_DE_INFORMACAO.Show
End Sub

Private Sub TELA_CADASTRO_FILTRO_Click()
    If TELA_CADASTRO_FILTRO.ListIndex <> -1 Then
        Dim splitValues() As String
        splitValues = Split(TELA_CADASTRO_FILTRO.List(TELA_CADASTRO_FILTRO.ListIndex), " - ")
        NOME_FILTRO = splitValues(0)
        RG_FILTRO = CLng(splitValues(1))
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not conexaoBD Is Nothing Then
        If conexaoBD.State = adStateOpen Then
            conexaoBD.Close
        End If
        Set conexaoBD = Nothing
    End If
End Sub


