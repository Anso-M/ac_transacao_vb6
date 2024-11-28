VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExportarExcel 
      Caption         =   "Exportar para Excel"
      Height          =   495
      Left            =   8280
      TabIndex        =   26
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdMostrarCategoria 
      Caption         =   "Mostrar Categoria"
      Height          =   495
      Left            =   8280
      TabIndex        =   23
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdLimparFiltros 
      Caption         =   "Limpar Filtros"
      Height          =   495
      Left            =   6600
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtFiltroValorTransacao 
      Height          =   285
      Left            =   3840
      TabIndex        =   18
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox txtFiltroDataTransacao 
      Height          =   285
      Left            =   3840
      TabIndex        =   17
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtFiltroNumeroCartao 
      Height          =   285
      Left            =   3840
      TabIndex        =   16
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtIdTransacao 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "Consultar"
      Height          =   495
      Left            =   6600
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   1920
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdInserir 
      Caption         =   "Inserir"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtDescricao 
      Height          =   1215
      Left            =   5640
      TabIndex        =   6
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtNumero 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   6271
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      Caption         =   "Limpa todos os filtros e a coluna de categoria."
      Height          =   375
      Left            =   6600
      TabIndex        =   25
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Limpa os campos acima."
      Height          =   375
      Left            =   1920
      TabIndex        =   24
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Filtrar por valor da transação"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Filtrar por data da transação"
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Filtrar por número do cartão"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "ID da transação (campo oculto)"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Data da transação"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Valor da transação"
      Height          =   255
      Left            =   5640
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Número do cartão"
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   2010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Form_Load()
    ' Configuração inicial da conexão com o banco de dados
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=SQLOLEDB;Data Source=localhost\SQLEXPRESS;Initial Catalog=Banco;User ID=adm;Password=12345;Trusted_Connection=True;Encrypt=False;"
    cn.Open
    ' Atualizar o DataGrid com os dados
    Call AtualizarGrid
End Sub

Private Sub AtualizarGrid()
    ' Preencher o DataGrid com os dados do banco, incluindo o nome do cliente
    Dim sql As String
    sql = "SELECT t.id_transacao, t.numero_cartao, t.valor_transacao, t.data_transacao, t.descricao, c.nome_cliente " & _
          "FROM transacoes t " & _
          "INNER JOIN clientes c ON t.id_cliente = c.id_cliente"
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = rs
End Sub


Private Sub DataGrid1_Click()
    On Error GoTo ErrorHandler

    ' Verifica se o Recordset tem registros
    If rs.EOF And rs.BOF Then
      MsgBox "Não há registros para exibir.", vbExclamation
      Exit Sub
    End If

    ' Formatar a data no formato dd/mm/yyyy
    Dim dataFormatada As String
    If Not DataGrid1.Columns("data_transacao") Is Nothing Then
        dataFormatada = Format(DataGrid1.Columns("data_transacao").Text, "dd/mm/yyyy")
    End If
    
    ' Preencher os campos de texto
    txtIdTransacao.Text = DataGrid1.Columns("id_transacao").Text
    txtNumero.Text = DataGrid1.Columns("numero_cartao").Text
    txtValor.Text = DataGrid1.Columns("valor_transacao").Text
    txtData.Text = dataFormatada
    txtDescricao.Text = DataGrid1.Columns("descricao").Text
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao tentar acessar os dados do DataGrid.", vbExclamation
End Sub


Private Function IsValidBrazilianDate(dateStr As String) As Boolean
    On Error GoTo ErrorHandler
    Dim dayPart As Integer
    Dim monthPart As Integer
    Dim yearPart As Integer
    Dim dateParts() As String
    Dim isLeapYear As Boolean
    Dim dataInserida As Date
    
    ' Dividir a string da data
    dateParts = Split(dateStr, "/")
    
    ' Verificar se temos exatamente 3 partes
    If UBound(dateParts) <> 2 Then
        IsValidBrazilianDate = False
        Exit Function
    End If
    
    ' Obter dia, mês e ano
    dayPart = CInt(dateParts(0))
    monthPart = CInt(dateParts(1))
    yearPart = CInt(dateParts(2))
    
    ' Verificar intervalo de mês
    If monthPart < 1 Or monthPart > 12 Then
        IsValidBrazilianDate = False
        Exit Function
    End If
    
    ' Verificar intervalo de dia para cada mês
    Select Case monthPart
        Case 4, 6, 9, 11 ' Abril, Junho, Setembro, Novembro têm 30 dias
            If dayPart < 1 Or dayPart > 30 Then
                IsValidBrazilianDate = False
                Exit Function
            End If
        Case 2 ' Fevereiro
            isLeapYear = (yearPart Mod 4 = 0 And yearPart Mod 100 <> 0) Or (yearPart Mod 400 = 0)
            If isLeapYear Then
                If dayPart < 1 Or dayPart > 29 Then
                    IsValidBrazilianDate = False
                    Exit Function
                End If
            Else
                If dayPart < 1 Or dayPart > 28 Then
                    IsValidBrazilianDate = False
                    Exit Function
                End If
            End If
        Case Else ' Meses com 31 dias
            If dayPart < 1 Or dayPart > 31 Then
                IsValidBrazilianDate = False
                Exit Function
            End If
    End Select
    
    ' Verificar o ano
    If yearPart < 1900 Or yearPart > Year(Now) Then
        IsValidBrazilianDate = False
        Exit Function
    End If
    
    ' Verificar se a data é posterior à data atual
    dataInserida = DateSerial(yearPart, monthPart, dayPart)
    If dataInserida > Date Then
      IsValidBrazilianDate = False
      Exit Function
    End If
    
    ' Se todas as verificações passarem, a data é válida
    IsValidBrazilianDate = True
    Exit Function

ErrorHandler:
    IsValidBrazilianDate = False
End Function


Private Sub cmdInserir_Click()
    Dim sql As String
    Dim rsCliente As ADODB.Recordset
    Dim idCliente As Long
    Dim dataTransacao As String

    ' Valida o formato do número do cartão
    If Len(txtNumero.Text) <> 16 Or Not IsNumeric(txtNumero.Text) Then
        MsgBox "Por favor, insira um número de cartão válido com 16 dígitos.", vbExclamation
        Exit Sub
    End If

    ' Valida o formato da data
    If Not IsValidBrazilianDate(txtData.Text) Then
      MsgBox "Por favor, insira uma data válida no formato dd/mm/yyyy.", vbExclamation
      Exit Sub
    End If

    ' Formata a data no formato yyyy-mm-dd
    dataTransacao = Format(CDate(txtData.Text), "yyyy-mm-dd")

    ' Valida o valor da transação
    If Not IsNumeric(txtValor.Text) Then
        MsgBox "Por favor, insira um valor de transação válido.", vbExclamation
        Exit Sub
    End If
    
    ' Substituir a vírgula por um ponto decimal no valor da transação
    valorTransacao = Replace(txtValor.Text, ",", ".")

    ' Obter o id_cliente baseado no numero_cartao
    sql = "SELECT id_cliente FROM clientes WHERE numero_cartao = '" & txtNumero.Text & "'"
    Set rsCliente = New ADODB.Recordset
    rsCliente.Open sql, cn, adOpenForwardOnly, adLockReadOnly

    ' Verifica se encontrou o cliente
    If Not rsCliente.EOF Then
        idCliente = rsCliente.Fields("id_cliente").Value
    Else
        MsgBox "Cliente não encontrado!"
        Exit Sub
    End If

    ' Inserção da nova transação
    sql = "INSERT INTO transacoes (numero_cartao, valor_transacao, data_transacao, descricao, id_cliente) VALUES ('" & txtNumero.Text & "', " & valorTransacao & ", '" & dataTransacao & "', '" & txtDescricao.Text & "', " & idCliente & ")"
    cn.Execute sql
    MsgBox "Transação inserida com sucesso!"
    Call AtualizarGrid
End Sub


Private Sub cmdEditar_Click()
    Dim sql As String
    Dim rsCliente As ADODB.Recordset
    Dim idCliente As Long
    Dim dataTransacao As Date
    
    ' Valida o formato do número do cartão
    If Len(txtNumero.Text) <> 16 Or Not IsNumeric(txtNumero.Text) Then
        MsgBox "Por favor, insira um número de cartão válido com 16 dígitos.", vbExclamation
        Exit Sub
    End If

    ' Valida o formato da data
    If Not IsValidBrazilianDate(txtData.Text) Then
      MsgBox "Por favor, insira uma data válida no formato dd/mm/yyyy.", vbExclamation
      Exit Sub
    End If

    ' Formata a data no formato yyyy-mm-dd
    dataTransacao = Format(CDate(txtData.Text), "yyyy-mm-dd")

    ' Valida o valor da transação
    If Not IsNumeric(txtValor.Text) Then
        MsgBox "Por favor, insira um valor de transação válido.", vbExclamation
        Exit Sub
    End If
    
    ' Substituir a vírgula por um ponto decimal
    valorTransacao = Replace(txtValor.Text, ",", ".")

    ' Obter o id_Cliente baseado no numero_cartao
    sql = "SELECT id_cliente FROM clientes WHERE numero_cartao = '" & txtNumero.Text & "'"
    Set rsCliente = New ADODB.Recordset
    rsCliente.Open sql, cn, adOpenForwardOnly, adLockReadOnly

    ' Verifica se encontrou o cliente
    If Not rsCliente.EOF Then
        idCliente = rsCliente.Fields("id_cliente").Value
    Else
        MsgBox "Cliente não encontrado!"
        Exit Sub
    End If

    ' Atualização da transação existente
    sql = "UPDATE transacoes SET numero_cartao = '" & txtNumero.Text & "', valor_transacao = " & valorTransacao & ", data_transacao = '" & dataTransacao & "', descricao = '" & txtDescricao.Text & "', id_cliente = " & idCliente & " WHERE id_transacao = " & CLng(txtIdTransacao.Text)
    cn.Execute sql
    MsgBox "Transação editada com sucesso!"
    Call AtualizarGrid
End Sub


Private Sub cmdExcluir_Click()
    Dim sql As String
    Dim idTransacao As Long

    ' Verifica se um id_transacao foi selecionado (pelo DataGrid)
    If txtIdTransacao.Text = "" Then
        MsgBox "Por favor, selecione uma transação para excluir.", vbExclamation
        Exit Sub
    End If

    ' Obtém o id_transacao do campo de texto oculto
    idTransacao = CLng(txtIdTransacao.Text)

    ' Confirmação de exclusão
    If MsgBox("Tem certeza que deseja excluir a transação selecionada?", vbYesNo + vbQuestion, "Confirmar Exclusão") = vbNo Then
        Exit Sub
    End If

    ' Exclusão da transação
    sql = "DELETE FROM transacoes WHERE id_transacao = " & idTransacao
    cn.Execute sql
    MsgBox "Transação excluída com sucesso!"
    Call AtualizarGrid
End Sub


Private Sub cmdLimpar_Click()
    ' Limpar todos os campos de texto da transação
    txtIdTransacao.Text = ""
    txtNumero.Text = ""
    txtValor.Text = ""
    txtData.Text = ""
    txtDescricao.Text = ""
End Sub


Private Sub cmdConsultar_Click()
    Dim sql As String
    Dim dataFiltro As String
    
    ' Inicia a consulta base
    sql = "SELECT t.id_transacao, t.numero_cartao, t.valor_transacao, t.data_transacao, t.descricao, c.nome_cliente " & _
          "FROM transacoes t " & _
          "INNER JOIN clientes c ON t.id_cliente = c.id_cliente WHERE 1=1"
    
    ' Adiciona cláusula para numero_cartao, se fornecido
    If txtFiltroNumeroCartao.Text <> "" Then
        sql = sql & " AND t.numero_cartao = '" & txtFiltroNumeroCartao.Text & "'"
    End If
    
    ' Adiciona cláusula para data_transacao, se fornecido
    If txtFiltroDataTransacao.Text <> "" Then
        ' Valida e formata a data
        If IsValidBrazilianDate(txtFiltroDataTransacao.Text) Then
            dataFiltro = Format(CDate(txtFiltroDataTransacao.Text), "yyyy-mm-dd")
            sql = sql & " AND t.data_transacao = '" & dataFiltro & "'"
        Else
            MsgBox "Por favor, insira uma data válida no formato dd/mm/yyyy.", vbExclamation
            Exit Sub
        End If
    End If
    
    ' Adiciona cláusula para valor_transacao, se fornecido
    If txtFiltroValorTransacao.Text <> "" Then
        ' Valida o valor numérico
        If IsNumeric(txtFiltroValorTransacao.Text) Then
            sql = sql & " AND t.valor_transacao = " & Replace(txtFiltroValorTransacao.Text, ",", ".")
        Else
            MsgBox "Por favor, insira um valor de transação válido.", vbExclamation
            Exit Sub
        End If
    End If

    ' Executa a consulta e atualiza o DataGrid
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = rs
End Sub


Private Sub cmdLimparFiltros_Click()
    ' Limpar todos os campos de filtro
    txtFiltroNumeroCartao.Text = ""
    txtFiltroDataTransacao.Text = ""
    txtFiltroValorTransacao.Text = ""
    ' Atualizar o DataGrid com todos os dados sem filtros
    Call AtualizarGrid
End Sub


Private Sub cmdMostrarCategoria_Click()
    On Error GoTo ErrorHandler

    ' Preencher o DataGrid com os dados do banco, incluindo a categoria da transação
    Dim sql As String
    sql = "SELECT t.id_transacao, t.numero_cartao, t.valor_transacao, t.data_transacao, t.descricao, " & _
          "c.nome_cliente, dbo.fn_CategoriaTransacao(t.valor_transacao) AS categoria " & _
          "FROM transacoes t " & _
          "INNER JOIN clientes c ON t.id_cliente = c.id_cliente"

    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = rs
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao tentar atualizar o DataGrid. Verifique se a função de categoria está criada no banco de dados.", vbExclamation
    Call AtualizarGrid
End Sub


Private Sub ExportarParaExcel()
    On Error GoTo ErrorHandler
    
    ' Criar objeto Excel
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
    
    Dim xlWorkbook As Object
    Set xlWorkbook = xlApp.Workbooks.Add
    
    Dim xlSheet As Object
    Set xlSheet = xlWorkbook.Sheets(1)
    
    ' Definir os cabeçalhos
    xlSheet.Cells(1, 1).Value = "numero_cartao"
    xlSheet.Cells(1, 2).Value = "valor_transacao"
    xlSheet.Cells(1, 3).Value = "data_transacao"
    xlSheet.Cells(1, 4).Value = "descricao"
    xlSheet.Cells(1, 5).Value = "categoria"
    
    ' Definir a formatação da coluna numero_cartao como texto
    xlSheet.Columns(1).NumberFormat = "@"
    
    ' Preencher o DataGrid com a consulta SQL
    Dim sql As String
    sql = "SELECT t.numero_cartao, t.valor_transacao, t.data_transacao, t.descricao, dbo.fn_CategoriaTransacao(t.valor_transacao) AS categoria " & _
          "FROM transacoes t " & _
          "INNER JOIN clientes c ON t.id_cliente = c.id_cliente " & _
          "WHERE t.data_transacao >= DATEADD(month, -1, GETDATE())"
    
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenStatic, adLockOptimistic
    
    ' Preencher o Excel com os dados do Recordset
    Dim i As Long
    i = 2
    Do While Not rs.EOF
        xlSheet.Cells(i, 1).Value = rs.Fields("numero_cartao").Value
        xlSheet.Cells(i, 2).Value = rs.Fields("valor_transacao").Value
        xlSheet.Cells(i, 3).Value = rs.Fields("data_transacao").Value
        xlSheet.Cells(i, 4).Value = rs.Fields("descricao").Value
        xlSheet.Cells(i, 5).Value = rs.Fields("categoria").Value
        rs.MoveNext
        i = i + 1
    Loop
    
    ' Exibir a caixa de diálogo para escolher o local de salvamento
    Dim saveDialog As Object
    Set saveDialog = CreateObject("MSComDlg.CommonDialog")
    saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    saveDialog.ShowSave
    
    Dim savePath As String
    savePath = saveDialog.FileName
    
    If savePath <> "" Then
        xlWorkbook.SaveAs savePath
        MsgBox "Relatório exportado com sucesso!", vbInformation
    Else
        MsgBox "Exportação cancelada.", vbExclamation
    End If
    
    ' Fechar o Excel
    xlWorkbook.Close False
    xlApp.Quit
    
    ' Limpar objetos
    Set xlSheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao exportar os dados para Excel. Verifique se a função de categoria está criada no banco de dados.", vbExclamation
End Sub



Private Sub cmdExportarExcel_Click()
    Call ExportarParaExcel
End Sub

