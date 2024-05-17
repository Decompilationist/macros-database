Sub GetMacroFromAPI()
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim user As String
    Dim pass As String
    Dim encodedCredentials As String
    Dim response As String
    Dim moduleName As String
    Dim startLine As Long

    ' Definir a URL da API
    url = "https://macros-database.onrender.com/macro"

    ' Definir as credenciais de autenticação
    user = "Control Tower" ' ou "Control Tower Formatar"
    pass = "PBIGustavo" ' Substitua pela sua senha

    ' Codificar as credenciais em Base64
    encodedCredentials = "Basic " & Base64Encode(user & ":" & pass)

    ' Criar o objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Fazer a solicitação HTTP GET
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", encodedCredentials
    http.send

    ' Verificar a resposta
    If http.Status = 200 Then
        response = http.responseText

        ' Criar um novo módulo VBA chamado "Módulo1"
        moduleName = "Módulo1"
        CreateNewModule moduleName

        ' Inserir a macro VBA no módulo recém-criado
        startLine = 1
        InsertMacroCode moduleName, response, startLine

        ' Criar botão para executar a macro
        CreateButtonToRunMacro moduleName, ExtractMacroName(response)
    Else
        MsgBox "Erro ao obter a macro da API: " & http.Status & " - " & http.statusText
    End If

    ' Limpar
    Set http = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro: " & Err.Description
End Sub

Sub CreateNewModule(moduleName As String)
    On Error GoTo ErrorHandler

    Dim vbProj As Object
    Dim vbComp As Object

    ' Obter o projeto VBA
    Set vbProj = ThisWorkbook.VBProject

    ' Adicionar um novo módulo VBA
    Set vbComp = vbProj.VBComponents.Add(1) ' 1 representa um módulo de código
    vbComp.Name = moduleName

    MsgBox "Novo módulo VBA '" & moduleName & "' criado com sucesso!"

    ' Limpar
    Set vbComp = Nothing
    Set vbProj = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao criar o novo módulo VBA: " & Err.Description
End Sub




Function Base64Encode(text As String) As String
    Dim arrData() As Byte
    Dim objXML As Object
    Dim objNode As Object

    ' Converte o texto em um array de bytes
    arrData = StrConv(text, vbFromUnicode)

    ' Cria um objeto DOMDocument
    Set objXML = CreateObject("MSXML2.DOMDocument.3.0")

    ' Cria um nó de elemento
    Set objNode = objXML.createElement("Base64Data")

    ' Define o DataType do nó como bin.base64
    objNode.DataType = "bin.base64"

    ' Atribui os dados em binário ao nó
    objNode.nodeTypedValue = arrData

    ' Obtém o texto codificado em Base64
    Base64Encode = objNode.text

    ' Limpeza
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Sub InsertMacroCode(moduleName As String, macroCode As String, startLine As Long)
    On Error GoTo ErrorHandler

    Dim vbProj As Object
    Dim vbComp As Object
    Dim codeModule As Object
    Dim lines() As String
    Dim i As Long

    ' Obter o projeto VBA
    Set vbProj = ThisWorkbook.VBProject

    ' Obter o módulo VBA
    Set vbComp = vbProj.VBComponents(moduleName)
    Set codeModule = vbComp.codeModule

    ' Dividir o código da macro em linhas
    lines = Split(macroCode, vbCrLf)

    ' Verificar se há linhas suficientes no módulo para a inserção
    If startLine > codeModule.CountOfLines + 1 Then
        MsgBox "Erro: A linha de início está fora do intervalo."
        Exit Sub
    End If

    ' Verificar se a divisão das linhas ocorreu corretamente
    If UBound(lines) < 0 Then
        MsgBox "Erro: Nenhuma linha de código encontrada para inserir."
        Exit Sub
    End If

    ' Inserir cada linha de código no módulo VBA
    For i = LBound(lines) To UBound(lines)
        codeModule.InsertLines startLine + i, lines(i)
    Next i

    MsgBox "Macro inserida com sucesso!"

    ' Limpar
    Set codeModule = Nothing
    Set vbComp = Nothing
    Set vbProj = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao inserir a macro: " & Err.Description
End Sub

Sub CreateButtonToRunMacro(moduleName As String, macroName As String)
    Dim ws As Worksheet
    Dim btn As Button
    Dim cell As Range
    
    ' Definir a célula D1 (ou ajuste conforme necessário)
    Set ws = ActiveSheet
    Set cell = ws.Range("D9")
    
    ' Adicionar botão na planilha ativa
    Set btn = ws.Buttons.Add(cell.Left, cell.Top, cell.Width, cell.Height)
    With btn
        .Caption = "Executar " & macroName
        .OnAction = "'" & ThisWorkbook.Name & "'!" & moduleName & "." & macroName
    End With

    MsgBox "Botão criado com sucesso na célula " & cell.Address & "!"
End Sub

Function ExtractMacroName(macroCode As String) As String
    Dim lines() As String
    lines = Split(macroCode, vbCrLf)
    Dim firstLine As String
    firstLine = Trim(lines(0))
    
    If InStr(1, firstLine, "Sub ") = 1 Then
        ExtractMacroName = Split(Split(firstLine, " ")(1), "(")(0)
    Else
        ExtractMacroName = "MacroDesconhecida"
    End If
End Function


Sub DeleteMacroModule(moduleName As String)
    On Error GoTo ErrorHandler

    Dim vbProj As Object

    ' Obter o projeto VBA
    Set vbProj = ThisWorkbook.VBProject

    ' Verificar se o módulo existe
    If ModuleExists(vbProj, moduleName) Then
        ' Remover o módulo VBA
        vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)

        MsgBox "Módulo VBA '" & moduleName & "' removido com sucesso!"
    Else
        MsgBox "O módulo VBA '" & moduleName & "' não existe neste projeto."
    End If

    ' Limpar
    Set vbProj = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao excluir o módulo VBA: " & Err.Description
End Sub

Function ModuleExists(vbProj As Object, moduleName As String) As Boolean
    Dim vbComp As Object
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(moduleName)
    On Error GoTo 0
    ModuleExists = Not vbComp Is Nothing
End Function

Sub DeleteMacroButtonClick()
    On Error GoTo ErrorHandler

    Dim moduleName As String
    moduleName = "Módulo1" ' Nome do módulo que deseja excluir

    ' Excluir o módulo VBA
    DeleteMacroModule moduleName

    ' Excluir todo o conteúdo da planilha, exceto na coluna D até o final
    DeleteSheetContentExceptColumnD

    ' Excluir o botão de execução de macro
    DeleteMacroButton

    Exit Sub

ErrorHandler:
    MsgBox "Erro: " & Err.Description
End Sub

Sub DeleteMacroButton()
    Dim btn As Button
    Dim ws As Worksheet

    ' Definir a planilha ativa
    Set ws = ThisWorkbook.ActiveSheet

    ' Loop através de cada botão na planilha
    For Each btn In ws.Buttons
        ' Verificar se o botão está vinculado à macro
        If InStr(btn.OnAction, "Módulo1") > 0 Then
            ' Excluir o botão
            btn.Delete
        End If
    Next btn

    MsgBox "Botão de execução de macro excluído com sucesso!"
End Sub


Sub DeleteSheetContentExceptColumnD()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim colD As Range
    
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayAlerts = False

    ' Defina a planilha ativa
    Set ws = ThisWorkbook.ActiveSheet

    ' Encontre a última linha e a última coluna com dados na planilha
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Defina a coluna D
    Set colD = ws.Range("D:D")

    ' Loop através de cada célula na planilha
    For Each cell In ws.Range("A1", ws.Cells(lastRow, lastColumn))
        ' Verifique se a célula não está na coluna D ou à esquerda dela
        If Intersect(cell, colD) Is Nothing Then
            ' Excluir o conteúdo da célula
            cell.ClearContents
        End If
    Next cell

    MsgBox "Conteúdo excluído da planilha, exceto na coluna D até o final!"

End Sub


