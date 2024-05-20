VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Banco de Dados - Macros Existentes"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' M�dulo do formul�rio UserForm1

Option Explicit

Private Sub UserForm_Initialize()
    ' Inicializa o formul�rio ao ser exibido

    ' Faz a solicita��o HTTP para obter os t�tulos das macros VBA
    Dim macroTitles As Variant
    macroTitles = GetMacroTitlesFromAPI()
    
    ' Preenche o ListBox com os t�tulos das macros VBA
    ListBox1.List = macroTitles
End Sub

Private Function GetMacroTitlesFromAPI() As Variant
    ' Faz uma solicita��o HTTP GET para a rota /macros
    ' Retorna os t�tulos das macros VBA obtidos da API
    
    Dim http As Object
    Dim url As String
    Dim encodedCredentials As String
    Dim response As String
    Dim macroTitles As Variant
    Dim cleanedResponse As String
    Dim tempArray As Variant
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Definir a URL da API
    url = "https://macros-database.onrender.com/macros" ' Ajuste conforme necess�rio
    
    ' Codificar as credenciais em Base64
    encodedCredentials = "Basic " & Base64Encode("admin:PASSWORD") ' Substitua "PASSWORD" pela sua senha
    
    ' Criar o objeto HTTP
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Fazer a solicita��o HTTP GET
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", encodedCredentials
    http.send
    
    ' Verificar a resposta
    If http.Status = 200 Then
        response = http.responseText
        ' Remover os caracteres especiais e converter para uma matriz
        cleanedResponse = Replace(Replace(Replace(response, "[", ""), "]", ""), """", "")
        tempArray = Split(cleanedResponse, ",")
        
        ' Remover espa�os em branco extras
        For i = LBound(tempArray) To UBound(tempArray)
            tempArray(i) = Trim(tempArray(i))
        Next i
        
        GetMacroTitlesFromAPI = tempArray
    Else
        MsgBox "Erro ao obter os t�tulos das macros da API: " & http.Status & " - " & http.statusText
        GetMacroTitlesFromAPI = Array()
    End If
    
    Exit Function

ErrorHandler:
    MsgBox "Erro: " & Err.Description
    GetMacroTitlesFromAPI = Array()
End Function

Private Function Base64Encode(text As String) As String
    ' Fun��o para codificar texto em Base64
    
    Dim arrData() As Byte
    Dim objXML As Object
    Dim objNode As Object
    
    ' Converte o texto em um array de bytes
    arrData = StrConv(text, vbFromUnicode)
    
    ' Cria um objeto DOMDocument
    Set objXML = CreateObject("MSXML2.DOMDocument.3.0")
    
    ' Cria um n� de elemento
    Set objNode = objXML.createElement("Base64Data")
    
    ' Define o DataType do n� como bin.base64
    objNode.DataType = "bin.base64"
    
    ' Atribui os dados em bin�rio ao n�
    objNode.nodeTypedValue = arrData
    
    ' Obt�m o texto codificado em Base64
    Base64Encode = objNode.text
    
    ' Limpeza
    Set objNode = Nothing
    Set objXML = Nothing
End Function

