' This is a set of functions that will allow you to communicate with the ChatGPT API within an Excel cell
'
'Just copy/paste this function into Excel following instructions in the Readme.md file.
' Don't forget to change the API key to your own.
' Author: Johann Dowa
' http://github.com/jddev273/chatgpt-excel-function
Function UnescapeString(ByVal str As String) As String
    Dim i As Integer
    Dim output As String
    For i = 1 To Len(str)
        If Mid(str, i, 2) = "\\" Then
            output = output & "\"
            i = i + 1
        ElseIf Mid(str, i, 2) = "\/" Then
            output = output & "/"
            i = i + 1
        ElseIf Mid(str, i, 2) = "\n" Then
            output = output & vbCrLf
            i = i + 1
        ElseIf Mid(str, i, 2) = "\r" Then
            output = output & vbCr
            i = i + 1
        ElseIf Mid(str, i, 2) = "\t" Then
            output = output & vbTab
            i = i + 1
        ElseIf Mid(str, i, 2) = "\" & Chr(34) Then
            output = output & """"
            i = i + 1
        Else
            output = output & Mid(str, i, 1)
        End If
    Next i
    UnescapeString = output
End Function


Function GetChatGPTResponse(prompt As String) As String
    Dim apiUrl As String
    Dim requestPayload As String
    Dim apiKey As String
    Dim httpRequest As Object
    Dim responseText As String
    Dim targetCell As Range
    Dim model As String
    Dim temperature As String
    Dim maxTokens As String
    
    apiUrl = "https://api.openai.com/v1/chat/completions"
    apiKey = "sk-YOUR-CHATGPT-KEY-HERE"
    
    model = "gpt-3.5-turbo"
    temperature = "0.5"
    maxTokens = 50

    ' Build the payload string
    requestPayload = "{""model"":""" & model & """,""messages"":[{""role"":""system"",""content"":""""},{""role"":""user"",""content"":""" & prompt & """}],""temperature"":" & temperature & ",""max_tokens"":" & maxTokens & "}"
    
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpRequest.Open "POST", apiUrl, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.setRequestHeader "Authorization", "Bearer " & apiKey
    On Error Resume Next
    httpRequest.send requestPayload
    On Error GoTo 0
    
    If httpRequest.Status <> 200 Then
        GetChatGPTResponse = "Error: " & httpRequest.Status & " " & httpRequest.StatusText
    Else
        responseText = httpRequest.responseText
        startPos = InStr(responseText, """content"":""") + 11
        endPos = InStr(responseText, """},""") - 1
        GetChatGPTResponse = Trim(UnescapeString(Mid(responseText, startPos, endPos - startPos + 1)))
    End If
    
    Set httpRequest = Nothing
End Function

Function ChatGPT(prompt As String) As String
    ChatGPT = GetChatGPTResponse(prompt)
End Function

Function ChatGPTQuickFill(Optional titleCell As Range = Nothing, Optional contextCell As Range = Nothing) As String
    Dim result As String
    Dim currentCell As Range
    Set currentCell = Application.Caller
    Dim prompt As String
    Dim titleRow As Integer
    Dim contextColumn As Integer
    
    ' Set default values for titleRow and contextColumn
    If titleCell Is Nothing Then
        titleRow = 1
    Else
        titleRow = title.row
    End If
    
    If contextCell Is Nothing Then
        contextColumn = 1
    Else
        contextColumn = context.Column
    End If
    
    result = GetContext(titleRow, contextColumn)


    prompt = "Provide {missing} value.  Use no extra words or punctuation.  Be specific.  Never explain anything.\n\n"
    prompt = prompt & "Country: Canada\nCapital: {missing}\nmissing=Ottawa\n\nPlanet: Mars\nCapital: {missing}\nmissing=Unknown\n\nCompany: Tesla\nTicker Symbol: {missing}\nmissing=TSLA\n\n"
    prompt = prompt & result & "{missing}\nmissing="

    result = GetChatGPTResponse(prompt)
    
    ChatGPTQuickFill = result
End Function

Function GetContext(Optional titleRow As Integer, Optional contextColumn As Integer) As String
    
    ' Get the active cell
    Dim activeCell As Range
    Set activeCell = Application.Caller
    
    ' Get the title
    Dim title As String
    title = Cells(titleRow, activeCell.Column).Value
    
    ' Get the context title
    Dim context_title As String
    context_title = Cells(titleRow, contextColumn).Value
    
    ' Get the context value
    Dim context_value As String
    context_value = Cells(activeCell.row, contextColumn).Value
    
    ' Return the results as a variant array
    GetContext = context_title & ": " & context_value & "\n" & title & ": "
    
End Function

Function ChatGPTList(topic As String, Optional horizontal As Boolean = False) As Variant
    Dim prompt As String
    Dim list As String
    Dim arr() As String

    prompt = "List values for topic.  Use no extra words or punctuation.  Be specific.  Never explain anything.  Each item in list will be in a new line without any formatting.\n\ntopic=3 largest countries in North America in land mass\nCanada\nUSA\nMexico\n\ntopic=5\nlargest cities on Mars\nUnknown\n\ntopic=founders of microsoft\nBill Gates\nPaul Allen\n\ntopic=" & topic
    list = GetChatGPTResponse(prompt)
    arr = Split(list, vbNewLine)
    
    If horizontal = False Then
        ChatGPTList = Application.Transpose(arr)
    Else
        ChatGPTList = arr
    End If
End Function