' This is a macro that will allow you to communicate with the ChatGPT API within an Excel cell
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