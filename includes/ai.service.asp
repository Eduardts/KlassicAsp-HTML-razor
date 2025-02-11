<%
Class AIService
    Private httpObj
    
    Private Sub Class_Initialize()
        Set httpObj = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    End Sub
    
    Public Function GetChatbotResponse(userInput)
        httpObj.Open "POST", Application("AIEndpoint"), False
        httpObj.setRequestHeader "Content-Type", "application/json"
        httpObj.setRequestHeader "Authorization", "Bearer " & Application("APIKey")
        
        Dim requestBody
        requestBody = "{""model"": ""gpt-3.5-turbo"", ""messages"": [{""role"": ""user"", ""content"": """ & userInput & """}]}"
        
        httpObj.send requestBody
        
        If httpObj.status = 200 Then
            Dim response
            response = ParseJSON(httpObj.responseText)
            GetChatbotResponse = response("choices")(0)("message")("content")
        Else
            GetChatbotResponse = "Sorry, I'm having trouble understanding right now."
        End If
    End Function
End Class
%>
