<%
Dim security
Set security = New SecurityService

If Not Session("UserAuthenticated") Then
    Response.Redirect "login.asp"
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim aiService
    Set aiService = New AIService
    
    Dim userInput
    userInput = Request.Form("message")
    
    Dim response
    response = aiService.GetChatbotResponse(userInput)
    
    Response.ContentType = "application/json"
    Response.Write "{""response"": """ & response & """}"
    Response.End
End If
%>
<!DOCTYPE html>
<html>
<head>
    <title>AI Chatbot Support</title>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <style>
        .chat-container {
            width: 400px;
            height: 500px;
            border: 1px solid #ccc;
            padding: 20px;
        }
        .chat-messages {
            height: 400px;
            overflow-y: auto;
        }
    </style>
</head>
<body>
    <div class="chat-container">
        <div class="chat-messages" id="messages"></div>
        <div class="chat-input">
            <input type="text" id="userInput" placeholder="Type your message...">
            <button onclick="sendMessage()">Send</button>
        </div>
    </div>
    
    <script>
        function sendMessage() {
            var input = $('#userInput');
            var message = input.val();
            input.val('');
            
            $.post('chatbot.asp', { message: message }, function(response) {
                $('#messages').append('<p><strong>You:</strong> ' + message + '</p>');
                $('#messages').append('<p><strong>Bot:</strong> ' + response.response + '</p>');
                $('#messages').scrollTop($('#messages')[0].scrollHeight);
            });
        }
    </script>
</body>
</html>

