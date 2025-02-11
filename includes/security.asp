<%
Class SecurityService
    Private db
    
    Private Sub Class_Initialize()
        Set db = New Database
    End Sub
    
    Public Function ValidateUser(username, password)
        Dim params(1)
        params(0) = Array("@Username", adVarChar, adParamInput, 50, username)
        
        Dim sql
        sql = "SELECT UserID, PasswordHash FROM Users WHERE Username = @Username"
        
        Dim rs
        Set rs = db.ExecuteQuery(sql, params)
        
        If Not rs.EOF Then
            If VerifyHash(password, rs("PasswordHash")) Then
                Session("UserID") = rs("UserID")
                Session("UserAuthenticated") = True
                ValidateUser = True
                LogActivity "LOGIN_SUCCESS", username
            Else
                ValidateUser = False
                LogActivity "LOGIN_FAILURE", username
            End If
        Else
            ValidateUser = False
        End If
    End Function
    
    Private Sub LogActivity(activityType, username)
        Dim params(3)
        params(0) = Array("@ActivityType", adVarChar, adParamInput, 50, activityType)
        params(1) = Array("@Username", adVarChar, adParamInput, 50, username)
        params(2) = Array("@IPAddress", adVarChar, adParamInput, 50, Request.ServerVariables("REMOTE_ADDR"))
        
        Dim sql
        sql = "INSERT INTO ActivityLog (ActivityType, Username, IPAddress, ActivityDate) " & _
              "VALUES (@ActivityType, @Username, @IPAddress, GETDATE())"
              
        db.ExecuteQuery sql, params
    End Sub
End Class
%>
