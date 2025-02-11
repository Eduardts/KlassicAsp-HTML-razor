<%
Class Database
    Private conn
    
    Private Sub Class_Initialize()
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.Open Application("dbConnection")
    End Sub
    
    Public Function ExecuteQuery(sql, params)
        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandText = sql
        
        If Not IsEmpty(params) Then
            For Each param in params
                cmd.Parameters.Append cmd.CreateParameter(param(0), param(1), param(2), param(3), param(4))
            Next
        End If
        
        Set ExecuteQuery = cmd.Execute()
    End Function
    
    Private Sub Class_Terminate()
        conn.Close
        Set conn = Nothing
    End Sub
End Class
%>

