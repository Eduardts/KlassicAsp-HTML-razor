<%
Class AnomalyDetector
    Private db
    
    Private Sub Class_Initialize()
        Set db = New Database
    End Sub
    
    Public Function DetectAnomalies(userID)
        Dim params(1)
        params(0) = Array("@UserID", adVarChar, adParamInput, 50, userID)
        
        Dim sql
        sql = "SELECT ActivityType, COUNT(*) as ActivityCount, " & _
              "DATEPART(hour, ActivityDate) as ActivityHour " & _
              "FROM ActivityLog " & _
              "WHERE UserID = @UserID " & _
              "AND ActivityDate >= DATEADD(day, -7, GETDATE()) " & _
              "GROUP BY ActivityType, DATEPART(hour, ActivityDate)"
        
        Dim rs
        Set rs = db.ExecuteQuery(sql, params)
        
        Dim anomalies
        Set anomalies = Server.CreateObject("Scripting.Dictionary")
        
        Do While Not rs.EOF
            If IsAnomalous(rs("ActivityCount"), rs("ActivityHour")) Then
                anomalies.Add rs("ActivityType"), _
                    Array(rs("ActivityCount"), rs("ActivityHour"))
            End If
            rs.MoveNext
        Loop
        
        Set DetectAnomalies = anomalies
    End Function
    
    Private Function IsAnomalous(count, hour)
        ' Simple threshold-based anomaly detection
        If hour >= 23 Or hour <= 4 Then
            IsAnomalous = (count > 5)
        Else
            IsAnomalous = (count > 20)
        End If
    End Function
End Class
%>
