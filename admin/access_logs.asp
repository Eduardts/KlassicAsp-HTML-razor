<%
If Not Session("UserAuthenticated") Then
    Response.Redirect "../login.asp"
End If

Dim db
Set db = New Database

Dim sql
sql = "SELECT TOP 1000 " & _
      "ActivityType, Username, IPAddress, ActivityDate " & _
      "FROM ActivityLog " & _
      "ORDER BY ActivityDate DESC"

Dim logs
Set logs = db.ExecuteQuery(sql, Empty)
%>
<!DOCTYPE html>
<html>
<head>
    <title>Access Logs</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        .chart {
            width: 800px;
            height: 400px;
            margin: 20px;
        }
    </style>
</head>
<body>
    <h1>Access Logs Analysis</h1>
    
    <div id="activityChart" class="chart"></div>
    <div id="timeChart" class="chart"></div>
    
    <table border="1">
        <tr>
            <th>Activity</th>
            <th>Username</th>
            <th>IP Address</th>
            <th>Date</th>
        </tr>
        <% Do While Not logs.EOF %>
        <tr>
            <td><%= logs("ActivityType") %></td>
            <td><%= logs("Username") %></td>
            <td><%= logs("IPAddress") %></td>
            <td><%= logs("ActivityDate") %></td>
        </tr>
        <% 
            logs.MoveNext
        Loop 
        %>
    </table>
    
    <script>
        // Create activity distribution chart
        var activityData = [
            {
                values: [<%= GetActivityCounts() %>],
                labels: [<%= GetActivityTypes() %>],
                type: 'pie'
            }
        ];
        
        Plotly.newPlot('activityChart', activityData);
        
        // Create time-based activity chart
        var timeData = [
            {
                x: [<%= GetActivityTimes() %>],
                y: [<%= GetActivityCounts() %>],
                type: 'scatter'
            }
        ];
        
        Plotly.newPlot('timeChart', timeData);
    </script>
</body>
</html>
