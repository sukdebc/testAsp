<%@LANGUAGE="VBSCRIPT" %>
<!DOCTYPE html>
<html>
<body>
<%
Response.Write("Hello in classic ASP World! Testing if we can execute few in-built method. <br>")
Response.Write "<br>"
Randomize()
randomNumber=Int(100 * Rnd())
response.write("A random number: <b>" & randomNumber & "</b><br>")
Response.Write "<br>"
Response.Write("Checking Session <br>")
Response.Write("Session.SessionID" & Session.SessionID & "<br>")
Response.Write "<br>"
Response.Write("Checking Authentication<br>")
Response.Write "LOGON_USER: " & Request.ServerVariables("LOGON_USER") & "<br>"
Response.Write "REMOTE_USER: " & Request.ServerVariables("REMOTE_USER") & "<br>"
Response.Write "AUTH_USER: " & Request.ServerVariables("AUTH_USER") & "<br>"
Response.Write "<br>"

Response.Write("Checking File System Access <br>")
Set fs = Server.CreateObject("Scripting.FileSystemObject")
strPath1 = Server.MapPath("test.asp")
Set rs=fs.GetFile(strPath1)
modified = rs.DateLastModified
response.write(Server.MapPath("test.asp"))
response.write("<br> File Modified on: " & modified)
Set rs=nothing
Set fs=nothing
Response.Write "<br><br>"
Response.Write("Checking for DB access using ODBC <br>")
Set conn = Server.CreateObject("ADODB.Connection")
'Below for Azure Postgres
conn.open "Driver={PostgreSQL ANSI};Server=testapp99.postgres.database.azure.com;Port=5432;Database=postgres;Uid=qadmin@testapp99;Pwd=Admin@123456;"
'Below for oracle
'conn.open "Provider=OraOLEDB.Oracle;User ID=dev1;Password=dev1;Data Source=XE;"
'conn.Open "Provider=MSDAORA;Data Source=XE;User Id=dev1;Password=dev1;"

sql="select ""FIRST_NAME"" from dev1.t_student" 

Set rs = conn.Execute(sql)

Do While Not rs.EOF
Response.Write "<tr>"
response.write("<td>Data From DB: " & rs(0) & "</td>")
Response.Write "</tr>"
rs.MoveNext
Loop

rs.close
conn.close

%>
</body>
