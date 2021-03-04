<%@LANGUAGE="VBSCRIPT" %>
<!DOCTYPE html>
<html>
<body>
<%
Response.Write("Hello World!")
Randomize()
randomNumber=Int(100 * Rnd())
response.write("A random number: <b>" & randomNumber & "</b>")
Response.Write("Session.SessionID" & Session.SessionID)

Set fs = Server.CreateObject("Scripting.FileSystemObject")
strPath1 = Server.MapPath("test.asp")
Set rs=fs.GetFile(strPath1)
modified = rs.DateLastModified
response.write(Server.MapPath("test.asp"))
response.write("<br> File Modified on: " & modified)
Set rs=nothing
Set fs=nothing
%>
</body>
