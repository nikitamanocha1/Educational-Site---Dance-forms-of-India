<%@language=vbscript%>
<HTML>
<HEAD>
<TITLE>Thanx! for visiting our site</TITLE>


</HEAD>

<BODY bgcolor=white text=orange>
<center>
<%
fn=request.form("vname")
ln=request.form("vname1")
if request.form ("R1")="V1" then
	gen="Male"
else
	gen="Female"
end if
age=request.form("aname")
Addr=request.form("add")
em=request.form("email")
ph=request.form("phno")
if request.form ("c1")="ON" then
	qedu=qedu+"Poor,"
end if

if request.form ("c2")="ON" then
	qedu=qedu+"Good"
end if
if request.form ("c3")="ON" then
	qedu=qedu+"Excellent,"
end if

com=request.form("S1")

set conn = server.createobject("adodb.connection")
database = "dance.mdb"
conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & server.mappath(database)

conn.execute("insert into feed values('"+fn+"','"+ln+"','"+gen+"',"+age+",'"+Addr+"','"+em+"',"+ph+",'"+qedu+"','"+com+"')")
Response.Write ("<br>")
Response.Write ("<br>")
Response.Write ("<br>")
Response.Write ("<br>")
Response.Write ("<br>")
%>
Thanx for <% response.write(fn)%> visiting our website. You will be redirected to home page in a five secs. To immediately go back to home page

<a href=home.html>Click here </a>
<META HTTP-EQUIV="refresh" CONTENT="5; URL=home.html">
</center>
</BODY>
</HTML>

