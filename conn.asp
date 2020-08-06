<%
Server.scriptTimeout="10"
connstr="DBQ="+server.mappath("bbs.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=Server.CreateObject("ADODB.connection")
conn.open connstr
%>