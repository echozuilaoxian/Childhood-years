<%@ CODEPAGE ="936"%>
<!--#include file="conn.asp"-->
<html>
<head>
<title>帖子管理--童真岁月论坛</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<br>
<table width="500" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="402" height="21" bgcolor="#F5F5F5"> 
      <div align="center"><font size="2" color="#000000">::::帖子标题::::</font></div>
    </td>
    <td width="98" height="21" bgcolor="#F5F5F5"><font size="2" color="#000000">::管理::</font></td>
  </tr>
</table>
<% idd=request.querystring("id")%>
<%
	set co=server.createobject("adodb.recordset")
	co.open "select count(id) as sum from wenzhang where boardid="&idd&"",conn,1,3
	if not co.eof then
%>       
<%
dim rs
dim sql
set rs = server.createobject("adodb.recordset")
sql = "select*from wenzhang where bid=0 and boardid="&idd&" order by id desc"
count=co("sum")
on error resume next
pagesetup=30
rs.Open sql,conn,1
If Count/pagesetup > (Count\pagesetup) then
TotalPage=(Count\pagesetup)+1
else TotalPage=(Count\pagesetup)
End If
PageCount= 0
RS.MoveFirst
if Request.QueryString("ToPage")<>"" then PageCount = cint(Request.QueryString("ToPage"))
if PageCount <=0 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
RS.Move (PageCount-1) * pagesetup
i=1
do while not rs.eof
%>
<table width="500" border="1" bordercolorlight="#FFFFFF" bordercolordark="#000000" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="398"><font color="#000000" size="2"><%=rs("title")%></font></td>
    <td width="96"> 
      <div align="center"><font size="2">[<a href="showmodify.asp?id=<%=rs("id")%>">修改</a>][<a href="showdel.asp?id=<%=rs("id")%>">删除</a>]</font></div>
    </td>
  </tr>
</table>
<% i=i+1 
if i>pagesetup then exit do 
rs.movenext 
loop 
rs.Close 
%>
<br>
<table width="451" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <div align="right"><font color="#FF0000"><b><font size="2"> 
        <% idd=request.querystring("id")%>
		<% 
ii=PageCount-5 
iii=PageCount+5 
if ii < 1 then 
ii=1 
end if 
if iii > TotalPage then 
iii=TotalPage 
end if 
if PageCount > 6 then 
Response.Write "<a href=?id="&idd&"&topage=1><font color=black>1</font></a> ... " 
end if  
for i=ii to iii 
If i<>PageCount then 
Response.Write "<a href=?id="&idd&"&topage="& i &"><font color=black>" & i & "</font></a> " 
else 
Response.Write " <font color=red><b>"&i&"</b></font> " 
end if 
next 
 
if TotalPage > PageCount+5 then 
Response.Write " ... <a href=?id="&idd&"&topage="&TotalPage&"><font color=black>"&TotalPage&"</font></a>" 
end if 
%>
<%
	end if
	co.close
	set co=nothing
	%>        </font></b></font></div>
    </td>
  </tr>
</table>
</body>
</html>
