<% 
'===================================================================================================
'
'********************  冰翎在线人数统计程序 V1.0  ******************************
'
'   本站程序由  『冰翎工作室』--冰淇淋剑客   独立制作！请尊重本人的劳动成果
'                                         　                           
'                   『冰翎工作室』─━╃→用实力打造个性                 
'   
'*  作者:冰淇淋剑客                                                           　
'*  网站:http://www.blbcn.com                                                　 
'*  电子邮件:bingqilinjianke@163.com                                         　 
'*  QQ:68156987                                                               
'*  住址：江苏无锡                                                            
'******************************************************************************
'
'     版权所有·抄袭挪用必究
'
'=================================================================================================== 


'---------定义变量
    dim rs
    dim ip
    dim timeout
    dim x
	dim conn
	dim dbpath
'---------定义变量结束

'---------建议数据库链接
   	set conn=server.createobject("adodb.connection")
	DBPath = Server.MapPath("zai.mdb")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
'---------建立数据库链接完成	

'---------新建数据库RS对象
set rs = server.createobject("adodb.recordset")
'---------建立数据库RS对象完成

'---------读取客户端IP地址
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")      '若是对方使用的是代理服务器上网的话,用Request.ServerVariables("HTTP_X_FORWARDED_FOR")可以得到对方的真实IP,若对方不是通过代理服务器上网的话,则IP的值为空
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")   '如果IP的值为空,则得到他的本地客户端地址
'---------IP读取完毕

'---------读出数据库内近20分钟所加入的新内容数,group by ip-表IP值相同的记作1
	sql="select ip from zai where time >= dateadd('n',-20,now()) group by ip"
	rs.Open sql,conn,1,1
	zai=rs.RecordCount
	rs.Close
'---------得到在线人数值

'---------查看数据库中是否已经有相同的值,有则 x="yes"没有则x="no"	   
	   sql="select ip from zai where ip='" & ip & "'"
	   rs.Open sql,conn,1,1
	   if rs.eof and rs.bof then          
            x="yes"
       else
	        x="no"
	   end if  
	   rs.close
'--------判断完毕

'--------如果数据库中没有相同的值则加入一个新值
	   if x="yes" then
       sql="select top 1 * from zai"
	   rs.Open sql,conn,3,2
	   rs.AddNew
	   rs("ip")=ip
	   rs("time")=now()
	   rs.update
	   rs.close 
	   end if
'--------判断加入完毕
	 

'--------删除20分钟以前所加入的值
timeout = dateadd("n", -20, now())
sql="delete * from zai where time < #" & timeout & "#"
conn.Execute sql
'--------删除完毕

'--------关闭数据对象
set rs=nothing
conn.Close 
set conn=nothing
%>
<%=zai%>