<% 
'===================================================================================================
'
'********************  ������������ͳ�Ƴ��� V1.0  ******************************
'
'   ��վ������  �����Ṥ���ҡ�--����ܽ���   ���������������ر��˵��Ͷ��ɹ�
'                                         ��                           
'                   �����Ṥ���ҡ����������ʵ���������                 
'   
'*  ����:����ܽ���                                                           ��
'*  ��վ:http://www.blbcn.com                                                �� 
'*  �����ʼ�:bingqilinjianke@163.com                                         �� 
'*  QQ:68156987                                                               
'*  סַ����������                                                            
'******************************************************************************
'
'     ��Ȩ���С���ϮŲ�ñؾ�
'
'=================================================================================================== 


'---------�������
    dim rs
    dim ip
    dim timeout
    dim x
	dim conn
	dim dbpath
'---------�����������

'---------�������ݿ�����
   	set conn=server.createobject("adodb.connection")
	DBPath = Server.MapPath("zai.mdb")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
'---------�������ݿ��������	

'---------�½����ݿ�RS����
set rs = server.createobject("adodb.recordset")
'---------�������ݿ�RS�������

'---------��ȡ�ͻ���IP��ַ
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")      '���ǶԷ�ʹ�õ��Ǵ��������������Ļ�,��Request.ServerVariables("HTTP_X_FORWARDED_FOR")���Եõ��Է�����ʵIP,���Է�����ͨ�����������������Ļ�,��IP��ֵΪ��
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")   '���IP��ֵΪ��,��õ����ı��ؿͻ��˵�ַ
'---------IP��ȡ���

'---------�������ݿ��ڽ�20�������������������,group by ip-��IPֵ��ͬ�ļ���1
	sql="select ip from zai where time >= dateadd('n',-20,now()) group by ip"
	rs.Open sql,conn,1,1
	zai=rs.RecordCount
	rs.Close
'---------�õ���������ֵ

'---------�鿴���ݿ����Ƿ��Ѿ�����ͬ��ֵ,���� x="yes"û����x="no"	   
	   sql="select ip from zai where ip='" & ip & "'"
	   rs.Open sql,conn,1,1
	   if rs.eof and rs.bof then          
            x="yes"
       else
	        x="no"
	   end if  
	   rs.close
'--------�ж����

'--------������ݿ���û����ͬ��ֵ�����һ����ֵ
	   if x="yes" then
       sql="select top 1 * from zai"
	   rs.Open sql,conn,3,2
	   rs.AddNew
	   rs("ip")=ip
	   rs("time")=now()
	   rs.update
	   rs.close 
	   end if
'--------�жϼ������
	 

'--------ɾ��20������ǰ�������ֵ
timeout = dateadd("n", -20, now())
sql="delete * from zai where time < #" & timeout & "#"
conn.Execute sql
'--------ɾ�����

'--------�ر����ݶ���
set rs=nothing
conn.Close 
set conn=nothing
%>
<%=zai%>