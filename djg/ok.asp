 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <meta http-equiv="Cache-Control" content="no-cache" /><!--ֻ�ǻ����������Ϣ���ܻ���-->
    <meta name="viewport" content="width=device-width" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" /><!--ǿ�����ĵ����豸�Ŀ�ȱ��� 1:1 ��
    �ĵ����Ŀ�ȱ�����1.0�� initial-scale ��ʼ�̶�ֵ�� maximum-scale ���̶�ֵ����user-scalable �����û��Ƿ�����ֶ����ţ� no Ϊ�����ţ���ʹҳ��̶��豸����Ĵ�С��-->
    <meta name="apple-mobile-web-app-capable" content="yes" /><!--��վ������ web app �����֧��-->
    <meta name="apple-mobile-web-app-status-bar-style" content="black" /><!--���ı䶥��״̬������ɫ��-->
    <link href="images/style.css" rel="stylesheet" type="text/css">  
  <title>���ӹ����S</title>

</head>
<body>


   
    
 
<% 
set conn=server.createobject("adodb.connection")
conn.open "driver={microsoft access driver (*.mdb)};dbq="&server.mappath("Databases/#wrtxcnqywz.mdb") 
%> 
 
 
<%if Request.QueryString("mark")="zc" then%>
<%
Set rec=Server.CreateObject("Adodb.Recordset")
sql="Select   * from  liuyan    "   
rec.open sql,conn,1,1


 
xm=Trim(Request("xm"))
dh=Trim(Request("dh"))
mc=Trim(Request("mc"))
sl=Trim(Request("sl"))
sfz=Trim(Request("sfz"))
dizhi=Trim(Request("dizhi"))
zhanghu=Trim(Request("zhanghu"))
 

  
If xm="" Then
response.write"<script>alert('�������ܠ���!');location.href='javascript:history.go(-1)';</script>"
response.end
end if 


If dh="" Then
response.write"<script>alert('�Ԓ���ܠ���!');location.href='javascript:history.go(-1)';</script>"
response.end
end if

 
 


Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from liuyan"
rs.open sql,conn,1,3
rs.addnew
 
rs("xm")=xm
rs("dh")=dh
rs("shijian")=now()
rs("mc")=mc
rs("bz1")=sl
rs("sfz")=sfz
rs("dizhi")=dizhi
rs("zhanghu")=zhanghu
rs.update
rs.close
response.write"<script>alert('������!');location.href='javascript:history.go(-1)';</script>"
response.end
end if
%>
 
    
    
 
 
</body>
</html>
 

