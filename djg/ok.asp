 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <meta http-equiv="Cache-Control" content="no-cache" /><!--只是或者请求的消息不能缓存-->
    <meta name="viewport" content="width=device-width" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=0" /><!--强制让文档与设备的宽度保持 1:1 ；
    文档最大的宽度比列是1.0（ initial-scale 初始刻度值和 maximum-scale 最大刻度值）；user-scalable 定义用户是否可以手动缩放（ no 为不缩放），使页面固定设备上面的大小；-->
    <meta name="apple-mobile-web-app-capable" content="yes" /><!--网站开启对 web app 程序的支持-->
    <meta name="apple-mobile-web-app-status-bar-style" content="black" /><!--（改变顶部状态条的颜色）-->
    <link href="images/style.css" rel="stylesheet" type="text/css">  
  <title>代加工工廠</title>

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
response.write"<script>alert('姓名不能爲空!');location.href='javascript:history.go(-1)';</script>"
response.end
end if 


If dh="" Then
response.write"<script>alert('電話不能爲空!');location.href='javascript:history.go(-1)';</script>"
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
response.write"<script>alert('審核中!');location.href='javascript:history.go(-1)';</script>"
response.end
end if
%>
 
    
    
 
 
</body>
</html>
 

