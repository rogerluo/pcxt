<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	courseorder = Trim(request.querystring("courseorder"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" or courseorder = "" or Not IsNumeric(courseorder) Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('请求页面错误！');window.location.href='adminlist.asp';</script>"
		response.end
	End If 

	strsql = "select * from [courseinfo] where courseid ='"+courseid+"' and coursedate ='"+coursedate+"' and courseorder = "+ courseorder
	Set rs = conn.execute(strsql)
	If rs.eof Or rs.bof Then
		rs.close
		response.write "<script>alert('请求页面错误！');window.location.href='adminlist.asp';</script>"
		responsw.End 
	Else 
		coursename = rs("coursename")
		coursetime = CStr(rs("coursetime"))
		courseteacher = cstr(rs("courseteacher"))
		coursetype = cstr(rs("coursetype"))
		courseheadcount = cstr(rs("courseheadcount"))
		coursecredit = cstr(rs("coursecredit"))
		courseteacher2 = cstr(rs("courseteacher2"))
	End If 

	'coursehost = rs("coursehost")
	'coursewhere = rs("coursewhere")
	rs.close()
	Set rs = Nothing 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META   HTTP-EQUIV="Pragma"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Cache-Control"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Expires"   CONTENT="0">    
<title>电子信息工程学院研究生教学状况调查系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
<%
  Response.Buffer   =   True    
  Response.ExpiresAbsolute   =   Now()   -   1    
  Response.Expires   =   0    
  Response.CacheControl   =   "no-cache"    
  Response.AddHeader   "Pragma",   "No-Cache"  
%>
</head>

<body>

<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生教学状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>，你好！</td>
    <td align="right">【<a href="adminlist.asp">课程列表</a>】&nbsp;&nbsp;【<a href="adminmain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="adminlogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>
<table width="800" border="0" align="center" cellpadding="5" cellspacing="0" class="tbbg"">
 <tr> 
      <td align="center" class="titlebg">课程详细信息&nbsp;&nbsp;【<a href="admintotalevaluate.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>">查看评价统计信息</a>】</td>
 </tr>
 <tr>
	<td>
		<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
			<tr bgcolor='#cccccc'>
				<td width=400>课程名&nbsp;&nbsp;：<%=coursename%></td>
				<td width=200>授课教师：<%=courseteacher%></td>
				<td>授课类型：<%=coursetype%></td>
			</tr>
			<tr bgcolor='#cccccc'>
				<td width=400>课序：<%=courseorder%>&nbsp;&nbsp;授课人数：<%=courseheadcount%></td>
				<td width=200>替课教师：<%=courseteacher2%></td>
				<td>学时：<%=coursetime%>&nbsp;&nbsp;学分：<%=coursecredit%></td>
			</tr>
		</table>
	</td>
 </tr>
  <tr> 
      <td class="txtbg"> 
	  <table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
		<tr  align="center" bgcolor=#cccccc>
			<td width=50>序号</td>
			<td width=100>学号</td>
			<td width=100>姓名</td>
			<td width=200>专业</td>
			<td width=150>院系</td>
			<td width=70>是否评价</td>
			<td>评价内容</td>
		</tr>
<%
	strsql = "select * from [coursemember] where courseid ='"+courseid+"' and coursedate = '"+coursedate+"' and courseorder ="+ courseorder
	Set rs = conn.execute(strsql)
	
	If rs.eof And rs.bof Then
		response.write ""
	End If 
	i = 1
	Do While Not rs.eof
		If i Mod 2 = 0 Then 
			response.write "<tr align=center bgcolor='#cccccc'>"
		Else 
			response.write "<tr align=center>"
		End If 
		response.write"<td align=center>"
		'====================
		If rs("isevaluate") = 0 Then 
			str = "否"
			response.write CStr(i) +"</td><td>"+ rs("stuid") +"</td><td>" + rs("stuname") +"</td><td>" +rs("stuobject")+"</td><td>"+rs("stucollege")+"</td><td>"+str+"</td><td><font color=gray>查看</font>"

		Else 
			str = "是"
			response.write CStr(i) +"</td><td>"+ rs("stuid") +"</td><td>" + rs("stuname") +"</td><td>" +rs("stuobject")+"</td><td>"+rs("stucollege")+"</td><td>"+str+"</td><td><a href='adminpersonevaluate.asp?courseid="+rs("courseid")+"&stuid="+rs("stuid")+"&coursedate="+cstr(coursedate)+"'>查看</a>"

		End If 
		'====================
		i = i + 1
		response.write"</td></tr>"
		rs.movenext
	Loop
	rs.close()
	Set rs = Nothing
	conn.close()
	Set conn = Nothing
%>		
	  </table>
	  </td>
 </tr>
</table>
</body>
</html>
