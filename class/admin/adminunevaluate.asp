<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	courseorder = Trim(request.querystring("courseorder"))
	target = Trim(request.querystring("target"))
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" or courseorder = "" or Not IsNumeric(courseorder) Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('����ҳ�����');history.back();</script>"
		response.end
	End If 

	strsql = "select * from [courseinfo] where courseid ='"+courseid+"' and coursedate = '"+coursedate+"' and courseorder="+courseorder
	Set rs = conn.execute(strsql)
	coursename = rs("coursename")
	coursetime = CStr(rs("coursetime"))
	courseteacher = cstr(rs("courseteacher"))
	coursetype = cstr(rs("coursetype"))
	courseheadcount = cstr(rs("courseheadcount"))
	coursecredit = cstr(rs("coursecredit"))
	courseteacher2 = cstr(rs("courseteacher2"))
	rs.close()
	'Set rs = Nothing 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META   HTTP-EQUIV="Pragma"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Cache-Control"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Expires"   CONTENT="0">    
<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
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
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>����ã�</td>
    <td align="right">��<a href="adminlist.asp">�γ��б�</a>��&nbsp;&nbsp;��<a href="adminmain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="adminlogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
<table width="800" border="1" align="center" cellpadding="5" cellspacing="0" class="tbbg"">
 <tr> 
      <td align="center" class="titlebg" bgcolor='#cccccc'>�γ���ϸ��Ϣ&nbsp;&nbsp;��<a href="admintotalevaluate.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>&courseorder=<%=courseorder%>">�鿴����ͳ����Ϣ</a>��</td>
 </tr>
 <tr>
	<td>
		<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
			<tr bgcolor='#cccccc'>
				<td width=400>�γ���&nbsp;&nbsp;��<%=coursename%></td>
				<td width=200>�ڿν�ʦ��<%=courseteacher%></td>
				<td>�ڿ����ͣ�<%=coursetype%></td>
			</tr>
			<tr bgcolor='#cccccc'>
				<td width=400>����<%=courseorder%>&nbsp;&nbsp;�ڿ�������<%=courseheadcount%></td>
				<td width=200>��ν�ʦ��<%=courseteacher2%></td>
				<td>ѧʱ��<%=coursetime%>&nbsp;&nbsp;ѧ�֣�<%=coursecredit%></td>
			</tr>
		</table>
	</td>
 </tr>
  <tr> 
      <td class="txtbg"> 
	  <table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
		<tr  align="center" bgcolor='#cccccc'>
			<td width=50>���</td>
			<td width=100>ѧ��</td>
			<td width=100>����</td>
			<td width=200>רҵ</td>
			<td width=150>Ժϵ</td>
			<td width=70>�Ƿ�����</td>
			<td>��������</td>
		</tr>
<%
	If target = 0 Then 
		strsql = "select * from [coursemember] where courseid ='"+courseid+"' and isevaluate=1 and coursedate = '"+coursedate+"' and courseorder="+courseorder
	Else 
		strsql = "select * from [coursemember] where courseid ='"+courseid+"' and isevaluate=0 and coursedate = '"+coursedate+"' and courseorder="+courseorder
	End If 
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
			str = "��"
			response.write CStr(i) +"</td><td>"+ rs("stuid") +"</td><td>" + rs("stuname") +"</td><td>" +rs("stuobject")+"</td><td>"+rs("stucollege")+"</td><td>"+str+"</td><td><font color=gray>�鿴</font>"

		Else 
			str = "��"
			response.write CStr(i) +"</td><td>"+ rs("stuid") +"</td><td>" + rs("stuname") +"</td><td>" +rs("stuobject")+"</td><td>"+rs("stucollege")+"</td><td>"+str+"</td><td><a href='adminpersonevaluate.asp?courseid="+rs("courseid")+"&stuid="+rs("stuid")+"&coursedate="+cstr(coursedate)+"&courseorder="+cstr(courseorder)+"'>�鿴</a>"

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