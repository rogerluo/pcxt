<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	'��ȡ�γ̺��Լ��γ̿���ʱ��
	courseid = Trim(request.querystring("courseid"))
	coursedate = Trim(request.querystring("coursedate"))
	'����У��
	If courseid = "" Or Not IsNumeric(courseid) or coursedate = "" Then 
		conn.close
		Set conn = Nothing 
		response.write "<script>alert('����ҳ�����');window.location.href='adminlist.asp';</script>"
		response.end
	End If 

	strsql = "select * from [courseinfo_gra] where courseid ='"+courseid+"' and coursedate ='"+coursedate+"'"
	Set rs = conn.execute(strsql)
	
	coursename = CStr(rs("coursename"))
	coursetime = CStr(rs("coursetime"))
	coursecredit = cstr(rs("coursecredit"))
	courseteacher = cstr(rs("courseteacher"))
	coursetype = cstr(rs("coursetype"))
	coursetype2 = cstr(rs("coursetype2"))

	courseheadcount = cstr(rs("courseheadcount"))

	rs.close()
	Set rs = Nothing 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META   HTTP-EQUIV="Pragma"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Cache-Control"   CONTENT="no-cache">    
<META   HTTP-EQUIV="Expires"   CONTENT="0">    
<title>������Ϣ����ѧԺ�о����γ�״������ϵͳ</title>
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
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о����γ�״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>����ã�</td>
    <td align="right">��<a href="adminlist.asp">�γ��б�</a>��&nbsp;&nbsp;��<a href="adminmain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="adminlogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
<table width="800" border="0" align="center" cellpadding="5" cellspacing="0" class="tbbg"">
 <tr> 
      <td align="center" class="titlebg">�γ���ϸ��Ϣ&nbsp;&nbsp;��<a href="admintotalevaluate.asp?courseid=<%=courseid%>&coursedate=<%=coursedate%>">�鿴����ͳ����Ϣ</a>��</td>
 </tr>
 <tr>
	<td>
		<table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
			<tr bgcolor=#cccccc>
				<td width=400>�γ���&nbsp;&nbsp;��<%=coursename%></td>
				<td width=200>�ڿν�ʦ��<%=courseteacher%></td>
				<td>�γ����<%=coursetype%></td>
			</tr>
			<tr bgcolor=#cccccc>
				<td width=400>�ڿ�������<%=courseheadcount%></td>
				<td width=200>���&nbsp;&nbsp;��<%=coursetype2%></td>
				<td>ѧʱ��<%=coursetime%>&nbsp;&nbsp;ѧ�֣�<%=coursecredit%></td>
			</tr>
		</table>
	</td>
 </tr>
  <tr> 
      <td class="txtbg"> 
	  <table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
		<tr  align="center" bgcolor=#cccccc>
			<td width=50>���</td>
			<td width=100>ѧ��</td>
			<td width=100>����</td>
			<td width=200>רҵ</td>
			<td width=150>Ժϵ</td>
			<td width=70>�Ƿ�����</td>
			<td>��������</td>
		</tr>
<%
	strsql = "select * from [coursemember_gra] where courseid ='"+courseid+"' and coursedate = '"+coursedate+"'"
	Set rs = conn.execute(strsql)	

	If rs.eof And rs.bof Then
		response.write ""
	End If 
	i = 1
	Do While Not rs.eof
		If i Mod 2 = 0 Then 
			response.write "<tr align=center bgcolor=#cccccc>"
			response.write"<td align=center>"
			'====================
			If rs("isevaluate") = 0 Then 
				str = "��"
				response.write CStr(i) +"</td><td>"+ rs("stuid") +"</td><td>" + rs("stuname") +"</td><td>" +rs("stuobject")+"</td><td>"+rs("stucollege")+"</td><td>"+str+"</td><td><font color=gray>�鿴</font>"

			Else 
				str = "��"
				response.write CStr(i) +"</td><td>"+ rs("stuid") +"</td><td>" + rs("stuname") +"</td><td>" +rs("stuobject")+"</td><td>"+rs("stucollege")+"</td><td>"+str+"</td><td><a href='adminpersonevaluate.asp?courseid="+rs("courseid")+"&stuid="+rs("stuid")+"&coursedate="+cstr(coursedate)+"'>�鿴</a>"

			End If 
		Else 
			response.write "<tr align=center>"
			response.write"<td align=center>"
			'====================
			If rs("isevaluate") = 0 Then 
				str = "��"
				response.write CStr(i) +"</td><td>"+ rs("stuid") +"</td><td>" + rs("stuname") +"</td><td>" +rs("stuobject")+"</td><td>"+rs("stucollege")+"</td><td>"+str+"</td><td><font color=gray>�鿴</font>"

			Else 
				str = "��"
				response.write CStr(i) +"</td><td>"+ rs("stuid") +"</td><td>" + rs("stuname") +"</td><td>" +rs("stuobject")+"</td><td>"+rs("stucollege")+"</td><td>"+str+"</td><td><a href='adminpersonevaluate.asp?courseid="+rs("courseid")+"&stuid="+rs("stuid")+"&coursedate="+cstr(coursedate)+"'>�鿴</a>"

			End If 

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
