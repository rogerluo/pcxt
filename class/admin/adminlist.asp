<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<%
	coursedate = trim(request.querystring("coursedate"))
	'response.write coursedate
%>
<body>

<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>����ã�</td>
    <td align="right">��<a href="adminmain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="adminlogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
<form action="" name="choosdate">
<table width="800" border="1" align="center" cellpadding="5" cellspacing="0" class="tbbg"">
 <tr> 
      <td align="center" class="titlebg" bgcolor=#cccccc>�γ��б�
<%
	dim datearray()
	strsql = "select distinct([courseinfo].coursedate) from [courseinfo]"
	set rs = conn.execute(strsql)
	if rs.eof and rs.bof then 
%>      	
      	
      </td>
 </tr>
<%
	else
		
		response.write "<tr><td>�γ̿���ʱ��ѡ��"
		response.write "<select name='date' style='width:140' onchange='top.location=this.value;'>"
		response.write "<option value='' selected='true' OnChange='javascript:var sid = this.options.value;doselect(sid);'>��ѡ��</option>"
		do while not rs.eof
			response.write "<option value=adminlist.asp?coursedate="+rs("coursedate")+">"+rs("coursedate")+"</option>"
			rs.movenext
		loop
		rs.close()
		set rs = nothing
		response.write "</select></td></tr>"
	end if
	if coursedate = "" then 
		strsql = "select * from [courseinfo]"
	else 
		strsql = "select * from [courseinfo] where coursedate ='"+coursedate+"'"
	end if
%>
  <tr> 
      <td bgcolor="#FFFFFF" class="txtbg"> 
	  <table  width="100%" border="1" align="center" cellpadding="3" cellspacing="0">
		<tr  align="center" bgcolor=#cccccc>
			<td width=70>����ʱ��</td>
			<td width=70>�γ̱��</td>
			<td width=240>�γ�����</td>
			<td width=30>����</td>
			<td width=30>ѧʱ</td>
			<td width=70>�ڿν�ʦ</td>
			<td width=80>�ڿ�����</td>
			<td width=50>�ڿ�����</td>
			<td>�γ���ϸ</td>
		</tr>
<%
	
	Set rs = conn.execute(strsql)	
	If rs.eof And rs.bof Then
		response.write ""
	End If 
	
	i = 0
	Do While Not rs.eof
		If i Mod 2 = 0 Then
			response.write "<tr>"
			response.write"<td align=center>"
			'====================
			response.write rs("coursedate")+"</td><td align=center>"+rs("courseid") +"</td><td>"+rs("coursename")+"</td><td align=center>"+Cstr(rs("courseorder"))+"</td><td align=center>"+CStr(rs("coursetime"))+"</td><td align=center>"+rs("courseteacher")+"</td><td align=center>"+rs("coursetype")+"</td><td align=center>"+cstr(rs("courseheadcount"))+"</td><td align=center>"+"<a href='adminlistdetail.asp?courseid="+rs("courseid")+"&coursedate="+rs("coursedate")+"&courseorder="+cstr(rs("courseorder"))+"'>�γ���ϸ</a>"
			'====================
			response.write"</td></tr>"
		Else 
			response.write "<tr bgcolor=#cccccc>"
			response.write"<td align=center>"
			'====================
			response.write rs("coursedate")+"</td><td align=center>"+rs("courseid") +"</td><td>"+rs("coursename")+"</td><td align=center>"+Cstr(rs("courseorder"))+"</td><td align=center>"+CStr(rs("coursetime"))+"</td><td align=center>"+rs("courseteacher")+"</td><td align=center>"+rs("coursetype")+"</td><td align=center>"+cstr(rs("courseheadcount"))+"</td><td align=center>"+"<a href='adminlistdetail.asp?courseid="+rs("courseid")+"&coursedate="+rs("coursedate")+"&courseorder="+cstr(rs("courseorder"))+"'>�γ���ϸ</a>"
			'====================
			response.write"</td></tr>"
		End If 
		rs.movenext
		i = i + 1
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
</form>
</body>
</html>
