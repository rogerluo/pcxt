<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="stuevaluate" method="post" action="stuchkevaluate.asp">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname")%>����ã�</td>
    <td align="right">��<a href="stumain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="stulogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
<%
	strsql = "select [courseinfo].coursedate,[courseinfo].coursename,[courseinfo].courseteacher,[courseinfo].coursetype,[courseinfo].courseid, [coursemember].courseorder,[coursemember].isevaluate from [coursemember] join [courseinfo] on ([courseinfo].courseid = [coursemember].courseid and [coursemember].stuid='"+session("stuid")+"' and [coursemember].courseorder = [courseinfo].courseorder)"
	Set rs = conn.execute(strsql)
	If Not (rs.bof And rs.eof) Then
%>
	<table width="600"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td class="titlebg" bgcolor="#cccccc">���޿γ��б�</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr align="center"  bgcolor="#cccccc">
						<td width=200 style="align">�γ�����</td>
						<td width=80>�ڿ���ʦ</td>
						<td width=80>�γ�����</td>
						<td width=80>�Ƿ�����</td>
						<td>��������</td>
					</tr>
					
<%
	i = 0
	Do While Not rs.eof
		If i Mod 2 = 0 Then
			response.write "<tr align='center'>"
		Else 
			response.write "<tr align='center' bgcolor='#cccccc'>"
		End If 
		If rs("isevaluate") = 0	Then 
			str = "��"
			response.write "<td align='left'>" + rs("coursename") + "</td><td>" + rs("courseteacher") + "</td><td>" + rs("coursetype") + "</td><td>" + str + "</td><td><a href='stuevaluatedetail.asp?courseid="+rs("courseid")+"&coursedate="+cstr(rs("coursedate"))+"&courseorder="+cstr(rs("courseorder"))+"'>����</a></td></tr>"
		Else 
			str = "��"
			response.write "<td align='left'>" + rs("coursename") + "</td><td>" + rs("courseteacher") + "</td><td>" + rs("coursetype") + "</td><td>" + str + "</td><td><font color=gray>����</font></td></tr>"

		End If 
		rs.movenext
	Loop 
%>
			
				</table>
			</td>
		</tr>
	</table>
</form>
</body>
</html>
<%
	Else 
%>
	<table width="550"  border="1" align="center" cellpadding="0" cellspacing="0">
		<tr align="center"><td bgcolor='#cccccc'>��û�����޿γ�</td></tr>
	</table>

</form>
</body>
</html>
<%
	End If 
%>
