<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о����γ�״������ϵͳ</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="stuevaluate" method="post" action="stuchkevaluate.asp">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о����γ�״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("stuname_gra")%>����ã�</td>
    <td align="right">��<a href="stumain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="stulogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
<%
	strsql = "select [courseinfo_gra].coursedate,[courseinfo_gra].coursename,[courseinfo_gra].courseteacher,[courseinfo_gra].coursetype,[courseinfo_gra].coursetype2,[courseinfo_gra].courseid,[coursemember_gra].isevaluate from [coursemember_gra] join [courseinfo_gra] on ([courseinfo_gra].courseid = [coursemember_gra].courseid and [coursemember_gra].stuid='"+session("stuid_gra")+"' and [coursemember_gra].isevaluate = 0)"
	Set rs = conn.execute(strsql)
	If Not (rs.bof And rs.eof) Then
%>
	<table width="600"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td class="titlebg">���޿γ��б�</td></tr>
		<tr>
			<td>
				<table width="100%"  border="1" align="center" cellpadding="3" cellspacing="0">
					<tr align="center" bgcolor=#cccccc>
						<td width=200>�γ�����</td>
						<td width=80>�ڿ���ʦ</td>
						<td width=80>�γ�����</td>
						<td width=40>���</td>
						<td width=80>�Ƿ�����</td>
						<td>��������</td>
					</tr>
					
<%
	i = 0
	Do While Not rs.eof
		If i Mod 2 = 0 Then 
			If rs("isevaluate") = 0	Then 
				str = "��"
				response.write "<tr align='center'><td>" + rs("coursename") + "</td><td>" + rs("courseteacher") + "</td><td>" + rs("coursetype") + "</td><td>"+rs("coursetype2")+"</td><td>" + str + "</td><td><a href='stuevaluatedetail.asp?courseid="+rs("courseid")+"&coursedate="+cstr(rs("coursedate"))+"'>����</a></td></tr>"
			Else 
				str = "��"
				response.write "<tr align='center'><td>" + rs("coursename") + "</td><td>" + rs("courseteacher") + "</td><td>" + rs("coursetype") + "</td><td>"+ rs("coursetype2") +"</td><td>" + str + "</td><td><font color=gray>����</font></td></tr>"
			End If 
		Else 
			If rs("isevaluate") = 0	Then 
				str = "��"
				response.write "<tr align='center' bgcolor=#cccccc><td>" + rs("coursename") + "</td><td>" + rs("courseteacher") + "</td><td>" + rs("coursetype") + "</td><td>"+rs("coursetype2")+"</td><td>" + str + "</td><td><a href='stuevaluatedetail.asp?courseid="+rs("courseid")+"&coursedate="+cstr(rs("coursedate"))+"'>����</a></td></tr>"
			Else 
				str = "��"
				response.write "<tr align='center' bgcolor=#cccccc><td>" + rs("coursename") + "</td><td>" + rs("courseteacher") + "</td><td>" + rs("coursetype") + "</td><td>"+ rs("coursetype2") +"</td><td>" + str + "</td><td><font color=gray>����</font></td></tr>"
			End If 

		End If
		rs.movenext
		i = i + 1
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
		<tr align="center"><td>��û�����޿γ�</td></tr>
	</table>

</form>
</body>
</html>
<%
	End If 
%>
