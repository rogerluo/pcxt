<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о����γ�״������ϵͳ</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>
<script language="javascript">
function ResetForm()
{
	var helpstudylen = document.stutotalevaluate.helpstudy.length;
	for (i = 0; i < helpstudylen ; i++)
	{
		document.stutotalevaluate.helpstudy[i].checked = false;
	}
	var helpjoblen = document.stutotalevaluate.helpjob.length;
	for (i = 0; i < helpjoblen ; i++ )
	{
		document.stutotalevaluate.helpjob[i].checked = false;
	}
}
</script>

<body>
<form name="stutotalevaluate" method="post" action="stuchktotalevaluate.asp">
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
	strsql = "select [coursemember_gra].coursedate,[coursemember_gra].coursename,[coursemember_gra].courseid from [coursemember_gra] join [courseinfo_gra] on ([coursemember_gra].stuid = '"+session("stuid_gra")+"' and [courseinfo_gra].courseid = [coursemember_gra].courseid and [courseinfo_gra].coursedate = [coursemember_gra].coursedate)"
	Set rs = conn.execute(strsql)

	'����checkbox��ѡ��
	strsql = "select coursestudy,coursejob from evaluate2_gra where stuid = '"+session("stuid_gra")+"'"
	Set rs2 = conn.execute(strsql)

	If rs2.eof Or rs2.bof Then
		helpstudy = ""
		helpjob = ""
	Else
		helpstudy = rs2("coursestudy")
		helpjob = rs2("coursejob")
		helpstudyarray = Split(helpstudy,",")
		helpjobarray = Split(helpjob,",")
	End If 
	rs2.close
	Set rs2 = Nothing

	If Not (rs.bof And rs.eof) Then
%>
	<table width="600"  border="1" align="center" cellpadding="5" cellspacing="0" >
		<tr><th colspan=2 align="center" class="titlebg">���޿γ��б�</th></tr>
		<tr>
			<td width=30%>�����������а����Ŀγ�</td>
			<td>
				<table width=100% border=1 align="center" cellpadding="3" cellspacing=0>
					<tr align=left>
						<td>
<%
	Do While Not rs.eof
		If helpstudy <> "" Then
			isSelected = 0
			For i = 0 To UBound(helpstudyarray)
				If Trim(helpstudyarray(i)) = Trim(rs("courseid")) Then 
					isSelected = 1
					Exit For
				End If 
			Next
		Else 
			isSelected = 0
		End If 
		If isSelected = 1 Then 
			response.write "<li type='1'>&nbsp;&nbsp;<input type='checkbox' name='helpstudy' value='"+rs("courseid")+"' checked='checked'>"+rs("coursename")+"</input><br>"
		Else 
			response.write "<li type='1'>&nbsp;&nbsp;<input type='checkbox' name='helpstudy' value='"+rs("courseid")+"'>"+rs("coursename")+"</input><br>"			
		End If 
		rs.movenext
	Loop
%>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		
		<tr>
			<td width=30%>���������а����Ŀγ�</td>
			<td>
				<table width=100% border=1 align="center" cellpadding="3" cellspacing=0>
					<tr align=left>
						<td>
<%
	rs.movefirst
	Do While Not rs.eof
		If helpstudy <> "" Then
			isSelected = 0
			For i = 0 To UBound(helpstudyarray)
				If Trim(helpstudyarray(i)) = Trim(rs("courseid")) Then 
					isSelected = 1
					Exit For
				End If 
			Next
		Else 
			isSelected = 0
		End If 
		If isSelected = 1 Then 
			response.write "<li type=1>&nbsp;&nbsp;<input type='checkbox' name='helpjob' value='"+rs("courseid")+"' checked='checked'>"+rs("coursename")+"</input><br>"
		Else 
			response.write "<li type=1>&nbsp;&nbsp;<input type='checkbox' name='helpjob' value='"+rs("courseid")+"'>"+rs("coursename")+"</input><br>"
		End If 
		rs.movenext
	Loop
	rs.close
	Set rs = Nothing
%>
						</td>
					</tr>
				</table>
			</td>
		</tr>

		<tr><th colspan=2 align="center">
			<input type="submit" name="submit" value="�ύ" style="width:80">&nbsp;&nbsp;
			<input type="button" name="reset" value="����" style="width:80" onClick="return ResetForm();">
		</th></tr>
	</table>
</form>
</body>
</html>
<%
	Else 
%>
	<table width="550"  border="1" align="center" cellpadding="5" cellspacing="0">
		<tr align="center"><td  class="titlebg">û���κογ̽�������</td></tr>
	</table>

</form>
</body>
</html>
<%
	End If 
%>
