<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function OnClicked(coursedate){
	alert(String(coursedate));
}
</script>
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
<table width="400" border="1" align="center" cellpadding="5" cellspacing="0">
	<tr>
		<td align="center" class="titlebg" bgcolor=#cccccc>ѧ�ڿγ��б�</td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" align="center" cellpadding="1" cellspacing="0">
				<tr align=center bgcolor=#cccccc>
					<td width=40>���</td>
					<td width=30%>ѧ������</td>
					<td>��������</td>
				</tr>
<%	
	count = 1
	strsql = "select distinct(coursedate) from courseinfo"
	Set rs = conn.execute(strsql)
	If (Not rs.eof) And (Not rs.bof) Then
		Do While Not rs.eof
			If count Mod 2 = 0 Then 
				response.write "<tr align=center  bgcolor=#cccccc><td>"+CStr(count)+"</td><td>"+rs("coursedate")
			Else 
				response.write "<tr align=center><td>"+CStr(count)+"</td><td>"+rs("coursedate")
			End If 
			response.write "</td><td><a href='adminchkalloutput.asp?coursedate="+rs("coursedate")+"'>��������</a></td></tr>"
			count = count + 1
			rs.movenext
		Loop
	End If 
	rs.close
	Set rs = Nothing
%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
