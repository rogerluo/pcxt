<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<%
	Set rs = server.CreateObject("Adodb.Recordset")
	strsql="select sysstate from [adminuser]  where adminname='"&session("adminname")&"'"
	Set rs = conn.execute(strsql)
	sysstate = rs("sysstate")
	If request.Form("state")<>"" Then 
		strsql="update [adminuser] set sysstate="&CInt(request.Form("state"))&" where adminname='"&session("adminname")&"'"
		conn.execute(strsql)
		session("sysstate") = CInt(request.Form("state"))
		rs.close
		Set rs=Nothing
		conn.close
		set conn = nothing
		response.write "<script>alert('�����޸ĳɹ���');window.location.href='adminmain.asp';</script>"
	Else
		rs.close
		Set rs=Nothing
		conn.close
		Set conn=nothing
	End If 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о����γ�״������ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="systemstate" method="post" action="">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">������Ϣ����ѧԺ�о����γ�״������ϵͳ</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>����ã�</td>
    <td align="right">��<a href="adminmain.asp">ϵͳ��ҳ</a>��&nbsp;&nbsp;��<a href="adminlogout.asp">�˳�ϵͳ</a>��&nbsp;</td>
  </tr>
</table>
<hr>
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td><li><a href="adminlist.asp">�γ��б�</a></td>
  </tr>
  <tr>
    <td><li><input type="submit" name="Submit" value="<%
		If sysstate = 1 Then
			response.write "�ر�"
		Else
			response.write "����"
		End If
	%>ϵͳ"><input name="state" type="hidden" id="state" value="<%
			If sysstate = 1 Then
				response.write "0" 
			Else
				response.write "1" 
			End If%>"></td>
  </tr>
  <tr>
    <td><li><a href="adminupload.asp">�ϴ��ļ�</td>
  </tr>
</table>
</form>
</body>
</html>
