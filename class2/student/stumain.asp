<!--#include file="chkstate.asp"-->
<!--#include file="../admin/conn.asp"-->
<%
	strsql = "select stuid from evaluate2_gra where stuid = '"+session("stuid_gra")+"'"
	Set rs = conn.execute(strsql)
	If rs.eof Or rs.bof Then 
		isevaluate = 0
	Else 
		isevaluate = 1
	End If 
	rs.close
	Set rs = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о����γ�״������ϵͳ</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>

<body>
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
<table width="600" border="0" align="center" cellpadding="6" cellspacing="0">
  <tr>
    <td>
		<li><a href="stutotalevaluate.asp">�γ���������</a>
<%
	If isevaluate = 0 Then 
		response.write "(��û����)"
	Else 
		response.write "(������)"
	End If 
%>	
	</td>
  </tr>	
  <tr>
    <td><li><a href="stuevaluate.asp">��û���ۿγ�</a></td>
  </tr>
  <tr>
    <td><li><a href="stulist.asp">�鿴�����γ�</td>
  </tr>
</table>
</body>
</html>