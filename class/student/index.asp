<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
<link href="../admin/style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{
	if(document.stulogin.stuname.value==""){
		alert("������������");
		document.form1.stuname.focus();
		return false;
	}
	if(document.stulogin.stuid.value == ""){
		alert("������ѧ�ţ�");
		document.form1.stuid.focus();
		return false;
	}
}
</script>
<body><br><br><br>
<form name="stulogin" method="post" action="login.asp" onSubmit="return CheckForm();">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="50" align="center" class="title">������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</td>
    </tr>
</table>
  <table width="300" border="0" align="center" cellpadding="5" cellspacing="0" class="tbbg">
    <tr> 
      <td bgcolor="#FFFFFF" class="txtbg"> 
	  <table  width="100%" border="0" align="center" cellpadding="6" cellspacing="1">
          <tr> 
            <td align="right">��&nbsp;����</td>
            <td><input name="stuname" type="text" id="stuname" style="width:150"></td>
          </tr>
          <tr> 
            <td align="right">ѧ&nbsp;�ţ�</td>
            <td><input name="stuid" type="text" id="stuid" style="width:150"></td>
          </tr>
          <tr> 
            <td height="40" colspan="2" align="center"><input type="submit" name="Submit" value="��&nbsp;¼">
                &nbsp; <input  type="reset" name="Submit2" value="��&nbsp;��"></td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
</body>
</html>