<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>电子信息工程学院研究生教学状况调查系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{
	if(document.adminlogin.adminname.value==""){
		alert("请输入用户名！");
		document.adminlogin.adminname.focus();
		return false;
	}
	if(document.adminlogin.password.value == ""){
		alert("请输入密码！");
		document.adminlogin.password.focus();
		return false;
	}
}
</script>
<body><br><br><br>
<form name="adminlogin" method="post" action="adminchklogin.asp" onSubmit="return CheckForm();">
<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="50" align="center" class="title">电子信息工程学院研究生教学状况调查系统</td>
    </tr>
</table>
  <table width="300" border="0" align="center" cellpadding="5" cellspacing="0" class="tbbg">
    <tr> 
      <td align="center" class="titlebg">管理员登录</td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF" class="txtbg"> 
	  <table  width="100%" border="0" align="center" cellpadding="4" cellspacing="1">
          <tr> 
            <td align="right">用户名：</td>
            <td><input name="adminname" type="text" id="adminname" style="width:140"></td>
          </tr>
          <tr> 
            <td align="right">密&nbsp;&nbsp;码：</td>
            <td><input name="password" type="password" id="password" style="width:140"></td>
          </tr>
          <tr> 
            <td height="40" colspan="2" align="center"><input type="submit" name="Submit" value="登&nbsp;录">
                &nbsp; <input  type="reset" name="Submit2" value="重&nbsp;填"></td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
</body>
</html>
