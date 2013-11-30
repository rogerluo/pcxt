<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="connxls.asp"-->
<%
	If session("sysstate") = 1 Then 
		conn.close()
		set conn = Nothing
		response.write "<script>alert('系统尚没关闭！');window.location.href='adminmain.asp';</script>"
		response.end
	End If

%>
	
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>电子信息工程学院研究生课程状况调查系统</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{
	if(document.upload.coursefile.value == ""){
		alert("请输入课程计划文件！");
		document.upload.coursefile.focus();
		return false;
	}
	if(document.upload.coursemember.value == ""){
		alert("请输入选课名单文件！");
		document.upload.coursemember.focus();
		return false;
	}
}
</script>

<body>

<table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="60" colspan="2" align="center" class="title">电子信息工程学院研究生课程状况调查系统</td>
  </tr>
  <tr>
    <td>&nbsp;<%=session("adminname")%>，你好！</td>
    <td align="right">【<a href="adminmain.asp">系统首页</a>】&nbsp;&nbsp;【<a href="adminlogout.asp">退出系统</a>】&nbsp;</td>
  </tr>
</table>
<hr>

<form name="upload" method="post" action="adminchkupload.asp" onSubmit="return CheckForm();" enctype="multipart/form-data">
  <table width="400" border="0" align="center" cellpadding="5" cellspacing="0" class="tbbg">
    <tr> 
      <td align="center" class="titlebg">课程计划上传</td>
    </tr>
	
    <tr> 
      <td bgcolor="#FFFFFF" class="txtbg"> 
	  <table  width="100%" border="1" align="center" cellpadding="4" cellspacing="1">
			<tr>
				<td align="right">上传路径：</td>
				<td align=left>
					<input name="coursefile" type="file" style="width:215">
				</td>
			</tr>
          
      </table>
      </td>
    </tr>
    <tr> 
      <td align="center" class="titlebg">选课名单上传</td>
    </tr>
    <tr>
    	<td>
    		<table  width="100%" border="1" align="center" cellpadding="4" cellspacing="1">
    			<tr>
						<td align="right">上传路径：</td>
						<td align=left>
							<input name="coursemember" type="file" style="width:215">
						</td>
					</tr>
    		</table>
    	</td>
  	</tr>
  	<tr> 
       <td height="40" colspan="2" align="center">
       		<input type="submit" name="Submit" value="上&nbsp;传">
          &nbsp; <input  type="reset" name="Submit2" value="重&nbsp;填">
       </td>
    </tr>
  </table>
</form>

</body>
</html>