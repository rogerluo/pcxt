<!--#include file="chkstate.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="connxls.asp"-->
<%
	If session("sysstate") = 1 Then 
		conn.close()
		set conn = Nothing
		response.write "<script>alert('ϵͳ��û�رգ�');window.location.href='adminmain.asp';</script>"
		response.end
	End If

%>
	
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������Ϣ����ѧԺ�о����γ�״������ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{
	if(document.upload.coursefile.value == ""){
		alert("������γ̼ƻ��ļ���");
		document.upload.coursefile.focus();
		return false;
	}
	if(document.upload.coursemember.value == ""){
		alert("������ѡ�������ļ���");
		document.upload.coursemember.focus();
		return false;
	}
}
</script>

<body>

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

<form name="upload" method="post" action="adminchkupload.asp" onSubmit="return CheckForm();" enctype="multipart/form-data">
  <table width="400" border="0" align="center" cellpadding="5" cellspacing="0" class="tbbg">
    <tr> 
      <td align="center" class="titlebg">�γ̼ƻ��ϴ�</td>
    </tr>
	
    <tr> 
      <td bgcolor="#FFFFFF" class="txtbg"> 
	  <table  width="100%" border="1" align="center" cellpadding="4" cellspacing="1">
			<tr>
				<td align="right">�ϴ�·����</td>
				<td align=left>
					<input name="coursefile" type="file" style="width:215">
				</td>
			</tr>
          
      </table>
      </td>
    </tr>
    <tr> 
      <td align="center" class="titlebg">ѡ�������ϴ�</td>
    </tr>
    <tr>
    	<td>
    		<table  width="100%" border="1" align="center" cellpadding="4" cellspacing="1">
    			<tr>
						<td align="right">�ϴ�·����</td>
						<td align=left>
							<input name="coursemember" type="file" style="width:215">
						</td>
					</tr>
    		</table>
    	</td>
  	</tr>
  	<tr> 
       <td height="40" colspan="2" align="center">
       		<input type="submit" name="Submit" value="��&nbsp;��">
          &nbsp; <input  type="reset" name="Submit2" value="��&nbsp;��">
       </td>
    </tr>
  </table>
</form>

</body>
</html>