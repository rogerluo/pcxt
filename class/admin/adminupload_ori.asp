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
<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<script language=javascript>
function CheckForm()
{

	if(document.upload.courseid.value==""){
		alert("������γ̺ţ�");
		document.upload.courseid.focus();
		return false;
	}
	if(document.upload.coursename.value == ""){
		alert("������γ����ƣ�");
		document.upload.coursename.focus();
		return false;
	}
	if(document.upload.coursetime.value == ""){
		alert("������ѧʱ��");
		document.upload.coursetime.focus();
		return false;
	}
	if(document.upload.courseteacher.value == ""){
		alert("�������ڿν�ʦ��");
		document.upload.courseteacher.focus();
		return false;
	}
	if(document.upload.coursetype.value == ""){
		alert("�������ڿ����ͣ�");
		document.upload.coursetype.focus();
		return false;
	}
	if(document.upload.coursehost.value == ""){
		alert("�����뿪��ϵ��");
		document.upload.coursehost.focus();
		return false;
	}
	if(document.upload.coursewhere.value == ""){
		alert("�������Ͽ�ʱ��ص㣡");
		document.upload.coursewhere.focus();
		return false;
	}

	if(document.upload.coursefile.value == ""){
		alert("�������ϴ�·����");
		document.upload.coursefile.focus();
		return false;
	}

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

<form name="upload" method="post" action="adminchkupload.asp" onSubmit="return CheckForm();">
  <table width="400" border="0" align="center" cellpadding="5" cellspacing="0" class="tbbg">
    <tr> 
      <td align="center" class="titlebg">�γ������ϴ�</td>
    </tr>
	
    <tr> 
      <td bgcolor="#FFFFFF" class="txtbg"> 
	  <table  width="100%" border="1" align="center" cellpadding="4" cellspacing="1">
		
          <tr> 
            <td align="right">�γ̺�&nbsp&nbsp��</td>
            <td><input name="courseid" type="text" id="courseid" style="width:140"></td>
          </tr>
          <tr> 
            <td align="right">�γ����ƣ�</td>
            <td><input name="coursename" type="text" id="coursename" style="width:140"></td>
          </tr>
			<tr>
				<td align="right">ѧʱ&nbsp&nbsp&nbsp&nbsp��</td>
				<td>
					<input name="coursetime" type="text" style="width:140">
				</td>
			</tr>
			<tr>
				<td align="right">�ڿ���ʦ��</td>
				<td>
					<input name="courseteacher" type="text" style="width:140">
				</td>
			</tr>
			<tr>
				<td align="right">�ڿ����ͣ�</td>
				<td align=left>
					<input name="coursetype" type="text" style="width:140">
				</td>
			</tr>
			<tr>
				<td align="right">����ϵ&nbsp&nbsp��</td>
				<td align=left>
					<input name="coursehost" type="text" style="width:140">
				</td>
			</tr>
			<tr>
				<td align="right">�Ͽεص㣺</td>
				<td align=left>
					<input name="coursewhere" type="text" style="width:140">
				</td>
			</tr>
		
			<tr>
				<td align="right">�ϴ�·����</td>
				<td align=left>
					<input name="coursefile" type="file" style="width:215">
				</td>
			</tr>
          <tr> 
            <td height="40" colspan="2" align="center"><input type="submit" name="Submit" value="��&nbsp;��">
                &nbsp; <input  type="reset" name="Submit2" value="��&nbsp;��"></td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>

</body>
</html>