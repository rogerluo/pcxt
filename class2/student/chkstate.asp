<%
	If IsEmpty(session("stuname_gra")) or session("stuname_gra")="" Then
		'response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
		response.write "<script>alert('��δ��¼���ߵ�½ʱ�䳬ʱ!');top.window.location.href='stulogin.asp';</script>"
		response.End
	end if
%>