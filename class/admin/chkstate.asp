<%
	If IsEmpty(session("adminname")) or session("adminname")="" then
		response.write "<script>alert('��δ��¼���ߵ�½ʱ�䳬ʱ��');top.window.location.href='adminlogin.asp';</script>"
		response.End
	end if
%>