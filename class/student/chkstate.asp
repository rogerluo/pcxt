<%
	If IsEmpty(session("stuname")) or session("stuname")="" then
		response.write "<script>alert('��δ��¼���ߵ�½ʱ�䳬ʱ��');top.window.location.href='stulogin.asp';</script>"
		response.End
	end if
%>