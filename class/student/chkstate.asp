<%
	If IsEmpty(session("stuname")) or session("stuname")="" then
		response.write "<script>alert('尚未登录或者登陆时间超时！');top.window.location.href='stulogin.asp';</script>"
		response.End
	end if
%>