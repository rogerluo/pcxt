<%
	If IsEmpty(session("adminname")) or session("adminname")="" then
		response.write "<script>alert('尚未登录或者登陆时间超时！');top.window.location.href='adminlogin.asp';</script>"
		response.End
	end if
%>