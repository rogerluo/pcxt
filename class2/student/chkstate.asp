<%
	If IsEmpty(session("stuname_gra")) or session("stuname_gra")="" Then
		'response.write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
		response.write "<script>alert('尚未登录或者登陆时间超时!');top.window.location.href='stulogin.asp';</script>"
		response.End
	end if
%>