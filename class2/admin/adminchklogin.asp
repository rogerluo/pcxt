<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%
	adminname=replace(trim(request.form("adminname")),"'","")
	password=md5(trim(request.form("password")))
	
	'response.write password+"<p>"
	'response.write adminname

	if adminname = "" Or password=""  then
		conn.close
		set conn = nothing
		response.Redirect "adminlogin.asp"
		response.end
	end if

	strsql = "select * from [adminuser]  where adminname='"&adminname&"' and password='"&password&"'"
	set rs = conn.execute(strsql)

	if Not (rs.eof And rs.bof) then
		session("adminname") = rs("AdminName")
		session("sysstate") = rs("sysstate")
		session.Timeout=30
		set rs=nothing
		conn.close
		set conn = nothing
		response.Redirect "adminmain.asp"
	else
		response.write "<script language='javascript'>alert('用户名或密码错误！');history.go(-1);</script>"
		set rs=nothing
		conn.close
		set conn = nothing
	End If
%>