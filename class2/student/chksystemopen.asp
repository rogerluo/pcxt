<%
	dim connsys,connstrsys
	connstr ="driver={SQL Server};Server=.;database=course;uid=sa;pwd=" 

'	On Error Resume Next
	Set connsys = Server.CreateObject("ADODB.Connection") 
	conn.open connstr

	If Err Then
		err.Clear
		Set connsys = Nothing
		Response.Write("数据库连接出错，请检查连接字串。")
		Response.End
	End If
	
	strsql = "select sysstate from [adminuser]"
	Set rs = connsys.execute(strsql)
	session("sysstate")
	rs.close
	If rs("sysstate") = 0 Then 
		response.write("<script>alert('服务器已经关闭！');top.window.location.href='stulogin.asp';</script>")

%>