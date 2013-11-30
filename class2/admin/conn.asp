<% 

	dim conn,connstr
	connstr ="driver={SQL Server};Server=.;database=course;uid=eaie2006;pwd=qwe@123" 

'	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection") 
	conn.open connstr

	If Err Then
		err.Clear
		Set conn = Nothing
		Response.Write("数据库连接出错，请检查连接字串。")
		Response.End
	End If
%>