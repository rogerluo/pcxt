<% 

	dim conn,connstr
	connstr ="driver={SQL Server};Server=.;database=course;uid=eaie2006;pwd=qwe@123" 

'	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection") 
	conn.open connstr

	If Err Then
		err.Clear
		Set conn = Nothing
		Response.Write("���ݿ����ӳ������������ִ���")
		Response.End
	End If
%>