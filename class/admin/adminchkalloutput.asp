<!--#include file="conn.asp"-->
<!--#include file="Cls\ExcelControl.asp"-->
<!--#include file="compress.asp"-->
<%
	'ɾ��ExcelFile�ļ��У�ɾ���������ݣ�
	strFolder = "ExcelFile"
	Call DeleteFolder(strFolder)
	Call CreateFolder(strFolder)

	coursedate = Trim(request.querystring("coursedate"))
	strsql = "select * from courseinfo where coursedate = '"+coursedate+"'"
	Set rs = conn.execute(strsql)
	'�жϼ�¼���Ƿ�Ϊ��
	If rs.eof Or rs.bof Then 
		rs.close
		Set rs = Nothing 
		response.End 
	End If
	'�������γ̴�����Ӧ�����ļ�
	'Set ObjExl=New ExcelMarker
	Do While Not rs.eof
		Set ObjExl=New ExcelMarker
		ObjExl.SetCourseid = CStr(rs("courseid"))
		ObjExl.SetCourseDate = CStr(rs("coursedate"))
		ObjExl.SetCourseOrder = CStr(rs("courseorder"))
		ObjExl.SetItemArray
		
		Set AdoRs=Server.CreateObject("ADODB.RecordSet")
		AdoRs.Open"select * From totalevaluate",conn,1,3

		ObjExl.ADOStreamExcel AdoRs

		rs.movenext
	Loop
	rs.close
	Set rs=Nothing

	'��ǰɾ��ѹ���ļ�
	Set ObjFso=Server.CreateObject("Scripting.FileSystemObject")
	mFileName = Server.Mappath(coursedate+"ѧ�ڿγ�����.rar")
	'�ж��ļ��Ƿ����
	If ObjFso.FileExists(mFileName) Then
		Debug "���������ļ���ɾ��"
		Debug mFileName
		ObjFso.DeleteFile(mFileName)  
	End If
	Set ObjFso = Nothing

	'ѹ��ExcelFile�ļ������ļ�
	Set ObjCompress = new COMPRESS_DECOMPRESS_FILES
	ObjCompress.compressPath=coursedate+"ѧ�ڿγ�����"
	ObjCompress.decompresspath="ExcelFile"
	ObjCompress.compress

	Function DeleteFolder(strFolder) 
	   '��ѡ�ж�Ҫ�������ļ����Ƿ��Ѿ����� 
		Dim strTestFolder, objFSO 
		strTestFolder = Server.Mappath(strFolder) 
		Set objFSO = CreateObject("Scripting.FileSystemObject") 
		' ����ļ����Ƿ���� 
		If objFSO.FolderExists(strTestFolder) Then 
			objFSO.DeleteFolder(strTestFolder) 
			Debug "ɾ���ɹ���"
		Else 
			Debug "���ļ��в����ڣ�"
		End If 
		Set objFSO = Nothing 
	End Function 
	Function CreateFolder(strFolder)
		set objFSO = server.CreateObject("scripting.filesystemobject")
		objFSO.createfolder(server.MapPath(strFolder))
		Debug "�����ɹ���"
	End Function 

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<title>������Ϣ����ѧԺ�о�����ѧ״������ϵͳ</title>
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
	response.write "��������"
	response.write "<a href='"+coursedate+"ѧ�ڿγ�����.rar'>Download</a>"
%>
</body>
</html>