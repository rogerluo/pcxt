<!--#include file="Cls\ExcelControl.asp"-->
<%
	strFolder = "ExcelFile"
	Call DeleteFolder(strFolder)
	Call CreateFolder(strFolder)
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