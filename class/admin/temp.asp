<!--#include file="Cls\ExcelControl.asp"-->
<%
	strFolder = "ExcelFile"
	Call DeleteFolder(strFolder)
	Call CreateFolder(strFolder)
	Function DeleteFolder(strFolder) 
	   '首选判断要建立的文件夹是否已经存在 
		Dim strTestFolder, objFSO 
		strTestFolder = Server.Mappath(strFolder) 
		Set objFSO = CreateObject("Scripting.FileSystemObject") 
		' 检查文件夹是否存在 
		If objFSO.FolderExists(strTestFolder) Then 
			objFSO.DeleteFolder(strTestFolder) 
			Debug "删除成功！"
		Else 
			Debug "该文件夹不存在！"
		End If 
		Set objFSO = Nothing 
	End Function 
	Function CreateFolder(strFolder)
		set objFSO = server.CreateObject("scripting.filesystemobject")
		objFSO.createfolder(server.MapPath(strFolder))
		Debug "创建成功！"
	End Function 
%> 