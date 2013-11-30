<%
class COMPRESS_DECOMPRESS_FILES 

	private version, copyright 
	private oWshShell, oFso 
	private sCompressPath, sDecompressPath  

	private sub class_initialize 
		version="BJTU PCXT Version1" 
		copyright="Writen By RogerLuo" 
		Set oFso=server.CreateObject("scripting.FileSystemObject") 
		Set oWshShell=server.CreateObject("Wscript.Shell") 
		'writeLn(version+"<br>"+copyright) 
	end Sub 
	private sub class_terminate 
		if isobject(oWshShell) then set oWshShell=nothing 
		if isobject(oFso) then set oFso=nothing 
	end Sub 
	private function physicalPath(byVal s) 
		physicalPath=server.mappath(s) 
	end Function 
	private sub validateFile(byVal s) 
		if oFso.FileExists(s) then exit sub 
		if oFso.FolderExists(s) then exit sub 
		callErr "file(folder) not exists!" 
	end Sub 
	private sub createFolder(byVal s) 
		if oFso.FolderExists(s) then exit Sub 
		oFso.createFolder(s)
	end Sub 
	private sub writeLn(byVal s) 
		response.write "<p>" + s + "</p>" + vbCrlf 
	end Sub 
	private sub callErr(byVal s) 
		writeLn "<p><b>ERROR:</b></p>" + s 
		response.End 
	end sub 
	private sub callSucc(byVal s) 
		writeLn "<p><b>SUCCESS:</b></p>" + s 
	end Sub 

	public sub compress 
		'validateFile(sCompressPath) 
		oWshShell.run(Server.MapPath("./")+"/WinRAR.exe A " + sCompressPath + " " + sDecompressPath & " ,1,False") 
		if Err.number>0 then callErr("compress lost!") 
		callSucc("compress <b>" + sDecompressPath + "</b> to <b>" + sCompressPath + ".rar</b> successfully!") 
	end Sub 
	public sub decompress 
		validateFile(sCompressPath) 
		createFolder(sDecompressPath) 
		oWshShell.run("WinRAR X " + sCompressPath + " " + sDecompressPath & "") 
		if Err.number>0 then callErr("decompress lost!") 
		'callSucc("decompress <b>" + sCompressPath + ".rar</b> to <b>" + sDecompressPath + "</b> successfully!") 
	end sub 

	public property Let compressPath(byVal s) 
		sCompressPath=physicalPath(s) 
	end property 
	public property Let decompressPath(byVal s) 
		sDecompressPath=physicalPath(s) 
	end property 

End class 
%> 