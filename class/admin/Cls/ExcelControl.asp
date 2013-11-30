<!--#include file="adovbs.inc"-->
<%
Function Debug(vData)
	response.write vData+"<br>"
End Function 
  Class ExcelMarker
     
     '--------'
     '公开属性'
     ''''''''''
     Private mFileName'Excel文件路径和文件名
     Private mSheetName'Excel工作表名称
     Private mTableName'数据库的表名称
     Private mConStr'Excel连接字符串
	 '--------'
	 '课程属性'
	 Private mCourseName
	 Private mCourseid
	 Private mCourseDate
	 Private mCourseOrder
	 Private mCourseTeacher
	 Private mCourseTime
	 Private mCourseType
	 Private mCourseHeadCount
	 Private mCourseTeacher2
	 Private mCourseCredit
	 Private mCourseHasEvaluate

	 Private mItemArray(17,5)
	 ''''''''''
     '--------'
     '内部属性'
     '''''''''' 
     Private ObjExcel'Excel对象
     Private ObjSheet'工作表对象
     Private ObjFso'FSO对象
     Private ExlHtml'Excel的HTML对象
     Private i,r,c'循环变量
     Private ExlCon'Excel连接对象
     Private ExlRs'Excel记录集对象
     Private AdoStream'ADODB.Stream对象
     '----------------'
     'Excel类属性初始化'
     ''''''''''''''''''     
    Private Sub Class_Initialize
         mFileName=""
         mSheetName=""
         mTableName=""
         ExlHtml=""
         '初始化连接字符串
         mConStr="Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source={FileName};Extended Properties='Excel 8.0;IMEX=1;HDR=NO';" 
    End Sub
	'初始化课程信息
	Public Sub InitCourseInfo()
		strsql = "select * from courseinfo where courseid = '"+mCourseid+"' and coursedate = '"+mCourseDate+"' and courseorder="+mCourseOrder
		Set clsrs = conn.execute(strsql)
		If clsrs.eof Or clsrs.bof Then
			clsrs.close
			Set clsrs = Nothing
		End If
		mCourseName = CStr(clsrs("coursename"))
		mCourseTeacher = CStr(clsrs("courseteacher"))
		mCourseTime = CStr(clsrs("coursetime"))
		mCourseType = CStr(clsrs("coursetype"))
		mCourseHeadCount = CStr(clsrs("courseheadcount"))
		mCourseTeacher2 = CStr(clsrs("courseteacher2"))
		mCourseCredit = CStr(clsrs("coursecredit"))
		clsrs.close
		Set clsrs = Nothing
		strsql = "select count(stuid) as evaluated from [coursemember] where courseid='"+mCourseid+"' and isevaluate =1 and coursedate='"+mCourseDate+"' and courseorder="+mCourseOrder
		Set clsrs = conn.execute(strsql)
		mCourseHasEvaluate = CStr(clsrs("evaluated"))
		clsrs.close
		
	 End Sub
	 '更新totalevaluate
	 Public Sub SetItemArray()
		'清除totalevaluate表内的数据
		strsql = "delete from totalevaluate"
		conn.execute(strsql)

		For i = 1 To 16 
			mItemArray(i,1) = 0
			mItemArray(i,2) = 0
			mItemArray(i,3) = 0	
			mItemArray(i,4) = 0
		Next

		strsql = "select * from [evaluate] where courseid='"+mCourseid+"' and coursedate='"+mCourseDate+"' and courseorder =" + mCourseOrder
		Set clsrs = conn.execute(strsql)
		If clsrs.eof And clsrs.bof Then 
		Else 
			Do While Not clsrs.eof
				For i = 1 To 16
					'response.write rs(i+3)+"<P>"
					Select Case clsrs(i+4)
					Case "A"
					mItemArray(i,1) = mItemArray(i,1) + 1
					Case "B"
					mItemArray(i,2) = mItemArray(i,2) + 1
					Case "C"
					mItemArray(i,3) = mItemArray(i,3) + 1
					Case "D"
					mItemArray(i,4) = mItemArray(i,4) + 1
					End Select
				Next	
				clsrs.movenext
			Loop
		End If 
		clsrs.close
		For i = 1 To 16 
			If mItemArray(i,1) = "" Then 
				mItemArray(i,1) = 0
			End If
			If mItemArray(i,2) = "" Then 
				mItemArray(i,2) = 0
			End If
			If mItemArray(i,3) = "" Then 
				mItemArray(i,3) = 0
			End If
			If mItemArray(i,4) = "" Then 
				mItemArray(i,4) = 0
			End If
		Next

		strsql = "select count(stuid) as evaluated from [coursemember] where courseid='"+mCourseid+"' and isevaluate =1 and coursedate='"+mCourseDate+"' and courseorder="+mCourseOrder
		Set clsrs = conn.execute(strsql)
		mCourseHasEvaluate = CStr(clsrs("evaluated"))
		clsrs.close
		if mCourseHasEvaluate = 0 Then 
			mCourseHasEvaluate = 1
		End If
		valueA = FormatNumber(mItemArray(1,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(1,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(1,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(1,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(1,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(1,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(1,4))+"("+CStr(valueD)+"%)"
		
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',1,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)
		
		valueA = FormatNumber(mItemArray(2,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(2,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(2,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(2,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(2,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(2,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(2,4))+"("+CStr(valueD)+"%)"

		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',2,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)
		
		valueA = FormatNumber(mItemArray(3,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(3,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(3,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(3,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(3,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(3,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(3,4))+"("+CStr(valueD)+"%)"

		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',3,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(4,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(4,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(4,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(4,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(1,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(4,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(4,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',4,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(5,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(5,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(5,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(5,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(5,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(5,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(5,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',5,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(6,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(6,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(6,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(6,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(6,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(6,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(6,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',6,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(7,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(7,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(7,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(7,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(7,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(7,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(7,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',7,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(8,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(8,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(8,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(8,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(8,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(8,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(8,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',8,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(9,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(9,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(9,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(9,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(9,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(9,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(9,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',9,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(12,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(12,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(12,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(12,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(12,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(12,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(12,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',10,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(14,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(14,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(14,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(14,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(14,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(14,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(14,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',11,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)

		valueA = FormatNumber(mItemArray(15,1)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionA=CStr(mItemArray(15,1))+"("+CStr(valueA)+"%)"	
					
		valueB = FormatNumber(mItemArray(15,2)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionB=CStr(mItemArray(15,2))+"("+CStr(valueB)+"%)"

		valueC = FormatNumber(mItemArray(15,3)/mCourseHasEvaluate*100, 0, -1, -1, 0)
		optionC=CStr(mItemArray(15,3))+"("+CStr(valueC)+"%)"
				
		'optionD=CStr(mItemArray(1,4))+"("+CStr(FormatNumber(mItemArray(1,4)/mCourseHasEvaluate*100,0,-1,-1,0))+"%)"
		valueD = FormatNumber(100 - valueA - valueB - valueC, 0, -1, -1, 0)
		optionD=CStr(mItemArray(15,4))+"("+CStr(valueD)+"%)"
		'更新totalevaluate表内的数据
		strsql = "insert into totalevaluate  (courseid,courseorder,coursedate,itemid,optionA,optionB,optionC,optionD) values ('"+mCourseid+"',"+mCourseOrder+",'"+mCourseDate+"',12,'"+optionA+"','"+optionB+"','"+optionC+"','"+optionD+"')"
		conn.execute(strsql)
	 End Sub
     '初始化文件名
	 Private Sub InitFileName()
		If cint(mCourseOrder) = 1 Then
			mFileName = Server.MapPath("ExcelFile/")+"\"+Replace(mCourseName,"/","")+".xls"
		Else
			mFileName = Server.MapPath("ExcelFile/")+"\"+Replace(mCourseName,"/","")+"_"+mCourseOrder+".xls"
		End If
	 End Sub 
     '--------'
     '内部过程
     '''''''''     
     '创建FSO对象
     Private Sub CreateFileSystemObject()
        Set ObjFso=Server.CreateObject("Scripting.FileSystemObject")
		'判断文件是否存在
		If ObjFso.FileExists(mFileName) then
			ObjFso.DeleteFile(mFileName)  
		End If
        Set ObjExcel=ObjFso.CreateTextFile(mFileName)
     End Sub
     
     '创建ADODB.Stream对象
     Private Sub CreateADODBStream()
        Set AdoStream=Server.CreateObject("ADODB.Stream")
        AdoStream.Type=2
        AdoStream.Open
     End Sub
     
     'ADODB.Stream对象保存文件方法
     Private Sub SaveADODBStream()
        AdoStream.SaveToFile mFileName,2
     End Sub

	 '生成Excel的HTML表格头部代码
     Private Sub MarkHtmlTBHead(ObjRs)
        ExlHtml="<table>"&Chr(13)
        ExlHtml=ExlHtml&"<tr>"&Chr(13)
		ExlHtml=ExlHtml&"<td align=center>"+mCourseName+"</td></tr>"&Chr(13)
		ExlHtml=ExlHtml&"<tr>"&Chr(13)
		ExlHtml=ExlHtml&"<td>课程基本信息</td></tr>"&Chr(13)

		ExlHtml=ExlHtml&"<tr>"&Chr(13)
		ExlHtml=ExlHtml&"<td><table><tr>"&Chr(13)
		ExlHtml=ExlHtml&"<td>授课老师："+mCourseTeacher+"</td><td>授课类型："+coursetype+"</td></tr>"&Chr(13)

		ExlHtml=ExlHtml&"<tr>"&Chr(13)
		ExlHtml=ExlHtml&"<td>替课老师："+mCourseTeacher2+"</td><td>学时："+mCourseTime+" 学分："+mCourseCredit+"</td></tr>"&Chr(13)
		

		ExlHtml=ExlHtml&"<tr>"&Chr(13)
		ExlHtml=ExlHtml&"<td>课序："+mCourseOrder+" 授课人数："+mCourseHeadCount+"</td><td>评价人数："+mCourseHasEvaluate+"</td></tr>"&Chr(13)

		ExlHtml=ExlHtml&"</table></td></tr>"&Chr(13)
		ExlHtml=ExlHtml&"<td><table><tr>"&Chr(13)
		ExlHtml = ExlHtml&"<td>评价表项</td><td>A</td><td>B</td><td>C</td><td>D</td>"&Chr(13)
        ExlHtml=ExlHtml&"</tr>"&Chr(13)
     End Sub
     
     '生成Excel的HTML表格内容代码
     Private Sub MarhHtmlTBBody(ObjRs)
        Do Until ObjRs.EOF
           ExlHtml=ExlHtml&"<tr>"&Chr(13)
		   'change by rogerluo
			itemtype = CStr(ObjRs("itemid"))
			Select Case itemtype
				Case "1"
				itemstr = "使用教材"
				Case "2"
				itemstr = "使用课件"
				Case "3"
				itemstr = "<font color=red>电子教案语种</font>"
				Case "4"
				itemstr = "<font color=red>授课语种</font>"
				Case "5"
				itemstr = "<font color=red>英语授课的比例</font>"
				Case "6"
				itemstr = "准备充分，上课精神饱满"
				Case "7"
				itemstr = "提供足够的信息量"
				Case "8"
				itemstr = "对教学内容熟悉，讲述清楚"
				Case "9"
				itemstr = "<font color=red>开课教师本人授课学时</font>"
				Case "10"
				itemstr = "<font color=red>停课次数</font>"
				Case "11"
				itemstr = "<font color=red>停课后补课次数</font>"
				Case "12"
				itemstr = "<font color=red>对课程总体评价</font>"
			End Select
			ExlHtml = ExlHtml&"<td>"&itemstr&"</td><td>"&ObjRs("optionA")&"</td><td>"&ObjRs("optionB")&"</td><td>"&ObjRs("optionC")&"</td><td>"&ObjRs("optionD")&"</td>"
           ExlHtml=ExlHtml&"</tr>"&Chr(13)
           ObjRs.MoveNext
        Loop
		ExlHtml=ExlHtml&"</table></td></tr></Table>"
     End Sub
     
     '释放对象方法
     Private Sub FreeObject(Obj)
        Set Obj=Nothing
     End Sub
     
     '--------'
     '公开过程'
     '--------' 
     'ADOStream方法生成Excel文件
     Public Sub ADOStreamExcel(ObjRs)
        CreateADODBStream
		InitCourseInfo
		InitFileName
        MarkHtmlTBHead ObjRs
        MarhHtmlTBBody ObjRs
        AdoStream.WriteText ExlHtml
        SaveADODBStream
        AdoStream.Close
        FreeObject AdoStream        
     End Sub

     '--------'
     '属性过程'
     '''''''''
     Public Property Let TableName(vData)
         mTableName=vData
     End Property
     Public Property Get TableName()
         TableName=mTableName
     End Property   
     Public Property Let SheetName(vData)
         mSheetName=vData
     End Property
     Public Property Get SheetName()
         SheetName=mSheetName
     End Property  
     Public Property Let FileName(vData)
         mFileName=vData
     End Property
     Public Property Get FileName()
         FileName=mFileName
     End Property    
     '课程信息'
	 Public Property Let SetCourseName(vData)
		mCourseName = vData
	 End Property 
	 Public Property Get GetCourseName()
		GetCourseName = mCourseName
	 End Property 
	 Public Property Let SetCourseid(vData)
		mCourseid = vData
	 End Property 
	 Public Property Let SetCourseDate(vData)
		mCourseDate = vData
	 End Property 
	 Public Property Let SetCourseOrder(vData)
		mCourseOrder = vData
	 End Property
	 Public Property Let SetCourseTeacher(vData)
		mCourseTeacher = vData
	 End Property 
	 Public Property Let SetCourseType(vData)
		mCourseType = vData
	 End Property 
	 Public Property Let SetCourseTime(vData)
		mCourseTime = vData
	 End Property 
	 Public Property Let SetCourseCredit(vData)
		mCourseCredit = vData
	 End Property 
	 Public Property Let SetCourseHeadCount(vData)
		mCourseHeadCount = vData
	 End Property 
	 Public Property Let SetCourseHasEvaluate(vData)
		mCourseHasEvaluate = vData
	 End Property        
  End Class
%>

