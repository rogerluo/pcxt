<%
	dim connxls,connxlsstr
	'connxlsstr ="DRIVER={Microsoft Excel Driver (*.xls)};ReadOnly=False;"
	connxlsstr ="Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='Excel 8.0;HDR=YES;IMEX=1';" 
	Set connxls=Server.CreateObject("ADODB.Connection") 
%>