<%
'---fix_img.asp?dbpath=../db/faq.mdb&col=ans&kw=/a1010517/html/&test=yes&table=faq


Function getMdb(mdbFile, mdbTable)
  Set mycnt = Server.CreateObject("ADODB.Connection")
  mycnt.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(mdbFile)
  Set myrs = Server.CreateObject("ADODB.Recordset")
  myrs.cursorLocation = 3 '�b�s�� RecordSet �ɡA�L�k�O�� RecordSet �C�����H�P�˪����ǥX�{�C
                          '�Y�n�ҥ� AbsolutePosition�A�������]�w���ϥΥΤ�� cursor�C
  myrs.open mdbTable, mycnt, 1, 3
  Set getMdb = myrs
End Function


'---- dbpath=��Ʈw���|    col=�J���N�����   kw=����r   test=�u��R��
sql="select * from "&request("table")&" where "&request("col")&" like '%"&request("kw")&"%'"

response.write sql
response.write "<hr>"
response.write request("db_path")



	set cmx=GetMdb(request("db_path"), sql)



	If Not cmx.eof Then
	
		While Not cmx.eof
		

			If request("test")<>"yes" Then

			cmx( Replace(request("col")," ","") ) = Replace(cmx( Replace(request("col")," ","") ),request("kw"),request("target"))
			response.write Replace(cmx( Replace(request("col")," ","") ),request("kw"),request("target"))
			cmx.update

			Else
			
			response.write Replace(cmx( Replace(request("col")," ","") ),request("kw"),request("target"))
			End if
			


		cmx.movenext
		wend

	End if

%>
