<%
'---fix_img.asp?dbpath=../db/faq.mdb&col=ans&kw=/a1010517/html/&test=yes&table=faq


Function getMdb(mdbFile, mdbTable)
  Set mycnt = Server.CreateObject("ADODB.Connection")
  mycnt.open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(mdbFile)
  Set myrs = Server.CreateObject("ADODB.Recordset")
  myrs.cursorLocation = 3 '在存取 RecordSet 時，無法保証 RecordSet 每次都以同樣的順序出現。
                          '若要啟用 AbsolutePosition，必須先設定為使用用戶端 cursor。
  myrs.open mdbTable, mycnt, 1, 3
  Set getMdb = myrs
End Function


'---- dbpath=資料庫路徑    col=遇取代之欄位   kw=關鍵字   test=真實刪除
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
