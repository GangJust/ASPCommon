<%
	'''''''''''''''''''''''''''
	' 基于现在所学、所用数据库封装，其他数据暂时未用到。
	' 连接方式为 OLEDB 至于如何设置数据源，烦请百度。
	'
	' @auther Gang
	'''''''''''''''''''''''''''


	'建立accdb格式的数据库连接
	function getACCDB(dbName)
		dim conn
		set conn = server.createobject("ADODB.Connection")
		conn.open "Provider=Microsoft.Ace.OLEDB.12.0; Data Source=" & server.mapPath("/") & "/" & dbName

		set getACCDB = conn
	end function

	'建立mdb格式的数据库连接
	function getMDB(dbName)
		dim conn
		set conn = server.createobject("ADODB.Connection")
		conn.open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & server.mapPath("/") & "/" & dbName

		set getMDB = conn
	end function

	'建立sqlserver数据库连接(未测试)
	function getSqlserver(location,dbName,userName,userPass)
		dim conn
		set conn = server.createobject("ADODB.Connection")
		conn.open "Provider=SQLOLEDB; Data Source=" & location & "; initial Catalog=" & dbName & "; uid=" & userName & "; pwd=" & userPass

		set getSqlserver = conn
	end function

%>


<%
	''打开表通用操作示例

	'打开样例表
	function getExampleTable(conn,where)
		dim sql

		sql = "select * from exampleTable"
		if strComp(trim(where),"") <> 0 then sql = sql & " " & where

		set getExampleTable = openTable(conn,sql)
	end function


	'打开表操作
	function openTable (conn,sql)
		set rec = server.createObject("ADODB.Recordset")
		rec.open sql,conn,1,2,1

		set openTable = rec
	end function

%>