<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<%
	dim rs
	set rs=conn.exec("select * from {prefix}custom")
	
	response.write "<table cellpadding=0 cellspacing=0 border=1 width=""500"" ><tr><td>字段名</td><td>字段类型</td><td>备注</td></tr>"
	For i=0 To rs.fields.count-1
		fieldname=rs.fields(i).name
		fieldtype=rs.fields(i).type
		response.write "<tr><td>"&fieldname&"</td><td>"&AccessTypeName(fieldtype)&"&nbsp;</td><td>&nbsp;</td></tr>"  'Access数据库
		'response.write "<tr><td>"&fieldname&"</td><td>"&SqlTypeName(fieldtype)&"&nbsp;</td><td>&nbsp;</td></tr>" 'sql数据库使用这句
	Next
	response.write "</table>

%>