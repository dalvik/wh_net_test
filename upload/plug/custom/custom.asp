<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<%
	dim rs
	set rs=conn.exec("select * from {prefix}custom")
	
	response.write "<table cellpadding=0 cellspacing=0 border=1 width=""500"" ><tr><td>�ֶ���</td><td>�ֶ�����</td><td>��ע</td></tr>"
	For i=0 To rs.fields.count-1
		fieldname=rs.fields(i).name
		fieldtype=rs.fields(i).type
		response.write "<tr><td>"&fieldname&"</td><td>"&AccessTypeName(fieldtype)&"&nbsp;</td><td>&nbsp;</td></tr>"  'Access���ݿ�
		'response.write "<tr><td>"&fieldname&"</td><td>"&SqlTypeName(fieldtype)&"&nbsp;</td><td>&nbsp;</td></tr>" 'sql���ݿ�ʹ�����
	Next
	response.write "</table>

%>