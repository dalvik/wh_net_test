<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<%
dim action : action=getForm("action","get")

Select case action	
	case "del" : DelTab
	case "add" : AddTab
	case "edit" : EditTab
	case "editsave" : EditTabSave
	case "save" :SaveTab	
	case "order" :updateOrder	
End Select

Dim TabID, TabName, TabOrder, TabField	
Dim TabOptions,TabDiversification,TabControlType,TabCategory,TabNotNull


Sub EditTab
	dim sql,rsObj
	sql = "select * from {prefix}TabSet where TabID="&CInt(getForm("ID","get"))
	Set rsObj=conn.Exec(sql,"r1")	
	if rsObj.eof then exit sub
	
	TabID = rsObj("TabID")
	TabName = rsObj("TabName")
	TabOrder = rsObj("TabOrder")
	TabField = rsObj("TabField")
	TabOptions = rsObj("TabOptions")
	TabDiversification = rsObj("TabDiversification")
	TabControlType = rsObj("TabControlType")
	TabCategory = rsObj("TabCategory")
	TabNotNull = rsObj("TabNotNull")
	rsObj.close	:Set rsObj = nothing
	
End Sub

Sub EditTabSave
dim f
for each f in request.Form
	echo "<h1>"&f&"="&request.Form(f)&"</h1>"
next
echo "修改功能添加中"
response.end
End Sub

Sub AddTab 


Dim rsObj,sql,tmpCount,i,fName

TabField=getForm("TabField","post")
TabOptions=getForm("TabOptions","post")
TabDiversification=getForm("TabDiversification","post")
TabControlType=getForm("TabControlType","post")
TabName=getForm("TabName","post")
TabCategory=getForm("TabCategory","post")
TabOrder=getForm("TabOrder","post")
TabNotNull=getForm("TabNotNull","post")



if TabNotNull = "on" then
TabNotNull = true
else
TabNotNull = false
end if

TabOptions = encode(TabOptions)

fName = TabField

if isnul(TabDiversification) then TabDiversification = 0


 
	if isnul(TabName) then alertMsgAndGo "参数名称不能为空，请修改","-1"
	if isnul(TabField) then alertMsgAndGo "字段名称不能为空，请修改","-1"
	
	sql = "select count(*) from {prefix}TabSet where TabField='"&TabField&"'"
	Set rsObj=conn.Exec(sql,"r1")
	tmpCount = rsObj(0)
	rsObj.close	:Set rsObj = nothing
	
	 
	
	sql = "select count(*) from {prefix}GuestBook"
	Set rsObj=conn.Exec(sql,"r1")
	for i=0 to rsObj.Fields.Count-1
		if rsObj.Fields(i).Name = fName then
			tmpCount = tmpCount + 1
		end if
	next
	rsObj.close	:Set rsObj = nothing
	
	if tmpCount > 0 then alertMsgAndGo "字段名称已存在，请修改","-1"
	
'0 文本,1 数字,2 编辑器,3 附件,4 日期,5 颜色,6 单选,7 多选
Select Case CInt(TabControlType)
	Case 2
	sql = "ALTER TABLE {prefix}GuestBook ADD column "&fName&" memo"
	Case Else
	sql = "ALTER TABLE {prefix}GuestBook ADD column "&fName&" Text(255)"

End Select
	conn.Exec sql,"exe"	
	if Err then alertMsgAndGo "添加字段失败,请联系管理员","-1"
	
	sql = "insert into {prefix}TabSet(TabName,TabField,TabOrder,TabControlType,TabCategory,TabOptions,TabNotNull,TabDiversification) values('"&TabName&"','"&TabField&"',"&TabOrder&",'"&TabControlType&"','"&TabCategory&"','"&TabOptions&"',"&TabNotNull&","&TabDiversification&")"
	conn.Exec sql,"exe"		
	
	alertMsgAndGo "添加成功","AspCms_Tab.asp"
End Sub	

Sub TabList
dim TabID, TabName, TabOrder, TabField	
dim sql, msg

	Dim rsObj	:	Set rsObj=conn.Exec("select TabID,TabName,TabField,TabOrder,TabCategory from {prefix}TabSet Order by TabOrder Asc,TabID","r1")
	If rsObj.Eof Then 
		echo"<tr bgcolor=""#FFFFFF"" align=""center"">"
		echo "<td colspan=""5"">没有数据</td>"
		echo "</tr>"
	Else
		Do while not rsObj.Eof 
		 echo "<tr bgcolor=""#FFFFFF"" align=""center"">"
		 echo "<td height=""28"">"
		 echo "<input type=""checkbox"" name=""id"" value="""&rsObj("TabID")&""" class=""checkbox"" onclick='silbingCheck(this)' />"
		 echo "<input type=""checkbox"" name=""TabField"" value=""'"&rsObj("TabField")&"'"" style=""display:none"" />"
		 echo "<input type=""hidden"" name=""TabIDs"" value="""&rsObj("TabID")&""" />"
		 echo "<input type=""hidden"" name=""TabCategory"" value="""&rsObj("TabCategory")&""" />"
		 echo "</td>"
		 echo "<td>"&rsObj("TabID")&"</td>"
		 echo "<td>"&rsObj("TabName")&"</td>"
		 if  rsObj("TabCategory") <> "" then
		 echo "<td>"&rsObj("TabField")&"</td>" 
		 else
		 echo "<td>"&rsObj("TabField")&"</td>" 
		 end if
		 echo "<td><input class=""input"" style=""width:30px"" name=""TabOrders"" value="""&rsObj("TabOrder")&"""/></td>"
		 echo "<td><a href=""?action=del&id="&rsObj("TabID")&"&TabField="&rsObj("TabField")&"""  onClick=""return confirm('确定要删除吗')"">删除</a></td>"
		 echo "</tr>"
		  rsObj.MoveNext
		Loop
	End If
	rsObj.close : Set rsObj = nothing
End Sub

Sub DelTab 

dim m_TabField
Dim sql,tmp
	dim ID 	:	ID = getForm("id","both")
	TabField=getForm("TabField","both")	
	TabCategory=getForm("TabCategory","both")
	
	
	m_TabField=replace(TabField,"&apos;","")
	 
	TabField = replace(TabField,"&apos;","")
	'die m_TabField
	tmp = m_TabField
	if TabCategory <> "" then tmp = m_TabField
	sql = "ALTER TABLE {prefix}GuestBook drop column "&m_TabField
	'die sql
	conn.Exec sql,"exe"
	
	
	if instr(m_TabField,"'")=0 then m_TabField = "'"&m_TabField&"'"
	sql = "Delete from {prefix}TabSet where TabID in ("&ID&")"
	'echo sql & "<br>"
	conn.Exec sql,"exe"
	alertMsgAndGo "删除成功","AspCms_Tab.asp"	
	
	if Err then echo err.description 'alertMsgAndGo "删除字段失败,请联系管理员","-1"
	
End Sub 

Sub updateOrder
dim sql, msg
	Dim ids				:	ids=split(getForm("TabIDs","post"),",")
	Dim orders		:	orders=split(getForm("TabOrders","post"),",")
	If Ubound(ids)=-1 Then 	'防止有值为空时下标越界
		ReDim ids(0)
		ids(0)=""
	End If	
	
	If Ubound(orders)=-1 Then
		ReDim orders(0)
		orders(0)=0
	End If
	Dim i
	
	For i=0 To Ubound(ids)		
		if isnum(trim(orders(i))) then
			Conn.Exec "update {prefix}TabSet Set TabOrder="&trim(orders(i))&" Where TabID="&trim(ids(i)),"exe"	
		else
			Conn.Exec "update {prefix}TabSet Set TabOrder=0 Where TabID="&trim(ids(i)),"exe"	
		end if
	Next
	
	
	alertMsgAndGo "更新排序成功","AspCms_Tab.asp"
End Sub

%>