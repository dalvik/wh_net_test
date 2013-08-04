<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<%CheckAdmin("AspCms_Spec.asp")%>
<%
dim action : action=getForm("action","get")
dim sortType:sortType=getForm("sortType","get")
dim sortTypeName,SpecCategory
Select case action	
	case "del" : DelSpec
	case "add" : AddSpec
	case "edit" : EditSpec
	case "editsave" : EditSpecSave
	case "save" :SaveSpec	
	case "order" :updateOrder	
End Select

select case sortType
	case "2"
		sortTypeName ="文章"
		SpecCategory = "C"
	case "3"
		sortTypeName ="产品"
		SpecCategory = "P"
	case "4"
		sortTypeName ="下载"
		SpecCategory = "DL"
	case "5"
		sortTypeName ="招聘"
		SpecCategory = "HR"
	case "6"
		sortTypeName ="相册"
		SpecCategory = "FO"

	case "8"
		sortTypeName = "视频"
		SpecCategory = "VI"
		
end select 
'单篇1,文章2,产品3,下载4,招聘5,相册6,链接7，视频8

Dim SpecID, SpecName, SpecOrder, SpecField	
Dim SpecOptions,SpecDiversification,SpecControlType,SpecNotNull

Sub EditSpec
	dim sql,rsObj
	sql = "select * from {prefix}SpecSet where SpecID="&CInt(filterPara((getForm("ID","get"))))
	Set rsObj=conn.Exec(sql,"r1")	
	if rsObj.eof then exit sub
	SpecID = rsObj("SpecID")
	SpecName = rsObj("SpecName")
	SpecField = rsObj("SpecField")

	SpecCategory = rsObj("SpecCategory")

	rsObj.close	:Set rsObj = nothing
	
End Sub

Sub EditSpecSave
	dim sql,rsObj
	SpecField=filterPara(getForm("SpecField","post"))
	SpecName=filterPara(getForm("SpecName","post"))
	SpecCategory=filterPara(getForm("SpecCategory","post"))
	SpecID =filterPara(getForm("SpecID","post"))
	sql = "select * from {prefix}SpecSet where SpecID="&SpecID
	Set rsObj=conn.Exec(sql,"r1")	
	
	
	sql = "Alter table aspcms_content drop Column "&rsObj("SpecCategory")&"_"&rsObj("SpecField") &""
	conn.Exec sql,"exe"	
	sql="ALTER TABLE aspcms_Content ADD column "& SpecCategory&"_"& SpecField &" text"
	conn.Exec sql,"exe"
	sql = "update {prefix}SpecSet set SpecName='"&SpecName&"',SpecField='"&SpecField&"',SpecCategory='"&SpecCategory&"' where SpecID="&SpecID
	conn.Exec sql,"exe"	
	alertMsgAndGo "修改成功","AspCms_Spec.asp"
End Sub

Sub AddSpec 
Dim rsObj,sql,tmpCount,i,fName

SpecField=filterPara(getForm("SpecField","post"))
SpecOptions=filterPara(getForm("SpecOptions","post"))
SpecDiversification=filterPara(getForm("SpecDiversification","post"))
SpecControlType=filterPara(getForm("SpecControlType","post"))
SpecName=filterPara(getForm("SpecName","post"))
SpecCategory=filterPara(getForm("SpecCategory","post"))
SpecOrder=filterPara(getForm("SpecOrder","post"))
SpecNotNull=filterPara(getForm("SpecNotNull","post"))



if SpecNotNull = "on" then
SpecNotNull = true
else
SpecNotNull = false
end if

SpecOptions = encode(SpecOptions)

fName = SpecCategory & "_" & SpecField

if isnul(SpecDiversification) then SpecDiversification = 0


 
	if isnul(SpecName) then alertMsgAndGo "参数名称不能为空，请修改","-1"
	if isnul(SpecField) then alertMsgAndGo "字段名称不能为空，请修改","-1"
	
	sql = "select count(*) from {prefix}SpecSet where SpecField='"&SpecField&"'"
	Set rsObj=conn.Exec(sql,"r1")
	tmpCount = rsObj(0)
	rsObj.close	:Set rsObj = nothing
	
	 
	
	sql = "select count(*) from {prefix}content"
	Set rsObj=conn.Exec(sql,"r1")
	for i=0 to rsObj.Fields.Count-1
		if rsObj.Fields(i).Name = fName then
			tmpCount = tmpCount + 1
		end if
	next
	rsObj.close	:Set rsObj = nothing
	
	if tmpCount > 0 then alertMsgAndGo "字段名称已存在，请修改","-1"
	
'0 文本,1 数字,2 编辑器,3 附件,4 日期,5 颜色,6 单选,7 多选
Select Case CInt(SpecControlType)
	Case 2
	sql = "ALTER TABLE {prefix}Content ADD column "&fName&" memo"
	Case Else
	sql = "ALTER TABLE {prefix}Content ADD column "&fName&" Text(255)"

End Select
	conn.Exec sql,"exe"	
	if Err then alertMsgAndGo "添加字段失败,请联系管理员","-1"
	
	sql = "insert into {prefix}SpecSet(SpecName,SpecField,SpecOrder,SpecControlType,SpecCategory,SpecOptions,SpecNotNull,SpecDiversification) values('"&SpecName&"','"&SpecField&"',"&SpecOrder&",'"&SpecControlType&"','"&SpecCategory&"','"&SpecOptions&"',"&SpecNotNull&","&SpecDiversification&")"
	conn.Exec sql,"exe"		
	
	alertMsgAndGo "添加成功","AspCms_Spec.asp"
End Sub	

Sub specList
dim SpecID, SpecName, SpecOrder, SpecField, sortType
sortType = filterPara(getForm("sortType","get"))
dim sql, msg, rsObj
	if sortType<>"" then
	Set rsObj=conn.Exec("select SpecID,SpecName,SpecField,SpecOrder,SpecCategory from {prefix}SpecSet where SpecCategory='"&SpecCategory&"'  Order by SpecOrder Asc,SpecID","r1")
	else
	Set rsObj=conn.Exec("select SpecID,SpecName,SpecField,SpecOrder,SpecCategory from {prefix}SpecSet  Order by SpecOrder Asc,SpecID","r1")
	end if
	If rsObj.Eof Then 
		echo"<tr bgcolor=""#FFFFFF"" align=""center"">"
		echo "<td colspan=""5"">没有数据</td>"
		echo "</tr>"
	Else
		Do while not rsObj.Eof 
		 echo "<tr bgcolor=""#FFFFFF"" align=""center"">"
		 echo "<td height=""28"">"
		 echo "<input type=""checkbox"" name=""id"" value="""&rsObj("SpecID")&""" class=""checkbox"" onclick='silbingCheck(this)' />"
		 echo "<input type=""checkbox"" name=""SpecField"" value=""'"&rsObj("SpecCategory")&"_"&rsObj("SpecField")&"'"" style=""display:none"" />"
		 echo "<input type=""hidden"" name=""SpecIDs"" value="""&rsObj("SpecID")&""" />"
		 echo "<input type=""hidden"" name=""SpecCategory"" value="""&rsObj("SpecCategory")&""" />"
		 
		 echo "</td>"
		 echo "<td>"&rsObj("SpecID")&"</td>"
		 echo "<td>"&rsObj("SpecName")&"</td>"
		 if  rsObj("SpecCategory") <> "" then
		 echo "<td>"&rsObj("SpecCategory")&"_"&rsObj("SpecField")&"</td>" 
		 else
		 echo "<td>"&rsObj("SpecField")&"</td>" 
		 end if
		 echo "<td><input class=""input"" style=""width:30px"" name=""SpecOrders"" value="""&rsObj("SpecOrder")&"""/></td>"
		 echo "<td> <a href=""AspCms_SpecEdit.asp?action=update&id="&rsObj("SpecID")&"""  onClick=""return confirm('字段修改必须删除原有字段进行重建，该字段数据将丢失，确认进入修改吗？')"">修改</a> | <a href=""?action=del&id="&rsObj("SpecID")&"&SpecField="&rsObj("SpecCategory")&"_"&rsObj("SpecField")&"""  onClick=""return confirm('确定要删除吗')"">删除</a></td>"
		 echo "</tr>"
		  rsObj.MoveNext
		Loop
	End If
	rsObj.close : Set rsObj = nothing
End Sub

Sub DelSpec 

dim m_SpecField
Dim sql,tmp
	dim ID 	:	ID = filterPara(getForm("id","both"))
	SpecField=getForm("SpecField","both")
	SpecCategory=getForm("SpecCategory","both")
	
	
	m_SpecField=replace(SpecField,"&apos;","")
	 
	SpecField = replace(SpecField,"&apos;","")
	'die m_SpecField
	tmp = m_SpecField
	
	if m_SpecField="P_Price" then alertMsgAndGo "该参数在订购中固定使用，不可删除","AspCms_Spec.asp"	
	if SpecCategory <> "" then tmp = SpecCategory & "_" & m_SpecField
	sql = "ALTER TABLE {prefix}Content drop column "&m_SpecField
	'die sql
	conn.Exec sql,"exe"
	
	if instr(m_SpecField,"'")=0 then m_SpecField = "'"&m_SpecField&"'"
	sql = "Delete from {prefix}SpecSet where SpecID in ("&ID&")"
	'echo sql & "<br>"
	conn.Exec sql,"exe"
	alertMsgAndGo "删除成功","AspCms_Spec.asp"	
	
	if Err then echo err.description 'alertMsgAndGo "删除字段失败,请联系管理员","-1"
	
End Sub 

Sub updateOrder
dim sql, msg
	Dim ids				:	ids=split(filterPara(getForm("SpecIDs","post")),",")
	Dim orders		:	orders=split(filterPara(getForm("SpecOrders","post")),",")
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
			Conn.Exec "update {prefix}SpecSet Set SpecOrder="&trim(orders(i))&" Where SpecID="&trim(ids(i)),"exe"	
		else
			Conn.Exec "update {prefix}SpecSet Set SpecOrder=0 Where SpecID="&trim(ids(i)),"exe"	
		end if
	Next
	
	
	alertMsgAndGo "更新排序成功","AspCms_Spec.asp"
End Sub

%>