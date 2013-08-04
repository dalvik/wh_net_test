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
		sortTypeName ="����"
		SpecCategory = "C"
	case "3"
		sortTypeName ="��Ʒ"
		SpecCategory = "P"
	case "4"
		sortTypeName ="����"
		SpecCategory = "DL"
	case "5"
		sortTypeName ="��Ƹ"
		SpecCategory = "HR"
	case "6"
		sortTypeName ="���"
		SpecCategory = "FO"

	case "8"
		sortTypeName = "��Ƶ"
		SpecCategory = "VI"
		
end select 
'��ƪ1,����2,��Ʒ3,����4,��Ƹ5,���6,����7����Ƶ8

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
	alertMsgAndGo "�޸ĳɹ�","AspCms_Spec.asp"
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


 
	if isnul(SpecName) then alertMsgAndGo "�������Ʋ���Ϊ�գ����޸�","-1"
	if isnul(SpecField) then alertMsgAndGo "�ֶ����Ʋ���Ϊ�գ����޸�","-1"
	
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
	
	if tmpCount > 0 then alertMsgAndGo "�ֶ������Ѵ��ڣ����޸�","-1"
	
'0 �ı�,1 ����,2 �༭��,3 ����,4 ����,5 ��ɫ,6 ��ѡ,7 ��ѡ
Select Case CInt(SpecControlType)
	Case 2
	sql = "ALTER TABLE {prefix}Content ADD column "&fName&" memo"
	Case Else
	sql = "ALTER TABLE {prefix}Content ADD column "&fName&" Text(255)"

End Select
	conn.Exec sql,"exe"	
	if Err then alertMsgAndGo "����ֶ�ʧ��,����ϵ����Ա","-1"
	
	sql = "insert into {prefix}SpecSet(SpecName,SpecField,SpecOrder,SpecControlType,SpecCategory,SpecOptions,SpecNotNull,SpecDiversification) values('"&SpecName&"','"&SpecField&"',"&SpecOrder&",'"&SpecControlType&"','"&SpecCategory&"','"&SpecOptions&"',"&SpecNotNull&","&SpecDiversification&")"
	conn.Exec sql,"exe"		
	
	alertMsgAndGo "��ӳɹ�","AspCms_Spec.asp"
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
		echo "<td colspan=""5"">û������</td>"
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
		 echo "<td> <a href=""AspCms_SpecEdit.asp?action=update&id="&rsObj("SpecID")&"""  onClick=""return confirm('�ֶ��޸ı���ɾ��ԭ���ֶν����ؽ������ֶ����ݽ���ʧ��ȷ�Ͻ����޸���')"">�޸�</a> | <a href=""?action=del&id="&rsObj("SpecID")&"&SpecField="&rsObj("SpecCategory")&"_"&rsObj("SpecField")&"""  onClick=""return confirm('ȷ��Ҫɾ����')"">ɾ��</a></td>"
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
	
	if m_SpecField="P_Price" then alertMsgAndGo "�ò����ڶ����й̶�ʹ�ã�����ɾ��","AspCms_Spec.asp"	
	if SpecCategory <> "" then tmp = SpecCategory & "_" & m_SpecField
	sql = "ALTER TABLE {prefix}Content drop column "&m_SpecField
	'die sql
	conn.Exec sql,"exe"
	
	if instr(m_SpecField,"'")=0 then m_SpecField = "'"&m_SpecField&"'"
	sql = "Delete from {prefix}SpecSet where SpecID in ("&ID&")"
	'echo sql & "<br>"
	conn.Exec sql,"exe"
	alertMsgAndGo "ɾ���ɹ�","AspCms_Spec.asp"	
	
	if Err then echo err.description 'alertMsgAndGo "ɾ���ֶ�ʧ��,����ϵ����Ա","-1"
	
End Sub 

Sub updateOrder
dim sql, msg
	Dim ids				:	ids=split(filterPara(getForm("SpecIDs","post")),",")
	Dim orders		:	orders=split(filterPara(getForm("SpecOrders","post")),",")
	If Ubound(ids)=-1 Then 	'��ֹ��ֵΪ��ʱ�±�Խ��
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
	
	
	alertMsgAndGo "��������ɹ�","AspCms_Spec.asp"
End Sub

%>