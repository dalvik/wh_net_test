<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<%
dim action : action=getForm("action","get")

Select case action	
	case "del" : DelForm
	case "add" : AddForm
	case "edit" : EditForm
	case "editsave" : EditFormSave
	case "save" :SaveForm	
	case "order" :updateOrder	
End Select

Dim FormID, FormName, FormOrder, FormField	
Dim FormOptions,FormDiversification,FormControlType,FormCategory,FormNotNull


Sub EditForm
	dim sql,rsObj
	sql = "select * from {prefix}FormSet where FormID="&CInt(getForm("ID","get"))
	Set rsObj=conn.Exec(sql,"r1")	
	if rsObj.eof then exit sub
	
	FormID = rsObj("FormID")
	FormName = rsObj("FormName")
	FormOrder = rsObj("FormOrder")
	FormField = rsObj("FormField")
	FormOptions = rsObj("FormOptions")
	FormDiversification = rsObj("FormDiversification")
	FormControlType = rsObj("FormControlType")
	FormCategory = rsObj("FormCategory")
	FormNotNull = rsObj("FormNotNull")
	rsObj.close	:Set rsObj = nothing
	
End Sub

Sub EditFormSave
dim f
for each f in request.Form
	echo "<h1>"&f&"="&request.Form(f)&"</h1>"
next
echo "�޸Ĺ��������"
response.end
End Sub

Sub AddForm 


Dim rsObj,sql,tmpCount,i,fName

FormField=getForm("FormField","post")
FormOptions=getForm("FormOptions","post")
FormDiversification=getForm("FormDiversification","post")
FormControlType=getForm("FormControlType","post")
FormName=getForm("FormName","post")
FormCategory=getForm("FormCategory","post")
FormOrder=getForm("FormOrder","post")
FormNotNull=getForm("FormNotNull","post")



if FormNotNull = "on" then
FormNotNull = true
else
FormNotNull = false
end if

FormOptions = encode(FormOptions)

fName = FormField

if isnul(FormDiversification) then FormDiversification = 0


 
	if isnul(FormName) then alertMsgAndGo "�������Ʋ���Ϊ�գ����޸�","-1"
	if isnul(FormField) then alertMsgAndGo "�ֶ����Ʋ���Ϊ�գ����޸�","-1"
	
	sql = "select count(*) from {prefix}FormSet where FormField='"&FormField&"'"
	Set rsObj=conn.Exec(sql,"r1")
	tmpCount = rsObj(0)
	rsObj.close	:Set rsObj = nothing
	
	 
	
	sql = "select count(*) from {prefix}Custom"
	Set rsObj=conn.Exec(sql,"r1")
	for i=0 to rsObj.Fields.Count-1
		if rsObj.Fields(i).Name = fName then
			tmpCount = tmpCount + 1
		end if
	next
	rsObj.close	:Set rsObj = nothing
	
	if tmpCount > 0 then alertMsgAndGo "�ֶ������Ѵ��ڣ����޸�","-1"
	
'0 �ı�,1 ����,2 �༭��,3 ����,4 ����,5 ��ɫ,6 ��ѡ,7 ��ѡ
Select Case CInt(FormControlType)
	Case 2
	sql = "ALTER TABLE {prefix}Custom ADD column "&fName&" memo"
	Case Else
	sql = "ALTER TABLE {prefix}Custom ADD column "&fName&" Text(255)"

End Select
	conn.Exec sql,"exe"	
	if Err then alertMsgAndGo "����ֶ�ʧ��,����ϵ����Ա","-1"
	
	sql = "insert into {prefix}FormSet(FormName,FormField,FormOrder,FormControlType,FormCategory,FormOptions,FormNotNull,FormDiversification) values('"&FormName&"','"&FormField&"',"&FormOrder&",'"&FormControlType&"','"&FormCategory&"','"&FormOptions&"',"&FormNotNull&","&FormDiversification&")"
	conn.Exec sql,"exe"		
	
	alertMsgAndGo "��ӳɹ�","AspCms_Form.asp"
End Sub	

Sub FormList
dim FormID, FormName, FormOrder, FormField	
dim sql, msg

	Dim rsObj	:	Set rsObj=conn.Exec("select FormID,FormName,FormField,FormOrder,FormCategory from {prefix}FormSet Order by FormOrder Asc,FormID","r1")
	If rsObj.Eof Then 
		echo"<tr bgcolor=""#FFFFFF"" align=""center"">"
		echo "<td colspan=""5"">û������</td>"
		echo "</tr>"
	Else
		Do while not rsObj.Eof 
		 echo "<tr bgcolor=""#FFFFFF"" align=""center"">"
		 echo "<td height=""28"">"
		 echo "<input type=""checkbox"" name=""id"" value="""&rsObj("FormID")&""" class=""checkbox"" onclick='silbingCheck(this)' />"
		 echo "<input type=""checkbox"" name=""FormField"" value=""'"&rsObj("FormField")&"'"" style=""display:none"" />"
		 echo "<input type=""hidden"" name=""FormIDs"" value="""&rsObj("FormID")&""" />"
		 echo "<input type=""hidden"" name=""FormCategory"" value="""&rsObj("FormCategory")&""" />"
		 
		 echo "</td>"
		 echo "<td>"&rsObj("FormID")&"</td>"
		 echo "<td>"&rsObj("FormName")&"</td>"
		 if  rsObj("FormCategory") <> "" then
		 echo "<td>"&rsObj("FormCategory")&"_"&rsObj("FormField")&"</td>" 
		 else
		 echo "<td>"&rsObj("FormField")&"</td>" 
		 end if
		 echo "<td><input class=""input"" style=""width:30px"" name=""FormOrders"" value="""&rsObj("FormOrder")&"""/></td>"
		 echo "<td> <a href=""?action=del&id="&rsObj("FormID")&"&FormField="&rsObj("FormField")&"""  onClick=""return confirm('ȷ��Ҫɾ����')"">ɾ��</a></td>"
		 echo "</tr>"
		  rsObj.MoveNext
		Loop
	End If
	rsObj.close : Set rsObj = nothing
End Sub

Sub DelForm 

dim m_FormField
Dim sql,tmp
	dim ID 	:	ID = getForm("id","both")
	FormField=getForm("FormField","both")	
	FormCategory=getForm("FormCategory","both")
	
	
	m_FormField=replace(FormField,"&apos;","")
	 
	FormField = replace(FormField,"&apos;","")
	'die m_FormField
	tmp = m_FormField
	if FormCategory <> "" then tmp =  m_FormField
	sql = "ALTER TABLE {prefix}Custom drop column "&m_FormField
	'die sql
	conn.Exec sql,"exe"
	
	
	if instr(m_FormField,"'")=0 then m_FormField = "'"&m_FormField&"'"
	sql = "Delete from {prefix}FormSet where FormID in ("&ID&")"
	'echo sql & "<br>"
	conn.Exec sql,"exe"
	alertMsgAndGo "ɾ���ɹ�","AspCms_Form.asp"	
	
	if Err then echo err.description 'alertMsgAndGo "ɾ���ֶ�ʧ��,����ϵ����Ա","-1"
	
End Sub 

Sub updateOrder
dim sql, msg
	Dim ids				:	ids=split(getForm("FormIDs","post"),",")
	Dim orders		:	orders=split(getForm("FormOrders","post"),",")
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
			Conn.Exec "update {prefix}FormSet Set FormOrder="&trim(orders(i))&" Where FormID="&trim(ids(i)),"exe"	
		else
			Conn.Exec "update {prefix}FormSet Set FormOrder=0 Where FormID="&trim(ids(i)),"exe"	
		end if
	Next
	
	
	alertMsgAndGo "��������ɹ�","AspCms_Form.asp"
End Sub

%>