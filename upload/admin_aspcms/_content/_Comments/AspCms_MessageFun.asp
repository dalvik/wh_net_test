<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<%
CheckAdmin("AspCms_Message.asp")
dim action : action=getForm("action","get")
'�������ID�������ؼ��ʣ�ҳ��������
dim SortID,keyword,page,order,pic,ID,tab

SortID =getForm("sort","get")	
keyword=getForm("keyword","post")
if isnul(keyword) then keyword=getForm("keyword","get")
page=getForm("page","get")
order=getForm("order","get")
pic=getForm("pic","get")
ID=getForm("id","get")


select case action 
	case "edit" : editFaq	
	case "del" : delFaq	
	case "enable" :Enable
	case "notenabled" :NotEnabled	
end select
Dim FaqID,FaqTitle,Contact,ContactWay,Content,Reply,AddTime,ReplyTime,FaqStatus,AuditStatus

Sub gettab(faqid)
Dim oFCKeditor,sql,k
Dim kvArr
Dim dicType
Set dicType = Server.CreateObject(DICTIONARY_OBJ_NAME)
	
	'ȡ��ֻ�в�Ʒ

		Dim rsObj,tabFields,rsObj1,i,tabNames
		tabFields=""
		sql = "select tabName,tabField,tabControlType from {prefix}tabSet order by tabOrder,tabID"
		kvArr = Conn.Exec(sql,"arr")
		if getDataCount("select Count(*) from {prefix}tabSet")>0 then
		for i=0 to ubound(kvArr,2)	
			if dicType.Exists(kvArr(1,i)) then				
				dicType(kvArr(1,i)) = kvArr(2,i)
				'echo "��rsObj(1)="&kvArr(1,i)&"��"
				'echo "��rsObj('tabControlType')=" & kvArr(2,i) & "��"
				'echo "��dicType('"&kvArr(1,i)&"')="&dicType(kvArr(1,i))&"��<br>"
			else
				'echo "���" & kvArr(1,i) & "<br>"
				dicType.add kvArr(1,i),kvArr(2,i)
			end if	
		
			if faqid=0 then
				'���
				echo "<tr>"
				echo "<td align=middle height=30><u>"&kvArr(0,i)&"</u></td>"
				echo "<td colspan='3'>"				
				EchoControlType kvArr(2,i),kvArr(1,i),""		
				echo "</td>"
				echo "</tr>"
			end if
			
			tabFields=tabFields&kvArr(1,i)&","
			tabNames=tabNames&kvArr(0,i)&","			
		next
		
		
		
		if faqid<>0 and not isnul(tabFields) then 
		sql = "select "&tabFields&"faqid from {prefix}GuestBook where faqid="&faqid
			Set rsObj1=Conn.Exec(sql,"r1")
			tabNames=split(tabNames,",")
			tabFields=split(tabFields,",")			
			Do While not rsObj1.Eof 
				for i=0 to ubound(tabNames)-1
					echo "<tr>"
					echo "<td align=middle width=100 height=30><u>"&trim(tabNames(i))&"��</u></td>"
					echo "<td colspan='3'>"		
'echo "<input type=""text"" class=""input"" maxlength=""100"" value="""&trim(rsObj1(i))&"""  name=""tab"" style=""width:200px"" />"

'echo "dicType('"&tabFields(i)&"')="&dicType(tabFields(i))
EchoControlType dicType(tabFields(i)),tabFields(i),trim(rsObj1(i))
					echo "</td>"
					echo "</tr>"
				next
				rsObj1.MoveNext
			Loop
			rsObj1.Close	: Set rsObj1=Nothing
		end if		
	end if
	set dicType = nothing
End Sub

'2011��7��21��
'by amysimple
'����ؼ�����
' ct ���� cn �ֶ��� cv ֵ
Sub EchoControlType(ct,cn,cv)
'0 �ı�,1 ����,2 �༭��,3 ����,4 ����,5 ��ɫ,6 ��ѡ,7 ��ѡ
select case ct
case 0
echo "<input type='text' class='input' maxlength='100' name='"&cn&"' id='"&cn&"' value='"&cv&"' style='width:200px' />"
case 1
	echo "<input type=""text"" class=""input"" maxlength=""100"" name='"&cn&"' onkeyup=""this.value=this.value.replace(/\D/g,'');"" style=""width:200px"" value='"&cv&"' />"
case 2
	Set oFCKeditor = New FCKeditor:oFCKeditor.BasePath="../../editor/":oFCKeditor.ToolbarSet="AdminMode":oFCKeditor.Width="615":oFCKeditor.Height="300":oFCKeditor.Value=cv:oFCKeditor.Create ""&cn&""
case 3
	echo "<input type='hidden' maxlength='100' name='"&cn&"' id='"&cn&"' value="""&cv&"""/>"
	echo "<iframe src='../../editor/upload.asp?sortType="&sortType&"&stype=file&Tobj="&cn&"' scrolling='no' topmargin='0' width='300px' height='24' marginwidth='0' marginheight='0' frameborder='0' align='left'></iframe>"
	if cv<>"" then echo "<a href='"&cv&"' target='_blank'>�鿴����</a>"
case 4
echo "<input type=""text"" onclick=""WdatePicker()"" readonly class=""input"" maxlength=""100"" name="""&cn&""" style=""width:200px"" value="""&cv&""" />"
case 5
echo "<input type=""text"" readonly class=""iColorPicker input"" maxlength=""100"" id="""&cn&""" name="""&cn&""" style=""width:200px"" value="""&cv&""" />"
case 6
	echo "<input type=""text"" class=""input"" maxlength=""100"" id="""&cn&""" name="""&cn&""" style=""width:200px"" value="""&cv&""" />"	
case 7
	echo "<input type=""text"" class=""input"" maxlength=""100"" name="""&cn&""" style=""width:200px"" value="""&cv&""" />"
case else
	echo "<input type=""text"" class=""input"" maxlength=""100"" name="""&cn&""" style=""width:200px"" value="""&cv&""" />"
end select	
End Sub



Sub delFaq	
	Dim id	:	id=getForm("id","both")
	if isnul(id) then alertMsgAndGo "��ѡ��Ҫɾ��������","-1"
	SortID =getForm("sort","get")
	keyword=getForm("keyword","get")
	page=getForm("page","get")
	order=getForm("order","get")
	pic=getForm("pic","get")
	Conn.Exec "delete from {prefix}GuestBook where FaqID in("&id&")","exe"
	alertMsgAndGo "ɾ���ɹ�","?page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword
End Sub

Sub getContent
	if not isnul(ID) then		
		Dim rs : Set rs = Conn.Exec("select * from {prefix}GuestBook where FaqID="&ID,"r1")
		if not rs.eof then			
			FaqID=rs("FaqID")		
			FaqTitle=rs("FaqTitle")		
			Contact=rs("Contact")		
			ContactWay=rs("ContactWay")		
			Content=decode(rs("Content"))		
			Reply=rs("Reply")		
			AddTime=rs("AddTime")		
			FaqStatus=rs("FaqStatus")		
			AuditStatus=rs("AuditStatus")			
		end if
	else		
		alertMsgAndGo "û��������¼","-1"
	end if
End Sub

Sub editFaq
	FaqID=getForm("FaqID","post")
	FaqTitle=getForm("FaqTitle","post")
	Contact=getForm("Contact","post")
	ContactWay=getForm("ContactWay","post")
	Content=encode(getForm("Content","post"))
	Reply=getForm("Reply","post")
	AddTime=now()
	ReplyTime=now()
	FaqStatus=getCheck(getForm("FaqStatus","post"))
	AuditStatus=getCheck(getForm("AuditStatus","post"))
		if isnul(getForm("tab","post")) then 
		tab=split(",",",")
	else
		tab=split(getForm("tab","post"),",")
	end if
	dim tabStr	: tabStr=""
	
	dim sql
		dim i	:i=0
		sql = "select tabField,tabControlType from {prefix}tabSet order by tabOrder,tabID"
		dim rsObj	:	Set rsObj=Conn.Exec(sql,"r1")
		Do While not rsObj.Eof 		
		
			if rsObj(1) = 2 then
			tabStr = tabStr & ","&rsObj(0)&"='" &encode(getForm(rsObj(0),"post")) & "'"
			else
			tabStr = tabStr & ","&rsObj(0)&"='" &getForm(rsObj(0),"post") & "'"
			end if
			
			i=i+1
			rsObj.MoveNext
		Loop		
		rsObj.Close	:	set rsObj=Nothing
	
	if not isnul(Reply) then AuditStatus=1 else AuditStatus=0 end if 
		
	Conn.Exec"update {prefix}GuestBook set FaqTitle='"&FaqTitle&"',Content='"&Content&"',Reply='"&Reply&"',ReplyTime='"&ReplyTime&"',FaqStatus="&FaqStatus&",AuditStatus="&AuditStatus&""&tabStr&" where FaqID="&FaqID,"exe"	
	alertMsgAndGo "�޸ĳɹ�","AspCms_Message.asp?page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword
End Sub

Sub FaqList	
	dim datalistObj,rsArray
	dim m,i,orderStr,whereStr,sqlStr,rsObj,allPage,allRecordset,numPerPage,searchStr
	numPerPage=10
	orderStr= " order by FaqID desc"
	if isNul(page) then page=1 else page=clng(page)
	if page=0 then page=1
	whereStr=" where 1=1 "
	if not isNul(SortID) then  whereStr=whereStr
	if not isNul(keyword) then 
		whereStr = whereStr&" and (FaqTitle like '%"&keyword&"%' or Content like '%"&keyword&"%')"
	end if
	sqlStr = "select FaqID,Contact,FaqTitle,AddTime,FaqStatus,AuditStatus,ContactWay,Content,Reply,ReplyTime from {prefix}GuestBook "&whereStr&orderStr
	
	set rsObj = conn.Exec(sqlStr,"r1")
	rsObj.pagesize = numPerPage
	allRecordset = rsObj.recordcount : allPage= rsObj.pagecount
	if page>allPage then page=allPage
	if allRecordset=0 then
		if not isNul(keyword) then
		    echo "<tr bgcolor=""#FFFFFF"" align=""center""><td colspan='8'>�ؼ��� <font color=red>"""&keyword&"""</font> û�м�¼</td></tr>" 
		else
		    echo "<tr bgcolor=""#FFFFFF"" align=""center""><td colspan='8'>��û�м�¼!</td></tr>"
		end if
	else  
		rsObj.absolutepage = page
		for i = 1 to numPerPage	
			 echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf& _
			 "<td height=""28""><input type=""checkbox"" name=""id"" value="""&rsObj(0)&""" class=""checkbox"" /><input type=""hidden"" name=""SpecIDs"" value="""&rsObj(0)&""" /></td>"&vbcrlf& _
				"<td>"&rsObj(0)&"</td>"&vbcrlf& _
				"<td>"&rsObj(1)&"</td>"&vbcrlf& _
				"<td>"&rsObj(2)&"</td>"&vbcrlf& _
				"<td>"&rsObj(3)&"</td>"&vbcrlf& _
				"<td>"&getStr(rsObj(4),"<a href=""?action=notenabled&id="&rsObj(0)&"&page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword&""" class=""txt_C1""><IMG src=""../../images/toolbar_ok.gif""></a>","<a href=""?action=enable&id="&rsObj(0)&"&page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword&""" class=""txt_C4""><IMG src=""../../images/toolbar_no.gif""></a>")&"</td>"&vbcrlf& _
				"<td>"&getStr(rsObj(5),"<IMG src=""../../images/toolbar_ok.gif"">","<IMG src=""../../images/toolbar_no.gif"">")&"</td>"&vbcrlf& _
				"<td><a href=""AspCms_MessageEdit.asp?id="&rsObj(0)&"&page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword&""" class=""txt_C1"">�ظ�</a> | <a href=""?action=del&id="&rsObj(0)&"&page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword&""" class=""txt_C1"" onClick=""return confirm('ȷ��Ҫɾ����')"">ɾ��</a></td>"&vbcrlf& _
			  "</tr>"&vbcrlf
			rsObj.movenext
			if rsObj.eof then exit for
		next
		echo"<tr bgcolor=""#FFFFFF"" class=""pagenavi"">"&vbcrlf& _
			"<td colspan=""8"" height=""28"" style=""padding-left:20px;"">"&vbcrlf& _			
			"ҳ����"&page&"/"&allPage&"  ÿҳ"&numPerPage &" �ܼ�¼��"&allRecordset&"�� <a href=""?page=1&order="&order&"&sort="&sortID&"&keyword="&keyword&""">��ҳ</a> <a href=""?page="&(page-1)&"&order="&order&"&sort="&sortID&"&keyword="&keyword&""">��һҳ</a> "&vbcrlf
		dim pageNumber
		pageNumber=makePageNumber_(page, 10, allPage, "guestlist","","","")
		echo pageNumber
		echo"<a href=""?page="&(page+1)&"&order="&order&"&sort="&sortID&"&keyword="&keyword&""">��һҳ</a> <a href=""?page="&allPage&"&order="&order&"&sort="&sortID&"&keyword="&keyword&""">βҳ</a>"&vbcrlf& _		
		"</td>"&vbcrlf& _
		"</tr>"&vbcrlf
	end if
	rsObj.close : set rsObj = nothing	
End Sub

Sub NotEnabled	
	SortID =getForm("sort","get")
	keyword=getForm("keyword","get")
	page=getForm("page","get")
	order=getForm("order","get")
	Dim id				:	id=getForm("ID","get")
	Conn.Exec"update {prefix}GuestBook set FaqStatus=0 Where FaqID="&id,"exe"
	response.Redirect getPageName()&"?page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword	
End Sub

Sub Enable
	SortID =getForm("sort","get")
	keyword=getForm("keyword","get")
	page=getForm("page","get")
	order=getForm("order","get")

	Dim id				:	id=getForm("ID","get")
	Conn.Exec"update {prefix}GuestBook set FaqStatus=1 Where FaqID="&id,"exe"
	response.Redirect getPageName()&"?page="&page&"&order="&order&"&sort="&sortID&"&keyword="&keyword
End Sub

%>