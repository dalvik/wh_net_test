<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="inc/AspCms_smsFuncs.asp" -->
<!--#include file="inc/AspCms_smsConfig.asp" -->
<%
CheckLogin()
dim action : action=getForm("action","get")
Select case action	
	case "add" : addSort
	case "addnum" : addnum
	case "quick" : quickAdd
	case "editnum" : editnum
	case "sendsort" : sendsortmsg
	case "sendname" : sendnamemsg		
	case "del" : delSort
	case "delnum" : delsmsnum
	case "move" : movesms
	case "copy" : copysms
	case "saveall" : saveAll
	case "on" : onOff "on", "Sort", "SortID", "SortStatus", "", getPageName()
	case "off" : onOff "off", "Sort", "SortID", "SortStatus", "", getPageName()
	case "move" : moveSort
End Select

dim SortID, LanguageID, ParentID, SortOrder, SortType, smsSortName, SortURL, SortLevel, AddTime, PageTitle, PageKeywords, PageDesc, SortPath, SortTemplate, ContentTemplate, SortFolder, ContentFolder, SortFileName, ContentFileName, SortStatus, TopSortID, GroupID, Exclusive, Content,indeximage,smsSortID,smsName,smsnum,remark,moveSortID
dim sql, msg,smsid
dim  page, psize, order, ordsc
page=getForm("page","get")
psize=getForm("psize","get")
if not isnum(page) then page=1
if not isnum(psize) then psize=10

Sub sendsortmsg
	dim	sendsort,sendmsg,recordmsg,sendmsgok,sendmsgno
	sendsort=getForm("sendsort","post")
	sendmsg=getForm("sendmsg","post")
	sendmsgok=0
	sendmsgno=0
	if isnul(sendsort) then alertMsgAndGo "请选择要操作的分类","-1"
	if isnul(sendmsg) or len(sendmsg)>500  then alertMsgAndGo "请输入正确的短信格式，不要为空也不能过长！","-1"

	Dim rs :set rs =Conn.Exec ("select * from {prefix}smsnum where smssortid in("& sendsort &")","r1")
	IF rs.eof or rs.bof Then
		alertMsgAndGo "该用户组不存在用户","-1"	
	Else
		Do While not rs.eof 
			recordmsg=mt(rs(2),sendmsg,"","","")
			if not InArray(recordmsg,errno) then 
				sendmsgok=sendmsgok+1
			else
				sendmsgno=sendmsgno+1
			end if
			rs.MoveNext
		Loop
	end if
	rs.close : set rs = nothing
	conn.exec "insert into {prefix}smsRecord(smsrtime,smsryes,smsrno,smsrcontent) values('"&now()&"',"&sendmsgok&","&sendmsgno&",'"&sendmsg&"')", "exe" 
	alertMsgAndGo "发送完成","AspCms_smssend.asp"
End Sub

Sub sendnamemsg
	dim	sendname,sendmsg,recordmsg,sendmsgok,sendmsgno
	sendname=getForm("sendname","post")

	sendmsg=getForm("sendmsg","post")
	sendmsgok=0
	sendmsgno=0
	if isnul(sendname) then alertMsgAndGo "请选择要操作的号码","-1"
	if isnul(sendmsg) or len(sendmsg)>500  then alertMsgAndGo "请输入正确的短信格式，不要为空也不能过长！","-1"

	Dim rs :set rs =Conn.Exec ("select * from {prefix}smsnum where smsid in("& sendname &")","r1")
	IF rs.eof or rs.bof Then
		alertMsgAndGo "该用户不存在","-1"	
	Else
		Do While not rs.eof 
			recordmsg=mt(rs(2),sendmsg,"","","")
			if not InArray(recordmsg,errno) then 
				sendmsgok=sendmsgok+1
			else
				sendmsgno=sendmsgno+1
			end if
			rs.MoveNext
		Loop
	end if
	rs.close : set rs = nothing
	conn.exec "insert into {prefix}smsRecord(smsrtime,smsryes,smsrno,smsrcontent) values('"&now()&"',"&sendmsgok&","&sendmsgno&",'"&sendmsg&"')", "exe" 
	alertMsgAndGo "发送完成,共 "&sendmsgok+sendmsgno&" 条，成功 "&sendmsgok&" 条，失败 "&sendmsgno&" 条！","AspCms_smssend.asp"
End Sub

 
Sub getsmsnum
	dim id : id=getForm("id","get")
	if not isnul(ID) then		
		sql ="select * from {prefix}smsnum where smsid="&id
		dim rs : set rs = conn.exec(sql,"r1")
		if rs.eof then 
			alertMsgAndGo "没有这条记录","-1"
		else
			smsid=rs("smsid")
			smsnum=rs("smsnum")
			smsname=rs("smsname")
			remark=rs("remark")
			smssortid=rs("smssortid")
		
		end if
		rs.close : set rs=nothing
	else		
		alertMsgAndGo "没有这条记录","-1"
	end if 
End Sub

Sub movesms
	dim id : id=getForm("id","post")
	if isnul(id) then alertMsgAndGo "请选择要操作的内容","-1"
	dim moveSortID 
	moveSortID=getForm("moveSortID","post")
	if moveSortID=0 then alertMsgAndGo "请选择要操作的分类","-1"
	conn.exec "update {prefix}smsnum set smsSortID="&moveSortID&" where smsid in("&id&")", "exe"	
	alertMsgAndGo "移动成功！", getPageName()&"?smsSortID="&smsSortID&"&page=1&psize="&psize&"&order="&order&"&ordsc="&ordsc
End Sub

Sub copysms
	dim id : id=getForm("id","post")
	if isnul(id) then alertMsgAndGo "请选择要操作的内容","-1"
	dim moveSortID,i
	moveSortID=getForm("moveSortID","post")
	if moveSortID=0 then alertMsgAndGo "请选择要操作的分类","-1"
	dim copyid:copyid=split(id,",")
	for i=0 to ubound(copyid)

	conn.exec "insert into {prefix}smsnum (smsname,smsnum,smssortid,remark) select smsname,smsnum,"&moveSortID&",remark from {prefix}smsnum where smsid="&copyid(i)&"", "exe"	
	next
	
	alertMsgAndGo "复制成功！",  getPageName()&"?smsSortID="&smsSortID&"&page=1&psize="&psize&"&order="&order&"&ordsc="&ordsc	
End Sub

Sub saveAll
	Dim ids				:	ids=split(getForm("SortIDs","post"),",")
	Dim SortNames		:	SortNames=split(getForm("SortNames","post"),",")
	
	If Ubound(ids)=-1 Then 	'防止有值为空时下标越界
		ReDim ids(0)
		ids(0)=""
	End If	
	If Ubound(SortNames)=-1 Then
		ReDim SortNames(0)
		SortNames(0)=""
	End If
	Dim i
	For i=0 To Ubound(ids)		
		
		Conn.Exec "update {prefix}smssort Set smsSortName='"&trim(SortNames(i))&"' Where smsSortID="&trim(ids(i)),"exe"		
	Next
	alertMsgAndGo "保存成功","AspCms_smsSort.asp"
End Sub	

function smssortList()
	Dim rs :set rs =Conn.Exec ("select * from {prefix}smsSort","r1")
	dim sortTypenames : sortTypenames=split(sortTypes,",")

	IF rs.eof or rs.bof Then
		echo "<tr bgcolor=""#ffffff"" align=""center""><td colspan=""4"">还没有记录</td></tr>"
	Else
		echo "<tr><td colspan=""8"" >"
		echo "<TABLE cellSpacing=0 cellPadding=0 width=""100%"" align=center border=0>"
		Do While not rs.eof 
			
			echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf
			echo "<td width=47  height=25><input type=""checkbox"" name=""id"" value="""&rs("smsSortID")&""" class=""checkbox"" /><input type=""hidden"" name=""SortIDs"" value="""&rs("smsSortID")&""" /></td>"&vbcrlf
			echo "<td width=70>"&rs("smsSortID")&"</td>"&vbcrlf
			
			echo "<td align=""left"" width=280><input name=""SortNames"" type=""text"" class=""input"" id=""SortNames"" value="""&rs("smsSortName")&""" maxlength=""200"" style=""width:120px;""/></td>"&vbcrlf
			
			echo "<td width=225>"
			if cint(rs("smsSortID")) > 2 then
			echo "<a href=""?action=del&id="&rs("smsSortID")&""" onClick=""return confirm('确定要删除吗')"">删除</a></td>"&vbcrlf			
			end if
			echo "</tr>"	&vbcrlf			
			
			
			rs.MoveNext
		Loop
		echo "</table>"	
		echo "</td></tr>"
		rs.close : set rs = nothing
	End If
End Function

function smssendList()
	dim i,orderStr,whereStr,sqlStr,rsObj,allPage,allRecordset,numPerPage
	numPerPage=psize
	if not isnum(numPerPage) then numPerPage=10	
	if isnul(order) then order="ContentID"
	if isnul(ordsc) then ordsc="desc"
	smsSortID=getForm("smsSortID", "get")

	if isNul(page) then page=1 else page=clng(page)
	if page=0 then page=1
	
	whereStr=" where 1=1 order by smsrid desc"

	sqlStr = "select * from {prefix}smsRecord "&whereStr


	set rsObj = conn.Exec(sqlStr,"r1")
	rsObj.pagesize = numPerPage
	allRecordset = rsObj.recordcount : allPage= rsObj.pagecount
	if page>allPage then page=allPage
	if allRecordset=0 then	
		    echo "<tr align=""center""><td colspan='8' height=""28"">还没有记录!</td></tr>"
	else  
		rsObj.absolutepage = page
		for i = 1 to numPerPage
			echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf
			echo "<td width=47  height=25>"&rsObj(0)&"</td>"&vbcrlf
			echo "<td align=""center"" >"&rsObj(1)&"</td>"&vbcrlf
			
			echo "<td align=""center"" >"&rsObj(2)+rsObj(3)&"</td>"&vbcrlf
			echo "<td align=""center"" >成功："&rsObj(2)&"/失败："&rsObj(3)&"</td>"&vbcrlf
			echo "<td width=225>"
			
			echo rsObj(4)&"</td>"&vbcrlf			
			
			echo "</tr>"	&vbcrlf			
			rsObj.movenext
			if rsObj.eof then exit for
		next
		echo"<tr bgcolor=""#FFFFFF"" class=""pagenavi"">"&vbcrlf& _
			"<td colspan=""8"" height=""28"" style=""padding-left:20px;"">"&vbcrlf& _			
			"页数："&page&"/"&allPage&"  每页"&numPerPage &" 总记录数"&allRecordset&"条 <a href=""?smsSortID="&smsSortID&"&page=1&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">首页</a> <a href=""?smsSortID="&smsSortID&"&page="&(page-1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">上一页</a> "&vbcrlf
		dim pageNumber
		pageNumber=makesmsPageNumber_(page, 10, allPage, "newslist",smsSortID, order)
		echo pageNumber
		echo"<a href=""?smsSortID="&smsSortID&"&page="&(page+1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">下一页</a> <a href=""?smsSortID="&smsSortID&"&page="&allPage&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">尾页</a>"&vbcrlf
			
		echo"每页显示"&getPageSize("10",psize)&getPageSize("20",psize)&getPageSize("30",psize)&getPageSize("50",psize)
				
		echo"条"&vbcrlf& _	
		"</td>"&vbcrlf& _
		"</tr>"&vbcrlf
	end if
	rsObj.close : set rsObj = nothing	
End Function


Sub addSort	
	smsSortName=getForm("smsSortName", "post")
	
	if isnul(smsSortName) then alertMsgAndGo "分类名称不能为空","-1"
	
	conn.exec "insert into {prefix}smsSort(smsSortName) values('"&smsSortName&"')", "exe" 

	alertMsgAndGo "添加成功","AspCms_smsSort.asp"
End Sub

Sub addnum
	smsName=getForm("smsName", "post")
	smsnum=getForm("smsnum", "post")
	moveSortID=getForm("moveSortID", "post")
	remark=getForm("remark", "post")
	if isnul(smsName) then alertMsgAndGo "姓名不能为空","-1"
	if isnul(smsnum) then alertMsgAndGo "号码不能为空","-1"
	
	conn.exec "insert into {prefix}smsnum(smsName,smsnum,smssortid,remark) values('"&smsName&"','"&smsnum&"',"&moveSortID&",'"&remark&"')", "exe" 
	alertMsgAndGo "添加成功","AspCms_smsnumlist.asp"
End Sub


Sub editnum
	smsid=getForm("smsid", "post")	
	smsName=getForm("smsName", "post")
	smsnum=getForm("smsnum", "post")
	moveSortID=getForm("moveSortID", "post")
	remark=getForm("remark", "post")
	if isnul(smsName) then alertMsgAndGo "姓名不能为空","-1"
	if isnul(smsnum) then alertMsgAndGo "号码不能为空","-1"

	conn.exec "update {prefix}smsnum set smsName='"&smsName&"',  smsnum='"&smsnum&"', smssortid="&moveSortID&", remark='"&remark&"' where smsid="&smsid, "exe"
	
	alertMsgAndGo "修改成功","AspCms_smsnumlist.asp"
End Sub	

Sub delSort
	dim id	:	id=getForm("id","both")
	if isnul(id) then alertMsgAndGo "请选择要操作的内容","-1"
	dim ids,i
	ids=split(id,",")
	for i=0 to ubound(ids)
		if not isnul(ids(i)) then 
			conn.exec "delete from {prefix}smsSort where smsSortID in ("&ids(i)&")","exe"
		end if
	next
	alertMsgAndGo "删除成功",getPageName()		
End Sub

Sub delsmsnum
	dim id	:	id=getForm("id","both")
	if isnul(id) then alertMsgAndGo "请选择要操作的内容","-1"
	dim ids,i
	ids=split(id,",")
	for i=0 to ubound(ids)
		if not isnul(ids(i)) then 
			conn.exec "delete from {prefix}smsnum where smsID in ("&ids(i)&")","exe"
		end if
	next
	alertMsgAndGo "删除成功",getPageName()		
End Sub


Sub smsnumlist
	dim i,orderStr,whereStr,sqlStr,rsObj,allPage,allRecordset,numPerPage
	numPerPage=psize
	if not isnum(numPerPage) then numPerPage=10	
	if isnul(order) then order="ContentID"
	if isnul(ordsc) then ordsc="desc"
	smsSortID=getForm("smsSortID", "get")

	if isNul(page) then page=1 else page=clng(page)
	if page=0 then page=1
	
	orderStr=" order by Contentorder asc,"&order&" "&ordsc &""
	whereStr=" where 1=1"

	if not isNul(smsSortID) and smsSortID<>"0" then  whereStr=whereStr&" and a.smsSortID in("&smsSortID&")"


	sqlStr = "select a.*,b.smssortname from {prefix}smsnum as a, {prefix}smsSort as b "&whereStr&" and a.smsSortID=b.smsSortID"


	set rsObj = conn.Exec(sqlStr,"r1")
	rsObj.pagesize = numPerPage
	allRecordset = rsObj.recordcount : allPage= rsObj.pagecount
	if page>allPage then page=allPage
	if allRecordset=0 then	
		    echo "<tr align=""center""><td colspan='8' height=""28"">还没有记录!</td></tr>"
	else  
		rsObj.absolutepage = page
		for i = 1 to numPerPage
		echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf& _
				"<td><input type=""checkbox"" name=""id"" value="""&rsObj(0)&""" class=""checkbox"" /></td>"&vbcrlf& _
				"<td>"&rsObj(0)&"</td>"&vbcrlf& _
				"<td align=""left"">"&rsObj(1)&"</td>"&vbcrlf& _	
				"<td>"&rsObj(2)&"</td>"&vbcrlf& _
				"<td>"&rsObj("smssortname")&"</td>"&vbcrlf& _			
				"<td>"&rsObj(4)&"</td>"&vbcrlf& _
				
				"<td><a href=""aspcms_smsnumedit.asp?id="&rsObj(0)&"""  >修改</a> | <a href=""?action=delnum&id="&rsObj(0)&"""  onClick=""return confirm('确定要删除吗')"">删除</a></td>"&vbcrlf& _
			  "</tr>"&vbcrlf
			rsObj.movenext
			if rsObj.eof then exit for
		next
		echo"<tr bgcolor=""#FFFFFF"" class=""pagenavi"">"&vbcrlf& _
			"<td colspan=""8"" height=""28"" style=""padding-left:20px;"">"&vbcrlf& _			
			"页数："&page&"/"&allPage&"  每页"&numPerPage &" 总记录数"&allRecordset&"条 <a href=""?smsSortID="&smsSortID&"&page=1&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">首页</a> <a href=""?smsSortID="&smsSortID&"&page="&(page-1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">上一页</a> "&vbcrlf
		dim pageNumber
		pageNumber=makesmsPageNumber_(page, 10, allPage, "newslist",smsSortID, order)
		echo pageNumber
		echo"<a href=""?smsSortID="&smsSortID&"&page="&(page+1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">下一页</a> <a href=""?smsSortID="&smsSortID&"&page="&allPage&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">尾页</a>"&vbcrlf
			
		echo"每页显示"&getPageSize("10",psize)&getPageSize("20",psize)&getPageSize("30",psize)&getPageSize("50",psize)
				
		echo"条"&vbcrlf& _	
		"</td>"&vbcrlf& _
		"</tr>"&vbcrlf
	end if
	rsObj.close : set rsObj = nothing	
End Sub

Function delList(Byval sortID)
	if not isnum(sortID) then exit function
	dim templateFile,page
	dim templateobj,TemplatePath : set templateobj=new TemplateClass
	dim rsObj,rsObjSmalltype,rsObjBigtype,channelTemplateName,tempStr,tempArr,pageStr,sql,sperStr,sperStrs,content,contentLink
	
	set rsobj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then delList="" : exit function	
	templateFile=rsObj("SortTemplate")
	if isnul(templateFile) then 
		select case rsObj("SortType")
			case "2"
				templateFile="newslist.html"
			case "3"
				templateFile="teacherlist.html"
			case "4"
				templateFile="downlist.html"
			case "5"
				templateFile="joblist.html"
			case "6"
				templateFile="albumlist.html"		
		end select
	end if
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/"	&templateFile	
	'die templatePath
	if not CheckTemplateFile(templatePath) then echo templateFile&err_16
	
	
	if not isnul(rsObj("GroupID")) then	
		if not ViewNoRight(rsObj("GroupID"),rsObj("Exclusive")) then exit function	
	end if	
	
	'开始解析标签
	templateObj.load(templatePath)
	tempstr=templateObj.content
	Dim objRegExp, Match, Matches, pages
	Set objRegExp=new Regexp 
	objRegExp.IgnoreCase=True 
	objRegExp.Global=True 	
	objRegExp.Pattern="{aspcms:list([\s\S]*?)}([\s\S]*?){/aspcms:list}"
	'进行匹配	
	set Matches=objRegExp.Execute(tempstr) 	
	for each Match in Matches 
		pages=templateObj.parseArr(Match.SubMatches(0))("size")
	next 
	'die pages
	set objRegExp=Nothing 	
	
	templateObj.parseHtml()	
	tempStr=templateObj.content 
	dim rs
	set rs =conn.exec("select * from {prefix}Content where ContentStatus=1 and SortID in ("&getSubSort(sortID, 1)&")","r1")
	dim pcount
	if isnul(pages) then 	
		rs.pagesize=1
		pcount=1
	else
		rs.pagesize=pages
		pcount=rs.pagecount
		if pcount=0 then pcount=1	
	end if
	'echo pages&"AA"&pcount&"<br>"
	for page=1 to pcount	
	
		dim sortFolder, sortFileName
		sortFolder=rsObj("sortFolder")
		sortFileName=rsObj("sortFileName")		
		sortFolder=replace(sortFolder, "{sitepath}", sitePath)	
		sortFileName=replace(sortFileName, "{sortid}", sortID)	
		
		sortFileName=replace(sortFileName, "{page}", page)
		if isExistFile(sortFolder&sortFileName&fileExt) then delFile sortFolder&sortFileName&fileExt		
		
		'echo sortFolder&sortFileName&fileExt&"<br>"
	next
		
	'makeList=templateObj.content 	
	rsObj.close : set rsObj=nothing
	set templateobj=nothing
End Function

Sub tempList(selectname,selectedtemp)
	dim path,fileListArray,i,fileAttr,folderListArray,folderAttr,parentPath
	path=sitePath&"/Templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath	
	if not isExistFolder(path) then createFolder path,"folderdir"
	
	echo "<select name="""&selectname&""" id="""&selectname&""">"
	if not isExistFolder(path) then	
		echo "<option value="""" >未找到模版</option>"
		response.end
	end if
	fileListArray= getFileList(path)
	if instr(fileListArray(0),",")>0 then	
		for  i = 0 to ubound(fileListArray)
			dim selectedstr:selectedstr=""
			fileAttr=split(fileListArray(i),",")
			if selectedtemp=fileAttr(0) then selectedstr=" selected"
			echo "<option value="""&fileAttr(0)&""" "&selectedstr&">"&fileAttr(0)&"</option>"
		next		
	end if	
	echo "</select>"
End Sub

Function makesmsPageNumber_(Byval currentPage,Byval pageListLen,Byval totalPages,Byval linkType,Byval smsSortID, Byval order)
	dim beforePages,pagenumber,page
	dim beginPage,endPage,strPageNumber
	if pageListLen mod 2 = 0 then beforePages = pagelistLen / 2 else beforePages = clng(pagelistLen / 2) - 1
	if  currentPage < 1  then currentPage = 1 else if currentPage > totalPages then currentPage = totalPages
	if pageListLen > totalPages then pageListLen=totalPages
	if currentPage - beforePages < 1 then
		beginPage = 1 : endPage = pageListLen
	elseif currentPage - beforePages + pageListLen > totalPages  then
		beginPage = totalPages - pageListLen + 1 : endPage = totalPages
	else 
		beginPage = currentPage - beforePages : endPage = currentPage - beforePages + pageListLen - 1
	end if	
'	die currentPage
	for pagenumber = beginPage  to  endPage
		if pagenumber=1 then page = "" else page = pagenumber
		if pagenumber=currentPage then
			strPageNumber=strPageNumber&"<span>"&pagenumber&"</span>"
		else
		
					strPageNumber=strPageNumber&"<a href='?smsSortID="&smsSortID&"&page="&pagenumber&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&"'>"&pagenumber&"</a>"
	
		end if	
	next
	makesmsPageNumber_=strPageNumber
End Function

Function loadsmsSelect(loadid)
	dim rs,sel
	sel=""
	echo "<select name=""moveSortID"" id=""moveSortID"">" & vbcr
	set rs =conn.Exec ("select * from {prefix}smssort","r1")
	
	Do While Not rs.Eof	
		if loadid=rs(0) then sel="selected"		
		echo "<option value="""& rs(0) &""" "&sel&">"&rs(1) &"</option>" & vbcr
		sel=""
		rs.MoveNext	
	Loop
	rs.Close	:	Set rs=Nothing
	echo "</select>" & vbcr
End Function

Function loadsmsbutton()
	dim rs
	set rs =conn.Exec ("select * from {prefix}smssort","r1")
	
	Do While Not rs.Eof			
		echo "<INPUT onClick=""location.href='"&getPageName()&"?"&"smsSortID="&rs(0)&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&"'"" type=""button"" value="""&rs(1)&""" class=""button"" />&nbsp;" & vbcr
		rs.MoveNext
	Loop
	rs.Close	:	Set rs=Nothing
	
End Function
%>