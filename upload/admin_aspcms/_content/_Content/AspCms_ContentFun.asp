<!--#include file="../../inc/AspCms_SettingClass.asp" -->
<!--#include file="../../editor/fckeditor.asp" -->
<%
'die debugmode
dim action : action=getForm("action","get")
dim ContentID, LanguageID, SortID, GroupID, Exclusive, Title, Title2, TitleColor, IsOutLink, OutLink, Author, ContentSource, ContentTag, Content, ContentStatus, IsTop, IsRecommend, IsImageNews, IsHeadline, IsFeatured, ContentOrder, IsGenerated, Visits, AddTime, ImagePath, IndexImage, DownURL, PageTitle, PageKeywords, PageDesc, PageFileName, spec, EditTime,DownGroupID,IsNoComment,Star,Timeing,TimeStatus,VideoGroupID,CHvalue,SpecCategory
'SpecCategory�����ж����ĸ����͵��Զ������
dim sortType, keyword, page, psize, order, ordsc, sortTypeName
sortType=getForm("sortType","get")
if isnul(sortType) then sortType=0	
sortid=getForm("sortid","post")	
if isnul(sortid) then sortid=getForm("sortid","get")
keyword=getForm("keyword","post")
if isnul(keyword) then keyword=getForm("keyword","get")
page=getForm("page","get")
psize=getForm("psize","get")
order=getForm("order","get")
ordsc=getForm("ordsc","get")
DownGroupID=getForm("DownGroupID","post")
VideoGroupID=getForm("VideoGroupID","post")
	
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
	
Select  case action
	case "add" : addContent	
	case "edit" : editContent	
	case "move" : moveContent
	case "copy" : copyContent
	case "rpost" : rpostContent			
	case "del" : delContent	
	case "recovery" : Recovery	
	case "tdel" : trueDelContent
	case "on" : onOff "on", "Content", "ContentID", "ContentStatus", "", getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc
	case "off" : onOff "off", "Content", "ContentID", "ContentStatus", "", getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc
	
	case "order" : UpdateOrder
	
End Select


Sub getSpec(Contentid)
Dim oFCKeditor,sql,k
Dim kvArr
Dim dicType
Set dicType = Server.CreateObject(DICTIONARY_OBJ_NAME)
	
	'ȡ��ֻ�в�Ʒ
		Dim rsObj,specFields,rsObj1,i,SpecNames
		specFields=""
		sql = "select SpecName,SpecCategory+'_'+SpecField,SpecControlType,SpecOptions from {prefix}SpecSet where SpecCategory = '"&SpecCategory&"' order by SpecOrder,SpecID"
		kvArr = Conn.Exec(sql,"arr")
		if getDataCount("select Count(*) from {prefix}SpecSet where SpecCategory = '"&SpecCategory&"'")>0 then
		for i=0 to ubound(kvArr,2)
			if dicType.Exists(kvArr(1,i)) then				
				dicType(kvArr(1,i)) = kvArr(2,i)
				'echo "��rsObj(1)="&kvArr(1,i)&"��"
				'echo "��rsObj('SpecControlType')=" & kvArr(2,i) & "��"
				'echo "��dicType('"&kvArr(1,i)&"')="&dicType(kvArr(1,i))&"��<br>"
			else
				'echo "���" & kvArr(1,i) & "<br>"
				dicType.add kvArr(1,i),kvArr(2,i)
			end if	
			if kvArr(3,i)="" then
				CHvalue=""
			else
				CHvalue=kvArr(3,i)
			end if
			if Contentid=0 then
				'���
				echo "<tr>"
				echo "<td align=middle height=30><u>"&kvArr(0,i)&"</u></td>"
				echo "<td colspan='3'>"				
				EchoControlType kvArr(2,i),kvArr(1,i),"",CHvalue	
				echo "</td>"
				echo "</tr>"
			end if
			
			specFields=specFields&kvArr(1,i)&","
			SpecNames=SpecNames&kvArr(0,i)&","			
		next
		
		
		if Contentid<>0 and not isnul(specFields) then 
		sql = "select "&specFields&"Contentid from {prefix}Content where Contentid="&Contentid
			Set rsObj1=Conn.Exec(sql,"r1")
			SpecNames=split(SpecNames,",")
			specFields=split(specFields,",")			
			Do While not rsObj1.Eof 
				for i=0 to ubound(SpecNames)-1
					if kvArr(3,i)="" then
						CHvalue=""
					else
						CHvalue=kvArr(3,i)
					end if
					echo "<tr>"
					echo "<td align=middle width=100 height=30><u>"&trim(SpecNames(i))&"��</u></td>"
					echo "<td colspan='3'>"		
'echo "<input type=""text"" class=""input"" maxlength=""100"" value="""&trim(rsObj1(i))&"""  name=""spec"" style=""width:200px"" />"

'echo "dicType('"&specFields(i)&"')="&dicType(specFields(i))
EchoControlType dicType(specFields(i)),specFields(i),trim(rsObj1(i)),CHvalue
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
Sub EchoControlType(ct,cn,cv,CHvalue)
'0 �ı�,1 ����,2 �༭��,3 ����,4 ����,5 ��ɫ,6 ��ѡ,7 ��ѡ
dim splitValue,i,ii,strr
select case ct
case 0
echo "<input type='text' class='input' maxlength='100' name='"&cn&"' id='"&cn&"' value='"&cv&"' style='width:200px' />"
case 1
	echo "<input type=""text"" class=""input"" maxlength=""100"" name='"&cn&"'style=""width:200px"" value='"&cv&"' />"
case 2
		dim oFCKeditor
		Set oFCKeditor = New FCKeditor:oFCKeditor.BasePath="../../editor/":oFCKeditor.ToolbarSet="AdminMode":oFCKeditor.Width="615":oFCKeditor.Height="300":oFCKeditor.Value=decodeHtml(cv):oFCKeditor.Create ""&cn&""
		
	'echo "<textarea id="""&cn&""" name="""&cn&""" style=""width:680px;height:300px;"">"&decodeHtml(cv)&"</textarea>"

case 3
	echo "<input type='hidden' maxlength='100' name='"&cn&"' id='"&cn&"' value="""&cv&"""/>"
	echo "<iframe src='../../editor/upload.asp?sortType="&sortType&"&stype=file&Tobj="&cn&"' scrolling='no' topmargin='0' width='300px' height='24' marginwidth='0' marginheight='0' frameborder='0' align='left'></iframe>"
	if cv<>"" then echo "<a href='"&cv&"' target='_blank'>�鿴����</a>"
case 4
echo "<input type=""text"" onclick=""WdatePicker()"" readonly class=""input"" maxlength=""100"" name="""&cn&""" style=""width:200px"" value="""&cv&""" />"
case 5
echo "<input type=""text"" readonly class=""iColorPicker input"" maxlength=""100"" id="""&cn&""" name="""&cn&""" style=""width:200px"" value="""&cv&""" />"
case 6
	if CHvalue<>"" then
		splitValue=split(CHvalue,"<br>")
	else
		splitValue=split(cv,"<br>")
	end if
	for i=0 to ubound(splitValue)
		if splitValue(i)=cv then
			strr="checked" 
		else
			strr=""
		end if
		echo "<input type=""radio"" class=""input"" maxlength=""100"" id="""&cn&""" name="""&cn&""" value="""&splitValue(i)&""" "&strr&"/>"&splitValue(i)&""
	next
case 7
	if CHvalue<>"" then
		splitValue=split(CHvalue,"<br>")	
	else
		splitValue=split(cv,"<br>")
	end if
	for i=0 to ubound(splitValue)
		if instr(cv,splitValue(i)) then
			strr="checked" 
		else
			strr=""
		end if
		echo "<input type=""checkbox"" class=""input"" maxlength=""100"" name="""&cn&""" value="""&splitValue(i)&""" "&strr&"/>"&splitValue(i)&""
	next
case else
	echo "<input type=""text"" class=""input"" maxlength=""100"" name="""&cn&""" style=""width:200px"" value="""&cv&""" />"
end select	
End Sub



Sub getContent
	dim id : id=getForm("id","get")
	if not isnul(ID) then		
		Dim rs : Set rs = Conn.Exec("select * from {prefix}Content where ContentID="&ID,"r1")
		if not rs.eof then
			ContentID=rs("ContentID")
			LanguageID=rs("LanguageID")
			SortID=rs("SortID")
			GroupID=rs("GroupID")
			Exclusive=rs("Exclusive")
			Title=rs("Title")
			Title2=rs("Title2")
			TitleColor=rs("TitleColor")
			IsOutLink=rs("IsOutLink")
			OutLink=rs("OutLink")
			Author=rs("Author")
			ContentSource=rs("ContentSource")
			'ContentTag=replace(replace(rs("ContentTag")," ",","),"��",",")
			ContentTag=getTags(rs("ContentTag"))
			Content=rs("Content")		
			ContentStatus=rs("ContentStatus")		
			IsTop=rs("IsTop")		
			IsRecommend=rs("IsRecommend")		
			IsImageNews=rs("IsImageNews")		
			IsHeadline=rs("IsHeadline")		
			IsFeatured=rs("IsFeatured")		
			ContentOrder=rs("ContentOrder")		
			IsGenerated=rs("IsGenerated")		
			Visits=rs("Visits")		
			AddTime=rs("AddTime")		
			ImagePath=rs("ImagePath")		
			IndexImage=rs("IndexImage")		
			DownURL=rs("DownURL")		
			PageTitle=rs("PageTitle")
			PageKeywords=rs("PageKeywords")		
			PageDesc=rs("PageDesc")		
			PageFileName=rs("PageFileName")		
			IsNoComment=rs("IsNoComment")			
			Star=rs("Star")		
			TimeStatus=rs("TimeStatus")		
			Timeing=rs("Timeing")
			DownGroupID=rs("DownGroupID")
			if DownGroupID = "" then DownGroupID=2
			VideoGroupID=rs("VideoGroupID")
			if VideoGroupID = "" then VideoGroupID=2	
		end if
	else		
		alertMsgAndGo "û��������¼","-1"
	end if
End Sub

Sub addContent
Dim sql

	LanguageID=cint(session("languageID"))
	SortID=getForm("SortID", "post")
	GroupID=getForm("GroupID", "post")
	Exclusive=getForm("Exclusive", "post")
	
	TitleColor=getForm("TitleColor", "post")
	Title=getForm("Title", "post")
	Title2=getForm("Title2", "post")
	Author=getForm("Author", "post")
	ContentSource=getForm("ContentSource", "post")
	Content=getForm("Content", "post")
	
	PageTitle=getForm("PageTitle", "post")
	PageKeywords=getForm("PageKeywords", "post")
	PageDesc=getForm("PageDesc", "post")	
	
	OutLink=getForm("OutLink", "post")
	IsOutLink=getCheck(getForm("IsOutLink", "post"))
	ContentStatus=getCheck(getForm("ContentStatus", "post"))
	IsTop=getCheck(getForm("IsTop", "post"))
	IsRecommend=getCheck(getForm("IsRecommend", "post"))
	IsImageNews=getCheck(getForm("IsImageNews", "post"))
	IsHeadline=getCheck(getForm("IsHeadline", "post"))
	IsFeatured=getCheck(getForm("IsFeatured", "post"))
	IsGenerated=getCheck(getForm("IsGenerated", "post"))
	IsNoComment=getCheck(getForm("IsNoComment", "post"))
	IsNoComment=getCheck(getForm("IsNoComment", "post"))
	
	TimeStatus=getCheck(getForm("TimeStatus", "post"))
	Timeing=getForm("Timeing", "post")
	
	Star=getForm("Star", "post")
	ContentOrder=getForm("ContentOrder", "post")
	Visits=getForm("Visits", "post")
	AddTime=getForm("AddTime", "post")
	EditTime=now()
	ImagePath=getForm("ImagePath", "post")
	IndexImage=getForm("IndexImage", "post")
	DownURL=getForm("DownURL", "post")
	PageFileName=getForm("PageFileName", "post")
	spec=split(getForm("spec","post"),",")
	DownGroupID=getForm("DownGroupID","post")
	if DownGroupID = "" then DownGroupID=2
	VideoGroupID=getForm("VideoGroupID","post")
	if VideoGroupID = "" then VideoGroupID=2
	
	dim specStr	: specStr=""
	dim specValue  : specValue=""
	dim valueStr	: valueStr=""
	
		dim i	:i=0
		dim rsObj
		sql = "select SpecName,SpecCategory+'_'+SpecField,SpecControlType,SpecOptions from {prefix}SpecSet where SpecCategory = '"&SpecCategory&"' order by SpecOrder,SpecID"
		sql = "select SpecCategory+'_'+SpecField,SpecControlType from {prefix}SpecSet order by SpecOrder,SpecID"
		Set rsObj=Conn.Exec(sql,"r1")
		'0 �ı�,1 ����,2 �༭��,3 ����,4 ����,5 ��ɫ,6 ��ѡ,7 ��ѡ
		Do While not rsObj.Eof 		
			specStr = specStr&","&rsObj(0)
			if rsObj(1) = 2 then
			valueStr = valueStr & ",'" &encode(getForm(rsObj(0),"post")) & "'"
			else
			valueStr = valueStr & ",'" &getForm(rsObj(0),"post") & "'"
			end if
			
			'if ubound(spec,1) < i then i = ubound(spec,1)
			'echo ubound(spec,1) &"=" &i & "[" &spec(5) & "]"
			'valueStr = valueStr&",'"&trim(spec(i))&"'"			
			i=i+1
			rsObj.MoveNext
		Loop
		rsObj.Close	:	set rsObj=Nothing

	if not isnum(LanguageID) then alertMsgAndGo"��ǰ�������Դ��������µ�½","-1"
	if isnul(Title) then alertMsgAndGo"����д����","-1"
	if not isnum(SortID) or SortID="0" then alertMsgAndGo"��ѡ�����","-1"
	if not isnum(ContentOrder) then ContentOrder=99
	if not isnum(Visits) then Visits=0
	if not isdate(AddTime) then AddTime=now()
	if not isdate(Timeing) then Timeing=now()

	ContentTag=addTags(replace(replace(getForm("ContentTag", "post")," ",","),"��",","))	
	
	'die specStr
	sql = "insert into {prefix}Content(VideoGroupID,DownGroupID, LanguageID, SortID, GroupID, Exclusive, TitleColor, Title, Title2, Author, ContentSource, [Content], ContentTag, PageTitle, PageKeywords, PageDesc, OutLink, IsOutLink, ContentStatus, IsTop, IsRecommend, IsImageNews, IsHeadline, IsFeatured, IsGenerated, ContentOrder, Visits, AddTime, EditTime, ImagePath, IndexImage, DownURL"&specStr&", PageFileName, IsNoComment, Star,TimeStatus,Timeing) values("&VideoGroupID&","&DownGroupID&","&LanguageID&", "&SortID&", "&GroupID&", '"&Exclusive&"', '"&TitleColor&"', '"&Title&"', '"&Title2&"', '"&Author&"', '"&ContentSource&"', '"&Content&"', '"&ContentTag&"', '"&PageTitle&"', '"&PageKeywords&"', '"&PageDesc&"', '"&OutLink&"', "&IsOutLink&", "&ContentStatus&", "&IsTop&", "&IsRecommend&", "&IsImageNews&", "&IsFeatured&", "&IsFeatured&", "&IsGenerated&", "&ContentOrder&", "&Visits&", '"&AddTime&"', '"&EditTime&"', '"&ImagePath&"',  '"&IndexImage&"',  '"&DownURL&"' "&valueStr&", '"&PageFileName&"', "&IsNoComment&", "&Star&", "&TimeStatus&", '"&Timeing&"')"
	
	'die sql
	
	conn.exec sql,"exe"
	sortType=getForm("sortType", "post")
	alertMsgAndGo"��ӳɹ�", "AspCms_ContentList.asp"&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc	
End Sub

Sub editContent
dim sql
	ContentID=getForm("ContentID", "post")
	LanguageID=cint(session("languageID"))
	SortID=getForm("SortID", "post")
	GroupID=getForm("GroupID", "post")
	Exclusive=getForm("Exclusive", "post")
	
	TitleColor=getForm("TitleColor", "post")
	Title=getForm("Title", "post")
	Title2=getForm("Title2", "post")
	Author=getForm("Author", "post")
	ContentSource=getForm("ContentSource", "post")
	Content=getForm("Content", "post")
	PageTitle=getForm("PageTitle", "post")
	PageKeywords=getForm("PageKeywords", "post")
	PageDesc=getForm("PageDesc", "post")
	
	OutLink=getForm("OutLink", "post")
	IsOutLink=getCheck(getForm("IsOutLink", "post"))
	ContentStatus=getCheck(getForm("ContentStatus", "post"))
	IsTop=getCheck(getForm("IsTop", "post"))
	IsRecommend=getCheck(getForm("IsRecommend", "post"))
	IsImageNews=getCheck(getForm("IsImageNews", "post"))
	IsHeadline=getCheck(getForm("IsHeadline", "post"))
	IsFeatured=getCheck(getForm("IsFeatured", "post"))
	IsGenerated=getCheck(getForm("IsGenerated", "post"))
	IsNoComment=getCheck(getForm("IsNoComment", "post"))
	
		
	TimeStatus=getCheck(getForm("TimeStatus", "post"))
	Timeing=getForm("Timeing", "post")
	Star=getForm("Star", "post")	
	ContentOrder=getForm("ContentOrder", "post")
	Visits=getForm("Visits", "post")
	AddTime=getForm("AddTime", "post")
	EditTime=now()
	ImagePath=getForm("ImagePath", "post")
	IndexImage=getForm("IndexImage", "post")
	DownURL=getForm("DownURL", "post")
	PageFileName=getForm("PageFileName", "post")
	DownGroupID=getForm("DownGroupID","post")
	if DownGroupID="" then DownGroupID=2
	VideoGroupID=getForm("VideoGroupID","post")
	if VideoGroupID="" then VideoGroupID=2
	
	if isnul(getForm("spec","post")) then 
		spec=split(",",",")
	else
		spec=split(getForm("spec","post"),",")
	end if
	dim specStr	: specStr=""
	
		dim i	:i=0
		sql = "select SpecCategory+'_'+SpecField,SpecControlType from {prefix}SpecSet order by SpecOrder,SpecID"
		dim rsObj	:	Set rsObj=Conn.Exec(sql,"r1")
		Do While not rsObj.Eof 	
			'specStr = specStr&","&rsObj(0)&"='"&trim(spec(i))&"'"
			if rsObj(1) = 2 then
			specStr = specStr & ","&rsObj(0)&"='" &getForm(rsObj(0),"post") & "'"
			else
			specStr = specStr & ","&rsObj(0)&"='" &getForm(rsObj(0),"post") & "'"
			end if
			
			i=i+1
			rsObj.MoveNext
		Loop		
		rsObj.Close	:	set rsObj=Nothing
	
	if not isnum(LanguageID) then alertMsgAndGo"��ǰ�������Դ��������µ�½","-1"
	if isnul(Title) then alertMsgAndGo"����д����","-1"
	if not isnum(SortID) or SortID="0" then alertMsgAndGo"��ѡ�����","-1"
	if not isnum(ContentOrder) then ContentOrder=99
	if not isnum(Visits) then Visits=0		
	ContentTag=addTags(replace(replace(getForm("ContentTag", "post")," ",","),"��",","))	
	if not isdate(AddTime) then AddTime=now()
	if not isdate(Timeing) then Timeing=now()
	
	sql = "update {prefix}Content  set SortID="&SortID&", GroupID="&GroupID&", Exclusive='"&Exclusive&"', TitleColor='"&TitleColor&"', Title='"&Title&"', Title2='"&Title2&"', Author='"&Author&"', ContentSource='"&ContentSource&"', [Content]='"&Content&"', ContentTag='"&ContentTag&"', PageTitle='"&PageTitle&"', PageKeywords='"&PageKeywords&"', PageDesc='"&PageDesc&"', OutLink='"&OutLink&"', IsOutLink="&IsOutLink&", ContentStatus="&ContentStatus&", IsTop="&IsTop&", IsRecommend="&IsRecommend&", IsImageNews="&IsImageNews&", IsHeadline="&IsHeadline&", IsFeatured="&IsFeatured&", IsGenerated="&IsGenerated&", ContentOrder="&ContentOrder&", Visits="&Visits&", AddTime='"&AddTime&"', EditTime='"&EditTime&"', ImagePath='"&ImagePath&"', IndexImage='"&IndexImage&"', DownURL='"&DownURL&"'"&specStr&", PageFileName='"&PageFileName&"', IsNoComment="&IsNoComment&", Star="&Star&",TimeStatus="&TimeStatus&",Timeing='"&Timeing&"',DownGroupID='"&DownGroupID&"',VideoGroupID='"&VideoGroupID&"' where ContentID="&ContentID
	
	conn.exec sql,"exe"
	sortType=getForm("sortType", "post")
	alertMsgAndGo"�޸ĳɹ�", "AspCms_ContentList.asp"&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc	
End Sub

Sub moveContent
	dim id : id=getForm("id","post")
	if isnul(id) then alertMsgAndGo "��ѡ��Ҫ����������","-1"
	dim moveSortID 
	moveSortID=getForm("moveSortID","post")
	if moveSortID=0 then alertMsgAndGo "��ѡ��Ҫ�����ķ���","-1"
	conn.exec "update {prefix}Content set SortID="&moveSortID&" where LanguageID="&session("languageID")&" and ContentID in("&id&")", "exe"	
	alertMsgAndGo "�ƶ��ɹ���", getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc	
End Sub

Sub copyContent
	dim id : id=getForm("id","post")
	if isnul(id) then alertMsgAndGo "��ѡ��Ҫ����������","-1"
	dim moveSortID,i,sql,specStr
	moveSortID=getForm("moveSortID","post")
	if moveSortID=0 then alertMsgAndGo "��ѡ��Ҫ�����ķ���","-1"
	dim copyid:copyid=split(id,",")

		sql = "select SpecCategory+'_'+SpecField,SpecControlType from {prefix}SpecSet order by SpecOrder,SpecID"
		dim rsObj	:	Set rsObj=Conn.Exec(sql,"r1")
		Do While not rsObj.Eof 		

			specStr = specStr & ","&rsObj(0)
			
			rsObj.MoveNext
		Loop		
		rsObj.Close	:	set rsObj=Nothing
	
	
	for i=0 to ubound(copyid)
		conn.exec "insert into {prefix}Content (VideoGroupID,DownGroupID, LanguageID,  GroupID, Exclusive, TitleColor, Title, Title2, Author, ContentSource, [Content], ContentTag, PageTitle, PageKeywords, PageDesc, OutLink, IsOutLink, ContentStatus, IsTop, IsRecommend, IsImageNews, IsHeadline, IsFeatured, IsGenerated, ContentOrder, Visits, AddTime, EditTime, ImagePath, IndexImage, DownURL, PageFileName, IsNoComment,TimeStatus,Timeing, Star"&specStr&",SortID) select VideoGroupID, DownGroupID, LanguageID, GroupID, Exclusive, TitleColor, Title, Title2, Author, ContentSource, [Content], ContentTag, PageTitle, PageKeywords, PageDesc, OutLink, IsOutLink, ContentStatus, IsTop, IsRecommend, IsImageNews, IsHeadline, IsFeatured, IsGenerated, ContentOrder, Visits, AddTime, EditTime, ImagePath, IndexImage, DownURL, PageFileName, IsNoComment,TimeStatus,Timeing, Star"&specStr&","&moveSortID&" from {prefix}Content where contentid="&copyid(i)&"", "exe"	
	next
	
	alertMsgAndGo "���Ƴɹ���", getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc	
End Sub

Sub rpostContent
	dim id : id=getForm("id","post")
	if isnul(id) then alertMsgAndGo "��ѡ��Ҫ����������","-1"
	dim moveSortID 
	Timeing=getForm("Timeing","post")
	if cdate(Timeing)<=now() then alertMsgAndGo "��ʱʱ�䲻��С�ڵ�ǰʱ��","-1"
	conn.exec "update {prefix}Content set TimeStatus=1,Timeing='"&Timeing&"' where LanguageID="&session("languageID")&" and ContentID in("&id&")", "exe"	
	alertMsgAndGo "��ʱ�ɹ���", getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc	
End Sub


Sub delContent
	dim id : id=getForm("id","both")
	if isnul(id) then alertMsgAndGo "��ѡ��Ҫ����������","-1"
	conn.exec "update {prefix}Content set ContentStatus=2 where ContentID in("&id&")","exe"
	alertMsgAndGo "ɾ���ɹ�",getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc		
End Sub

Sub Recovery
	dim id : id=getForm("id","both")
	if isnul(id) then alertMsgAndGo "��ѡ��Ҫ����������","-1"
	conn.exec "update {prefix}Content set ContentStatus=1 where ContentID in("&id&")","exe"
	alertMsgAndGo "�ָ��ɹ�",getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc		
End Sub
	
Sub trueDelContent
	dim id : id=getForm("id","both")
	if isnul(id) then alertMsgAndGo "��ѡ��Ҫ����������","-1"
	if runmode=1 then
		dim rs, sql, filepath
		dim templateobj : set templateobj=new TemplateClass
		sql="select ContentID,Title,sortType,SortFolder,a.GroupID,ContentFolder,ContentFileName,a.AddTime,a.PageFileName,a.SortID,b.GroupID from {prefix}Content as a, {prefix}Sort as b where a.LanguageID="&session("languageID")&" and a.SortID=b.SortID and ContentStatus=2 and ContentID in("&id&")"
		set rs=conn.exec(sql,"r1")		
		do while not rs.eof
			
			filepath=templateobj.getContentLink(rs("SortID"),rs("ContentID"),rs("SortFolder"),rs("a.GroupID"),rs("ContentFolder"),rs("ContentFileName"),rs("AddTime"),rs("PageFileName"),rs("b.GroupID"))
			if isExistFile(filepath) then delFile filepath		
			'echo filepath&"<br>"
			rs.movenext
		loop
	end if
	conn.exec "delete from {prefix}Content where ContentStatus=2 and ContentID in("&id&")","exe"
	alertMsgAndGo "����ɾ���ɹ�",getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc
End Sub
		
Sub listContent
	dim i,orderStr,whereStr,sqlStr,rsObj,allPage,allRecordset,numPerPage
	numPerPage=psize
	if not isnum(numPerPage) then numPerPage=10	
	if isnul(order) then order="ContentID"
	if isnul(ordsc) then ordsc="desc"

	if isNul(page) then page=1 else page=clng(page)
	if page=0 then page=1
	
	orderStr=" order by Contentorder asc,"&order&" "&ordsc &""
	whereStr=" where a.LanguageID="&session("languageID")
	
	if not isNul(sortid) and sortid<>"0" then  whereStr=whereStr&" and a.SortID in("&getSubSort(sortID, 1)&")"
	if not isNul(keyword) then whereStr = whereStr&" and Title like '%"&keyword&"%'"

	sqlStr = "select ContentID,Title,visits,ContentStatus,a.SortID,a.AddTime,IsTop,a.IndexImage,ContentOrder,SortName,IsRecommend,IsImageNews,IsHeadline,IsFeatured,IsOutLink,OutLink,TitleColor,TimeStatus,Timeing from {prefix}Content as a, {prefix}Sort as b"&whereStr&" and a.SortID=b.SortID and ContentStatus<>2 and b.SortType="&sortType&" "&orderStr	


	set rsObj = conn.Exec(sqlStr,"r1")
	rsObj.pagesize = numPerPage
	allRecordset = rsObj.recordcount : allPage= rsObj.pagecount
	if page>allPage then page=allPage
	if allRecordset=0 then
		if not isNul(keyword) then
		    echo "<tr align=""center""><td colspan='8' height=""28"">�ؼ��� <font color=red>"""&keyword&"""</font> û�м�¼</td></tr>" 
		else
		    echo "<tr align=""center""><td colspan='8' height=""28"">��û�м�¼!</td></tr>"
		end if
	else  
		dim getIndexImage,rpoststr
		rsObj.absolutepage = page
		for i = 1 to numPerPage
			getIndexImage=""
			rpoststr=""
			if rsObj("TimeStatus")=1 then rpoststr="[<font color=""red"">��ʱ����:"& rsObj("Timeing") &"</font>]"
			if not isnul(rsObj("IndexImage")) then getIndexImage="<img src=""../../images/image_s.gif"" class=""imgTip"" title='<img src="&rsObj("IndexImage")&"  width=135 />'/>"
			 echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf& _
				"<td><input type=""checkbox"" name=""id"" value="""&rsObj(0)&""" class=""checkbox"" /></td>"&vbcrlf& _
				"<td>"&rsObj(0)&"</td>"&vbcrlf& _
				"<td align=""left""><a href=""AspCms_ContentEdit.asp?id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" ><font color="""&rsObj("TitleColor")&""">"&rsObj(1)&"</font></a>"&getStr(rsObj("IsOutLink"),"<a href="""&rsObj("OutLink")&""" target=""_blank"" title=""�鿴""><img src=""../../images/view.gif"" border=0 /></a>","<a href="""&sitePath&setting.languagepath&"content/?"&rsObj(0)&".html"" target=""_blank"" title=""�鿴""><img src=""../../images/view.gif"" border=0 /></a>")&getStr(rsObj("IsHeadline"),"[<font color=""red"">ͷ</font>]","")&getStr(rsObj("IsFeatured"),"[<font color=""red"">��</font>]","")&getStr(rsObj("IsTop"),"[<font color=""red"">��</font>]","")&getStr(rsObj("IsRecommend"),"[<font color=""red"">��</font>]","")&getStr(rsObj("IsImageNews"),"[<font color=""red"">ͼ</font>]","")&getStr(rsObj("IsOutLink"),"[<font color=""red"">��</font>]","")&getIndexImage&rpoststr&"</td>"&vbcrlf& _	
				"<td>"&rsObj("SortName")&"</td>"&vbcrlf& _			
				"<td>"&rsObj(5)&"</td>"&vbcrlf& _
				"<td><input type=""hidden"" name=""nid"" value="""&rsObj(0)&""" style=""width:40px;""><input type=""text"" name=""order"" value="""&rsObj(8)&""" style=""width:40px; text-align:center;""/></td>"&vbcrlf& _
				"<td>"&getStr(rsObj(3),"<a href=""?action=off&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" title=""����"" ><IMG src=""../../images/toolbar_ok.gif""></a>","<a href=""?action=on&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" title=""����"" ><IMG src=""../../images/toolbar_no.gif""></a>")&"</td>"&vbcrlf& _
				"<td><a href=""AspCms_ContentEdit.asp?id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" >�޸�</a> | <a href=""?action=del&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&"""  onClick=""return confirm('ȷ��Ҫɾ����')"">ɾ��</a></td>"&vbcrlf& _
			  "</tr>"&vbcrlf
			rsObj.movenext
			if rsObj.eof then exit for
		next
		echo"<tr bgcolor=""#FFFFFF"" class=""pagenavi"">"&vbcrlf& _
			"<td colspan=""8"" height=""28"" style=""padding-left:20px;"">"&vbcrlf& _			
			"ҳ����"&page&"/"&allPage&"  ÿҳ"&numPerPage &" �ܼ�¼��"&allRecordset&"�� <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page=1&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">��ҳ</a> <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&(page-1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">��һҳ</a> "&vbcrlf
		dim pageNumber
		pageNumber=makePageNumber_(page, 10, allPage, "newslist",sortID, order,keyword)
		echo pageNumber
		echo"<a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&(page+1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">��һҳ</a> <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&allPage&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">βҳ</a>"&vbcrlf
			
		echo"ÿҳ��ʾ"&getPageSize("10",psize)&getPageSize("20",psize)&getPageSize("30",psize)&getPageSize("50",psize)
				
		echo"��"&vbcrlf& _	
		"</td>"&vbcrlf& _
		"</tr>"&vbcrlf
	end if
	rsObj.close : set rsObj = nothing	
End Sub

Sub listContent_

	dim i,orderStr,whereStr,sqlStr,rsObj,allPage,allRecordset,numPerPage
	numPerPage=psize
	if not isnum(numPerPage) then numPerPage=10	
	if isnul(order) then order="ContentID"
	if isnul(ordsc) then ordsc="desc"

	if isNul(page) then page=1 else page=clng(page)
	if page=0 then page=1
	
	orderStr=" order by "&order&" "&ordsc
	whereStr=" where a.LanguageID="&session("languageID")
	
	if not isNul(sortid) and sortid<>"0" then  whereStr=whereStr&" and a.SortID="&SortID
	if not isNul(keyword) then whereStr = whereStr&" and Title like '%"&keyword&"%'"

	sqlStr = "select ContentID,Title,visits,ContentStatus,a.SortID,a.AddTime,IsTop,a.IndexImage,ContentOrder,SortName,IsRecommend,TitleColor from {prefix}Content as a, {prefix}Sort as b"&whereStr&" and a.SortID=b.SortID and ContentStatus=2 "&orderStr	

	set rsObj = conn.Exec(sqlStr,"r1")
	rsObj.pagesize = numPerPage
	allRecordset = rsObj.recordcount : allPage= rsObj.pagecount
	if page>allPage then page=allPage
	if allRecordset=0 then
		if not isNul(keyword) then
		    echo "<tr align=""center""><td colspan='8' height=""28"">�ؼ��� <font color=red>"""&keyword&"""</font> û�м�¼</td></tr>" 
		else
		    echo "<tr align=""center""><td colspan='8' height=""28"">��û�м�¼!</td></tr>"
		end if
	else  
		rsObj.absolutepage = page
		for i = 1 to numPerPage
			echo "<tr bgcolor=""#ffffff"" align=""center"" onMouseOver=""this.bgColor='#CDE6FF'"" onMouseOut=""this.bgColor='#FFFFFF'"">"&vbcrlf& _
				"<td><input type=""checkbox"" name=""id"" value="""&rsObj(0)&""" class=""checkbox"" /></td>"&vbcrlf& _
				"<td>"&rsObj(0)&"</td>"&vbcrlf& _
				"<td><font color="""&rsObj("TitleColor")&""">"&rsObj(1)&"</font>"&"<a href=""/"&sitePath&"/?"&rsObj(4)&"_"&rsObj(0)&".html"" target=""_blank"" title=""�鿴""><img src=""../../images/view.gif"" border=0 /></a>"&"</td>"&vbcrlf& _	
				"<td>"&rsObj("SortName")&"</td>"&vbcrlf& _			
				"<td>"&rsObj(5)&"</td>"&vbcrlf& _
				"<td><input type=""hidden"" name=""nid"" value="""&rsObj(0)&""" style=""width:40px;""><input type=""text"" name=""order"" value="""&rsObj(8)&""" style=""width:40px; text-align:center;""/></td>"&vbcrlf& _
				"<td>"&"<IMG src=""../../images/toolbar_del.gif"">"&"</td>"&vbcrlf& _
				"<td> <a href=""?action=recovery&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""" >�ָ�</a> <a href=""?action=tdel&id="&rsObj(0)&"&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&"""  onClick=""return confirm('ȷ��Ҫɾ����')"">ɾ��</a></td>"&vbcrlf& _
			  "</tr>"&vbcrlf
			rsObj.movenext
			if rsObj.eof then exit for
		next
		echo"<tr bgcolor=""#FFFFFF"" class=""pagenavi"">"&vbcrlf& _
			"<td colspan=""8"" height=""28"" style=""padding-left:20px;"">"&vbcrlf& _			
			"ҳ����"&page&"/"&allPage&"  ÿҳ"&numPerPage &" �ܼ�¼��"&allRecordset&"�� <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page=1&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">��ҳ</a> <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&(page-1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">��һҳ</a> "&vbcrlf
		dim pageNumber
		pageNumber=makePageNumber_(page, 10, allPage, "newslist",sortID, order,keyword)
		echo pageNumber
		echo"<a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&(page+1)&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">��һҳ</a> <a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&allPage&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""">βҳ</a>"&vbcrlf
			
		echo"ÿҳ��ʾ"&getPageSize("10",psize)&getPageSize("20",psize)&getPageSize("30",psize)&getPageSize("50",psize)
				
		echo"��"&vbcrlf& _	
		"</td>"&vbcrlf& _
		"</tr>"&vbcrlf
	end if
	rsObj.close : set rsObj = nothing	
End Sub
'"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&""

Function getPageSize(ps,nps)
	if isnul(nps) then nps="10"
	if ps=nps then 
		getPageSize="<span>"&ps&"</span>"
	else
		getPageSize="<a href=""?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&ps&"&order="&order&"&ordsc="&ordsc&""">"&ps&"</a>"
	end if
End Function

Sub updateOrder
	Dim ids				:	ids=split(getForm("nid","post"),",")
	Dim orders		:	orders=split(getForm("order","post"),",")
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
			Conn.Exec "update {prefix}Content Set ContentOrder="&trim(orders(i))&" Where ContentID="&trim(ids(i)),"exe"	
		else
			Conn.Exec "update {prefix}Content Set ContentOrder=0 Where ContentID="&trim(ids(i)),"exe"	
		end if
	Next
	
	
	alertMsgAndGo "��������ɹ�",getPageName()&"?sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc
End Sub


Function addTags(tags)
	dim sql,rs,tag,tagIDs
	tags=split(tags,",")
	for each tag in tags
		if not isnul(tag) then 	
			sql="select TagID from {prefix}Tag where TagName='"&tag&"'"
			set rs=conn.exec(sql,"r1")
			if not rs.eof then 
				tagIDs=tagIDs&"{"&rs(0)&"}"
			else
				sql="insert into {prefix}Tag(TagName, TagCount, SortType, SortID, TagVisits, LanguageID, AddTime) values('"&tag&"',0,"&SortType&","&SortID&",0,"&LanguageID&",'"&now()&"')"
				conn.exec sql,"exe"
				tagIDs=tagIDs&"{"&conn.exec("select @@identity","r1")(0)&"}"
			end if
		end if		
	next
	addTags=tagIDs	
End Function
%>