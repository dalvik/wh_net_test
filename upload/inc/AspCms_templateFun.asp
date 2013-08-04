<%
Function makeAbout(Byval sortID, Byval page, Byval isMakeHtml)
	dim templateobj,templatePath : set templateobj=new TemplateClass
	dim rsObj,rsObjSmalltype,rsObjBigtype,channelTemplateName,tempStr,tempArr,pageStr,content
	dim templateFile
	
	set rsObj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")
	if rsObj.eof then makeAbout="" : exit function

	templateFile=rsObj("SortTemplate")
	if isnul(templateFile) then templateFile="about.html"
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/"	&templateFile
	
	if not CheckTemplateFile(templatePath) then echo templateFile&err_16 : exit function 
	if not isnul(rsObj("GroupID")) and isMakeHtml<>1 then	
		if not ViewNoRight(rsObj("GroupID"),rsObj("Exclusive")) then 
			echoErr err_17,"17",err_17
			response.end()
		end if		
	end if	
	
	'开始解析标签
	templateObj.load(templatePath)
	templateObj.parseHtml()
	templateObj.content=replace(templateObj.content,"{aspcms:sortname}",rsObj("SortName"))
	templateObj.content=replace(templateObj.content,"{aspcms:parentsortid}",rsObj("parentid"))		
	templateObj.content=replace(templateObj.content,"{aspcms:sortid}",SortID)	
	templateObj.content=replace(templateObj.content,"{aspcms:topsortid}",rsObj("topsortid"))
	if isnul(rsObj("PageKeywords")) then 
		templateObj.content=replace(templateObj.content,"[about:keyword]",setting.siteKeyWords)	
	else
		templateObj.content=replace(templateObj.content,"[about:keyword]",rsObj("PageKeywords"))	
	end if
	if isnul(rsObj("PageDesc")) then
		templateObj.content=replace(templateObj.content,"[about:desc]",setting.sitedesc)
	else
		templateObj.content=replace(templateObj.content,"[about:desc]",rsObj("PageDesc"))
	end if
	if not isNul(rsObj("IndexImage")) then 
		if instr(rsObj("IndexImage"),"http://")>0 then 
			templateObj.content = replace(templateObj.content,"[about:IndexImage]",rsObj("IndexImage"))
		else
			templateObj.content = replace(templateObj.content,"[about:IndexImage]",rsObj("IndexImage"))
		end if
	else
		templateObj.content = replace(templateObj.content,"[about:IndexImage]",sitePath&"/"&"Images/nopic.gif")			
	end if
	if isnul(rsObj("PageTitle")) then 
		templateObj.content=replace(templateObj.content,"[about:title]",rsObj("SortName"))	
	else
		templateObj.content=replace(templateObj.content,"[about:title]",rsObj("PageTitle"))
	end if
	if isnul(rsObj("IndexImage")) then
		templateObj.content=replace(templateObj.content,"[about:pic]",sitePath&"/images/nopic.gif" )
	else 
		if instr(rsObj("IndexImage"),"http://")>0 then 
			templateObj.content=replace(templateObj.content,"[about:pic]",rsObj("IndexImage"))
		else
			templateObj.content=replace(templateObj.content,"[about:pic]",rsObj("IndexImage"))
		end if
	end if
	templateObj.parsePosition(SortID)
	templateObj.parseCommon
	
	content=decodeHtml(rsObj("SortContent"))
	
	tempStr=templateObj.content
	if isExistStr(content,"{aspcms:page}") then		
		dim parr ,j
		tempArr=split(content,"{aspcms:page}")
		parr=ubound(tempArr)
		if parr=-1 then parr=0
		dim htmlfilename,htmlPath
		htmlfilename="?"&sortID
		if runMode=1 then htmlfilename=replace(rsObj("SortFolder"), "{sitepath}", sitePath)&replace(rsObj("SortFileName"), "{sortid}", sortID)
		htmlPath=replace(rsObj("SortFolder"), "{sitepath}", sitePath)

		for j=0 to parr						
			Page=clng(Page)
			pageStr=""
			if Page<1 then Page=1 : end if
			if Page>ubound(tempArr)+1 then Page=ubound(tempArr)+1 : end if			
			
			if Page>2 then					
				pageStr=pageStr+"<div class='pages'><a href='"&htmlfilename&"_"&Page-1&FileExt&"'>上一页</a>"
			else
				pageStr=pageStr+"<div class='pages'><a href='"&htmlfilename&FileExt&"'>上一页</a>"
			end if
			
			
			pageStr=pageStr+makePageNumber(Page,10,ubound(tempArr)+1,"about",sortID,htmlfilename)
		
			if Page=ubound(tempArr)+1 then
				pageStr=pageStr+"<a href='"&htmlfilename&"_"&ubound(tempArr)+1&FileExt&"'>下一页</a></div>"
			else
				pageStr=pageStr+"<a href='"&htmlfilename&"_"&Page+1&FileExt&"'>下一页</a></div>"
			end if
			
			if runMode=1 and isMakeHtml=1 then	
				templateObj.content=replace(tempStr,"[about:info]",tempArr(j)+pageStr)	
				if j=0 then 
					echo htmlfilename&FileExt&"<br>"
					createTextFile  templateObj.content, htmlfilename&FileExt, ""
				else
					echo htmlfilename&"_"&j+1&FileExt&"<br>"
					createTextFile  templateObj.content, htmlfilename&"_"&j+1&FileExt, ""				
				end if		
			else
				templateObj.content=replace(templateObj.content,"[about:info]",tempArr(Page-1)+pageStr)
				makeAbout=templateObj.content
			end if
			Page=Page+1
		next
	else 
		templateObj.content=replace(templateObj.content,"[about:info]",content)
		makeAbout=templateObj.content
		if isMakeHtml then 
			dim htmlfilepath
			htmlfilepath=templateObj.getSortLink(rsObj("sortType"), rsObj("sortID"), rsObj("sortUrl"), rsObj("sortFolder"), rsObj("sortFileName"),rsObj("GroupID"),rsObj("Exclusive"))
			echo htmlfilepath&"<br>"
			createTextFile  templateObj.content,htmlfilepath,""
		end if
		
	end if
	
	rsObj.close : set rsObj=nothing
	
	set templateobj=nothing
End Function

Function makeUser(Byval userid, Byval page, Byval isMakeHtml)
	dim templateobj,templatePath : set templateobj=new TemplateClass
	dim rsObj,rsObjSmalltype,rsObjBigtype,channelTemplateName,tempStr,tempArr,pageStr,content
	dim templateFile
	set rsObj=conn.exec("select * from {prefix}user where userid="&userid, "exe")
	if rsObj.eof then makeUser="" : exit function
	
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/userinfo.html"
	'die templatePath
	if not CheckTemplateFile(templatePath) then echo "userinfo.html"&err_16 : exit function 

	'开始解析标签
	templateObj.load(templatePath)
	templateObj.parseHtml()

		templateObj.content=replace(templateObj.content,"[about:keyword]",setting.siteKeyWords)	

		templateObj.content=replace(templateObj.content,"[about:desc]",setting.sitedesc)
	
	templateObj.content=replace(templateObj.content,"[user:userid]",rsObj("UserID"))
	templateObj.content=replace(templateObj.content,"[user:username]",rsObj("LoginName"))		
	templateObj.content=replace(templateObj.content,"[user:UserStatus]",rsObj("UserStatus"))
	if not isnul(rsObj("RegTime")) then 		
	templateObj.content=replace(templateObj.content,"[user:RegTime]",rsObj("RegTime"))
		else
	templateObj.content=replace(templateObj.content,"[user:RegTime]","null")	
	end if
	if not isnul(rsObj("LastLoginIP")) then 	
	templateObj.content=replace(templateObj.content,"[user:LastLoginIP]",rsObj("LastLoginIP"))
		else
	templateObj.content=replace(templateObj.content,"[user:LastLoginIP]","null")	
	end if
	if not isnul(rsObj("LastLoginTime")) then 		
	templateObj.content=replace(templateObj.content,"[user:LastLoginTime]",rsObj("LastLoginTime"))
		else
	templateObj.content=replace(templateObj.content,"[user:LastLoginTime]","null")	
	end if	
	if not isnul(rsObj("LoginCount")) then 	
	templateObj.content=replace(templateObj.content,"[user:LoginCount]",rsObj("LoginCount"))
		else
	templateObj.content=replace(templateObj.content,"[user:LoginCount]","null")	
	end if	
	if not isnul(rsObj("Gender")) then 	
	templateObj.content=replace(templateObj.content,"[user:Gender]",rsObj("Gender"))
		else
	templateObj.content=replace(templateObj.content,"[user:Gender]","null")	
	end if
	
	if not isnul(rsObj("TrueName")) then 	
	templateObj.content=replace(templateObj.content,"[user:TrueName]",rsObj("TrueName"))
		else
	templateObj.content=replace(templateObj.content,"[user:TrueName]","null")	
	end if
	if not isnul(rsObj("Birthday")) then 	
	templateObj.content=replace(templateObj.content,"[user:Birthday]",rsObj("Birthday"))	
	else
	templateObj.content=replace(templateObj.content,"[user:Birthday]","null")	
	end if
	if not isnul(rsObj("Country")) then 	
	templateObj.content=replace(templateObj.content,"[user:Country]",rsObj("Country"))	
	else
	templateObj.content=replace(templateObj.content,"[user:Country]","null")	
	end if
	if not isnul(rsObj("Province")) then 
	templateObj.content=replace(templateObj.content,"[user:Province]",rsObj("Province"))
	else
	templateObj.content=replace(templateObj.content,"[user:Province]","null")
	end if
	if not isnul(rsObj("City")) then 	
	templateObj.content=replace(templateObj.content,"[user:City]",rsObj("City"))
	else
	templateObj.content=replace(templateObj.content,"[user:City]","null")
	end if
	if not isnul(rsObj("Address")) then 	
	templateObj.content=replace(templateObj.content,"[user:Address]",rsObj("Address"))
	else
	templateObj.content=replace(templateObj.content,"[user:Address]","null")
	end if
	if not isnul(rsObj("PostCode")) then 	
	templateObj.content=replace(templateObj.content,"[user:PostCode]",rsObj("PostCode"))
	else
	templateObj.content=replace(templateObj.content,"[user:PostCode]","null")
	end if
	if not isnul(rsObj("Phone")) then 	
	templateObj.content=replace(templateObj.content,"[user:Phone]",rsObj("Phone"))
	else
	templateObj.content=replace(templateObj.content,"[user:Phone]","null")
	end if
	if not isnul(rsObj("Mobile")) then 
	templateObj.content=replace(templateObj.content,"[user:Mobile]",rsObj("Mobile"))
	else
	templateObj.content=replace(templateObj.content,"[user:Mobile]","null")
	end if
	if not isnul(rsObj("Email")) then 	
	templateObj.content=replace(templateObj.content,"[user:Email]",rsObj("Email"))
	else
	templateObj.content=replace(templateObj.content,"[user:Email]","null")
	end if
	if not isnul(rsObj("QQ")) then 	
	templateObj.content=replace(templateObj.content,"[user:QQ]",rsObj("QQ"))
	else
	templateObj.content=replace(templateObj.content,"[user:QQ]","null")
	end if
	if not isnul(rsObj("MSN")) then 	
	templateObj.content=replace(templateObj.content,"[user:MSN]",rsObj("MSN"))	
	else
	templateObj.content=replace(templateObj.content,"[user:MSN]","null")	
	end if
	'die templateObj.content
	templateObj.parseCommon
	makeUser=templateObj.content 
	rsObj.close : set rsObj=nothing
	set templateobj=nothing
End Function



function makeContent(Byval contentID, Byval page, Byval isMakeHtml)

	dim str 	
	dim templateFile
	dim templateobj,TemplatePath : set templateobj=new TemplateClass
	dim rsObj,rsObjSmalltype,rsObjBigtype,channelTemplateName,tempStr,tempArr,pageStr,sql,sperStr,sperStrs,content,contentLink
	dim ParentID,SortName,topsortid
	sperStrs=conn.exec("select SpecCategory+'_'+SpecField from {prefix}SpecSet Order by SpecOrder Asc,SpecID", "arr")
	dim spec	
	if isarray(sperStrs) then
		for each spec in sperStrs
			sperStr=sperStr&","&spec
		next
	end if
	set rsobj=conn.exec("select ContentID,a.SortID,a.GroupID,a.Exclusive,b.GroupID,b.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,a.PageKeywords,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,SortTemplate,ParentID,TopSortID,ContentTemplate,b.GroupID,b.Exclusive,b.GroupID,IsNoComment "&sperStr&",a.DownGroupID,a.VideoGroupID from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&"and a.SortID=b.SortID and ContentStatus=1 and TimeStatus=0 and IsOutLink=0 and ContentID="&ContentID, "exe")							
	if rsObj.eof then makeContent="" : exit function	
	templateFile=rsObj("ContentTemplate")
	if isnul(templateFile) then 
		select case rsObj("SortType")
			case "1"
				templateFile="about.html"
			case "2"
				templateFile="news.html"
			case "3"
				templateFile="product.html"
			case "4"
				templateFile="down.html"
			case "5"
				templateFile="job.html"
			case "6"
				templateFile="album.html"		
		end select
	end if	
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/"	&templateFile	
	'die templatePath
	if not CheckTemplateFile(templatePath) then echo templateFile&err_16 : exit function
	
	if not isnul(rsObj("a.GroupID")) then	
		if not ViewNoRight(rsObj("a.GroupID"),rsObj("a.Exclusive")) or not ViewNoRight(rsObj("b.GroupID"),rsObj("b.Exclusive")) then 
			if isMakeHtml=1 then 
				exit function
			else
				echoErr err_17,"17",err_17
				response.end()
			end if
		end if		
	end if	
	templateObj.load(TemplatePath)
	
	
	
	'加载评论模板
	'die sitePath&"plug/comment/comment.html"
	if SwitchComments=1 and rsObj("IsNoComment")=0 then
		templateObj.content=replace(templateObj.content,"{aspcms:comment}",loadFile(sitePath&"/plug/comment/comment.html"))	
	else
		templateObj.content=replace(templateObj.content,"{aspcms:comment}","")
	end if
	
	
	
	templateObj.parseHtml()
	templateObj.content=replace(templateObj.content,"{aspcms:sortid}",rsObj("SortID"))
	templateObj.content=replace(templateObj.content,"{aspcms:sortname}",rsObj("SortName"))
	templateObj.content=replace(templateObj.content,"{aspcms:parentsortid}",rsObj("ParentID"))
	templateObj.content=replace(templateObj.content,"{aspcms:topsortid}",rsObj("TopSortID"))
	templateObj.parsePosition(rsObj("SortID")) 
	
	
	templateObj.parsePrevAndNext contentID, rsObj("SortID")	
	if isExistStr(templateObj.content,"[about:") then str="about"
	if isExistStr(templateObj.content,"[content:") then str="content"
	if isExistStr(templateObj.content,"[news:") then str="news"
	if isExistStr(templateObj.content,"[product:") then str="product"
	if isExistStr(templateObj.content,"[down:") then str="down"
	if isExistStr(templateObj.content,"[pic:") then str="pic"
	
	'die isExistStr(templateObj.content,"[news:") &"AAAAAAAAAAA"&str
	templateObj.content=replace(templateObj.content,"["&str&":id]",contentID)
	
	templateObj.content=replace(templateObj.content,"["&str&":title]",rsObj("Title"))
	templateObj.content=replace(templateObj.content,"["&str&":titlecolor]",rsObj("TitleColor"))
	templateObj.content=replace(templateObj.content,"["&str&":author]",repnull(rsObj("Author")))
	templateObj.content=replace(templateObj.content,"["&str&":source]",repnull(rsObj("ContentSource")))
	templateObj.content=replace(templateObj.content,"["&str&":videourl]",repnull(rsObj("ContentSource")))
	templateObj.content=replace(templateObj.content,"["&str&":sortname]",rsObj("SortName"))
	templateObj.content=replace(templateObj.content,"["&str&":sortlink]",templateObj.getSortLink(rsObj("sortType"),rsObj("sortID"),rsObj("sortUrl"),rsObj("sortFolder"),rsObj("sortFileName"),rsObj("b.GroupID"),rsObj("b.Exclusive")))
	templateObj.content=replace(templateObj.content,"["&str&":date]",rsObj("Addtime"))
	templateObj.content=replace(templateObj.content,"["&str&":visits]","<script src="""&sitePath&"/inc/AspCms_Visits.asp?id="&contentID&"""></script>")
	templateObj.content=replace(templateObj.content,"["&str&":tag]",getTags(rsObj("ContentTag")))	
	templateObj.content=replace(templateObj.content,"["&str&":linktag]",tagsLink(getTags(rsObj("ContentTag"))))		
	templateObj.parseLoop("aboutcontent")	
	templateObj.content=replace(templateObj.content,"["&str&":istop]",rsObj("istop"))
	templateObj.content=replace(templateObj.content,"["&str&":isrecommend]",rsObj("isrecommend"))
	templateObj.content=replace(templateObj.content,"["&str&":isimage]",rsObj("IsImageNews"))
	templateObj.content=replace(templateObj.content,"["&str&":isfeatured]",rsObj("isfeatured"))	
	templateObj.content=replace(templateObj.content,"["&str&":isheadline]",rsObj("isheadline"))	
	contentLink=templateObj.getContentLink(rsObj("SortID"),rsObj("ContentID"),rsObj("SortFolder"),rsObj("a.GroupID"),rsObj("ContentFolder"),rsObj("ContentFileName"),rsObj("AddTime"),rsobj("PageFileName"),rsObj("b.GroupID"))
	templateObj.content=replace(templateObj.content,"["&str&":link]","http://"&setting.siteUrl&"/"&contentLink)
	templateObj.content=replace(templateObj.content,"["&str&":downurl]","<a href="&repnull(rsObj("DownURL"))&" >点击下载</a>")
	if isnul(rsObj("PageKeywords")) then 
		templateObj.content=replace(templateObj.content,"["&str&":keyword]",setting.siteKeyWords)	
	else
		templateObj.content=replace(templateObj.content,"["&str&":keyword]",rsObj("PageKeywords"))	
	end if
	if isnul(rsObj("PageDesc")) then
		templateObj.content=replace(templateObj.content,"["&str&":desc]",left(dropHtml(rsObj("Content")),100))
	else
		templateObj.content=replace(templateObj.content,"["&str&":desc]",rsObj("PageDesc"))
	end if
	if isnul(rsObj("IndexImage")) then
		templateObj.content=replace(templateObj.content,"["&str&":pic]",sitePath&"/images/nopic.gif" )
	else 
		if instr(rsObj("IndexImage"),"http://")>0 then 
			templateObj.content=replace(templateObj.content,"["&str&":pic]",rsObj("IndexImage"))
		else
			templateObj.content=replace(templateObj.content,"["&str&":pic]",rsObj("IndexImage"))
		end if
	end if
	dim imagepath
	if not isnul(rsObj("imagepath")) then 
		dim i,images
		images=split(rsObj("imagepath"),"|")
		for i=0 to ubound(images)
			if not isnul(images(i)) then
				if instr(rsObj("IndexImage"),"http://")>0 then 			
					imagepath=imagepath&"{showtit:'',showtxt:'',smallpic:'"&images(i)&"','bigpic':'"&images(i)&"'}"
				else
					imagepath=imagepath&"{showtit:'',showtxt:'',smallpic:'"&images(i)&"','bigpic':'"&images(i)&"'}"
				end if	
				if i<>ubound(images) then imagepath=imagepath&","
			end if
		next
	end if

	templateObj.content=replace(templateObj.content,"["&str&":pics]",ImagePath)
	'die rsObj("SortType")
	'if rsObj("SortType")="3" then
		if isarray(sperStrs) then
			for each spec in sperStrs
				templateObj.content=replace(templateObj.content,"["&str&":"&spec&"]",decodeHtml(rsObj(spec)))
			next
		end if
	'end if
	'parseCommon 内含有if标记，若在参数没有准备做if解析会发生错误
	templateObj.parseCommon()
	templateobj.parseLoop("aboutart")	
	content=decodeHtml(rsObj("Content"))	
	tempStr=templateObj.content	
	  
	if isExistStr(content,"{aspcms:page}") then
		dim parr ,j
		tempArr=split(content,"{aspcms:page}")
		parr=ubound(tempArr)
		if parr=-1 then parr=0
		for j=0 to parr			
			Page=clng(Page)
			pageStr=""
			if Page<1 then Page=1 : end if
			if Page>ubound(tempArr)+1 then Page=ubound(tempArr)+1 : end if			
			dim htmlfilename
			htmlfilename="?"&contentID
			if runMode=1 then htmlfilename= replace(contentLink,FileExt,"")
	
			if Page>2 then					
				pageStr=pageStr+"<div class='pages'><a href='"&htmlfilename&"_"&Page-1&FileExt&"'>上一页</a>"
			else
				pageStr=pageStr+"<div class='pages'><a href='"&htmlfilename&FileExt&"'>上一页</a>"
			end if			
			
			pageStr=pageStr+makePageNumber(Page,10,ubound(tempArr)+1,"about",contentID,htmlfilename)
		
			if Page=ubound(tempArr)+1 then
				pageStr=pageStr+"<a href='"&htmlfilename&"_"&ubound(tempArr)+1&FileExt&"'>下一页</a></div>"
			else
				pageStr=pageStr+"<a href='"&htmlfilename&"_"&Page+1&FileExt&"'>下一页</a></div>"
			end if
			pageStr=pageStr&"<script src="""&sitePath&"/inc/AspCms_VisitsAdd.asp?id="&contentID&"""></script>"
			if runMode=1 and isMakeHtml=1 then
				templateObj.content=replace(tempStr,"["&str&":info]",tempArr(Page-1)+pageStr)
				if j=0 then 
					echo replace(contentLink,FileExt,"")&FileExt&"<br>"
					createTextFile makeContentImages(templateObj.content), replace(contentLink,FileExt,"")&FileExt, ""
				else
					echo replace(contentLink,FileExt,"")&"_"&j+1&FileExt&"<br>"
					createTextFile makeContentImages(templateObj.content), replace(contentLink,FileExt,"")&"_"&j+1&FileExt, ""				
				end if	
			else
				makeContent=makeContentImages(templateObj.content)			
				templateObj.content=replace(templateObj.content,"["&str&":info]",tempArr(Page-1)+pageStr)				
			end if
			Page=Page+1
		next
	else 
		templateObj.content=replace(templateObj.content,"["&str&":info]",content&"<script src="""&sitePath&"/inc/AspCms_VisitsAdd.asp?id="&contentID&"""></script>")
		makeContent=makeContentImages(templateObj.content)
		if isMakeHtml then createTextFile  makeContent, contentLink, "" :echo contentLink&"<br>"
	end if	
	rsObj.close : set rsObj=nothing	
	set templateobj=nothing 	
End Function


Function makeContentImages(sContent)
'{aspcms:cimages count=5 contentid=1}
'[cimages:src]
'{/aspcms:cimages}

'if instr(sConetnt,"{aspcms:cimages") then
	dim rs,sql,img,imgs
	dim m_labelRule,m_labelRuleField
	dim regExpObj
	dim match,matches
	dim m_contentid
	dim m_maxcount,iCount
	dim soutput
	
	soutput = ""
	m_contentid = empty
	iCount = 0
	sql = "select * from {prefix}Content where contentid="
	'set rs = conn.exec(sql,"r1")
	set regExpObj= new RegExp	
	
	m_labelRule="{aspcms:cimages([\s\S]*?)}([\s\S]*?){/aspcms:cimages}"

	regExpObj.Pattern=m_labelRule
	set matches=regExpObj.Execute(sContent)


	for each match in matches
		'echo "ci"
		m_contentid = parseArr(match.SubMatches(0))("contentid")
		
		m_maxcount = parseArr(match.SubMatches(0))("count")
		if isnul(m_maxcount) or not isnumeric(m_maxcount) then m_maxcount = 9999		
		
		if not isnul(m_contentid) then
		if not isnumeric(m_contentid) then m_contentid=-1
			sql = "select imagepath from {prefix}Content where ContentID=" & m_contentid
			set rs = conn.exec(sql,"r1")		
			if not rs.eof then
				img=rs(0)
			end if
			rs.close
			set rs = nothing
			
			imgs = split(img,"|")
			
			for each img in imgs
			if iCount < cint(m_maxcount) then
				soutput = soutput & match.SubMatches(1)
				'soutput = replace(soutput,"[aspcms:cimagesitem]","")
				'soutput = replace(soutput,"[/aspcms:cimagesitem]","")
				soutput =  replace(soutput,"[cimages:src]",img)
				soutput = replace(soutput,"[cimages:i]",iCount+1)
				'echo img & "<br>"
			end if
			iCount = iCount + 1
			next
		end if
		
		
		'die soutput
		
		
		
		sContent=replaceStr(sContent,match.value,soutput)
		'die sConetnt
	next

	if instr(sContent,"{aspcms:cimages") > 0 then
	'die "停止"
	makeContentImages sContent
	end if


'end if
 
makeContentImages = sContent
End Function
'makeContentImages附属方法，临时
Function parseArr(Byval attr)
	dim attrStr,attrArray,attrDictionary,i,singleAttr,singleAttrKey,singleAttrValue
	dim strDictionary
	
	if not isobject(strDictionary) then set strDictionary=server.CreateObject(DICTIONARY_OBJ_NAME)
	
	attrStr = regExpReplace(attr,"[\s]+",chr(32))
	attrStr = trim(attrStr)
	attrArray = split(attrStr,chr(32))
	for i=0 to ubound(attrArray)
		singleAttr = split(attrArray(i),chr(61))
		singleAttrKey =  singleAttr(0) : singleAttrValue =  singleAttr(1)
		if not strDictionary.Exists(singleAttrKey) then strDictionary.add singleAttrKey,singleAttrValue else strDictionary(singleAttrKey) = singleAttrValue
	next
	set parseArr = strDictionary
End Function
'makeContentImages附属方法，临时
Function regExpReplace(contentstr,patternstr,replacestr)
dim regExpObj
if not isobject(regExpObj) then set regExpObj = new RegExp
	regExpObj.Pattern=patternstr
	regExpReplace=regExpObj.replace(contentstr,replacestr)
End Function



Function strList(Byval sortID, Byval page)
	dim templateFile
	dim templateobj,TemplatePath : set templateobj=new TemplateClass
	dim rsObj,rsObjSmalltype,rsObjBigtype,channelTemplateName,tempStr,tempArr,pageStr,sql,sperStr,sperStrs,content,contentLink
	
	set rsobj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")							
	if rsObj.eof then strList="" : exit function	
	templateFile=rsObj("SortTemplate")
	
	'echo templateFile
	if isnul(templateFile) then 
		select case rsObj("SortType")
			case "2"
				templateFile="newslist.html"
			case "3"
				templateFile="productlist.html"
			case "4"
				templateFile="downlist.html"
			case "5"
				templateFile="joblist.html"
			case "6"
				templateFile="albumlist.html"		
		end select
	end if
	'die templateFile
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/"	&templateFile	
	'die templatePath
	if not CheckTemplateFile(templatePath) then echo templateFile&err_16 : exit function
	if not isnul(rsObj("GroupID")) then		
		if not ViewNoRight(rsObj("GroupID"),rsObj("Exclusive")) then
			echoErr err_17,"17",err_17
			response.end()
		end if		
	end if	

	
	'开始解析标签
	templateObj.load(templatePath)	
	templateObj.parseHtml()
	templateObj.content=replace(templateObj.content,"{aspcms:sortname}",rsObj("SortName"))
	templateObj.content=replace(templateObj.content,"{aspcms:parentsortid}",rsObj("parentid"))		
	templateObj.content=replace(templateObj.content,"{aspcms:sortid}",sortID)	
	templateObj.content=replace(templateObj.content,"{aspcms:topsortid}",rsObj("topsortid"))
	
	
	'die replace(templateObj.content,"{aspcms:sortkeyword}",setting.siteKeyWords)
	
	if isnul(rsObj("PageKeywords")) then 
		templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",setting.siteKeyWords)	
	else
		templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",rsObj("PageKeywords"))	
	end if
	if isnul(rsObj("PageDesc")) then
		templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",setting.siteDesc)
	else
		templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",rsObj("PageDesc"))
	end if
	if isnul(rsObj("PageTitle")) then 
		templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("SortName"))	
	else
		templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("PageTitle"))
	end if
	'templateObj.content=replace(templateObj.content,"{aspcms:SortContent}",rsObj("SortContent"))
	if isnul(rsObj("SortContent")) then 
		templateObj.content=replace(templateObj.content,"{aspcms:SortContent}","")	
	else
		templateObj.content=replace(templateObj.content,"{aspcms:SortContent}",rsObj("SortContent"))
	end if
	
	if not isNul(rsObj("IndexImage")) then 
		if instr(rsObj("IndexImage"),"http://")>0 then 
			templateObj.content = replace(templateObj.content,"{aspcms:indeximage}",rsObj("IndexImage"))
		else
			templateObj.content = replace(templateObj.content,"{aspcms:indeximage}",rsObj("IndexImage"))
		end if
	else
		templateObj.content = replace(templateObj.content,"{aspcms:indeximage}",sitePath&"/"&"Images/nopic.gif")			
	end if
	
	templateObj.parsePosition(sortID)
	
	dim sortFolder, sortFileName
	sortFolder=rsObj("sortFolder")
	sortFileName=rsObj("sortFileName")		
	sortFolder=replace(sortFolder, "{sitepath}", sitePath)	
	sortFileName=replace(sortFileName, "{sortid}", sortID)			
	
	templateObj.parseList sortID,page,"list","",sortFolder&sortFileName&fileExt	
	templateObj.parseList sortID,page,"newslist","",sortFolder&sortFileName&fileExt	
	templateObj.parseList sortID,page,"piclist","",sortFolder&sortFileName&fileExt	
	templateObj.parseList sortID,page,"productlist","",sortFolder&sortFileName&fileExt	
	templateObj.parseList sortID,page,"downlist","",sortFolder&sortFileName&fileExt	
		
	'templateObj.parseList sortID,page,"list","","ff"
	templateObj.parseCommon		
	strList=templateObj.content 	
	rsObj.close : set rsObj=nothing
	set templateobj=nothing
End Function

Function struserbuy(Byval userid, Byval page)
	dim templateFile
	dim templateobj,TemplatePath : set templateobj=new TemplateClass
	dim rsObj,rsObjSmalltype,rsObjBigtype,channelTemplateName,tempStr,tempArr,pageStr,sql,sperStr,sperStrs,content,contentLink

	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/userbuy.html"
	'die templatePath
	if not CheckTemplateFile(templatePath) then echo templateFile&err_16 : exit function
	
	'开始解析标签
	templateObj.load(templatePath)	
	templateObj.parseHtml()
		'templateObj.content=replace(templateObj.content,"{aspcms:sortname}",rsObj("SortName"))
	'templateObj.content=replace(templateObj.content,"{aspcms:parentsortid}",rsObj("parentid"))		
	'templateObj.content=replace(templateObj.content,"{aspcms:sortid}",sortID)	
	templateObj.content=replace(templateObj.content,"{aspcms:topsortid}","0")
	'die replace(templateObj.content,"{aspcms:sortkeyword}",setting.siteKeyWords)
	

		templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",setting.siteKeyWords)	

		templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",setting.siteDesc)			
	if isnul(page) then page=1
	templateObj.parseList userid,page,"userbuylist","",""
		
	'templateObj.parseList sortID,page,"list","","ff"
	templateObj.parseCommon		
	struserbuy=templateObj.content 	
	
	set templateobj=nothing
End Function


Function makeList(Byval sortID)
	dim templateFile,page
	dim templateobj,TemplatePath : set templateobj=new TemplateClass
	dim rsObj,rsObjSmalltype,rsObjBigtype,channelTemplateName,tempStr,tempArr,pageStr,sql,sperStr,sperStrs,content,contentLink

	
	set rsobj=conn.exec("select * from {prefix}Sort where SortID="&sortID, "exe")
	if rsObj.eof then strList="" : exit function	
	templateFile=rsObj("SortTemplate")
	if isnul(templateFile) then 
		select case rsObj("SortType")
			case "2"
				templateFile="newslist.html"
			case "3"
				templateFile="productlist.html"
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
	templateObj.content=replace(templateObj.content,"{aspcms:sortname}",rsObj("SortName"))
	templateObj.content=replace(templateObj.content,"{aspcms:parentsortid}",rsObj("parentid"))		
	templateObj.content=replace(templateObj.content,"{aspcms:sortid}",sortID)	
	templateObj.content=replace(templateObj.content,"{aspcms:topsortid}",rsObj("topsortid"))
	if isnul(rsObj("PageKeywords")) then 
		templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",setting.siteKeyWords)	
	else
		templateObj.content=replace(templateObj.content,"{aspcms:sortkeyword}",rsObj("PageKeywords"))	
	end if
	if isnul(rsObj("PageDesc")) then
		templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",setting.siteDesc)
	else
		templateObj.content=replace(templateObj.content,"{aspcms:sortdesc}",rsObj("PageDesc"))
	end if
	if isnul(rsObj("PageTitle")) then 
		templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("SortName"))	
	else
		templateObj.content=replace(templateObj.content,"{aspcms:sorttitle}",rsObj("PageTitle"))
	end if
	templateObj.parsePosition(sortID)
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
		templateObj.content=tempStr
		'.parseList typeIds,page,str&"list","",str	
		dim sortFolder, sortFileName
		sortFolder=rsObj("sortFolder")
		sortFileName=rsObj("sortFileName")		
		sortFolder=replace(sortFolder, "{sitepath}", sitePath)	
		sortFileName=replace(sortFileName, "{sortid}", sortID)			
		
		templateObj.parseList sortID,page,"list","",sortFolder&sortFileName&fileExt	
		templateObj.parseList sortID,page,"newslist","",sortFolder&sortFileName&fileExt	
		templateObj.parseList sortID,page,"piclist","",sortFolder&sortFileName&fileExt	
		templateObj.parseList sortID,page,"productlist","",sortFolder&sortFileName&fileExt	
		templateObj.parseList sortID,page,"downlist","",sortFolder&sortFileName&fileExt					
		templateObj.parseCommon	
	
		sortFileName=replace(sortFileName, "{page}", page)
		createTextFile templateObj.content, sortFolder&sortFileName&fileExt, ""
		
		echo sortFolder&sortFileName&fileExt&"<br>"
	next
		
	'makeList=templateObj.content 	
	rsObj.close : set rsObj=nothing
	set templateobj=nothing
End Function



%>