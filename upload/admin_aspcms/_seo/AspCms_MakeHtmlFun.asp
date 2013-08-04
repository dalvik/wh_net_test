<!--#include file="../inc/AspCms_SettingClass.asp" -->

<%
CheckLogin()
Server.ScriptTimeOut=36000
dim action : action=getForm("action","get")
dim actType : actType=getform("actType","get")
select case action
	case "day" : checkRunMode : makeByDay
	case "all" : checkRunMode : makeAll
	case "index" : checkRunMode : makeIndex
	case "about" : checkRunMode : makeAllAbout
	case "alllist" : checkRunMode : makeAllList
	case "allcontent" : checkRunMode : makeAllContent
	case "list" : checkRunMode : makeListBySort
	case "content" : checkRunMode : makeContentBySortID	
	case "baidu" : makeBaiduMap
	case "google" : makeGoogleMap
	case "rss" : makeRss
	case "site" : makeSiteMap
end select

Sub checkRunMode
	if runmode<>1 then die"<br><br>当前模式不是静态模式，不能生成！<a href=""../_system/AspCms_SiteSetting.asp"">点击此处修改运行模式</a><br><a href=""javascript:history.go(-1)"">点击此处返回</a><br><br>"
End Sub

Sub makeAll()
	'delAllHtml		
	'if isExistFile(sitePath&"/"&"index.html") then delFile sitePath&"/"&"index.html"	'删除首页	
	'if not isnul(htmlDir) and isExistFolder(sitePath&"/"&htmlDir) then delFolder(sitePath&"/"&htmlDir)
	makeIndex()
	makeAllAbout()
	makeAllList()
	makeAllContent()
	alertMsgAndGo "生成全站成功","AspCms_MakeHtml.asp?actType="&actType	
End Sub

Sub makeByDay
	dim rs,sql,EditTime,SortIDs
	EditTime=getForm("EditTime", "post")
	if not isdate(EditTime) then alertMsgAndGo "请输入正确的日期","-1"	
	sql="select * from {prefix}Content where ContentStatus=1 and LanguageID="&cint(session("languageID"))&" and DateDiff('d', EditTime, '"&EditTime&"')<=0 order by ContentID"
	set rs=conn.Exec(sql,"r1")	
	Do While not rs.Eof
		'echo "生成"""&rs("ContentID")&"""成功<br>"
		makeContent rs("ContentID"), 1, 1	
		rs.movenext
	Loop
	
	sql="select SortPath from {prefix}Sort where SortID in(select SortID from {prefix}Content where ContentStatus=1 and LanguageID="&cint(session("languageID"))&" and DateDiff('d', EditTime, '"&EditTime&"')<=0 group by SortID)"	
	set rs=conn.Exec(sql,"r1")	
	Do While not rs.Eof
		SortIDs=SortIDs&rs(0)
		rs.moveNext
	Loop
	if not isnul(SortIDs) then 
		sql="select SortID from {prefix}Sort where SortID in("&SortIDs&")"
		set rs=conn.Exec(sql,"r1")		
		Do While not rs.Eof	
			makeList( rs(0))
			rs.moveNext
		Loop	
	end if
	rs.close : set rs =nothing
	makeIndex()
	alertMsgAndGo "更新成功","AspCms_MakeHtml.asp?actType="&actType	
End Sub

'生成所有列表页
Sub makeAllList()
	dim rs,sql
	sql = "select SortID,SortType,SortName from {prefix}Sort where SortType<>1 and SortType<>7 and GroupID<3 and LanguageID="&cint(session("languageID"))&" order by SortID"
	set rs=conn.Exec(sql,"r1")	
	Do While not rs.Eof
		makeList( rs("SortID"))
		'echo "生成"""&rs("SortID")&rs("SortName")&"""成功<br>"
		rs.movenext
	Loop
	rs.close : set rs =nothing		
	if action<>"all" then alertMsgAndGo "生成列表页成功","AspCms_MakeHtml.asp?actType="&actType	
End Sub

'按条件生成列表页
Sub makeListBySort()
	dim rs,sql,SortID
	SortID=getForm("lsortid","post")
	if SortID=0 then alertMsgAndGo "请选择栏目","AspCms_MakeHtml.asp?actType="&actType	
	if not isnum(SortID) then alertMsgAndGo "分类ID不正确","AspCms_MakeHtml.asp?actType="&actType	
	dim sortids	
	sortids=getSubSort(SortID, 1)
	sql="select SortID,SortType,SortName from {prefix}Sort where SortType<>1 and SortType<>7 and GroupID<3 and LanguageID="&cint(session("languageID"))&" and SortID in ("&sortids&") order by SortID"
	set rs=conn.Exec(sql,"r1")
	Do While not rs.Eof
		makeList( rs("SortID"))
		'echo "生成"""&rs("SortName")&"""成功<br>"
		rs.movenext
	Loop
	rs.close : set rs =nothing
	
	'for each SortID in sortids
'		if isnum(SortID) then makeList(SortID)
'	next
	alertMsgAndGo "生成指定列表页成功","AspCms_MakeHtml.asp?actType="&actType	
End Sub

'生成首页
Sub makeIndex
	dim templateobj,templatePath : set templateobj = new TemplateClass	
	templatePath=sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/index.html"	
	'die templatePath
	if not CheckTemplateFile(templatePath) then echo "index.html"&err_16	
	with templateObj 
		.content=loadFile(templatePath) 
		.parseHtml()		
		.parseCommon		
		createTextFile  .content, sitePath&setting.LanguagePath&"index"&FileExt,""
	end with	
	set templateobj =nothing
	
	if action<>"all" and action<>"day" then alertMsgAndGo "生成首页成功","AspCms_MakeHtml.asp?actType="&actType		
End Sub

'生成所有内容页
Sub makeAllContent
	dim rs,sql
	sql="select * from {prefix}Content where ContentStatus=1 and GroupID<3 and LanguageID="&cint(session("languageID"))&" order by ContentID"
	set rs=conn.Exec(sql,"r1")	
	Do While not rs.Eof
		'echo "生成"""&rs("ContentID")&"""成功<br>"
		makeContent rs("ContentID"), 1, 1
		rs.movenext
	Loop
	rs.close : set rs =nothing
	
	if action<>"all" then alertMsgAndGo "生成内容页成功","AspCms_MakeHtml.asp?actType="&actType	
End Sub

'按分类生成
Sub makeContentBySortID
	dim rs,sql,SortID
	SortID=getForm("csortid","post")
	if SortID=0 then alertMsgAndGo "请选择栏目","AspCms_MakeHtml.asp?actType="&actType	
	sql="select * from {prefix}Content where ContentStatus=1 and GroupID<3 and SortID in("&getSubSort(SortID, 1)&") and LanguageID="&cint(session("languageID"))
	set rs=conn.Exec(sql,"r1")	
	Do While not rs.Eof		
		'echo "生成"""&rs("ContentID")&"""成功<br>"
		makeContent rs("ContentID"), 1, 1	
		rs.movenext
	Loop
	rs.close : set rs =nothing	
	if action<>"all" then alertMsgAndGo "生成指定内容页成功","AspCms_MakeHtml.asp?actType="&actType	
End Sub

'生成所有单页
Sub makeAllAbout
	dim rs,sql
	sql = "select SortID,SortType from {prefix}Sort where SortType=1 and GroupID<3 and LanguageID="&cint(session("languageID"))
	set rs=conn.exec(sql,"r1")
	do while not rs.Eof
		makeAbout rs("SortID"), 0, 1
		rs.movenext
	loop
	rs.close : set rs =nothing	
	if action<>"all" then alertMsgAndGo "生成所有单页成功","AspCms_MakeHtml.asp?actType="&actType
End Sub

'生成Baidu 站点地图
'by amysimple 2011-06-24
Sub makeBaiduMap
	Dim m_timespanB,m_timespanE:m_timespanB=timer
	Dim resultUrl,tempUrl
	Dim resultMsg
	Dim makenum
	
	if isNul(makenum) then makenum=1000 else makenum=clng(makenum)
	
	allmakenum=getForm("allmakenum","get") : if isNul(allmakenum) then allmakenum=5000 else allmakenum=clng(allmakenum)
	dim vDes,vName,i,j,rsObj,baiduStr,allmakenum,pagenum,xmlUrl,dt
	
	dim TemplateObj : set TemplateObj= new TemplateClass
	set rsObj=conn.Exec("select top "&allmakenum&" ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&"and a.SortID=b.SortID and ContentStatus=1  order by a.addtime desc","r1")
	
	rsObj.pagesize=makenum
	pagenum=rsObj.pagecount
	redim resultUrl(pagenum-1)
	for i=1 to pagenum
		rsObj.absolutepage=i
baiduStr =  "<?xml version=""1.0"" encoding=""utf-8"" ?><document><webSite>http://"&setting.siteUrl&sitePath&setting.languagePath&"</webSite><webMaster>"&setting.siteTitle&"</webMaster><updatePeri>1800</updatePeri>"
		for j=1 to rsObj.pagesize
			vDes=rsObj("ContentTag") : if isNul(vDes) then vDes=""
			vName=rsObj("Title") : if isNul(vName) then vName=""
			Dim link
				link="http://"&setting.siteUrl&TemplateObj.getContentLink(rsObj("SortID"),rsObj("ContentID"),rsObj("SortFolder"),rsObj("a.GroupID"),rsObj("ContentFolder"),rsObj("ContentFileName"),rsObj("AddTime"),rsobj("PageFileName"),rsObj("b.GroupID"))
	
baiduStr = baiduStr & "<item><title>"&server.htmlencode(vName)&"</title><link>"&link&"</link><text>"&server.htmlencode(filterStr(left(vDes, 300),"html"))&"</text>"
			If left(rsObj("ImagePath"), 7) = "http://"  Then
				baiduStr = baiduStr & "<image>"&rsObj("ImagePath")&"</image>"&vbcrlf
			Else
				baiduStr = baiduStr & "<image>http://"&setting.siteUrl&sitePath&rsObj("ImagePath")&"</image>"&vbcrlf
			End If
			dt=rsObj("AddTime")
baiduStr = baiduStr & "<keywords>"&rsObj("Title")&","&rsObj("Author")&"</keywords><author>"&setting.siteTitle&"</author><source>"&setting.siteTitle&"</source><pubDate>"&formatDate(dt,1)&"</pubDate></item>"
			rsObj.movenext
			if rsObj.eof then exit for
		next
		baiduStr = baiduStr + "</document>"
		tempUrl = sitePath&setting.languagePath&"baidu"&xmlUrl&"_"&i&".xml"
		createTextFile baiduStr,tempUrl,"utf-8"
		'echo "http://"&setting.siteUrl&sitePath&setting.languagePath&"baidu"&xmlUrl&".xml"&" 生成完毕 <a target='_blank' href='"&sitePath&setting.languagePath&"baidu"&xmlUrl&".xml'><font color=red>浏览</font></a><br>"
		resultUrl(i-1)="http://"&setting.siteUrl&tempUrl
	next
	rsObj.close : set rsObj=nothing
	
baiduStr = ""
	for i=0 to ubound(resultUrl,1)
		baiduStr = baiduStr & "sitemap:"&resultUrl(i)&Chr(13)
		
	next
	
	createTextFile baiduStr,sitePath&setting.languagePath&"baidu.txt","utf-8"		
	
	m_timespanE = timer
	resultMsg = "<dl>"
	resultMsg = resultMsg & "<dd>生成完毕！用时"&m_timespanE-m_timespanB&"秒</dd>"
	resultMsg = resultMsg & "<dd>已生成"&ubound(resultUrl,1)+1&"页</dd>"
	
	for i=0 to ubound(resultUrl,1)
		resultMsg = resultMsg & "<dd>"&resultUrl(i)&" 生成完毕 <a target='_blank' style='text-decoration:underline;color:red' href='"&resultUrl(i)&"'>浏览</a></dd>"
	next
	
	resultMsg = resultMsg & "<dd>"&"http://"&setting.siteUrl&"/baidu.txt 索引页生成完毕 <a target='_blank' style='text-decoration:underline;color:red' href='http://"&setting.siteUrl&"/baidu.txt'>浏览</a></dd>"
	resultMsg = resultMsg & "<dd>说明：百度目前不支持类似Google的索引格式，不过将baidu.txt内容复制到robots.txt中对百度有效</dd>"
	 
	resultMsg = resultMsg & "<dd><a style='text-decoration:underline;color:red' href='http://news.baidu.com/newsop.html' target='_blank'>立刻提交到Baidu!</a></dd>"
	resultMsg = resultMsg & "</dl>"
	ResultUI resultMsg
	
	
	'echo "生成完毕"
	'echo  "请通过<a href='http://news.baidu.com/newsop.html' target='_blank'>http://news.baidu.com/newsop.html</a>提交!"		
	'die"<br><a href=""javascript:history.go(-1)"">返回</a>！"
	
	
End Sub



'生成Google 站点地图
'by amysimple 2011-06-24
Sub makeGoogleMap
	Dim m_timespanB,m_timespanE:m_timespanB=timer
	Dim makenum
	Dim resultUrl,tempUrl
	Dim resultMsg,sortxml
	if isNul(makenum) then makenum=1000 else makenum=clng(makenum)
	
	allmakenum=getForm("allmakenum","get") : if isNul(allmakenum) then allmakenum=5000 else allmakenum=clng(allmakenum)
	
	dim i,j,rsObj,googleStr,allmakenum,pagenum,xmlUrl,googleDateArray,googleDate 
	
	dim TemplateObj : set TemplateObj= new TemplateClass
	sortxml=""
	Dim linkArray : linkArray =conn.Exec("select SortName,SortType,SortURL,sortID,(select count (*) from {prefix}Sort as a where a.ParentID=b.sortID) as subcount,SortFolder,SortFileName,GroupID,Exclusive from {prefix}Sort as b  where LanguageID="&setting.languageID&" and SortStatus=1 and sorttype<>7 order by SortOrder asc","arr")
	
	for i=0 to ubound(linkArray,2)
		sortxml = sortxml & "<url><loc>"&"http://"&setting.siteUrl&getSortLink(linkArray(1,i),linkArray(3,i),linkArray(2,i),linkArray(5,i),linkArray(6,i),linkArray(7,i),linkArray(8,i))&"</loc><lastmod>"&now()&"</lastmod></url>"	
		
	next
	
	set rsObj=conn.Exec("select top "&allmakenum&" ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&"and a.SortID=b.SortID and ContentStatus=1  order by a.addtime desc","r1")
	rsObj.pagesize=makenum
	pagenum=rsObj.pagecount
	redim resultUrl(pagenum-1)
	for i=1 to pagenum
		rsObj.absolutepage=i
		Dim link	
'此处优化节省1.5秒以上
'by amysimple
googleStr =  "<?xml version=""1.0"" encoding=""UTF-8""?><urlset xmlns=""http://www.google.com/schemas/sitemap/0.84""><url><loc>http://"&setting.siteUrl&sitePath&setting.languagePath&"</loc></url>" & sortxml
		for j=1 to rsObj.pagesize
				googleDateArray=rsObj("AddTime")
				googleDate=formatDate(googleDateArray,1)				
		link="http://"&setting.siteUrl&TemplateObj.getContentLink(rsObj("SortID"),rsObj("ContentID"),rsObj("SortFolder"),rsObj("a.GroupID"),rsObj("ContentFolder"),rsObj("ContentFileName"),rsObj("AddTime"),rsobj("PageFileName"),rsObj("b.GroupID"))
		
googleStr = googleStr & "<url><loc>"&link&"</loc><lastmod>"&googleDate&"</lastmod></url>"
				rsObj.movenext
				if rsObj.eof then exit for
		next
		
		googleStr = googleStr &"</urlset>"	
		tempUrl = sitePath&setting.languagePath&"google"&xmlUrl&"_"&i&".xml"
		createTextFile googleStr,tempUrl,"utf-8"
		resultUrl(i-1)="http://"&setting.siteUrl&tempUrl		
	next	
	rsObj.close : set rsObj=nothing
	
googleStr = "<?xml version='1.0' encoding='UTF-8'?><sitemapindex xmlns='http://www.sitemaps.org/schemas/sitemap/0.9'>"

	for i=0 to ubound(resultUrl,1)
		googleStr = googleStr & "<sitemap><loc>"&resultUrl(i)&"</loc><lastmod>"&googleDate&"</lastmod></sitemap>"		
	next
	
googleStr = googleStr & "</sitemapindex>"	
	createTextFile googleStr,sitePath&setting.languagePath&"google.xml","utf-8"	
	
	m_timespanE = timer
	resultMsg = "<dl>"
	resultMsg = resultMsg & "<dd>生成完毕！用时"&m_timespanE-m_timespanB&"秒</dd>"
	resultMsg = resultMsg & "<dd>已生成"&ubound(resultUrl,1)+1&"页</dd>"
	
	for i=0 to ubound(resultUrl,1)
		resultMsg = resultMsg & "<dd>"&resultUrl(i)&" 生成完毕 <a target='_blank' style='text-decoration:underline;color:red' href='"&resultUrl(i)&"'>浏览</a></dd>"
	next
	
	resultMsg = resultMsg & "<dd>"&"http://"&setting.siteUrl&"/google.xml 索引页生成完毕 <a target='_blank' style='text-decoration:underline;color:red' href='http://"&setting.siteUrl&"/google.xml'>浏览</a></dd>"
	
	 
	resultMsg = resultMsg & "<dd><a style='text-decoration:underline;color:red' href='http://www.google.com/webmasters/tools/' target='_blank'>立刻提交到Google!</a></dd>"
	resultMsg = resultMsg & "</dl>"
	ResultUI resultMsg
End Sub


Sub makeSiteMap	
	dim templateobj,templatePath : set templateobj=new TemplateClass	
	templatePath=sitePath&"/templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/sitemap.html"	
	if not CheckTemplateFile(templatePath) then echo templatePath&err_16 : exit Sub
	with templateObj 
		.content=loadFile(templatePath) 
		.parseHtml()		
		.parseCommon		
		createTextFile .content ,sitePath&setting.languagePath&"sitemap.html",""
	end with
	set templateobj =nothing : terminateAllObjects
	alertMsgAndGo "SiteMap生成成功!","-1" 
End Sub

Sub makeRssMap	
	dim templateobj,templatePath : set templateobj=new TemplateClass	
	templatePath=sitePath&"/templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/rssmap.html"	
	if not CheckTemplateFile(templatePath) then echo templatePath&err_16 : exit Sub

	with templateObj 
		.content=loadFile(templatePath) 
		.parseHtml()	
		.parseRssList("")	
		.parseCommon		
		createTextFile .content ,sitePath&setting.languagePath&"rssmap.html",""
	end with
	set templateobj =nothing 
End Sub

Sub makeRss
	Dim rssStr,rssStr1,rssStr2,rssStr3,sortid
	Dim rs	
	dim templateobj,templatePath : set templateobj=new TemplateClass	

	set rs=conn.exec("select SortID,SortName,SortURL,SortType,SortFolder,PageKeyWords,PageDesc from {prefix}Sort where SortType<>7 and LanguageID="&session("LanguageID"),"r1")

	do while not rs.eof
		rssStr1="<?xml version=""1.0"" encoding=""gbk"" ?><rss version=""2.0""><channel><title><![CDATA["&rs("SortName")&"_"&setting.siteTitle&"]]></title><link>http://"&setting.siteUrl&"</link><description><![CDATA["&setting.siteDesc&"]]></description><language>zh-cn</language><generator><![CDATA[Rss Powered By "&setting.siteUrl&"]]></generator><webmaster>"&setting.companyEmail&"</webmaster>"
		rssStr3="</channel></rss>"

		Dim rsObj,sql
		rssStr2=""
		sql="select top 100 ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&"and a.SortID=b.SortID and a.SortID="&rs("SortID")&" and ContentStatus=1 order by a.addtime desc"
		set rsObj=conn.Exec(sql,"r1")
		do while not rsObj.eof
rssStr2=rssStr2&"<item><title><![CDATA["&rsObj("Title")&"]]></title><link>http://"&setting.siteUrl&TemplateObj.getContentLink(rsObj("SortID"),rsObj("ContentID"),rsObj("SortFolder"),rsObj("a.GroupID"),rsObj("ContentFolder"),rsObj("ContentFileName"),rsObj("AddTime"),rsobj("PageFileName"),rsObj("b.GroupID"))&"</link><description><![CDATA["&left(dropHtml(decodeHtml(rsObj("Content"))),100)&"]]></description><pubDate>"&formatDate(rsObj("AddTime"),1)&"</pubDate><category>"&rs("SortName")&"</category><author>"&rsObj("Author")&"</author><comments>"&setting.siteTitle&"</comments></item>"
			rsObj.movenext
		loop
		rssStr=rssStr1&rssStr2&rssStr3
		createTextFile rssStr,sitePath&setting.languagePath&"rss/"&rs("SortID")&".xml",""	
		rs.movenext
	loop
	rsObj.close : set rsObj=nothing
	rs.close : set rs=nothing
	
	set templateobj =nothing 
	makeRssMap()	
	alertMsgAndGo "RSS生成成功!","-1" 
End Sub



Sub delAllHtml
	dim styles,style
	styles=split("news,down,pic,product,about,",",")	
	if isExistFile(sitePath&"/"&"index.html") then delFile sitePath&"/"&"index.html"	'删除首页	
	for each style in styles
		if style="news" or style="down" or style="pic" or style="product" then Delhtml(style&"list")	'删除列表页
		Delhtml(style)			'删除详细页				
	next	
	'删除指定的生成目录
	styles=""
	styles=conn.exec("select SortFolder from {prefix}Sort where not isnull(SortFolder)","arr")
	for each style in styles
		'if not isnul(style) then echo "/"&sitePath&style&"<br>"
		'Delhtml(style)			'删除详细页			
		if not isnul(style) and isExistFolder("/"&sitePath&style) then delFolder("/"&sitePath&style)	
	next
End Sub


'根据目录删除html文件
Sub delhtml(str)
	dim fileListArray,fileAttr,i
	fileListArray= getFileList("/"&sitePath&str)
	if instr(fileListArray(0),",")>0 then		
		for  i = 0 to ubound(fileListArray)
			fileAttr=split(fileListArray(i),",")	
			if GetExtend(fileAttr(0))=replace(FileExt,".","") then delFile fileAttr(4)	
		next		
	end if
End Sub

	Function getSortLink(sortType, sortID, sortUrl, sortFolder, sortFileName, GroupID, Exclusive)	
		sortFolder=replace(repnull(sortFolder), "{sitepath}", sitePath)	
		sortFileName=replace(repnull(sortFileName), "{sortid}", sortID)	
		sortFileName=replace(sortFileName, "{page}", "1")	
		if sortType="7" then 
				if isurl(sortUrl) then
					getSortLink=sortUrl
				else
					getSortLink=sitePath&sortUrl
				end if 				
		else	
			if runMode=1 and viewNoRight(GroupID, Exclusive) then
					getSortLink=sortFolder&sortFileName&fileExt
			else
				Select  case sortType					
					case "1" 			
						getSortLink=sitePath&setting.languagePath&""&"about"&"/?"&sortID&fileExt
					case else
						getSortLink=sitePath&setting.languagePath&""&"list"&"/?"&sortID&"_1"&fileExt											
				End Select
			end if
		end if
	End Function

Sub ResultUI(resultHTML)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../images/style.css" type=text/css rel=stylesheet>
</HEAD>

<BODY>
<DIV class=formzone>
<DIV class=namezone>结果页面</DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>

<TR>						
<TD width="96%" align="left">
<%=resultHTML%>
</TD>
</TR>
<tr>
<TD align=middle width=100 height=1></TD>        
</tr>
</TBODY>
</TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 
</DIV></DIV>

</BODY></HTML>
<%
die ""
End Sub
%>
