<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit

Const DebugMode = false
Dim AppSpan:AppSpan=timer

'Response.Charset = "gbk"

dim CONN_OBJ_NAME,RECORDSET_OBJ_NAME,DICTIONARY_OBJ_NAME,JPEG_OBJ_NAME,FSO_OBJ_NAME,STREAM_OBJ_NAME
CONN_OBJ_NAME="ADODB.CONNECTION"
RECORDSET_OBJ_NAME="ADODB.RECORDSET"
DICTIONARY_OBJ_NAME="SCRIPTING.DICTIONARY"
JPEG_OBJ_NAME="PERSITS.JPG"
FSO_OBJ_NAME="SCRI"&"PTING.FILES"&"YSTEMOBJECT"
STREAM_OBJ_NAME="ADOD"&"B.ST"&"REAM"

dim startTime,endTime
setStartTime

dim conn : set conn=new DBClass 
conn.databaseType=dbType

dim objFso,objStream
initAllObjects
%>

<!--#include file="../config/AspCms_Config.asp" -->
<!--#include file="../config/AspCms_PayMentConfig.asp" -->
<!--#include file="AspCms_CommonFun.asp" -->
<!--#include file="AspCms_Version.asp" -->
<!--#include file="AspCms_Getpy.asp" -->
<!--#include file="md5.asp" -->
<%
Class SettingClass
	Public languageID, languageName, languagePath, Alias, defaultTemplate, htmlFilePath, siteTitle, additionTitle, siteLogoUrl, siteUrl, companyName, companyAddress, companyPostCode, companyContact, companyPhone, companyMobile, companyFax, companyEmail, companyICP, statisticalCode, copyRight, siteKeywords, siteDesc, languageOrder,errid,errdes
	
	Private Sub Class_Initialize
		SettingInit
		Rposting
	End Sub
	
	Public Sub SettingInit
		dim rs
		if LanguageAlias="" then 
			set rs=conn.exec("select * from {prefix}Language where IsDefault=1","r1")
		else
			set rs=conn.exec("select * from {prefix}Language where Alias='"&LanguageAlias&"'","r1")
		end if
		if not rs.eof then
			languageID=rs("languageID")
			languageName=rs("languageName")
			languagePath=rs("languagePath")
			Alias=rs("Alias")
			defaultTemplate=rs("defaultTemplate")
			htmlFilePath=rs("htmlFilePath")
			siteTitle=rs("siteTitle")
			additionTitle=rs("additionTitle")
			siteLogoUrl=rs("siteLogoUrl")
			siteUrl=rs("siteUrl")
			companyName=rs("companyName")
			companyAddress=rs("companyAddress")
			companyPostCode=rs("companyPostCode")
			companyContact=rs("companyContact")
			companyPhone=rs("companyPhone")
			companyMobile=rs("companyMobile")
			companyFax=rs("companyFax")
			companyEmail=rs("companyEmail")
			companyICP=rs("companyICP")
			statisticalCode=rs("statisticalCode")
			copyRight=rs("copyRight")
			siteKeywords=rs("siteKeywords")
			siteDesc=rs("siteDesc")
			languageOrder=rs("languageOrder")
		else
			echoErr err_02,errid,errdes		
		end if
	End Sub
	
	Public Sub Class_Terminate
	
	End Sub
End Class

Class DBClass
	public dbConn,dbRs,isConnect,fetchCount
	private connStr,vqueryCount,vdbType
	private errid,errdes

	Private Sub Class_Initialize
		isConnect=false
		vqueryCount=0
		fetchCount=0
	End Sub
	
	Public Property Get queryCount
		queryCount=vqueryCount
	End Property

	Public Property Let databaseType(byval pType)
		vdbType=pType
	End Property

	Private Sub getConnStr()
		if vdbType="1" then
			connStr="Provider=Sqloledb;Data Source="&databaseServer&";Initial Catalog="&databaseName&";User ID="&databaseUser&";Password="&databasePwd&";"
		elseif vdbType="0" then
			connStr="Provider=Microsoft.Jet.OLEdb.4.0;Data Source="&server.mappath(sitePath&"/"&accessFilePath)
		end if
	End Sub

	Public Sub connect()
		getConnStr
		if isObject(dbConn)=false or isConnect=false then
			On Error Resume Next
			set dbConn=server.CreateObject(CONN_OBJ_NAME)
			dbConn.open connStr
			isConnect=true
			if Err then errid=Err.number:errdes=Err.description:Err.Clear:dbConn.close:set dbConn=nothing:isConnect=false:echoErr err_01,errid,errdes
		end if
	End Sub

	Function exec(byval sqlStr,byval sqlType)
	
		if not isConnect=true then connect
		If Not DebugMode Then On Error Resume Next
		sqlStr=replace(sqlStr,"{prefix}",tablePrefix)		
		set exec=server.CreateObject(RECORDSET_OBJ_NAME)
		if isnul(sqlStr) then exit function
		If DebugMode Then echo sqlStr &"<hr>"
		select case sqlType
			case "exe"
				err.clear
				set Exec=dbConn.execute(sqlStr)
			case "r1"
				exec.open sqlStr,dbConn,1,1
			case "r3"
				exec.open sqlStr,dbConn,3,3
			csae "openschema"
				exec.OpenSchema(16)
			case "arr"
				exec.open sqlStr,dbConn,1,1				
				if not exec.eof then
					if fetchCount=0 then  exec=exec.getRows() else exec=exec.getRows(fetchCount)
				end if
		end select
		vqueryCount=vqueryCount+1
		If Not DebugMode Then						
		if Err then
			errid=Err.number:errdes=Err.description:Err.Clear:dbConn.close:set dbConn=nothing:isConnect=false
			echoErr err_03,errid,errdes&"sql="&sqlStr
		end if
		end if 
	End Function

	Public Sub Class_Terminate()
		if isObject(dbRs) then set dbRs=nothing
		if isConnect then dbConn.close:set dbConn=nothing:isConnect=false
	End Sub
End Class

Class DataListClass
	Public primaryField,tableStr
	Public orderStr,whereStr,dataSortType,fieldsStr,dataPageSize,dataCurrentPage
	Public recordsCount,pagesCount
	private tempTableCount,sqlstr,topCount,whereStr2,whereStr3
	private m,n
	
	Public Sub Class_Initialize
		 dataSortType="desc"
	End Sub

	Public Sub Class_Terminate

	End Sub

	Public Function getDataList()
		dim order
		if isNul(dataPageSize) then dataPageSize=100 else dataPageSize=clng(dataPageSize)
		if not isNul(whereStr) then whereStr= " where "&whereStr else whereStr=""
		if isNul(tableStr) then die err_08
		if isNul(fieldsStr) then fieldsStr=" * "  else  fieldsStr=" "&fieldsStr&" "
		if not isNul(orderStr) then order=" order by "&orderStr&" "&dataSortType  else order=" "
		sqlstr="select top "&dataPageSize&fieldsStr&" from "&tableStr&" "&whereStr&order
		getDataList=conn.db(sqlstr,"array")
	End Function

	Public Function getPageList()
		dim order
		if isNul(dataPageSize) then dataPageSize=30 else dataPageSize=clng(dataPageSize)
		if not isNul(whereStr) then whereStr2=" where "&whereStr : whereStr3=" and "&whereStr else whereStr2="":whereStr3=""
		recordsCount=conn.db("select count(*) from "&tableStr&whereStr2,"array")(0,0)
		m=recordsCount mod dataPageSize
		n=int(recordsCount / dataPageSize)
		if m=0 then pagesCount=n else pagesCount=n + 1  
		if isNul(primaryField) then die err_07
		if isNul(tableStr) then die err_08
		if isNul(orderStr) then orderStr=primaryField
		if isNul(fieldsStr) then fieldsStr=" * "  else  fieldsStr=" "&fieldsStr&" "
		if dataCurrentPage > pagesCount   then dataCurrentPage=pagesCount
		if isNul(dataCurrentPage)  then 
			dataCurrentPage=1 
		else 
			if dataCurrentPage <= 0 then dataCurrentPage=1 else dataCurrentPage=clng(dataCurrentPage)
		end if
		order=" "&orderStr&" "&dataSortType 
		if dataSortType="desc" then
			if dataCurrentPage=1 then 
				sqlstr="select top "&dataPageSize&fieldsStr&" from "&tableStr&" "&whereStr2&" order by "&order
			else
				sqlstr="select top "&dataPageSize&fieldsStr&" from "&tableStr&" where "&primaryField&"<(select min("&primaryField&") from (select top "&(dataCurrentPage - 1) * dataPageSize&" "&primaryField&" from "&tableStr &" "&whereStr2& " order by  "&order&") as temptable)  "& whereStr3& " order by "&order
			end if
		else dataSortType="asc" 
			if dataCurrentPage=1 then 
				sqlstr="select top "&dataPageSize&fieldsStr&" from "&tableStr&" "&whereStr2&" order by "&order
			else
				sqlstr="select top "&dataPageSize&fieldsStr&" from "&tableStr&" where "&primaryField&">(select max("&primaryField&") from (select top "&(dataCurrentPage - 1) * dataPageSize &" "&primaryField&" from "&tableStr &" "& whereStr& " order by  "&order&") as temptable)  "& whereStr3& " order by "&order
			end if
		end if
		getPageList=conn.db(sqlstr,"array")
	End Function
End Class

Class XmlClass
	Public xmlDocument,xmlPath,xmlDomObj,xmlstr
	Private xmlDomVer,xmlFileSavePath

	Public Sub Class_Initialize()
		xmlDomVer=getXmlDomVer()
		createXmlDomObj
	End Sub

	Public Sub Class_Terminate()
		If IsObject(xmlDomObj) Then Set xmlDomObj=Nothing
	End Sub

	Public Function getXmlDomVer()
		dim i,xmldomVersions,xmlDomVersion
		getXmlDomVer=false
		xmldomVersions=Array("Microsoft.2MLDOM","MSXML2.DOMDocument","MSXML2.DOMDocument.3.0","MSXML2.DOMDocument.4.0","MSXML2.DOMDocument.5.0")
		for i=0 to ubound(xmldomVersions)
			xmlDomVersion=xmldomVersions(i)
			if isInstallObj(xmlDomVersion) then getXmlDomVer=xmlDomVersion : Exit Function
		next
	End Function

	Private Sub createXmlDomObj
		set xmlDomObj=server.CreateObject(xmlDomVer)
		xmlDomObj.validateonparse=true 
		xmlDomObj.async=false 
	End Sub

	Public Function load(Byval xml,Byval xmlType)
		dim xmlUrl,xmlfilePath
		select case xmlType 
			case "xmlfile"
				xmlfilePath=server.mappath(xml)   
		 		xmlDomObj.load(xmlfilePath)   
			case "xmldocument"
				xmlUrl=xml
				xmlstr=getRemoteContent(xmlUrl,"text")
				If left(xmlstr, 5) <> "<?xml" then die err_xml else xmlDomObj.loadXML(xmlstr)
			case "transfer"
				xmlUrl=xml
				xmlstr=bytesToStr(getRemoteContent(xmlUrl,"body"),"gbk")
				If left(xmlstr, 5) <> "<?xml" then die err_xml else xmlDomObj.loadXML(xmlstr)
		end select
	End Function

	Public Function isExistNode(nodename)
        dim node
        isExistNode=True
        set node=xmlDomObj.getElementsByTagName(nodename)
        If node.Length=0 Then isExistNode=False : set node=nothing
    End Function

	Public Function getNodeValue(nodename, itemId)
		if isNul(itemId) then  itemId=0
		getNodeValue=xmlDomObj.getElementsByTagName(nodename).Item(itemId).Text
	End Function

	Public Function getNodeLen(nodename)
		getNodeLen=xmlDomObj.getElementsByTagName(nodename).Length
	End Function
	
	Public Function getNodes(nodename)
		Set getNodes=xmlDomObj.getElementsByTagName(nodename)
	End Function
	
	Public Function getNode(nodename, itemId)
		Set getNode=xmlDomObj.getElementsByTagName(nodename).Item(itemId)
	End Function

	Public Function getAttributes(nodeName, attrName, itemId)
	dim xmlAttributes, i
		if isNul(itemId) then  itemId=0
		err.clear:on error resume next
		getAttributes=xmlDomObj.getElementsByTagName(nodeName).Item(itemId).getAttributeNode(attrName).nodevalue
		if err then getAttributes="":err.clear
	End Function

	Public Function getAttributesByNode(node, attrName)
		err.clear:on error resume next
		getAttributesByNode=node.getAttributeNode(attrName).nodevalue
		if err then getAttributesByNode="":err.clear
	End Function

	Public Function setXmlNodeValue(Byval nodename, Byval itemId, Byval str,Byval savePath)
	dim node
		xmlFileSavePath=savePath
		Set node=xmlDomObj.getElementsByTagName(nodename).Item(itemId)
		node.childNodes(0).text=str
		xmlDomObj.save Server.MapPath(xmlFileSavePath)
		set node=nothing
	End Function
End Class

Class TemplateClass
	Public content,allPages,currentPage,currentType
	Private cacheName,labelRule,regExpObj,strDictionary
	'初始化类
	Public Sub Class_Initialize()
		set regExpObj= new RegExp
		regExpObj.ignoreCase=false
		regExpObj.Global=true
		set strDictionary=server.CreateObject(DICTIONARY_OBJ_NAME)
	End Sub
	
	Public Sub Class_Terminate()
		set regExpObj=nothing
		set strDictionary=nothing
	End Sub

	'加载文件
	Public Function load(Byval filePath)
		content=loadFile(filePath)
	End Function

	'解析头部和底部
	Public Function parseTopAndFoot()
		if isExistStr(content,"{aspcms:top}") then content=replaceStr(content,"{aspcms:top}",loadFile(sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/head.html"))
		if isExistStr(content,"{aspcms:foot}") then content=replaceStr(content,"{aspcms:foot}","<script type=""text/javascript"" src="""&sitePath&"/"&"inc/AspCms_Statistics.asp""></script>"&loadFile(sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/foot.html"))	
	End Function
	
	'解析辅助模板
	Public Function parseAuxiliaryTemplate()
		Dim labelRuleRuxiliaryTemplate,matchesRuxiliary,matchRuxiliary,srcTemplate
		labelRuleRuxiliaryTemplate = "\{aspcms:template([\s\S]*?)\}"
		regExpObj.Pattern = labelRuleRuxiliaryTemplate
		set matchesRuxiliary = regExpObj.Execute(content)
		for each matchRuxiliary in matchesRuxiliary			
			srcTemplate = parseArr(matchRuxiliary.SubMatches(0))("src")			
			content=replaceStr(content,"{aspcms:template src="&srcTemplate&"}",loadFile(sitePath&"/"&"templates/"&setting.defaultTemplate&"/"&setting.htmlFilePath&"/"&srcTemplate))
		next
		set matchesRuxiliary = nothing	
	End Function
	
	'解析全局标签
	Public Function parseGlobal()		
		content=replaceStr(content,"{aspcms:sitelogo}",setting.siteLogoUrl)
		content=replaceStr(content,"{aspcms:companyname}",setting.companyName)
		content=replaceStr(content,"{aspcms:additiontitle}",setting.additionTitle)
		content=replaceStr(content,"{aspcms:companyaddress}",setting.companyAddress)
		content=replaceStr(content,"{aspcms:companypostcode}",setting.companyPostCode)
		content=replaceStr(content,"{aspcms:companycontact}",setting.companyContact)
		content=replaceStr(content,"{aspcms:companyphone}",setting.companyPhone)
		content=replaceStr(content,"{aspcms:companymobile}",setting.companyMobile)
		content=replaceStr(content,"{aspcms:companyfax}",setting.companyFax)
		content=replaceStr(content,"{aspcms:companyemail}",setting.companyEmail)
		content=replaceStr(content,"{aspcms:companyicp}",setting.companyICP)
		content=replaceStr(content,"{aspcms:statisticalcode}",decodehtml(setting.statisticalCode))
		content=replaceStr(content,"{aspcms:username}",session("loginName"))
		content=replaceStr(content,"{aspcms:siteurl}",setting.siteUrl)
		content=replaceStr(content,"{aspcms:sitepath}",sitePath)

		content=replaceStr(content,"{aspcms:languagepath}",setting.languagepath)
		content=replaceStr(content,"{aspcms:defaulttemplate}",setting.defaultTemplate)
		content=replaceStr(content,"{aspcms:sitetitle}",setting.siteTitle)
		'content=replaceStr(content,"{aspcms:copyright}",decodeHtml(setting.copyRight) & "- Powered by <a href='http://www.aspcms.com' target='_blank'><font color='#5C5C5C'>AspCms2</font></a>")
		'去版权
		content=replaceStr(content,"{aspcms:copyright}",decodeHtml(setting.copyRight))
		content=replaceStr(content,"{aspcms:sitedesc}",decodeHtml(setting.siteDesc))
		content=replaceStr(content,"{aspcms:sitekeywords}",setting.siteKeyWords)
		content=replaceStr(content,"{aspcms:slide}",getslide)
		content=replaceStr(content,"{aspcms:slidea}",getslide)
		content=replaceStr(content,"{aspcms:slideb}",getslideb)
		content=replaceStr(content,"{aspcms:slidec}",getslidec)
		content=replaceStr(content,"{aspcms:slided}",getslided)
		content=replaceStr(content,"{aspcms:kf}",getkf)
		content=replaceStr(content,"{aspcms:onlineservice}",getonlineservice)	
		
		'
		content=replaceStr(content,"{aspcms:googlemapkey}",GoogleAPIKey)
		content=replaceStr(content,"{aspcms:googlemaplat}",GoogleMapLat)
		content=replaceStr(content,"{aspcms:googlemaplng}",GoogleMapLng)
		if session("GroupID")="1" then
			content=replaceStr(content,"{aspcms:userright}",0)
		end if
		if session("GroupID")="2"or session("GroupID")="" then
			content=replaceStr(content,"{aspcms:userright}",2)
		end if
		if session("GroupID")="3" then
			content=replaceStr(content,"{aspcms:userright}",1)
		end if
		
		
				
		content=replaceStr(content,"{visits:today}","<script type=""text/javascript"" src="""&sitePath&"/"&"inc/AspCms_aStatistics.asp?act=t""></script>")
		content=replaceStr(content,"{visits:yesterday}","<script type=""text/javascript"" src="""&sitePath&"/"&"inc/AspCms_aStatistics.asp?act=y""></script>")
		content=replaceStr(content,"{visits:month}","<script type=""text/javascript"" src="""&sitePath&"/"&"inc/AspCms_aStatistics.asp?act=m""></script>")
		content=replaceStr(content,"{visits:all}","<script type=""text/javascript"" src="""&sitePath&"/"&"inc/AspCms_aStatistics.asp?act=a""></script>")
		
	End Function
		
			
	'获取可用标签参数
	Public Function parseArr(Byval attr)
		dim attrStr,attrArray,attrDictionary,i,singleAttr,singleAttrKey,singleAttrValue
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

	Public Function regExpReplace(contentstr,patternstr,replacestr)
		regExpObj.Pattern=patternstr
		regExpReplace=regExpObj.replace(contentstr,replacestr)
	End Function
	
	'解析导航栏
	Public Function parseNavList(str)	
		if not isExistStr(content,"{aspcms:"&str&"navlist") then Exit Function
		dim match,matches,matchfield,matchesfield
		dim labelAttrLinklist,loopstrLinklist,loopstrLinklistNew,loopstrTotal
		dim vtype,vnum,whereStr,linkArray
		dim fieldName,fieldAttr,fieldNameAndAttr,fieldAttrLen
		dim i,labelRuleField
		dim m,namelen,deslen,m_des
		labelRule="{aspcms:"&str&"navlist([\s\S]*?)}([\s\S]*?){/aspcms:"&str&"navlist}"
		labelRuleField="\["&str&"navlist:([\s\S]+?)\]"
		regExpObj.Pattern=labelRule
		set matches=regExpObj.Execute(content)
		
		for each match in matches
			labelAttrLinklist=match.SubMatches(0)
			loopstrLinklist=match.SubMatches(1)
			vtype=parseArr(labelAttrLinklist)("type") 	
			if isnul(vtype) then vtype=0	
				
			linkArray=conn.Exec("select SortName,SortType,SortURL,sortID,(select count (*) from {prefix}Sort as a where a.ParentID=b.sortID) as subcount,SortFolder,SortFileName,GroupID,Exclusive,indeximage from {prefix}Sort as b  where LanguageID="&setting.languageID&" and SortStatus=1 and ParentID="&vtype&" order by SortOrder asc","arr")
			
			if not isarray(linkArray) then  vnum=-1  else vnum=ubound(linkArray,2)
			regExpObj.Pattern=labelRuleField
			set matchesfield=regExpObj.Execute(loopstrLinklist)
			loopstrTotal=""
			for i=0 to vnum
				loopstrLinklistNew=loopstrLinklist
				for each matchfield in matchesfield
					fieldNameAndAttr=regExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
					fieldNameAndAttr=trimOuter(fieldNameAndAttr)
					m=instr(fieldNameAndAttr,chr(32))
					if m > 0 then 
						fieldName=left(fieldNameAndAttr,m - 1)
						fieldAttr =	right(fieldNameAndAttr,len(fieldNameAndAttr) - m)
					else
						fieldName=fieldNameAndAttr
						fieldAttr =	""
					end if
					select case fieldName
						case "name"
							namelen=parseArr(fieldAttr)("len") 
							if isNul(namelen) then 
								loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(0,i))
							else 
								namelen=clng(namelen)
								loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,left(linkArray(0,i),namelen)&"..")
							end if
						case "link"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,getSortLink(linkArray(1,i),linkArray(3,i),linkArray(2,i),linkArray(5,i),linkArray(6,i),linkArray(7,i),linkArray(8,i)))						
						case "sortid"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(3,i))
						case "subcount"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(4,i))
						case "desc"						
							m_des=decodeHtml(linkArray(3,i)):deslen=parseArr(fieldAttr)("len")
							if isNul(deslen) then deslen=100
							if len(m_des) > clng(deslen) then  m_des=left(m_des,clng(deslen)-1)&".."
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,m_des)
						case "pic"
							if not isNul(linkArray(9,i)) then 
								if instr(linkArray(9,i),"http://")>0 then 
									loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(9,i))
								else
									loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(9,i))
								end if
							else	
								loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,sitePath&"/"&"Images/nopic.gif")		
							end if
						case "i"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,i+1)
						case "cursortid"
						If runMode = 0 and request.QueryString<>"" Then 
							dim m_SortAndID
							m_SortAndID=split(replaceStr(request.QueryString,FileExt,""),"_")
							if IsArray(m_SortAndID) then
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,m_SortAndID(0))
							End If 
						End If
						case "pic"
							if not isNul(linkArray(9,i)) then 
									if instr(linkArray(9,i),"http://")>0 then 
										loopstrLinklistNew = replace(loopstrLinklistNew,matchfield.value,linkArray(9,i))
									else
										loopstrLinklistNew = replace(loopstrLinklistNew,matchfield.value,linkArray(9,i))
									end if
								else
									loopstrLinklistNew = replace(loopstrLinklistNew,matchfield.value,sitePath&"/"&"Images/nopic.gif")			
							end if
					end select
				next
				loopstrTotal=loopstrTotal&loopstrLinklistNew
			next
			set matchesfield=nothing
			content=replaceStr(content,match.value,loopstrTotal)
			strDictionary.removeAll
		next
		set matches=nothing
		if str="" then str=0
		if instr(content,"{aspcms:"& str+1 &"navlist")>0 then parseNavList(str+1) else Exit Function
	End Function
	
	
	'解析子导航栏
	Public Function parseSubNavList()
	Dim sql
		if not isExistStr(content,"{aspcms:subnavlist") then Exit Function

		dim match,matches,matchfield,matchesfield
		dim labelAttrLinklist,loopstrLinklist,loopstrLinklistNew,loopstrTotal
		dim vtype,vnum,whereStr,linkArray,v_maxlevel
		dim fieldName,fieldAttr,fieldNameAndAttr,fieldAttrLen
		dim i,labelRuleField
		dim m,namelen,deslen,m_des
		
		dim HeadTemplateRule,sHeadTemplate
		dim FootTemplateRule,sFootTemplate
		dim ItemItemplateRule,sItemTemplate		
		
		dim minLevel:minLevel = 2 '定义输出最小层级，兼容作用
		Dim m_type
		
		labelRule="{aspcms:subnavlist([\s\S]*?)}([\s\S]*?){/aspcms:subnavlist}"
		
		HeadTemplateRule="\[subnavtemplate:head]([\s\S]*?)\[/subnavtemplate:head]"
		FootTemplateRule="\[subnavtemplate:foot]([\s\S]*?)\[/subnavtemplate:foot]"
		ItemItemplateRule="\[subnavtemplate:item]([\s\S]*?)\[/subnavtemplate:item]"
		
		'response.ContentType = "text/plain"
		'response.clear		
		regExpObj.Pattern=HeadTemplateRule
		Set Matches=regExpObj.Execute(content)
		For Each Match in Matches ' 遍历匹配集合。			
			'echo "匹配位置:" & Match.FirstIndex &vbCRLF
			'echo "匹配字符:" & Match.Value &vbCRLF
			'echo "匹配条件:" & HeadTemplateRule &vbCRLF			
			'echo "匹配值:"&match.SubMatches(0) &vbCRLF
			sHeadTemplate = match.SubMatches(0)
			Content=replaceStr(content,Match.Value,"")
		Next
		regExpObj.Pattern=FootTemplateRule
		Set Matches=regExpObj.Execute(content)
		For Each Match in Matches:sFootTemplate = match.SubMatches(0):Content=replaceStr(content,Match.Value,""):Next
		
		regExpObj.Pattern=ItemItemplateRule
		Set Matches=regExpObj.Execute(content)
		For Each Match in Matches
			sItemTemplate = match.SubMatches(0)		
			'Content=replaceStr(content,Match.Value,"")
		Next
		labelRuleField="\[subnavlist:([\s\S]+?)\]"
		regExpObj.Pattern=labelRule
		set matches=regExpObj.Execute(content)
		
		
		for each match in matches
			labelAttrLinklist=match.SubMatches(0)
			loopstrLinklist=match.SubMatches(1)
			'获取指定的pid
			vtype=parseArr(labelAttrLinklist)("type")
			v_maxlevel = parseArr(Match.SubMatches(0))("maxlevel")
			vnum = parseArr(Match.SubMatches(0))("count")
			if isnul(vnum) then vnum = 65000
			if isnul(vtype) then vtype=0				
			if isnul(v_maxlevel) or v_maxlevel < minLevel then v_maxlevel = minLevel
			
			loopstrTotal=""
			'*****************
			'echo loopstrLinklist
			loopstrTotal = parseSubNavListProc(vnum,v_maxlevel,sHeadTemplate,sFootTemplate,sItemTemplate,vtype)
			'*****************
			
			content=replaceStr(content,match.value,loopstrTotal)
			strDictionary.removeAll
		next
		set matches=nothing		
	End Function
	
	'调用节点
	Private Function parseSubNavListProc(vnum,maxlevel,tHead,tFoot,tItem,pid)
	dim sTranslate,i,labelRuleField
	dim mf,mfs
	dim sNameAndAttr,sName,sAttr
	dim m
	
	dim namelen
	dim sql,marr
	
		
	
	sql ="select SortName,SortType,SortURL,sortID,IcoImage,(select count (*) from {prefix}Sort as a where a.ParentID=b.sortID) as subcount,SortFolder,SortFileName,GroupID,Exclusive,parentid,sortlevel from {prefix}Sort as b  where LanguageID="&setting.languageID&" and SortStatus=1 and ParentID="&pid&" order by SortOrder asc" 
	marr=conn.Exec(sql,"arr")	

	if not isarray(marr) then exit function
	for i=0 to ubound(marr,2)			
		If CInt(marr(10,i)) = CInt(pid) Then '父ID匹配
		If CInt(marr(11,i)) <= CInt(maxlevel) Then '最大深度匹配
		If i <= vnum then ' 单个计数5个
			parseSubNavListProc=parseSubNavListProc&SubNavListAttTranslate(tHead,i,marr)				
			parseSubNavListProc=parseSubNavListProc&SubNavListAttTranslate(tItem,i,marr)
	
			sTranslate=parseSubNavListProc(vnum,maxlevel,tHead,tFoot,tItem,marr(3,i))
			parseSubNavListProc = parseSubNavListProc & sTranslate
	
			parseSubNavListProc=parseSubNavListProc&SubNavListAttTranslate(tFoot,i,marr)	
		end if
		end if
		end if
	next
				
	End Function
	'节点属性翻译
	Function SubNavListAttTranslate(tItem,i,marr)
	Dim labelRuleField
	Dim mfs,mf,m
	Dim sTranslate
	Dim sName,sAttr,sNameAndAttr
	Dim namelen
		labelRuleField="\[subnavlist:([\s\S]+?)\]"
	
		regExpObj.Pattern=labelRuleField
		'echo titem&"<hr>"
		set mfs=regExpObj.Execute(tItem)
				
		sTranslate=tItem
		for each mf in mfs
		sNameAndAttr=regExpReplace(mf.SubMatches(0),"[\s]+",chr(32))
		'echo sNameAndAttr & "<br>"
		sNameAndAttr=trimOuter(sNameAndAttr)
		m=instr(sNameAndAttr,chr(32))
		if m > 0 then 
			sName=left(sNameAndAttr,m - 1)
			sAttr =	right(sNameAndAttr,len(sNameAndAttr) - m)
		else
			sName=sNameAndAttr
			sAttr =	""
		end if
		select case sName
			case "name"
				namelen=parseArr(sAttr)("len") 
				if isNul(namelen) then 
				sTranslate=replaceStr(sTranslate,mf.value,marr(0,i))
				else 
				namelen=clng(namelen)
				sTranslate=replaceStr(sTranslate,mf.value,left(marr(0,i),namelen)&"..")
				end if
			case "link"
				sTranslate=replaceStr(sTranslate,mf.value,getSortLink(marr(1,i),marr(3,i),marr(2,i),marr(5,i),marr(6,i),marr(7,i),marr(8,i)))						
			case "sortid"
				sTranslate=replaceStr(sTranslate,mf.value,marr(3,i))
			case "subcount"
				sTranslate=replaceStr(sTranslate,mf.value,marr(4,i))
			case "desc"
				m_des=decodeHtml(marr(3,i)):deslen=parseArr(sAttr)("len")
				if isNul(deslen) then deslen=100
				if len(m_des) > clng(deslen) then  m_des=left(m_des,clng(deslen)-1)&".."
				sTranslate=replaceStr(sTranslate,mf.value,m_des)
			case "i"
				sTranslate=replaceStr(sTranslate,mf.value,i+1)							
			case "cursortid"
				If runMode = 0 Then 
				dim m_SortAndID
				m_SortAndID=split(replaceStr(request.QueryString,FileExt,""),"_")
				if IsArray(m_SortAndID) then
				sTranslate=replaceStr(sTranslate,mf.value,m_SortAndID(0))
				end if
				End If
			case "level"
				sTranslate=replaceStr(sTranslate,mf.value,marr(10,i))
			case "pid"
				sTranslate=replaceStr(sTranslate,mf.value,marr(9,i))
		end select
		next
		set mfs=nothing
		SubNavListAttTranslate = sTranslate
	End Function
	
	
	'解析RSS
	Public Function parseRssList(str)	
		if not isExistStr(content,"{aspcms:"&str&"rsslist") then Exit Function
		dim match,matches,matchfield,matchesfield
		dim labelAttrLinklist,loopstrLinklist,loopstrLinklistNew,loopstrTotal
		dim vtype,vnum,whereStr,linkArray
		dim fieldName,fieldAttr,fieldNameAndAttr,fieldAttrLen
		dim i,labelRuleField
		dim m,namelen,deslen,m_des
		labelRule="{aspcms:"&str&"rsslist([\s\S]*?)}([\s\S]*?){/aspcms:"&str&"rsslist}"
		labelRuleField="\["&str&"rsslist:([\s\S]+?)\]"
		regExpObj.Pattern=labelRule
		set matches=regExpObj.Execute(content)
		
		for each match in matches
			labelAttrLinklist=match.SubMatches(0)
			loopstrLinklist=match.SubMatches(1)
			vtype=parseArr(labelAttrLinklist)("type") 	
			if isnul(vtype) then vtype=0		
			linkArray=conn.Exec("select SortName,SortType,SortURL,sortID,IcoImage,(select count (*) from {prefix}Sort as a where a.ParentID=b.sortID) as subcount,SortFolder,SortFileName from {prefix}Sort as b  where LanguageID="&setting.languageID&" and SortStatus=1 and ParentID="&vtype&" order by SortOrder asc","arr")
			if not isarray(linkArray) then  vnum=-1  else vnum=ubound(linkArray,2)
			regExpObj.Pattern=labelRuleField
			set matchesfield=regExpObj.Execute(loopstrLinklist)
			loopstrTotal=""
			for i=0 to vnum
				loopstrLinklistNew=loopstrLinklist
				for each matchfield in matchesfield
					fieldNameAndAttr=regExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
					fieldNameAndAttr=trimOuter(fieldNameAndAttr)
					m=instr(fieldNameAndAttr,chr(32))
					if m > 0 then 
						fieldName=left(fieldNameAndAttr,m - 1)
						fieldAttr =	right(fieldNameAndAttr,len(fieldNameAndAttr) - m)
					else
						fieldName=fieldNameAndAttr
						fieldAttr =	""
					end if
					select case fieldName
						case "name"
							namelen=parseArr(fieldAttr)("len") 
							if isNul(namelen) then 
								loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(0,i))
							else 
								namelen=clng(namelen)
								loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,left(linkArray(0,i),namelen)&"..")
							end if
						case "link"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,sitePath&setting.LanguagePath&"rss/"&linkArray(3,i)&".xml")							
						case "sortid"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(3,i))
						case "subcount"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(4,i))
						case "desc"
							m_des=decodeHtml(linkArray(3,i)):deslen=parseArr(fieldAttr)("len")
							if isNul(deslen) then deslen=100
							if len(m_des) > clng(deslen) then  m_des=left(m_des,clng(deslen)-1)&".."
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,m_des)
						case "i"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,i+1)
					end select
				next
				loopstrTotal=loopstrTotal&loopstrLinklistNew
			next
			set matchesfield=nothing
			content=replaceStr(content,match.value,loopstrTotal)
			strDictionary.removeAll
		next
		set matches=nothing
		if instr(content,"{aspcms:subrsslist")>0 then  parseRssList("sub") else Exit Function
	End Function
	
	'获取导航栏链接
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
	
	
	'内容页链接链接
	Function getContentLink(Byval SortID,Byval Id,Byval SortFolder,Byval GroupID,Byval ContentFolder,Byval ContentFileName,Byval ContentTime,Byval pageFileName,Byval SortGroupID)		
		ContentFolder=replace(ContentFolder, "{sitepath}", sitePath)	
		ContentFileName=replace(ContentFileName, "{sortid}", sortID)	
		ContentFileName=replace(ContentFileName, "{id}", Id)	
		ContentFileName=replace(ContentFileName, "{y}", year(ContentTime))		
		ContentFileName=replace(ContentFileName, "{m}", month(ContentTime))		
		ContentFileName=replace(ContentFileName, "{d}", day(ContentTime))			
		
		dim linkStr,rsObj
		if isnul(GroupID) or isnull(GroupID) then GroupID=0
		if runMode=1 and not isnul(SortFolder) then 
			if GroupID>2 or SortGroupID>2 then 
				getContentLink=sitePath&setting.languagePath&"content/"&"?"&Id&fileExt
			elseif not isnul(PageFileName) then
				getContentLink=decodeHtml(ContentFolder&pageFileName&fileExt)			
			else
				getContentLink=decodeHtml(ContentFolder&ContentFileName&fileExt)				
			end if
		else
			getContentLink=sitePath&setting.languagePath&"content/"&"?"&Id&fileExt			
		end if
	End Function		
	
	'替换循环标签
	Public Function parseLoop(Byval str)
		dim sortArr,sortStr,sortI,labelRuleField,matches,match,labelStr,loopStr,labelArr,lnum,ltype,lsort,lorder,ltime,whereType,whereSort,orderStr,whereTime,sql,DateArray,matchesfield,loopstrTotal,i,sperStrs,spec,sperStr,aboutkey,title,lstar,contentlen,pagecontent
		labelRule = "{aspcms:"&str&"([\s\S]*?)}([\s\S]*?){/aspcms:"&str&"}"
		labelRuleField = "\["&str&":([\s\S]+?)\]"
		regExpObj.Pattern = labelRule
		set matches = regExpObj.Execute(content)
		for each match in matches
		
		    labelStr = match.SubMatches(0)
			'echo str & labelStr
			loopStr = match.SubMatches(1)
			set labelArr = parseArr(labelStr)
			lnum = labelArr("num") : ltype = labelArr("type") : lsort = labelArr("sort") : lorder = labelArr("order") : ltime = labelArr("time") : aboutkey = labelArr("tag") : lstar=labelArr("star")

			if isNul(ltype) then ltype="all"
			if ltype="all" then 			
				whereType=""
			end if
			if isNul(lnum) then lnum = 10  else lnum = cint(lnum)
			sortStr=""
			
			if isNul(lsort) then lsort="all"
			whereSort=""
			if lsort <> "all" then 		
				whereSort=" and a.SortID in ("&getSubSort(lsort, 1)&")"				
			end if
			
			if isnum(lstar) then				
				whereSort=whereSort&" and a.Star="&lstar
			end if
			
			if not isnul(aboutkey) then		
			
				aboutkey=getTagID(aboutkey)
				'die aboutkey
				if not isnul(aboutkey) then 
				
					aboutkey=replace(getTagID(aboutkey),"}{","%' or ContentTag like '%")
					aboutkey=replace(aboutkey,"{"," (ContentTag like '%")
					aboutkey=replace(aboutkey,"}","%')")
					aboutkey=replace(aboutkey,"%'","}%'")
					aboutkey=replace(aboutkey,"'%","'%{")
	
					whereSort=whereSort&" and "&aboutkey
				end if
			end if
			
			if isNul(lorder) then lorder = "time"
			select case lorder           
				case "id" : orderStr =" order by ContentID desc"
				case "visits" : orderStr =" order by Visits desc"
				case "time" : orderStr =" order by a.AddTime desc"
				case "order" : orderStr =" order by IsTop desc,isrecommend desc,ContentOrder,a.AddTime desc"				
				case "istop" : orderStr =" and IsTop order by ContentOrder,a.AddTime desc"
				case "isrecommend" : orderStr =" and isrecommend order by ContentOrder,a.AddTime desc" 
				case "isimagenews" : orderStr =" and IsImageNews order by ContentOrder,a.AddTime desc" 
				case "isheadline" : orderStr =" and IsHeadline order by ContentOrder,a.AddTime desc" 
				case "isfeatured" : orderStr =" and IsFeatured order by ContentOrder,a.AddTime desc" 
			end select
			
			select case ltime
				case "day" : whereTime=" and  DateDiff('d',a.AddTime,'"&now()&"')=0"
				case "week" : whereTime=" and  DateDiff('w',a.AddTime,'"&now()&"')=0"
				case "month" : whereTime=" and  DateDiff('m',a.AddTime,'"&now()&"')=0"
				case else : whereTime=""
			end select	
								
			set labelArr = nothing
			if str="content" or str="news" or str="product" or str="down" or str="pic" then			
				sperStrs =conn.Exec("select SpecCategory+'_'+SpecField from {prefix}SpecSet Order by SpecOrder Asc,SpecID", "arr")				
				if isarray(sperStrs) then
					sperStr=""
					for each spec in sperStrs
						sperStr = sperStr&","&spec						
					next
				end if
				sql="select top "&lnum&" ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.ImagePath,a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID,b.IcoImage,b.Exclusive"&sperStr&" from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&"  and a.SortID=b.SortID and ContentStatus=1 and TimeStatus=0 "&whereType&whereSort&whereTime&orderStr								
			elseif str="about" or str="type" then			
				sql="select SortType,SortID,SortURL,SortFolder,SortFileName,SortName,SortContent,GroupID,ImagePath,IcoImage,IndexImage,Exclusive, PageDesc from {prefix}Sort where SortStatus=1 and SortID="&lsort&""
			elseif 	str="userbuylist" then
					sql="select * from {prefix}AspCms_Order2 where userid="&TypeId&" order by orderno desc"
			elseif str="gbook" then	
				if SwitchFaqStatus=0 then 			
					sql="select FaqID,FaqTitle,Contact,ContactWay,Content,Reply,AddTime,ReplyTime,FaqStatus,AuditStatus from {prefix}GuestBook order by AddTime"
				else
					sql="select FaqID,FaqTitle,Contact,ContactWay,Content,Reply,AddTime,ReplyTime,FaqStatus,AuditStatus from {prefix}GuestBook where FaqStatus order by AddTime"
				end if
			elseif str="comment" then	
				if switchCommentsStatus=0 then 			
					sql="select CommentsID,contentID,Commentator,CommentContent,AddTime,CommentIP from {prefix}Comments order by AddTime"
				else
					sql="select CommentsID,contentID,Commentator,CommentContent,AddTime,CommentIP from {prefix}Comments where CommentStatus order by AddTime"
				end if
			'elseif str="tag" then			
			'		sql="select top "&lnum&" NewsTag from {prefix}Content where  NOT isNULL(NewsTag) and ContentStatus=1 and TimeStatus=0 "&whereType&whereSort&whereTime&orderStr	
			elseif str="aboutart" then				
				dim ltypestr: ltypestr=""
					if not isnul(ltype) and not ltype="all" then ltypestr=" and sortType="&ltype	
					dim aboutkeystr,aboutkeys,ak
					
					if Instr(aboutkey,",") > 0 then
						aboutkey = Split(aboutkey,",")
						aboutkeystr = aboutkeystr &"("
						For i = 0 to Ubound(aboutkey)
							aboutkeystr = aboutkeystr &" ContentTag like '%"& aboutkey(i) &"%'"
							if i = Ubound(aboutkey) then
								aboutkeystr = aboutkeystr &") "
							else
								aboutkeystr = aboutkeystr &" Or "
							end if
						Next
					else
						aboutkeystr = aboutkeystr &" ContentTag like '%"& aboutkey &"%' "
					end if					
					 	
					sql="select top "&lnum&" ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID,b.Exclusive,b.IcoImage"&sperStr&" from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&"and a.SortID=b.SortID and ContentStatus=1 and TimeStatus=0 "&ltypestr&" and "&aboutkeystr&whereType&whereSort&whereTime&orderStr
					
	'die sql
			end if
			conn.fetchCount=lnum
			DateArray = conn.Exec(sql,"arr")
			dim rsObj
			set rsObj = conn.Exec(sql,"r1")
			
			conn.fetchCount=0
			regExpObj.Pattern = labelRuleField
			set matchesfield = regExpObj.Execute(loopStr)
			loopstrTotal = ""
			if isArray(DateArray) then lnum = ubound(DateArray,2) else lnum=-1
			dim nloopstr,matchfield,fieldNameArr,m,fieldName,fieldArr,infolen,namelen,timestyle, desclen
			
			for i = 0 to lnum
			    nloopstr=loopStr
			    for each matchfield in matchesfield
					fieldNameArr = regExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
					fieldNameArr = trim(fieldNameArr)
					m = instr(fieldNameArr,chr(32))
					if  m > 0 then 
						fieldArr=""
						fieldName = left(fieldNameArr,m - 1)
						fieldArr =	right(fieldNameArr,len(fieldNameArr) - m)
					else
						fieldName = fieldNameArr
						fieldArr =	""
					end if
					
					
					if str="content" or str="aboutcontent" or str="news" or str="product" or str="down" or str="pic"  then							
						if isarray(sperStrs) then
							for each spec in sperStrs			
								nloopstr = replace(nloopstr,"["&str&":"&spec&"]",repnull(rsObj(spec)))								
							next
						end if
					
						select case fieldName
							case "id"
								nloopstr = replace(nloopstr,matchfield.value,DateArray(0,i))	
							case "i"
								nloopstr = replace(nloopstr,matchfield.value,i+1)
							case "isoutlink"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("isoutlink"))									
							case "link"	
								'跳转链接						
								if DateArray(7,i)=1 then 
									nloopstr = replace(nloopstr,matchfield.value,DateArray(8,i))								
								else									
									nloopstr = replace(nloopstr,matchfield.value,getContentLink(rsObj("SortID"),rsObj("ContentID"),rsObj("SortFolder"),rsObj("a.GroupID"),rsObj("ContentFolder"),rsObj("ContentFileName"),rsObj("AddTime"),rsobj("PageFileName"),rsObj("b.GroupID")))
								end if							    
							case "title"
								namelen = parseArr(fieldArr)("len")
								title=DateArray(4,i)
								if not isNul(fieldArr) then   								
									namelen=cint(namelen)
									if len(DateArray(4,i))>namelen then title=left(DateArray(4,i),namelen)&"..." 										
								end if	
									nloopstr = replace(nloopstr,matchfield.value,title)	
							case "content"
								contentlen = parseArr(fieldArr)("len") 
								pagecontent=rsObj("content")
								if not isNul(contentlen) then   								
									contentlen=cint(contentlen)
									if len(pagecontent)>contentlen then pagecontent=left(pagecontent,contentlen)&"..." 
								end if	
								nloopstr = replace(nloopstr,matchfield.value,pagecontent)		
							case "titlecolor"
								nloopstr = replace(nloopstr,matchfield.value,DateArray(6,i))
							case "sortname"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("SortName"))
							case "sortlink"
								nloopstr = replace(nloopstr,matchfield.value,getSortLink(rsObj("sortType"),rsObj("sortID"),rsObj("sortUrl"),rsObj("sortFolder"),rsObj("sortFileName"),rsObj("b.GroupID"),rsObj("b.Exclusive")))
							case "date"
								timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
								select case timestyle
									case "yy-m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj("AddTime"),1))
									case "y-m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj("AddTime"),2))
									case "m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj("AddTime"),3))
								end select
							case "visits"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("Visits"))									
							case "author"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("Author"))
								
							case "source"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("ContentSource"))
							case "videourl"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("ContentSource"))
							case "downurl"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("DownURL"))								
							case "tag"
								nloopstr = replace(nloopstr,matchfield.value,getTags(rsObj("ContentTag")))								
							case "istop" '置顶
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsTop"))								
							case "isrecommend" '推荐
								nloopstr = replace(nloopstr,matchfield.value,rsObj("Isrecommend"))										
							case "isimage" '图片新闻
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsImageNews"))							
							case "isfeatured" '特别推荐
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsFeatured"))							
							case "isheadline" '头条
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsHeadline"))								
							case "desc"
								if not isnul(rsObj("PageDesc")) then
									infolen = parseArr(fieldArr)("len") : if isNul(infolen) then infolen = 200 else infolen=cint(infolen)
									nloopstr = replace(nloopstr,matchfield.value,left(filterStr(decodeHtml(replace(rsObj("PageDesc"),"{aspcms:page}","")),"html"),infolen))
								else
									infolen = parseArr(fieldArr)("len") : if isNul(infolen) then infolen = 200 else infolen=cint(infolen)
									nloopstr = replace(nloopstr,matchfield.value,left(filterStr(decodeHtml(replace(rsObj("content"),"{aspcms:page}","")),"html"),infolen))
								end if

							case "pic"
								if not isNul(rsObj("IndexImage")) then 
									if instr(rsObj("IndexImage"),"http://")>0 then 
										nloopstr = replace(nloopstr,matchfield.value,rsObj("IndexImage"))
									else
										nloopstr = replace(nloopstr,matchfield.value,rsObj("IndexImage"))
									end if
								else
									nloopstr = replace(nloopstr,matchfield.value,sitePath&"/"&"Images/nopic.gif")			
								end if
						end select
					elseif str="type" then
						select case fieldName
							case "i"
								nloopstr = replace(nloopstr,matchfield.value,i+1)
							case "link"
								nloopstr = replace(nloopstr,matchfield.value,getSortLink(DateArray(0,i),DateArray(1,i),DateArray(2,i),DateArray(3,i),DateArray(4,i),DateArray(7,i),DateArray(8,i)))	
							case "name"
								namelen = parseArr(fieldArr)("len")
								title=DateArray(5,i)
								if not isNul(namelen) then   								
									namelen=cint(namelen)
									if len(title)>namelen then title=left(title,namelen)&"..." 
								end if	
								nloopstr = replace(nloopstr,matchfield.value,title)					
							case "info"		
								infolen = parseArr(fieldArr)("len") 
								if isNul(infolen) then 
									nloopstr = replace(nloopstr,matchfield.value,replace(decodeHtml(DateArray(6,i)),"{aspcms:page}",""))	
								else
								 	infolen=cint(infolen)
									if len(decodeHtml(DateArray(6,i)))>infolen then 
										nloopstr = replace(nloopstr,matchfield.value,left(replace(decodeHtml(DateArray(6,i)),"{aspcms:page}",""),infolen)&"…")	
									else
										nloopstr = replace(nloopstr,matchfield.value,left(replace(decodeHtml(DateArray(6,i)),"{aspcms:page}",""),infolen))												
									end if
								end if								
							case "title"
								namelen = parseArr(fieldArr)("len")
								title=DateArray(5,i)
								if not isNul(fieldArr) then   								
									namelen=cint(namelen)
									if len(title)>namelen then title=left(title,namelen)&"..." 
								end if	
								nloopstr = replace(nloopstr,matchfield.value,title)
							case "ico"
								if not isNul(rsObj("IcoImage")) then 
									if instr(rsObj("IcoImage"),"http://")>0 then 
										nloopstr = replace(nloopstr,matchfield.value,rsObj("IcoImage"))
									else
										nloopstr = replace(nloopstr,matchfield.value,rsObj("IcoImage"))
									end if
								else
									nloopstr = replace(nloopstr,matchfield.value,sitePath&"/"&"Images/nopic.gif")			
								end if	
						end select
					elseif str="about" then
						select case fieldName
							case "i"
								nloopstr = replace(nloopstr,matchfield.value,i+1)
							case "link"
								nloopstr = replace(nloopstr,matchfield.value,getSortLink(DateArray(0,i),DateArray(1,i),DateArray(2,i),DateArray(3,i),DateArray(4,i),DateArray(7,i),DateArray(8,i)))	
							case "name"
								namelen = parseArr(fieldArr)("len")
								title=DateArray(5,i)
								if not isNul(namelen) then   								
									namelen=cint(namelen)
									if len(title)>namelen then title=left(title,namelen)&"..." 
								end if	
								nloopstr = replace(nloopstr,matchfield.value,title)					
							case "info"		
								infolen = parseArr(fieldArr)("len") 
								if isNul(infolen) then 
									nloopstr = replace(nloopstr,matchfield.value,replace(decodeHtml(DateArray(6,i)),"{aspcms:page}",""))	
								else
								 	infolen=cint(infolen)
									if len(decodeHtml(DateArray(6,i)))>infolen then 
										nloopstr = replace(nloopstr,matchfield.value,left(replace(decodeHtml(DateArray(6,i)),"{aspcms:page}",""),infolen)&"…")	
									else
										nloopstr = replace(nloopstr,matchfield.value,left(replace(decodeHtml(DateArray(6,i)),"{aspcms:page}",""),infolen))												
									end if
								end if								
							case "title"
								namelen = parseArr(fieldArr)("len")
								title=DateArray(5,i)
								if not isNul(fieldArr) then   								
									namelen=cint(namelen)
									if len(title)>namelen then title=left(title,namelen)&"..." 
								end if	
								nloopstr = replace(nloopstr,matchfield.value,title)
							case "pic"
								if not isNul(rsObj("IndexImage")) then 
									if instr(rsObj("IndexImage"),"http://")>0 then 
										nloopstr = replace(nloopstr,matchfield.value,rsObj("IndexImage"))
									else
										nloopstr = replace(nloopstr,matchfield.value,rsObj("IndexImage"))
									end if
								else
									nloopstr = replace(nloopstr,matchfield.value,sitePath&"/"&"Images/nopic.gif")			
								end if
							case "desc"
								desclen = parseArr(fieldArr)("len") 
								if isNul(infolen) then 
									nloopstr = replace(nloopstr,matchfield.value,replace(decodeHtml(DateArray(12,i)),"{aspcms:page}",""))	
								else
								 	infolen=cint(infolen)
									if len(decodeHtml(DateArray(12,i)))>infolen then 
										nloopstr = replace(nloopstr,matchfield.value,left(replace(decodeHtml(DateArray(12,i)),"{aspcms:page}",""),desclen)&"…")	
									else
										nloopstr = replace(nloopstr,matchfield.value,left(replace(decodeHtml(DateArray(12,i)),"{aspcms:page}",""),desclen))												
									end if
								end if
						end select
					elseif str="tag" then
						select case fieldName			
							case "tag"
								Dim tagStrs,tagStr,tags
								tagStrs=split(replace(replace(DateArray(0,i)," ",","),"，",","),",")
								tags=""
								for each tagStr in tagStrs
									tags=tags&"<a href="&sitePath&"/"&"tag.asp?key="&tagStr&"&searchType=-1"">"&tagStr&"</a> "
								next
								nloopstr = replace(nloopstr,matchfield.value,tags)
						end select
					elseif str="gbook" then
						select case fieldName		
							case "i"
								nloopstr = replace(nloopstr,matchfield.value,i+1)
							case "link"
								'if rsObj(5)=1 then nloopstr = replace(nloopstr,matchfield.value,rsObj(9)) : else nloopstr = replace(nloopstr,matchfield.value,getContentLink(DateArray(0,i),DateArray(0,i),showType))
							case "title"
								namelen = parseArr(fieldArr)("len") 
								title=filterDirty(DateArray(1,i))
								if not isNul(fieldArr) then   								
									namelen=cint(namelen)
									if len(title)>namelen then title=left(title,namelen)&"..." 
								end if	
								nloopstr = replace(nloopstr,matchfield.value,title)	
							case "name"									
								nloopstr = replace(nloopstr,matchfield.value,repNull(DateArray(2,i)))
							case "status"									
								nloopstr = replace(nloopstr,matchfield.value,DateArray(8,i))							
							case "winfo"
								nloopstr = replace(nloopstr,matchfield.value,filterDirty(repNull(DateArray(4,i))))
							case "rinfo"
								nloopstr = replace(nloopstr,matchfield.value,repNull(DateArray(5,i)))
							case "wdate"
								timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
								 select case timestyle
									case "yy-m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(6,i),1))
									case "y-m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(6,i),2))
									case "m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(6,i),3))
								end select	
							case "rdate"
								timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
								 select case timestyle
									case "yy-m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(7,i),1))
									case "y-m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(7,i),2))
									case "m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(7,i),3))
								end select	
						end select	
					elseif str="comment" then
						select case fieldName		
								case "i"
									nloopstr = replace(nloopstr,matchfield.value,i+1)
								case "link"
									dim commentsql,commentrsObj
									commentsql="select top 1 ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID,b.Exclusive from {prefix}Content as a,{prefix}Sort as b where ContentID="&DateArray(1,i)
									'die commentsql
									
									set commentrsObj = conn.Exec(commentsql,"r1")
									nloopstr = replace(nloopstr,matchfield.value,getContentLink(commentrsObj("SortID"),commentrsObj("ContentID"),commentrsObj("SortFolder"),commentrsObj("a.GroupID"),commentrsObj("ContentFolder"),commentrsObj("ContentFileName"),commentrsObj("AddTime"),commentrsObj("PageFileName"),commentrsObj("b.GroupID")))
								case "name"									
									nloopstr = replace(nloopstr,matchfield.value,repNull(DateArray(2,i)))
								case "info"
									infolen = parseArr(fieldArr)("len") 
									if isNul(infolen) then 
										nloopstr = replace(nloopstr,matchfield.value,replace(decodeHtml(DateArray(3,i)),"{aspcms:page}",""))	
									else
										infolen=cint(infolen)
										if len(decodeHtml(DateArray(3,i)))>infolen then 
											nloopstr = replace(nloopstr,matchfield.value,left(replace(decodeHtml(DateArray(3,i)),"{aspcms:page}",""),infolen))&"…"		
										else
											nloopstr = replace(nloopstr,matchfield.value,left(replace(decodeHtml(DateArray(3,i)),"{aspcms:page}",""),infolen))												
										end if
									end if	
								case "ip"									
									nloopstr = replace(nloopstr,matchfield.value,repNull(DateArray(5,i)))
								case "date"
									timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
									 select case timestyle
										case "yy-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(4,i),1))
										case "y-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(4,i),2))
										case "m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(DateArray(4,i),3))
									end select	
							end select		
									
					end if	
				next
				loopstrTotal = loopstrTotal & nloopstr
			rsObj.movenext
			next
			set matchesfield = nothing
			content = replace(content,match.value,loopstrTotal)
			strDictionary.removeAll
		next
		
		set matches = nothing
	End Function
	

	
	'替换List循环标签
	Public Function parseList(typeIds,currentPage,pageListType,keys,showType)	
	    dim lenPagelist,TypeId,strPagelist,lsize,rsObj,labelRuleField,labelRulePagelist,matches,match,labelStr,loopStr,labelArr,lorder,orderStr,sql,matchesfield,sperStrs,spec,sperStr,title,aboutkey
		labelRule = "{aspcms:"&pageListType&"([\s\S]*?)}([\s\S]*?){/aspcms:"&pageListType&"}"
		labelRuleField = "\["&pageListType&":([\s\S]+?)\]"
		labelRulePagelist = "\["&pageListType&":pagenumber([\s\S]*?)\]"
		regExpObj.Pattern = labelRule
		set matches = regExpObj.Execute(content)
		


		for each match in matches
			
		    labelStr = match.SubMatches(0)
			loopStr = match.SubMatches(1)
			set labelArr = parseArr(labelStr)
			lsize = cint(labelArr("size")) : lorder = labelArr("order") : aboutkey = labelArr("tag") :
			if isNul(lsize) then lsize = 12 
			if isNul(lorder) then lorder = "time"
			select case lorder
				case "id" : orderStr =" order by ContentID desc"
				case "visits" : orderStr =" order by Visits desc"
				case "time" : orderStr =" order by a.AddTime desc"
				case "order" : orderStr =" order by IsHeadline,IsTop,IsFeatured,isrecommend,ContentOrder,a.AddTime desc"				
				case "istop" : orderStr =" and IsTop order by ContentOrder,a.AddTime desc"
				case "isrecommend" : orderStr =" and isrecommend order by ContentOrder,a.AddTime desc" 
				case "isimagenews" : orderStr =" and IsImageNews order by ContentOrder,a.AddTime desc" 
				case "isheadline" : orderStr =" and IsHeadline order by ContentOrder,a.AddTime desc" 
				case "isfeatured" : orderStr =" and IsFeatured order by ContentOrder,a.AddTime desc" 
			end select
			
			set labelArr = nothing
			
			if pageListType="list" or pageListType="newslist" or pageListType="productlist" or pageListType="downlist" or pageListType="piclist"  or pageListType="searchlist" then				
				
				sperStrs =conn.Exec("select SpecCategory+'_'+SpecField from {prefix}SpecSet Order by SpecOrder Asc,SpecID", "arr")				
				if isarray(sperStrs) then
					sperStr=""
					for each spec in sperStrs
						sperStr = sperStr&","&spec
					next
				end if
			
				if isNul(keys) then					
					if not isnul(aboutkey) then		
						aboutkey=getTagID(aboutkey)
						if not isnul(aboutkey) then 		
							aboutkey=replace(aboutkey,"}{","%' or ContentTag like '%")
							aboutkey=replace(aboutkey,"{"," (ContentTag like '%")
							aboutkey=replace(aboutkey,"}","%')")
							aboutkey=replace(aboutkey,"%'","}%'")
							aboutkey=replace(aboutkey,"'%","'%{") 
							sql="select ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID,b.Exclusive,b.GroupID "&sperStr&" from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&" and a.SortID=b.SortID and ContentStatus=1 and TimeStatus=0  and "&aboutkey&orderStr
							'die sql
						end if						
					else
						sql="select ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID,b.Exclusive,b.GroupID "&sperStr&" from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&"and a.SortID=b.SortID and ContentStatus=1 and TimeStatus=0 and a.SortID in ("&getSubSort(typeIds, 1)&")"&orderStr
					end if
				else		
					dim typeStr: typeStr=""
					dim searchType
					searchType=filterPara(getForm("searchType","get"))
					if isnul(searchType) then searchType="0"
					if  not "0"=searchType  then typeStr=" and a.SortID in (select {prefix}Sort.sortid from {prefix}Sort where sortType="&searchType&") " 	
					sql="select ContentID,a.SortID,a.GroupID,a.Exclusive,Title,Title2,TitleColor,IsOutLink,OutLink,Author,ContentSource,ContentTag,Content,ContentStatus,IsTop,Isrecommend,IsImageNews,IsHeadline,IsFeatured,ContentOrder,IsGenerated,Visits,a.AddTime,a.[ImagePath],a.IndexImage,a.DownURL,a.PageFileName,a.PageDesc,SortType,SortURL,SortFolder,SortFileName,SortName,ContentFolder,ContentFileName,b.GroupID "&sperStr&" from {prefix}Content as a,{prefix}Sort as b where a.LanguageID="&setting.languageID&"and a.SortID=b.SortID and ContentStatus=1 and TimeStatus=0 and a.SortID in ("&getSubSort(typeIds, 1)&") and Title like '%"&keys&"%'"&typeStr&orderStr
				end if
			elseif 	pageListType="gbooklist" then
				select case lorder           
					case "id" : orderStr =" order by FaqID desc"
					case "time" : orderStr =" order by AddTime desc"
				end select
				
				if SwitchFaqStatus=0 then 
					sql="select FaqID,FaqTitle,Contact,ContactWay,Content,Reply,AddTime,ReplyTime,FaqStatus,AuditStatus from {prefix}GuestBook where LanguageID="&setting.languageID&""&orderStr
				else
					sql="select FaqID,FaqTitle,Contact,ContactWay,Content,Reply,AddTime,ReplyTime,FaqStatus,AuditStatus from {prefix}GuestBook where LanguageID="&setting.languageID&"and FaqStatus=1 "&orderStr					
				end if
			elseif 	pageListType="userbuylist" then
					sql="select * from {prefix}Order2 where userid="&typeIds&" order by orderno desc"
			elseif 	pageListType="taglist" then
				select case lorder           
					case "id" : orderStr =" order by TagID desc"
					case "time" : orderStr =" order by AddTime desc"
					case "visits" : orderStr =" order by Tagvisits desc"
				end select		
				sql="select TagID, TagName, TagCount, SortType, SortID, TagVisits, LanguageID, AddTime from {prefix}Tag where LanguageID="&setting.languageID&orderStr	
			end if	
			'die sql	
			regExpObj.Pattern = labelRuleField
			set matchesfield = regExpObj.Execute(loopStr)
			'die sql
			
			set rsObj=conn.Exec(sql,"r1")
			
			
			Dim loopstrTotal,i,nloopstr,matchfield,fieldNameArr,m,fieldName,fieldArr,namelen,infolen,timestyle,matchesPagelist,matchPagelist,contentlen,pagecontent
			if rsObj.eof then 
			    if isNul(keys) then
				    if pageListType="gbooklist" then loopstrTotal=str_10 else loopstrTotal=str_08
				else
				    loopstrTotal=str_09&"<font color='red'>"&keys&"</font>"&str_10
				end if
			else
				rsObj.pagesize = lsize			
			
				if cint(currentPage)>rsObj.pagecount then currentPage=rsObj.pagecount
				'die currentPage
				rsObj.absolutepage=currentPage
				loopstrTotal = ""
				for i = 1 to lsize
		
				
					nloopstr=loopStr
					for each matchfield in matchesfield
						fieldNameArr = regExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
						fieldNameArr = trim(fieldNameArr)
						m = instr(fieldNameArr,chr(32))
						if  m > 0 then 
							fieldName = left(fieldNameArr,m - 1)
							fieldArr =	right(fieldNameArr,len(fieldNameArr) - m)
						else
							fieldName = fieldNameArr
							fieldArr =	""
						end if
						if pageListType="list" or pageListType="newslist" or pageListType="productlist" or pageListType="downlist" or pageListType="piclist"  or pageListType="searchlist" then
		
							if isarray(sperStrs) then
								for each spec in sperStrs			
									nloopstr = replace(nloopstr,"[list:"&spec&"]",repnull(rsObj(spec)))
								next
							end if
					
						select case fieldName
							case "id"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("ContentID"))
							case "i"
								nloopstr = replace(nloopstr,matchfield.value,i)								
							case "link"	
								'跳转链接						
								if rsObj("IsOutLink")=1 then 
									nloopstr = replace(nloopstr,matchfield.value,rsObj("OutLink"))								
								else									
									nloopstr = replace(nloopstr,matchfield.value,getContentLink(rsObj("SortID"),rsObj("ContentID"),rsObj("SortFolder"),rsObj("a.GroupID"),rsObj("ContentFolder"),rsObj("ContentFileName"),rsObj("AddTime"),rsobj("PageFileName"),rsObj("b.GroupID")))
								end if							    
							case "title"
								namelen = parseArr(fieldArr)("len") 
								title=rsObj("Title")
								if not isNul(fieldArr) then   								
									namelen=cint(namelen)
									if len(title)>namelen then title=left(title,namelen)&"..." 
								end if	
								nloopstr = replace(nloopstr,matchfield.value,title)		
							case "content"
								contentlen = parseArr(fieldArr)("len") 
								pagecontent=decodeHtml(rsObj("content"))
								if not isNul(contentlen) then   								
									contentlen=cint(contentlen)
									'if len(pagecontent)>contentlen then pagecontent=left(LeftH(pagecontent),contentlen)&"..." 
									if len(pagecontent)>contentlen then pagecontent=left(dropHtml(pagecontent),contentlen)&"..." 
								end if	
								nloopstr = replace(nloopstr,matchfield.value,pagecontent)		
							case "titlecolor"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("titlecolor"))
							case "sortname"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("SortName"))
							case "sortlink"
								nloopstr = replace(nloopstr,matchfield.value,getSortLink(rsObj("sortType"),rsObj("sortID"),rsObj("sortUrl"),rsObj("sortFolder"),rsObj("sortFileName"),rsObj("b.GroupID"),rsObj("b.Exclusive")))
							case "date"
								timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
								select case timestyle
									case "yy-m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj("AddTime"),1))
									case "y-m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj("AddTime"),2))
									case "m-d"
										nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj("AddTime"),3))
								end select
							case "visits"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("Visits"))									
							case "author"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("Author"))
							case "source"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("ContentSource"))
							case "videourl"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("ContentSource"))					
							case "tag"
								nloopstr = replace(nloopstr,matchfield.value,getTags(rsObj("ContentTag")))								
							case "isoutlink"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsOutLink"))	
							case "downurl"
								nloopstr = replace(nloopstr,matchfield.value,rsObj("DownURL"))										
							case "istop" '置顶
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsTop"))								
							case "isrecommend" '推荐
								nloopstr = replace(nloopstr,matchfield.value,rsObj("Isrecommend"))										
							case "isimage" '图片新闻
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsImageNews"))							
							case "isfeatured" '特别推荐
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsFeatured"))							
							case "isheadline" '头条
								nloopstr = replace(nloopstr,matchfield.value,rsObj("IsHeadline"))								
							case "desc"
								if not isnul(rsObj("PageDesc")) then
									nloopstr = replace(nloopstr,matchfield.value,rsObj("PageDesc"))
								else
									infolen = parseArr(fieldArr)("len") : if isNul(infolen) then infolen = 200 else infolen=cint(infolen)
									nloopstr = replace(nloopstr,matchfield.value,left(filterStr(decodeHtml(rsObj("Content")),"html"),infolen))
								end if
							case "pic"
								if not isNul(rsObj("IndexImage")) then 
									if instr(rsObj("IndexImage"),"http://")>0 then 
										nloopstr = replace(nloopstr,matchfield.value,rsObj("IndexImage"))
									else
										nloopstr = replace(nloopstr,matchfield.value,rsObj("IndexImage"))
									end if
								else
									nloopstr = replace(nloopstr,matchfield.value,sitePath&"/"&"Images/nopic.gif")			
								end if
							end select	
						elseif pageListType="gbooklist" then	
							select case fieldName
								case "i"
									nloopstr = replace(nloopstr,matchfield.value,i)
								case "link"
								case "title"
									namelen = parseArr(fieldArr)("len") 
									title=filterDirty(rsObj(1))
									if not isNul(fieldArr) then   								
										namelen=cint(namelen)
										if len(title)>namelen then title=left(title,namelen)&"..." 
									end if	
									nloopstr = replace(nloopstr,matchfield.value,title)																		
									'if len(rsObj(1))>namelen then title=left(rsObj(1),namelen)&"..." else title=rsObj(1)
									'nloopstr = replace(nloopstr,matchfield.value,title)
								case "name"									
									nloopstr = replace(nloopstr,matchfield.value,filterDirty(repNull(rsObj(2))))
								case "status"									
									nloopstr = replace(nloopstr,matchfield.value,rsObj(9))							
								case "winfo"
									nloopstr = replace(nloopstr,matchfield.value,filterDirty(repNull(rsObj(4))))
								case "rinfo"
									nloopstr = replace(nloopstr,matchfield.value,repNull(rsObj(5)))
								case "wdate"
									timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
									 select case timestyle
										case "yy-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(6),1))
										case "y-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(6),2))
										case "m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(6),3))
									end select	
								case "rdate"
									timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
									 select case timestyle
										case "yy-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(7),1))
										case "y-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(7),2))
										case "m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(7),3))
									end select	
							end select
						elseif pageListType="userbuylist" then
						
							select case fieldName
								case "i"
									nloopstr = replace(nloopstr,matchfield.value,i)
								case "OrderNo"
									nloopstr = replace(nloopstr,matchfield.value,filterDirty(repNull(rsObj(0))))
								case "OrderLink"
									nloopstr = replace(nloopstr,matchfield.value,"orderinfo.asp?orderno="&filterDirty(repNull(rsObj(0))))
								case "OrderTime"									
												timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
									 select case timestyle
										case "yy-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(2),1))
										case "y-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(2),2))
										case "m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(2),3))
										end select	
								case "State"
									dim statecn
									if filterDirty(repNull(rsObj(3)))=0 then
										statecn="未支付"
									elseif filterDirty(repNull(rsObj(3)))=1 then
										statecn="已支付"
									end if
									nloopstr = replace(nloopstr,matchfield.value,filterDirty(repNull(statecn)))					
								case "Payment"
									nloopstr = replace(nloopstr,matchfield.value,filterDirty(repNull(rsObj(4))))
							end select
						elseif pageListType="taglist" then	
							select case fieldName
								case "i"
									nloopstr = replace(nloopstr,matchfield.value,i)
								case "link"									
									nloopstr = replace(nloopstr,matchfield.value,sitePath&setting.languagePath&"taglist.asp?tag="&rsObj(1))	
								case "title","name"	
									namelen = parseArr(fieldArr)("len") 
									title=filterDirty(rsObj(1))
									if not isNul(namelen) then   								
										namelen=cint(namelen)
										if len(title)>namelen then title=left(title,namelen)&"..." 
									end if	
									nloopstr = replace(nloopstr,matchfield.value,title)	
																	
									'if len(rsObj(1))>namelen then title=left(rsObj(1),namelen)&"..." else title=rsObj(1)
									'nloopstr = replace(nloopstr,matchfield.value,title)								
								case "date"
									timestyle = parseArr(fieldArr)("style") : if isNul(timestyle) then timestyle = "m-d"
									 select case timestyle
										case "yy-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(6),1))
										case "y-m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(6),2))
										case "m-d"
											nloopstr = replace(nloopstr,matchfield.value,FormatDate(rsObj(7),3))
									end select								
							end select
						end if
					next
					loopstrTotal = loopstrTotal & nloopstr
					rsObj.movenext
					if rsObj.eof then exit for
				next
			end if
			
			content = replace(content,match.value,loopstrTotal)
			regExpObj.Pattern = labelRulePagelist
			set matchesPagelist = regExpObj.Execute(content)
			for each matchPagelist in matchesPagelist
				if rsObj.pagecount=0 then
					content = replace(content,matchPagelist.value,"")
				else
					lenPagelist = parseArr(matchPagelist.SubMatches(0))("len")
					if isNul(lenPagelist) then lenPagelist = 10 else lenPagelist = cint(lenPagelist)
					if isExistStr(TypeIds,",") then TypeId=split(TypeIds,",")(0) : else TypeId=TypeIds
					
					strPagelist = pageNumberLinkInfo(currentPage,lenPagelist,rsObj.pagecount,pageListType,TypeId,showType)
					
					content = replace(content,matchPagelist.value,strPagelist)
				end if
			next			
			set matchesPagelist = nothing
			set matchesfield = nothing
			strDictionary.removeAll
		next
		set matches = nothing
	End Function

	

	Public Function parseLinkList()
		if not isExistStr(content,"{aspcms:linklist") then Exit Function
		dim match,matches,matchfield,matchesfield
		dim labelAttrLinklist,loopstrLinklist,loopstrLinklistNew,loopstrTotal
		dim vtype,vnum,whereStr,linkArray,lgroup
		dim fieldName,fieldAttr,fieldNameAndAttr,fieldAttrLen
		dim i,labelRuleField
		dim m,namelen,deslen,m_des
		labelRule="{aspcms:linklist([\s\S]*?)}([\s\S]*?){/aspcms:linklist}"
		labelRuleField="\[linklist:([\s\S]+?)\]"
		regExpObj.Pattern=labelRule
		set matches=regExpObj.Execute(content)
		for each match in matches
			labelAttrLinklist=match.SubMatches(0)
			loopstrLinklist=match.SubMatches(1)
			vtype=parseArr(labelAttrLinklist)("type")
			lgroup=parseArr(labelAttrLinklist)("group")
			if isNul(vtype) then vtype=0
			select case vtype
				case "font" : whereStr=chr(32)&"LinkType=0 and LinkStatus"&chr(32)
				case "pic" : whereStr=chr(32)&"LinkType=1 and LinkStatus"&chr(32)
				case else : whereStr=chr(32)&"LinkStatus"&chr(32)
			end select
			if not isNul(lgroup) then
				 whereStr=chr(32)&"LinkGroup="&lgroup&chr(32)
			end if
			linkArray=conn.Exec("select LinkText,ImageURL,LinkURL,LinkDesc from {prefix}Links  where "&whereStr&" order by LinkOrder asc","arr")
			if not isarray(linkArray) then  vnum=-1  else vnum=ubound(linkArray,2)
			regExpObj.Pattern=labelRuleField
			set matchesfield=regExpObj.Execute(loopstrLinklist)
			loopstrTotal=""
			for i=0 to vnum
				loopstrLinklistNew=loopstrLinklist
				for each matchfield in matchesfield
					fieldNameAndAttr=regExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
					fieldNameAndAttr=trimOuter(fieldNameAndAttr)
					m=instr(fieldNameAndAttr,chr(32))
					if m > 0 then 
						fieldName=left(fieldNameAndAttr,m - 1)
						fieldAttr =	right(fieldNameAndAttr,len(fieldNameAndAttr) - m)
					else
						fieldName=fieldNameAndAttr
						fieldAttr =	""
					end if
					select case fieldName
						case "name"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(0,i))
						case "link"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(2,i))
						case "pic"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(1,i))
						case "des"
							m_des=decodeHtml(linkArray(3,i)):deslen=parseArr(fieldAttr)("len")
							if isNul(deslen) then deslen=100
							if len(m_des) > clng(deslen) then  m_des=left(m_des,clng(deslen)-1)&".."
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,m_des)
						case "i"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,i+1)
					end select
				next
				loopstrTotal=loopstrTotal&loopstrLinklistNew
			next
			set matchesfield=nothing
			content=replaceStr(content,match.value,loopstrTotal)
			strDictionary.removeAll
		next
		set matches=nothing
	End Function

	'解析if
	Public Function parseIf()
		if not isExistStr(content,"{if:") then Exit Function		
		dim matchIf,matchesIf,strIf,strThen,strThen1,strElse1,labelRule2,labelRule3
		dim ifFlag,elseIfArray,elseIfSubArray,elseIfArrayLen,resultStr,elseIfLen,strElseIf,strElseIfThen,elseIfFlag
		labelRule="{if:([\s\S]+?)}([\s\S]*?){end\s+if}":labelRule2="{elseif":labelRule3="{else}":elseIfFlag=false
		regExpObj.Pattern=labelRule
		set matchesIf=regExpObj.Execute(content)
		for each matchIf in matchesIf 			
			strIf=matchIf.SubMatches(0):strThen=matchIf.SubMatches(1)
			'echo "<h1>"&strIf &"</h1>"			
			if instr(strThen,labelRule2)>0 then
				elseIfArray=split(strThen,labelRule2):elseIfArrayLen=ubound(elseIfArray):elseIfSubArray=split(elseIfArray(elseIfArrayLen),labelRule3)
				resultStr=elseIfSubArray(1)
				
				Execute("if "&strIf&" then resultStr=elseIfArray(0)")
				for elseIfLen=1 to elseIfArrayLen-1
					strElseIf=getSubStrByFromAndEnd(elseIfArray(elseIfLen),":","}","")
					strElseIfThen=getSubStrByFromAndEnd(elseIfArray(elseIfLen),"}","","start")
					Execute("if "&strElseIf&" then resultStr=strElseIfThen")
					Execute("if "&strElseIf&" then elseIfFlag=true else  elseIfFlag=false")
					if elseIfFlag then exit for
				next
				Execute("if "&getSubStrByFromAndEnd(elseIfSubArray(0),":","}","")&" then resultStr=getSubStrByFromAndEnd(elseIfSubArray(0),""}"","""",""start""):elseIfFlag=true")
				content=replace(content,matchIf.value,resultStr)
			else 
				if instr(strThen,"{else}")>0 then 
					strThen1=split(strThen,labelRule3)(0)
					strElse1=split(strThen,labelRule3)(1)
					Execute("if "&strIf&" then ifFlag=true else ifFlag=false")
					if ifFlag then content=replace(content,matchIf.value,strThen1) else content=replace(content,matchIf.value,strElse1)
				else	
					Execute("if "&strIf&" then ifFlag=true else ifFlag=false")
					if ifFlag then content=replace(content,matchIf.value,strThen) else content=replace(content,matchIf.value,"")
				end if
			end if
			elseIfFlag=false
		next
		set matchesIf=nothing
	End Function
	
	
	'解析留言
	Public Function parseGbook()
		Dim gbook
		gbook="<div id=""faqbox"">"&vbcrlf& _
		"<form action=""save.asp?action=add"" method=""post"">"&vbcrlf& _
		"    <div class=""faqline"">"&vbcrlf& _
		"        <span class=""faqtit"">问题：</span>"&vbcrlf& _
		"        <input name=""FaqTitle"" type=""text"" /><font color=""#FF0000"">*</font>"&vbcrlf& _
		"    </div>    "&vbcrlf& _
		"    <div class=""Content"">"&vbcrlf& _
		"        <span class=""faqtit"">内容：</span>"&vbcrlf& _
		"        <textarea name=""Content"" cols=""60"" rows=""5""></textarea><font color=""#FF0000"">*</font>"&vbcrlf& _
		"    </div>"&vbcrlf& _
		"   <div class=""faqline"">"&vbcrlf& _
		"        <span class=""faqtit"">联系人：</span>"&vbcrlf& _
		"        <input name=""Contact"" type=""text"" /><font color=""#FF0000"">*</font>"&vbcrlf& _
		"    </div>"&vbcrlf& _
		"   <div class=""faqline"">"&vbcrlf& _
		"        <span class=""faqtit"">联系方式：</span>"&vbcrlf& _
		"        <input name=""ContactWay"" type=""text"" /> <font color=""#FF0000"">*</font> 请注明是手机、电话、QQ、Email,方便我们和您联系"&vbcrlf& _
		"    </div>"&vbcrlf& _
		"  <div class=""faqline"">"&vbcrlf& _
		"        <span class=""faqtit"">验证码：</span>"&vbcrlf& _
		"        <input name=""code"" type=""text"" class=""login_verification"" id=""verification"" size=""6"" maxlength=""6""/><font color=""#FF0000"">*</font>"&vbcrlf& _
		"        <img src=""../inc/checkcode.asp"" alt=""看不清验证码?点击刷新!"" onClick=""this.src='../inc/checkcode.asp'"" />"&vbcrlf& _
		"    </div>"&vbcrlf& _
		"   <div class=""faqline"">"&vbcrlf& _
		"        <span class=""faqtit"">&nbsp;</span>"&vbcrlf& _
		"        <input type=""submit"" value="" 提交 ""/>"&vbcrlf& _
		"    </div>"&vbcrlf& _
		"</form>"&vbcrlf& _
		"</div>"&vbcrlf
		content=replaceStr(content,"{aspcms:gbook}",gbook)
	End Function
	
	
	Function getSlide
		Dim Str,sTexts,sFiles,sLinks
		if slidestyle=0 then
			sTexts="var texts ;"
			sFiles="var files ;"
			sLinks="var links ;"
			if slideTextStatus then 
				if len(slideTexts) <> 0 then
				sTexts="var texts='"&replace(replace(left(slideTexts,len(slideTexts)-1),",","|")," ","")&"';"
				end if			
			end if
			if len(slideImgs) <> 0 then
			sFiles="var files='"&replace(replace(left(slideImgs,len(slideImgs)-1),",","|")," ","")&"';"
			end if
			if len(slideLinks) <> 0 then		
			sLinks="var links='"&replace(replace(left(slideLinks,len(slideLinks)-1),",","|")," ","")&"';"
			end if
			
			
			Str="<SCRIPT language=JavaScript type=text/javascript>"&vbcrlf& _		
			"var swf_width='"&slideWidth&"';"&vbcrlf& _
			"var swf_height='"&slideHeight&"';"&vbcrlf& _
			"var configtg='0xffffff:文字颜色|2:文字位置|0x000000:文字背景颜色|30:文字背景透明度|0xffffff:按键文字颜色|0x4f6898:按键默认颜色|0x000033:按键当前颜色|8:自动播放时间|3:图片过渡效果|1:是否显示按钮|_blank:打开新窗口';"&vbcrlf& _
			sFiles&vbcrlf& sLinks &vbcrlf& sTexts&vbcrlf& _
			"document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ swf_width +'"" height=""'+ swf_height +'"">');"&vbcrlf& _
			"document.write('<param name=""movie"" value="""&sitePath&"/flash/slideflash.swf""><param name=""quality"" value=""high"">');"&vbcrlf& _
			"document.write('<param name=""menu"" value=""false""><param name=wmode value=""opaque"">');"&vbcrlf& _
			"document.write('<param name=""FlashVars"" value=""bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'"">');"&vbcrlf& _
			"document.write('<embed src="""&sitePath&"/flash/slideflash.swf"" wmode=""opaque"" FlashVars=""bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'& menu=""false"" quality=""high"" width=""'+ swf_width +'"" height=""'+ swf_height +'"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />'); document.write('</object>'); "&vbcrlf& _
			"</SCRIPT>"	
		elseif  slidestyle=1 then
		Str="<SCRIPT src=""/js/slide1.js"" type=text/javascript></SCRIPT>"&vbcrlf& _	
			"<DIV id=sasFlashFocus27></DIV>"&vbcrlf& _
			"<SCRIPT>"&vbcrlf& _	
			"var sohuFlash2 = new sohuFlash(""/flash/slide1.swf"", ""20"", "&slideWidth&", "&slideHeight&", ""7"");"&vbcrlf& _	
			"sohuFlash2.addParam(""quality"", ""high"");"&vbcrlf& _	
			"sohuFlash2.addParam(""wmode"", ""opaque"");"&vbcrlf& _	
			"sohuFlash2.addVariable(""image"","""&replace(replace(left(slideImgs,len(slideImgs)-1),",","|")," ","")&""");"&vbcrlf& _	
			"sohuFlash2.addVariable(""url"","""&replace(replace(left(slideLinks,len(slideLinks)-1),",","|")," ","")&""");"&vbcrlf& _	
			"//sohuFlash2.addVariable(""info"", """&replace(replace(left(slideTexts,len(slideTexts)-1),",","|")," ","")&""");"&vbcrlf& _	
			"sohuFlash2.addVariable(""stopTime"",""5000"");"&vbcrlf& _	
			"sohuFlash2.write(""sasFlashFocus27"");"&vbcrlf& _	
			"</SCRIPT>"
		end if
		
		getSlide=Str
	End Function
	
		Function getSlideb
		Dim Str,sTexts,sFiles,sLinks
		if slidestyleb=0 then
			sTexts="var texts ;"
			sFiles="var files ;"
			sLinks="var links ;"
			if slideTextStatusb then 
				if len(slideTextsb) <> 0 then
				sTexts="var texts='"&replace(replace(left(slideTextsb,len(slideTextsb)-1),",","|")," ","")&"';"
				end if			
			end if
			if len(slideImgsb) <> 0 then
			sFiles="var files='"&replace(replace(left(slideImgsb,len(slideImgsb)-1),",","|")," ","")&"';"
			end if
			if len(slideLinksb) <> 0 then		
			sLinks="var links='"&replace(replace(left(slideLinksb,len(slideLinksb)-1),",","|")," ","")&"';"
			end if
			
			
			Str="<SCRIPT language=JavaScript type=text/javascript>"&vbcrlf& _		
			"var swf_width='"&slideWidthb&"';"&vbcrlf& _
			"var swf_height='"&slideHeightb&"';"&vbcrlf& _
			"var configtg='0xffffff:文字颜色|2:文字位置|0x000000:文字背景颜色|30:文字背景透明度|0xffffff:按键文字颜色|0x4f6898:按键默认颜色|0x000033:按键当前颜色|8:自动播放时间|3:图片过渡效果|1:是否显示按钮|_blank:打开新窗口';"&vbcrlf& _
			sFiles&vbcrlf& sLinks &vbcrlf& sTexts&vbcrlf& _
			"document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ swf_width +'"" height=""'+ swf_height +'"">');"&vbcrlf& _
			"document.write('<param name=""movie"" value="""&sitePath&"/flash/slideflash.swf""><param name=""quality"" value=""high"">');"&vbcrlf& _
			"document.write('<param name=""menu"" value=""false""><param name=wmode value=""opaque"">');"&vbcrlf& _
			"document.write('<param name=""FlashVars"" value=""bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'"">');"&vbcrlf& _
			"document.write('<embed src="""&sitePath&"/flash/slideflash.swf"" wmode=""opaque"" FlashVars=""bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'& menu=""false"" quality=""high"" width=""'+ swf_width +'"" height=""'+ swf_height +'"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />'); document.write('</object>'); "&vbcrlf& _
			"</SCRIPT>"	
		elseif  slidestyleb=1 then
		Str="<SCRIPT src=""/js/slide1.js"" type=text/javascript></SCRIPT>"&vbcrlf& _	
			"<DIV id=sasFlashFocus27></DIV>"&vbcrlf& _
			"<SCRIPT>"&vbcrlf& _	
			"var sohuFlash2 = new sohuFlash(""/flash/slide1.swf"", ""20"", "&slideWidthb&", "&slideHeightb&", ""7"");"&vbcrlf& _	
			"sohuFlash2.addParam(""quality"", ""high"");"&vbcrlf& _	
			"sohuFlash2.addParam(""wmode"", ""opaque"");"&vbcrlf& _	
			"sohuFlash2.addVariable(""image"","""&replace(replace(left(slideImgsb,len(slideImgsb)-1),",","|")," ","")&""");"&vbcrlf& _	
			"sohuFlash2.addVariable(""url"","""&replace(replace(left(slideLinksb,len(slideLinksb)-1),",","|")," ","")&""");"&vbcrlf& _	
			"//sohuFlash2.addVariable(""info"", """&replace(replace(left(slideTextsb,len(slideTextsb)-1),",","|")," ","")&""");"&vbcrlf& _	
			"sohuFlash2.addVariable(""stopTime"",""5000"");"&vbcrlf& _	
			"sohuFlash2.write(""sasFlashFocus27"");"&vbcrlf& _	
			"</SCRIPT>"
		end if
		
		getSlideb=Str
	End Function
	
		Function getSlidec
		Dim Str,sTexts,sFiles,sLinks
		if slidestylec=0 then
			sTexts="var texts ;"
			sFiles="var files ;"
			sLinks="var links ;"
			if slideTextStatusc then 
				if len(slideTextsc) <> 0 then
				sTexts="var texts='"&replace(replace(left(slideTextsc,len(slideTextsc)-1),",","|")," ","")&"';"
				end if			
			end if
			if len(slideImgsc) <> 0 then
			sFiles="var files='"&replace(replace(left(slideImgsc,len(slideImgsc)-1),",","|")," ","")&"';"
			end if
			if len(slideLinksc) <> 0 then		
			sLinks="var links='"&replace(replace(left(slideLinksc,len(slideLinksc)-1),",","|")," ","")&"';"
			end if
			
			
			Str="<SCRIPT language=JavaScript type=text/javascript>"&vbcrlf& _		
			"var swf_width='"&slideWidthc&"';"&vbcrlf& _
			"var swf_height='"&slideHeightc&"';"&vbcrlf& _
			"var configtg='0xffffff:文字颜色|2:文字位置|0x000000:文字背景颜色|30:文字背景透明度|0xffffff:按键文字颜色|0x4f6898:按键默认颜色|0x000033:按键当前颜色|8:自动播放时间|3:图片过渡效果|1:是否显示按钮|_blank:打开新窗口';"&vbcrlf& _
			sFiles&vbcrlf& sLinks &vbcrlf& sTexts&vbcrlf& _
			"document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ swf_width +'"" height=""'+ swf_height +'"">');"&vbcrlf& _
			"document.write('<param name=""movie"" value="""&sitePath&"/flash/slideflash.swf""><param name=""quality"" value=""high"">');"&vbcrlf& _
			"document.write('<param name=""menu"" value=""false""><param name=wmode value=""opaque"">');"&vbcrlf& _
			"document.write('<param name=""FlashVars"" value=""bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'"">');"&vbcrlf& _
			"document.write('<embed src="""&sitePath&"/flash/slideflash.swf"" wmode=""opaque"" FlashVars=""bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'& menu=""false"" quality=""high"" width=""'+ swf_width +'"" height=""'+ swf_height +'"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />'); document.write('</object>'); "&vbcrlf& _
			"</SCRIPT>"	
		elseif  slidestylec=1 then
		Str="<SCRIPT src=""/js/slide1.js"" type=text/javascript></SCRIPT>"&vbcrlf& _	
			"<DIV id=sasFlashFocus27></DIV>"&vbcrlf& _
			"<SCRIPT>"&vbcrlf& _	
			"var sohuFlash2 = new sohuFlash(""/flash/slide1.swf"", ""20"", "&slideWidthc&", "&slideHeightc&", ""7"");"&vbcrlf& _	
			"sohuFlash2.addParam(""quality"", ""high"");"&vbcrlf& _	
			"sohuFlash2.addParam(""wmode"", ""opaque"");"&vbcrlf& _	
			"sohuFlash2.addVariable(""image"","""&replace(replace(left(slideImgsc,len(slideImgsc)-1),",","|")," ","")&""");"&vbcrlf& _	
			"sohuFlash2.addVariable(""url"","""&replace(replace(left(slideLinksc,len(slideLinksc)-1),",","|")," ","")&""");"&vbcrlf& _	
			"//sohuFlash2.addVariable(""info"", """&replace(replace(left(slideTextsc,len(slideTextsc)-1),",","|")," ","")&""");"&vbcrlf& _	
			"sohuFlash2.addVariable(""stopTime"",""5000"");"&vbcrlf& _	
			"sohuFlash2.write(""sasFlashFocus27"");"&vbcrlf& _	
			"</SCRIPT>"
		end if
		
		getSlidec=Str
	End Function
	
		Function getSlided
		Dim Str,sTexts,sFiles,sLinks
		if slidestyled=0 then
			sTexts="var texts ;"
			sFiles="var files ;"
			sLinks="var links ;"
			if slideTextStatusd then 
				if len(slideTextsd) <> 0 then
				sTexts="var texts='"&replace(replace(left(slideTextsd,len(slideTextsd)-1),",","|")," ","")&"';"
				end if			
			end if
			if len(slideImgsd) <> 0 then
			sFiles="var files='"&replace(replace(left(slideImgsd,len(slideImgsd)-1),",","|")," ","")&"';"
			end if
			if len(slideLinksd) <> 0 then		
			sLinks="var links='"&replace(replace(left(slideLinksd,len(slideLinksd)-1),",","|")," ","")&"';"
			end if
			
			
			Str="<SCRIPT language=JavaScript type=text/javascript>"&vbcrlf& _		
			"var swf_width='"&slideWidthd&"';"&vbcrlf& _
			"var swf_height='"&slideHeightd&"';"&vbcrlf& _
			"var configtg='0xffffff:文字颜色|2:文字位置|0x000000:文字背景颜色|30:文字背景透明度|0xffffff:按键文字颜色|0x4f6898:按键默认颜色|0x000033:按键当前颜色|8:自动播放时间|3:图片过渡效果|1:是否显示按钮|_blank:打开新窗口';"&vbcrlf& _
			sFiles&vbcrlf& sLinks &vbcrlf& sTexts&vbcrlf& _
			"document.write('<object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"" width=""'+ swf_width +'"" height=""'+ swf_height +'"">');"&vbcrlf& _
			"document.write('<param name=""movie"" value="""&sitePath&"/flash/slideflash.swf""><param name=""quality"" value=""high"">');"&vbcrlf& _
			"document.write('<param name=""menu"" value=""false""><param name=wmode value=""opaque"">');"&vbcrlf& _
			"document.write('<param name=""FlashVars"" value=""bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'"">');"&vbcrlf& _
			"document.write('<embed src="""&sitePath&"/flash/slideflash.swf"" wmode=""opaque"" FlashVars=""bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'& menu=""false"" quality=""high"" width=""'+ swf_width +'"" height=""'+ swf_height +'"" type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" />'); document.write('</object>'); "&vbcrlf& _
			"</SCRIPT>"	
		elseif  slidestyled=1 then
		Str="<SCRIPT src=""/js/slide1.js"" type=text/javascript></SCRIPT>"&vbcrlf& _	
			"<DIV id=sasFlashFocus27></DIV>"&vbcrlf& _
			"<SCRIPT>"&vbcrlf& _	
			"var sohuFlash2 = new sohuFlash(""/flash/slide1.swf"", ""20"", "&slideWidthd&", "&slideHeightd&", ""7"");"&vbcrlf& _	
			"sohuFlash2.addParam(""quality"", ""high"");"&vbcrlf& _	
			"sohuFlash2.addParam(""wmode"", ""opaque"");"&vbcrlf& _	
			"sohuFlash2.addVariable(""image"","""&replace(replace(left(slideImgsd,len(slideImgsd)-1),",","|")," ","")&""");"&vbcrlf& _	
			"sohuFlash2.addVariable(""url"","""&replace(replace(left(slideLinksd,len(slideLinksd)-1),",","|")," ","")&""");"&vbcrlf& _	
			"//sohuFlash2.addVariable(""info"", """&replace(replace(left(slideTextsd,len(slideTextsd)-1),",","|")," ","")&""");"&vbcrlf& _	
			"sohuFlash2.addVariable(""stopTime"",""5000"");"&vbcrlf& _	
			"sohuFlash2.write(""sasFlashFocus27"");"&vbcrlf& _	
			"</SCRIPT>"
		end if
		
		getSlided=Str
	End Function
	
	Function getFloatAD
		if adStatus=0 then exit function
		Dim Str
		Str="<div id=""img"" style=""position:absolute;""> <a href="""&adLink&""" target=""_blank""> <img src="""&adImagePath&""" width="""&adImgWidth&""" height="""&adImgHeight&""" border=""0""></a> </div>"&vbcrlf& _
		"<script type=""text/javascript"" language=""JavaScript""> "&vbcrlf& _
		"<!-- "&vbcrlf& _
		"var xPos = 20; "&vbcrlf& _
		"var yPos = document.body.clientHeight; "&vbcrlf& _
		"var step = 1; "&vbcrlf& _
		"var delay = 30; "&vbcrlf& _
		"var height = 0; "&vbcrlf& _
		"var Hoffset = 0; "&vbcrlf& _
		"var Woffset = 0; "&vbcrlf& _
		"var yon = 0; "&vbcrlf& _
		"var xon = 0; "&vbcrlf& _
		"var pause = true; "&vbcrlf& _
		"var interval; "&vbcrlf& _
		"img.style.top = yPos; "&vbcrlf& _
		"function changePos() { "&vbcrlf& _
		"width = document.body.clientWidth; "&vbcrlf& _
		"height = document.body.clientHeight; "&vbcrlf& _
		"Hoffset = img.offsetHeight; "&vbcrlf& _
		"Woffset = img.offsetWidth; "&vbcrlf& _
		"img.style.left = xPos + document.body.scrollLeft; "&vbcrlf& _
		"img.style.top = yPos + document.body.scrollTop; "&vbcrlf& _
		"if (yon) { "&vbcrlf& _
		"yPos = yPos + step; "&vbcrlf& _
		"} "&vbcrlf& _
		"else { "&vbcrlf& _
		"yPos = yPos - step; "&vbcrlf& _
		"} "&vbcrlf& _
		"if (yPos < 0) { "&vbcrlf& _
		"yon = 1; "&vbcrlf& _
		"yPos = 0;"&vbcrlf& _
		"} "&vbcrlf& _
		"if (yPos >= (height - Hoffset)) { "&vbcrlf& _
		"yon = 0; "&vbcrlf& _
		"yPos = (height - Hoffset); "&vbcrlf& _
		"} "&vbcrlf& _
		"if (xon) { "&vbcrlf& _
		"xPos = xPos + step; "&vbcrlf& _
		"} "&vbcrlf& _
		"else { "&vbcrlf& _
		"xPos = xPos - step; "&vbcrlf& _
		"} "&vbcrlf& _
		"if (xPos < 0) { "&vbcrlf& _
		"xon = 1; "&vbcrlf& _
		"xPos = 0;"&vbcrlf& _
		"} "&vbcrlf& _
		"if (xPos >= (width - Woffset)) { "&vbcrlf& _
		"xon = 0; "&vbcrlf& _
		"xPos = (width - Woffset); "&vbcrlf& _
		"} "&vbcrlf& _
		"} "&vbcrlf& _
		"function ad() { "&vbcrlf& _
		"img.visibility = ""visible""; "&vbcrlf& _
		"interval = setInterval('changePos()', delay); "&vbcrlf& _
		"} "&vbcrlf& _
		"ad(); "&vbcrlf& _
		"//For more,visit:www.helpor.net "&vbcrlf& _
		"--> "&vbcrlf& _
		"</script>"&vbcrlf	
		getFloatAD=Str
	End Function
	
	Function getOnlineservice
		if serviceStatus=1 then
			if serviceStyle=1 then getOnlineservice=getqqkf1
			if serviceStyle=2 then getOnlineservice=getqqkf2
		end if
	End Function
		
	Function getKf
		if servicekfStatus=1 then
			getKf=decodeHtml(servicekf)
		end if
		if service53kfStatus=1 then
			getKf=getKf&"<script type='text/javascript' src='http://chat.53kf.com/kf.php?arg="& service53kfaccount &"&style=1'></script>"
		end if	
	End Function
	
	
	
	Function getqqkf1
		Dim Str	,i,tempstr
		Str="<LINK rev=stylesheet href="""&sitePath&"/"&"Images/qq/qqkf1/default.css"" type=text/css rel=stylesheet>"&vbcrlf& _
		"<DIV id=kefu_pannel style=""Z-INDEX: 30000; FILTER: alpha(opacity=85); "&serviceLocation&": 0px; POSITION: absolute; TOP: 120px"">"&vbcrlf& _
		"<TABLE cellSpacing=0 cellPadding=0 border=0>"&vbcrlf& _
		"<THEAD id=kefu_pannel_top>"&vbcrlf& _
		"<TR>"&vbcrlf& _
		"<TH class=kefu_Title><SPAN class=kefu_shut id=kefu_ctrl onclick=HideKefu()></SPAN>"&vbcrlf& _
		"<H2 class=txtCut>在线客服</H2></TH></TR></THEAD>"&vbcrlf& _
		"<TBODY id=kefu_pannel_mid>"&vbcrlf& _
		"<TR>"&vbcrlf& _
		"<TD height=3></TD></TR>"&vbcrlf& _
		"<TR>"&vbcrlf& _
		"<TD>"&vbcrlf

		if not isnul(serviceQQ) then
			tempstr = split(serviceQQ," ")
			for i=0 to ubound(tempstr)
				if isExistStr(tempstr(i),"|") then		
					Str=Str&"<div class=kefu_box ><span  style=""padding-left:10px;"">"&split(tempstr(i),"|")(0)&"</span></div><DIV class=kefu_box onmouseover=""this.className='kefu_boxOver'"" onmouseout=""this.className='kefu_box'""><SPAN class=kefu_image><IMG src="""&sitePath&"/"&"Images/qq/qqkf1/icon_person_stat_online.gif""></SPAN><a target=""_blank"" href=""http://wpa.qq.com/msgrd?v=3&uin="&split(tempstr(i),"|")(1)&"&amp;site=qq&menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=2:"&split(tempstr(i),"|")(1)&":47"" alt=""点击这里给我发消息"" title=""点击这里给我发消息""/></a></DIV>"&vbcrlf

				else				
					Str=Str&"<DIV class=kefu_box onmouseover=""this.className='kefu_boxOver'"" onmouseout=""this.className='kefu_box'""><SPAN class=kefu_image><IMG src="""&sitePath&"/"&"Images/qq/qqkf1/icon_person_stat_online.gif""></SPAN><A class=kefu_Type_qq href=""http://wpa.qq.com/msgrd?v=3&uin="&split(tempstr(i),"|")(1)&"&amp;site=qq&menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=2:"&split(tempstr(i),"|")(1)&":47"" alt=""点击这里给我发消息"" title=""点击这里给我发消息""></A></DIV>"&vbcrlf
				end if
			next
		end if
		
		if not isnul(serviceWangWang) then		
			tempstr = split(serviceWangWang," ")
			for i=0 to ubound(tempstr)			
				if isExistStr(tempstr(i),"|") then	
					Str=Str&"<div class=kefu_box><span  style=""padding-left:10px;"">"&split(tempstr(i),"|")(0)&"</span></div><DIV class=kefu_box onmouseover=""this.className='kefu_boxOver'"" onmouseout=""this.className='kefu_box'""><SPAN class=kefu_image><IMG src="""&sitePath&"/"&"Images/qq/qqkf1/icon_person_stat_online.gif""></SPAN><A target=""_blank"" class=kefu_Type_msn href=""http://amos1.taobao.com/msg.ww?v=2&uid="&split(tempstr(i),"|")(1)&"&s=1""><img border=""0"" src=""http://amos1.taobao.com/online.ww?v=2&uid="&split(tempstr(i),"|")(1)&"&s=1"" alt=""点击这里给我发消息"" /></A></DIV>"&vbcrlf					
				else
					Str=Str&"<DIV class=kefu_box onmouseover=""this.className='kefu_boxOver'"" onmouseout=""this.className='kefu_box'""><SPAN class=kefu_image><IMG src="""&sitePath&"/"&"Images/qq/qqkf1/icon_person_stat_online.gif""></SPAN><A target=""_blank"" class=kefu_Type_msn href=""http://amos1.taobao.com/msg.ww?v=2&uid="&trim(tempstr(i))&"&s=1""><img border=""0"" src=""http://amos1.taobao.com/online.ww?v=2&uid="&trim(tempstr(i))&"&s=1"" alt=""点击这里给我发消息"" /></A></DIV>"&vbcrlf	
				end if
			next
		end if		
		Str=Str&"</TD></TR>"&vbcrlf& _
		"<TR>"&vbcrlf& _
		"<TD height=3></TD></TR></TBODY>"&vbcrlf& _
		"<TFOOT id=kefu_pannel_btm>"&vbcrlf& _
		"<TR style=""CURSOR: hand"" onclick=""parent.location='"&serviceContact&"';"">"&vbcrlf& _
		"<TD class=kefu_other></TD></TR></TFOOT></TABLE></DIV>"&vbcrlf& _
		"<SCRIPT language=JavaScript src="""&sitePath&"/"&"Images/qq/qqkf1/qqkf.js""></SCRIPT>"&vbcrlf
		getqqkf1=Str
	End Function
	
	
	Function getqqkf2
		Dim Str	,i,tempstr
		Str="<LINK rev=stylesheet href="""&sitePath&"/"&"Images/qq/qqkf2/kf.css"" type=text/css rel=stylesheet>"&vbcrlf& _
		"<div onmouseout=""toSmall()"" onmouseover=""toBig()"" id=""qq_Kefu"" style=""top: 1363.9px; left: -143px; position: absolute; "">"&vbcrlf& _
		"  <table cellspacing=""0"" cellpadding=""0"" border=""0"">"&vbcrlf& _
		"    <tbody><tr>"&vbcrlf& _
		"      <td><table cellspacing=""0"" cellpadding=""0"" border=""0"">"&vbcrlf& _
		"          <tbody>"&vbcrlf& _
		"          <tr>"&vbcrlf& _
		"          	<td background="""&sitePath&"/"&"Images/qq/qqkf2/Kf_bg03_01.gif"" height=""32""></td>"&vbcrlf& _
		"          </tr>"&vbcrlf& _
		"          <tr>"&vbcrlf& _
		"            <td background="""&sitePath&"/"&"Images/qq/qqkf2/Kf_bg03_02.gif"" align=""left""  width=""143"">"&vbcrlf& _
		"            <div class=""kfbg"">"&vbcrlf
		if not isnul(serviceQQ) then
			tempstr = split(serviceQQ," ")
			for i=0 to ubound(tempstr)				
				if isExistStr(tempstr(i),"|") then	
					Str=Str&" <div class=kefu_box><span  style=""padding-left:10px;"">"&split(tempstr(i),"|")(0)&"</span></div><DIV class=kefu_box onmouseover=""this.className='kefu_boxOver'"" onmouseout=""this.className='kefu_box'""><A class=kefu_Type_qq href=""http://wpa.qq.com/msgrd?v=3&uin="&split(tempstr(i),"|")(1)&"&amp;site=qq&menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=2:"&split(tempstr(i),"|")(1)&":47"" alt=""点击这里给我发消息"" title=""点击这里给我发消息""></A></DIV>"&vbcrlf
				else
					Str=Str&" <div class=kefu_box><span  style=""padding-left:10px;"">销售代表：</span></div><DIV class=kefu_box onmouseover=""this.className='kefu_boxOver'"" onmouseout=""this.className='kefu_box'""><A class=kefu_Type_qq href=""http://wpa.qq.com/msgrd?v=3&uin="&split(tempstr(i),"|")(1)&"&amp;site=qq&menu=yes""><img border=""0"" src=""http://wpa.qq.com/pa?p=2:"&split(tempstr(i),"|")(1)&":47"" alt=""点击这里给我发消息"" title=""点击这里给我发消息""></A></DIV>"&vbcrlf				
				end if
			next
		end if
		
		if not isnul(serviceWangWang) then		
			tempstr = split(serviceWangWang," ")
			for i=0 to ubound(tempstr)	
				if isExistStr(tempstr(i),"|") then
					Str=Str&"<div class=kefu_box><span  style=""padding-left:10px;"">"&split(tempstr(i),"|")(0)&"</span></div><DIV class=kefu_box onmouseover=""this.className='kefu_boxOver'"" onmouseout=""this.className='kefu_box'""><A target=""_blank"" class=kefu_Type_msn href=""http://amos1.taobao.com/msg.ww?v=2&uid="&trim(split(tempstr(i),"|")(1))&"&s=1""><img border=""0"" src=""http://amos1.taobao.com/online.ww?v=2&uid="&trim(split(tempstr(i),"|")(1))&"&s=1"" alt=""点击这里给我发消息"" /></A></DIV>"&vbcrlf					
				else
					Str=Str&"<DIV class=kefu_box onmouseover=""this.className='kefu_boxOver'"" onmouseout=""this.className='kefu_box'""><A target=""_blank"" class=kefu_Type_msn href=""http://amos1.taobao.com/msg.ww?v=2&uid="&trim(tempstr(i))&"&s=1""><img border=""0"" src=""http://amos1.taobao.com/online.ww?v=2&uid="&trim(tempstr(i))&"&s=1"" alt=""点击这里给我发消息"" /></A></DIV>"&vbcrlf		
				end if
			next
		end if	

         Str=Str&"</div>   "&vbcrlf& _
		"            </td>"&vbcrlf& _
		"          </tr>"&vbcrlf& _
		"          <tr  style=""CURSOR: hand"" onclick=""parent.location='"&serviceContact&"';"" >"&vbcrlf& _
		"            <td height=""46"" background="""&sitePath&"/"&"Images/qq/qqkf2/Kf_bg03_03.gif""><div class=""Kefu_Work""></div></td>"&vbcrlf& _
		"          </tr>"&vbcrlf& _
		"        </tbody></table></td>"&vbcrlf& _
		"      <td class=""kfbut"" width=""27"" rowspan=""3"" class=""Kefu_Little""></td>"&vbcrlf& _
		"    </tr>"&vbcrlf& _
		"  </tbody></table>"&vbcrlf& _
		"</div>"&vbcrlf& _
		"<script src="""&sitePath&"/"&"Images/qq/qqkf2/kefu.js"" type=""text/javascript""></script>"&vbcrlf
		getqqkf2=Str
	End Function	
	
	
	
	Public Function parsePrevAndNext(Id,SortID)
		Dim rsObjPrev,rsObjNext,tempStr,linkStr
		set rsObjPrev = conn.Exec("select top 1 ContentID,Title,sortType,SortFolder,a.GroupID,ContentFolder,ContentFileName,a.AddTime,a.PageFileName,a.SortID,b.GroupID from {prefix}Content as a,{prefix}Sort as b where a.SortID=b.SortID and ContentStatus=1 and TimeStatus=0 and ContentID<"&Id&" and a.SortID="&SortID&" order by ContentID desc","r1")
		if rsObjPrev.bof then 
			linkStr ="没有了!"
			content = replace(content,"{aspcms:prevlink}","#")
		else
			linkStr=getContentLink(rsObjPrev("SortID"),rsObjPrev("ContentID"),rsObjPrev("SortFolder"),rsObjPrev("a.GroupID"),rsObjPrev("ContentFolder"),rsObjPrev("ContentFileName"),rsObjPrev("AddTime"),rsObjPrev("PageFileName"),rsObjPrev("b.GroupID"))
			content = replace(content,"{aspcms:prevlink}",linkStr)
			linkStr="<a href="""&linkStr&""">"&rsObjPrev(1)&"</a>"
			content = replace(content,"{aspcms:prevtitle}",rsObjPrev(1))
		end if
		content = replace(content,"{aspcms:prev}",linkStr)
		
		rsObjPrev.close : set rsObjPrev = nothing
		
		set rsObjNext = conn.Exec("select top 1 ContentID,Title,sortType,SortFolder,a.GroupID,ContentFolder,ContentFileName,a.AddTime,a.PageFileName,a.SortID,b.GroupID from {prefix}Content as a,{prefix}Sort as b where a.SortID=b.SortID and ContentStatus=1 and TimeStatus=0 and ContentID>"&Id&" and a.SortID="&SortID&" order by ContentID asc","r1")
		if rsObjNext.eof then 
			linkStr = "没有了!"
			content = replace(content,"{aspcms:nextlink}","#")
			content = replace(content,"{aspcms:nexttitle}","没有了!")
		else			
			linkStr=getContentLink(rsObjNext("SortID"),rsObjNext("ContentID"),rsObjNext("SortFolder"),rsObjNext("a.GroupID"),rsObjNext("ContentFolder"),rsObjNext("ContentFileName"),rsObjNext("AddTime"),rsObjNext("PageFileName"),rsObjNext("b.GroupID"))
			content = replace(content,"{aspcms:nextlink}",linkStr)
			linkStr="<a href="""&linkStr&""">"&rsObjNext(1)&"</a>"
			content = replace(content,"{aspcms:nexttitle}",rsObjNext(1))
		end if	
		content = replace(content,"{aspcms:next}",linkStr)
		rsObjNext.close	: set rsObjNext = nothing
	End Function
	
	Function getArrt(str,tag,arr)
		Dim labelRule,match,matches
		labelRule = "\["&str&":"&tag&"([\s\S]*?)\]"
		regExpObj.Pattern = labelRule
		set matches = regExpObj.Execute(content)
		for each match in matches
			getArrt = parseArr(match.SubMatches(0))(arr)			
		next
		set matches = nothing
		strDictionary.removeAll
	End Function
	
	Function getTopType(SortID)
		Dim tempStr,rsObj
		set rsObj = conn.Exec("select * from {prefix}Sort where SortID="&SortID&"","r1")
			tempStr=tempStr&"<a href="""&getSortLink(rsObj("sortType"),rsObj("sortID"),rsObj("sortUrl"),rsObj("sortFolder"),rsObj("sortFileName"),rsObj("GroupID"),rsObj("Exclusive"))&""">"&rsObj("SortName")&"</a>,"
			if rsObj("ParentID")<>0 then tempStr=tempStr&getTopType(rsObj("ParentID"))
		rsObj.close : set rsObj=nothing
		getTopType=tempStr
	End Function 
	
	
	Public Function parsePosition(SortID)
		Dim rsObjSmalltype
		set rsObjSmalltype = conn.Exec("select SortName from {prefix}Sort where SortID="&SortID&"","r1")
			content = replace(content,"{aspcms:sortname}",rsObjSmalltype(0))
		rsObjSmalltype.close : set rsObjSmalltype=nothing		
		if not isExistStr(content,"{aspcms:position") then Exit Function
		dim match,matches,matchfield,matchesfield,arrlen
		dim labelAttrLinklist,loopstrLinklist,loopstrLinklistNew,loopstrTotal
		dim vtype,vnum,whereStr,linkArray
		dim fieldName,fieldAttr,fieldNameAndAttr,fieldAttrLen
		dim i,labelRuleField
		dim m,namelen,deslen,m_des
		labelRule="{aspcms:position([\s\S]*?)}([\s\S]*?){/aspcms:position}"
		labelRuleField="\[position:([\s\S]+?)\]"
		regExpObj.Pattern=labelRule
		set matches=regExpObj.Execute(content)
		for each match in matches
		
			linkArray=Split(getTopType(SortID),",")
			arrlen=ubound(linkArray)
			for i=0 to arrlen-1				
				loopstrLinklist=match.SubMatches(1)
				regExpObj.Pattern=labelRuleField
				set matchesfield=regExpObj.Execute(loopstrLinklist)
				loopstrLinklistNew=loopstrLinklist
				for each matchfield in matchesfield
					fieldNameAndAttr=regExpReplace(matchfield.SubMatches(0),"[\s]+",chr(32))
					fieldNameAndAttr=trimOuter(fieldNameAndAttr)				
					
					m=instr(fieldNameAndAttr,chr(32))
					'die m
					if m > 0 then 
						fieldName=left(fieldNameAndAttr,m - 1)
						fieldAttr =	right(fieldNameAndAttr,len(fieldNameAndAttr) - m)
					else
						fieldName=fieldNameAndAttr
						fieldAttr =	""
					end if					
					select case fieldName
						case "link"
							loopstrLinklistNew=replaceStr(loopstrLinklistNew,matchfield.value,linkArray(arrlen-1-i))					
					end select
				next
				loopstrTotal=loopstrTotal&loopstrLinklistNew
			next				
			set matchesfield=nothing
			content=replaceStr(content,match.value,loopstrTotal)
			strDictionary.removeAll
		next
		set matches=nothing 		
	End Function
	
	Public Function parseLabel()
		if not isExistStr(content,"{label:") then Exit Function

		dim matches,match,labelName,selfLabelArray,selfLabelLen,sql,singleAttrKey,singleAttrValue
		sql="select LabelName,LabelContent from {prefix}Labels"
		labelRule="{label:([\s\S]+?)}"		
		selfLabelArray=conn.exec(sql,"arr")
		if isArray(selfLabelArray) then 
			for selfLabelLen=0 to ubound(selfLabelArray,2)
				singleAttrKey=selfLabelArray(0,selfLabelLen)
				singleAttrValue=decodeHtml(selfLabelArray(1,selfLabelLen))
				if not strDictionary.Exists(singleAttrKey) then strDictionary.add singleAttrKey,singleAttrValue  else  strDictionary(singleAttrKey)=singleAttrValue
			next
		end if
		regExpObj.Pattern=labelRule
		set matches=regExpObj.Execute(content)
		for each match in matches
			labelName=trim(match.SubMatches(0))
			if strDictionary.Exists(labelName) then  content=replace(content,match.value,strDictionary(labelName))
		next
		strDictionary.RemoveAll
		set matches=nothing
	End Function		
	
Public Function parseAdvList()
		'if not isExistStr(content,"{aspcms:adv") then Exit Function
		dim lenPagelist,TypeId,strPagelist,lsize,rsObj,labelRuleField,labelRulePagelist,matches,match,labelStr,loopStr,labelArr,lorder,orderStr,sperStrs,spec,sperStr,title,whereStr,AdvArray,loopstrAdv,arri,arrl,loopstrAdvpf,loopstrAdvdl,loopstrAdvtc,i
		labelRule = "{aspcms:adv([\s\S]*?)}"
		regExpObj.Pattern = labelRule
		set matches = regExpObj.Execute(content)
		loopstrAdv=""
		for each match in matches
			whereStr=chr(32)&"AdvStatus=1 and AdvID="&match.SubMatches(0)
			AdvArray=conn.Exec("select AdvName,AdvClass,AdvImg,AdvLink,AdvWidth,AdvHeight,AdvStime,AdvEtime,AdvContent from {prefix}Adv  where "&whereStr&"","arr")
			if isarray(AdvArray) then
				if now()>AdvArray(6,0) and now()<AdvArray(7,0) then
				
					loopstrAdv="<script src="""&sitePath&"/inc/AspCms_AdvJs.asp?id="& match.SubMatches(0) &""" language=""JavaScript""></script>"
				
				end if
			end if
			content=replaceStr(content,match.value,loopstrAdv)
		next
		content=replaceStr(content,"{aspcms:floatad}","<script src="""&sitePath&"/js/piaofu.js"" language=""JavaScript""></script><script src="""&sitePath&"/inc/AspCms_AdvJs.asp?type=pf"" language=""JavaScript""></script>")	
		content=replaceStr(content,"{aspcms:coupletad}","<script src="""&sitePath&"/inc/AspCms_AdvJs.asp?type=dl"" language=""JavaScript""></script>")		
		content=replaceStr(content,"{aspcms:windowad}","<script src="""&sitePath&"/inc/AspCms_AdvJs.asp?type=tc"" language=""JavaScript""></script>")
	End Function
			
	Public Function parseHtml()		
		parseTopAndFoot()
		parseAuxiliaryTemplate() 
		parseLabel() 
	End Function
	
	Public Function parseCommon()
		parseGlobal()
		'echo "<h1>parseGlobal()</h1>"		
		parseNavList("")		
		'echo "<h1>parseNavList("""")</h1>"
		parseSubNavList()	
		'echo "<h1>parseSubNavList()</h1>"
		parseLinkList()
		'echo "<h1>parseLinkList()</h1>"
		parseLoop("news")	'for 1.5
		'echo "<h1>parseLoop(""news"")</h1>"
		parseLoop("down")	'for 1.5
		'echo "<h1>parseLoop(""down"")</h1>"
		parseLoop("pic")	'for 1.5
		'echo "<h1>parseLoop(""pic"")</h1>"
		parseLoop("product")	'for 1.5
		'echo "<h1>parseLoop(""product"")</h1>"
		parseLoop("aboutart")	'for 1.5
		'echo "<h1>parseLoop(""aboutart"")</h1>"		
		parseLoop("type")
		'echo "<h1>parseLoop(""type"")</h1>"
		parseLoop("about")
		'echo "<h1>parseLoop(""about"")</h1>"
		parseLoop("content")	
		'echo "<h1>parseLoop(""content"")</h1>"
		parseLoop("tag")	
		'echo "<h1>parseLoop(""tag"")</h1>"
		parseLoop("gbook")	
		'echo "<h1>parseLoop(""gbook"")</h1>"
		parseLoop("tag")
		
		parseLoop("comment")
		'echo "<h1>parseLoop(""tag"")</h1>"
		parseAdvList()
		'echo "<h1>parseAdvList()</h1>"		
		'echo "<hr>"&content&"<hr>"
		
		dim searchtype : searchtype=filterPara(getForm("searchtype","get"))
		if isnul(searchtype) then searchtype=0
		
		content=replaceStr(content,"{aspcms:keys}",filterPara(getForm("keys","get")))
		content=replaceStr(content,"{aspcms:searchtype}",searchtype)	
		content=replaceStr(content,"{aspcms:searchstyle}",searchtype)	'for 1.5
		
		parseGbook()
		'echo "<hr>"&content&"<hr>"
		parseIf()
		'echo "parseCommon() 执行完毕"
		
		
				
	End Function

	
End Class
%>