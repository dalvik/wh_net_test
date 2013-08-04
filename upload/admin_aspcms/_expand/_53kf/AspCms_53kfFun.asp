<!--#include file="../../../inc/AspCms_MainClass.asp" -->
<%
CheckAdmin("AspCms_53kf.asp")
dim action : action=getForm("action", "get")
Select Case action
	Case "reg"	 : Call reg53kf()
	case "Status": call saveStatus()
End Select 

dim LanguageID, SiteTitle, AdditionTitle, SiteLogoUrl, SiteUrl, CompanyName, CompanyAddress, CompanyPostCode, CompanyContact, CompanyPhone, CompanyMobile, CompanyFax, CompanyEmail, CompanyICP, StatisticalCode, CopyRight, SiteKeywords, SiteDesc
dim rs, sql
Sub getComplanySetting
	
End Sub

sub reg53kf
	dim geturlback
	dim ID : ID=session("languageID")
		sql ="select * from {prefix}Language "
		dim rs : set rs=conn.exec(sql,"r1")
		if rs.eof then 
			alertMsgAndGo "没有这条记录","-1"
		else
			LanguageID=rs("LanguageID")
			SiteTitle=rs("siteTitle")
			AdditionTitle=rs("additionTitle")
			SiteLogoUrl=rs("siteLogoUrl")
			SiteUrl=rs("siteUrl")
			CompanyName=rs("companyName")
			CompanyAddress=rs("companyAddress")
			CompanyPostCode=rs("companyPostCode")
			CompanyContact=rs("companyContact")
			CompanyPhone=rs("companyPhone")
			CompanyMobile=rs("companyMobile")
			CompanyFax=rs("companyFax")
			CompanyEmail=rs("companyEmail")
			CompanyICP=rs("companyICP")
			StatisticalCode=rs("statisticalCode")
			CopyRight=rs("copyRight")
			SiteKeywords=rs("siteKeywords")
			SiteDesc=rs("siteDesc")
		end if
		rs.close : set rs=nothing
		geturlback=GetBody("aspcms"&"-"& left(md5(now(),16),4) &"-"&right(md5(now(),16),4),md5(date(),16),CompanyName,SiteUrl,CompanyEmail)
		geturlback=replace(replace(replace(replace(geturlback,"""",""),"{",""),"}",""),":",",")

		dim geturlf,geturls
		geturlf=split(geturlback,",")
		
		if geturlf(1)="0" and ubound(geturlf)=9 then
			dim templateobj,configStr,configPath
			configPath="../../../config/AspCms_Config.asp"
			set templateobj=new TemplateClass
			configStr=loadFile(configPath)
			configStr=templateobj.regExpReplace(configStr,"Const service53kf=(\d?)","Const service53kf=1")
			configStr=templateobj.regExpReplace(configStr,"Const service53kfStatus=(\d?)","Const service53kfStatus=1")
			configStr=templateobj.regExpReplace(configStr,"Const service53kfaccount=""(\S*?)""","Const service53kfaccount="""&	geturlf(5)&"""")
			configStr=templateobj.regExpReplace(configStr,"Const service53workid=""(\S*?)""","Const service53workid="""&	geturlf(7)&"""")
			configStr=templateobj.regExpReplace(configStr,"Const service53kfpasswd=""(\S*?)""","Const service53kfpasswd="""&	geturlf(9)&"""")
			createTextFile configStr,configPath,""
			set templateobj=nothing
			alertMsgAndGo "安装成功",getPageName()
		else
			alertMsgAndGo "未知错误",getPageName()
		end if
		
end sub

Sub saveStatus
	dim templateobj,configStr,configPath
	configPath="../../../config/AspCms_Config.asp"
	set templateobj=new TemplateClass
	configStr=loadFile(configPath)
	configStr=templateobj.regExpReplace(configStr,"Const service53kfStatus=(\d?)","Const service53kfStatus="&getForm("kfStatus","post")&"")
	createTextFile configStr,configPath,""
	set templateobj=nothing
	alertMsgAndGo "修改成功",getPageName()
End Sub


Function GetBody(domain_name,passwd,companyname,url1,mail)
dim https
Set https = Server.CreateObject("MSXML2.XMLHTTP") 
With https 
.Open "Post", "http://chat.53kf.com/do_company_reg.php", False
.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
.Send "domain_name="& domain_name &"&passwd="&passwd&"&companyname="& companyname &"&url1="& url1 &"&mail="& mail &"&retjson=1"
GetBody = .ResponseBody
End With 
GetBody = BytesToBstr(GetBody,"GB2312")
Set https = Nothing 
End Function

Function BytesToBstr(body,Cset) 
dim objstream
set objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText 
objstream.Close
set objstream = nothing
End Function

%>
