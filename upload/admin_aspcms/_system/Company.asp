<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim LanguageID, SiteTitle, AdditionTitle, SiteLogoUrl, SiteUrl, CompanyName, CompanyAddress, CompanyPostCode, CompanyContact, CompanyPhone, CompanyMobile, CompanyFax, CompanyEmail, CompanyICP, StatisticalCode, CopyRight, SiteKeywords, SiteDesc
dim rs, sql
Sub getComplanySetting
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
End Sub

getComplanySetting
%>
企业名称:<%=CompanyName%><br>
企业地址:<%=CompanyAddress%><br>
邮政编码:<%=CompanyPostCode%><br>
联系人:<%=CompanyContact%><br>
电话号码:<%=CompanyPhone%><br>
手机号码:<%=CompanyMobile%><br>
公司传真:<%=CompanyFax%><br>
电子邮箱:<%=CompanyEmail%><br>
备案号:<%=CompanyICP%><br>
统计代码:<%=decode(StatisticalCode)%><br>
版权信息:<%=decode(CopyRight)%><br>
站点关键词:<%=SiteKeywords%><br>
站点描述:<%=SiteDesc%>