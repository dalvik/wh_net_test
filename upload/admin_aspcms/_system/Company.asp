<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim LanguageID, SiteTitle, AdditionTitle, SiteLogoUrl, SiteUrl, CompanyName, CompanyAddress, CompanyPostCode, CompanyContact, CompanyPhone, CompanyMobile, CompanyFax, CompanyEmail, CompanyICP, StatisticalCode, CopyRight, SiteKeywords, SiteDesc
dim rs, sql
Sub getComplanySetting
	dim ID : ID=session("languageID")
		sql ="select * from {prefix}Language "
		dim rs : set rs=conn.exec(sql,"r1")
		if rs.eof then 
			alertMsgAndGo "û��������¼","-1"
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
��ҵ����:<%=CompanyName%><br>
��ҵ��ַ:<%=CompanyAddress%><br>
��������:<%=CompanyPostCode%><br>
��ϵ��:<%=CompanyContact%><br>
�绰����:<%=CompanyPhone%><br>
�ֻ�����:<%=CompanyMobile%><br>
��˾����:<%=CompanyFax%><br>
��������:<%=CompanyEmail%><br>
������:<%=CompanyICP%><br>
ͳ�ƴ���:<%=decode(StatisticalCode)%><br>
��Ȩ��Ϣ:<%=decode(CopyRight)%><br>
վ��ؼ���:<%=SiteKeywords%><br>
վ������:<%=SiteDesc%>