<!--#include file="AspCms_SettingFun.asp" -->
<%CheckAdmin("AspCms_ComplanySetting.asp")%>
<%getComplanySetting%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<LINK href="../images/style.css" type=text/css rel=stylesheet>
<script type="text/javascript" src="../js/jquery.min.js"></script>
<script type="text/javascript" src="../js/jquery.tinyTips.js"></script>
<script type="text/javascript">
		$(document).ready(function() {
			$('a.tTip').tinyTips('title');
			$('a.imgTip').tinyTips('title');
			$('img.tTip').tinyTips('title');
			$('h1.tagline').tinyTips('tinyTips are totally awesome!');
		});
		</script>
<link rel="stylesheet" type="text/css" media="screen" href="../css/tinyTips.css" />
</HEAD>
<BODY>
<DIV class=formzone>
<FORM name="" action="?action=savec" method="post">
<DIV class=tablezone>

<TABLE cellSpacing=0 cellPadding=8 width="100%" align=center border=0>
<TBODY>
    <TR>
    <TD width="225" class=innerbiaoti><STRONG>��վ��Ϣ</STRONG></TD>
        <TD class=innerbiaoti width=661 height=28><STRONG>����</STRONG></TD>
        </TR>
        <TR class=list>
        <TD><STRONG>��վ���� : </STRONG></TD>
        <TD width=661><INPUT class="input" size="30" style="width:300px;"  value="<%=SiteTitle%>" name="SiteTitle"/> <img src="../images/help.gif" class="tTip" title="��վ������<br>���ñ�ǩ��{aspcms:sitetitle}"></TD>
    </TR>
    <TR class=list>
        <TD><STRONG>��ҳ���ӱ��� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=AdditionTitle%>" name="AdditionTitle"/> <img src="../images/help.gif" class="tTip" title="��վ������<br>���ñ�ǩ��{aspcms:additiontitle}"></TD>
    </TR>
    <TR class=list>
        <TD><STRONG>LOGOͼƬURL : </STRONG></TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;"  value="<%=SiteLogoUrl%>" name="SiteLogoUrl" id="SiteLogoUrl"/> <img src="../images/help.gif" class="tTip" title="��վLOGOͼƬ<br>���ñ�ǩ��{aspcms:sitelogo}">
        </TD>
    </TR>
    <TR class=list>   
        <TD valign="top"><STRONG>�ϴ� : </STRONG></TD>
        <TD align="left"><iframe src="../editor/upload.asp?sortType=12&stype=file&Tobj=SiteLogoUrl" scrolling="no" topmargin="0" height="40" width="100%" marginwidth="0" marginheight="0" frameborder="0" align="left"></iframe>
        </TD>
    </TR>
	 <TR class=list>
        <TD><STRONG>icoͼƬURL : </STRONG></TD>
        <TD align="left"><INPUT class="input" size="30" style="width:300px;"  value="favicon.ico" name="SiteicoUrl" id="SiteicoUrl" disabled/> 
        </TD>
    </TR>
    <TR class=list>   
        <TD valign="top"><STRONG>�ϴ� : </STRONG></TD>
        <TD align="left"><iframe src="../editor/upload.asp?sortType=15&stype=file&Tobj=SiteicoUrl" scrolling="no" topmargin="0" height="40" width="100%" marginwidth="0" marginheight="0" frameborder="0" align="left"></iframe>
        </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>վ����ַ : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=SiteUrl%>" name="SiteUrl"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>��ҵ���� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyName%>" name="CompanyName"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>��ҵ��ַ : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyAddress%>" name="CompanyAddress"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>�������� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyPostCode%>" name="CompanyPostCode"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>��ϵ�� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyContact%>" name="CompanyContact"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>�绰���� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyPhone%>" name="CompanyPhone"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>�ֻ����� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyMobile%>" name="CompanyMobile"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>��˾���� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyFax%>" name="CompanyFax"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>�������� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyEmail%>" name="CompanyEmail"/> </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>������ : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:300px;"  value="<%=CompanyICP%>" name="CompanyICP"/> <img src="../images/help.gif" class="tTip" title="��վ������<br>���ñ�ǩ��{aspcms:companyicp}"></TD>
    </TR>
   
    <TR class=list>
        <TD><STRONG>ͳ�ƴ��� : </STRONG></TD>
        <TD>        
        <textarea class="textarea" name="StatisticalCode" cols="60" style="width:80%" rows="3"><%=decodeHtml(StatisticalCode)%></textarea> <img src="../images/help.gif" class="tTip" title="������ͳ�ƴ���<br>���ñ�ǩ��{aspcms:companypostcode}">
		</TD>
    </TR>
    <TR class=list>
        <TD><STRONG>�ײ���Ϣ : </STRONG></TD>
        <TD>
        <textarea class="textarea" name="CopyRight" cols="60" style="width:80%" rows="3"><%=decodeHtml(CopyRight)%></textarea> <img src="../images/help.gif" class="tTip" title="��Ȩ��Ϣ<br>���ñ�ǩ��{aspcms:copyright}">
         </TD>
    </TR>
    <TR class=list>
        <TD><STRONG>վ��ؼ��� : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:80%;"  value="<%=SiteKeywords%>" name="SiteKeywords"/> <img src="../images/help.gif" class="tTip" title="��վ�ؼ���<br>���ñ�ǩ��{aspcms:sitekeywords}"></TD>
    </TR>
    <TR class=list>
        <TD><STRONG>վ������ : </STRONG></TD>
        <TD><INPUT class="input" size="30" style="width:80%;"  value="<%=SiteDesc%>" name="SiteDesc"/> <img src="../images/help.gif" class="tTip" title="��վ��Ϣ����<br>���ñ�ǩ��{aspcms:sitedesc}"></TD>
    </TR>
</TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT type="hidden" value="<%=LanguageID%>" name="LanguageID" />
<INPUT class="button" type="submit" value="����" />
<INPUT class="button" type="button" value="����" onClick="history.go(-1)"/> 
 </DIV></FORM>
</DIV>
</BODY>
</HTML>
