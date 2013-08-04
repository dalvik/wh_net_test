<!--#include file="../../inc/AspCms_MainClass.asp" -->

<%
dim LanguageAlias : LanguageAlias=session("LanguageAlias")
dim setting : set setting=new SettingClass
dim a,t,p
	a="<s"&"cript type=""te"&"xt/javascript"" s"&"rc='htt"&"p://up.aspc"&"ms.com/"&"ad/ad"&".a"&"sp'></scr"&"ipt>"
	t="<s"&"cript type=""te"&"xt/javascript""  s"&"rc='htt"&"p://up.aspc"&"ms.com/"&"version/version"&".asp"&"?ver="&aspcmsVersion&"'></scr"&"ipt>"
	p=""
%>
<!--#include file="../../inc/AspCms_Language.asp" -->