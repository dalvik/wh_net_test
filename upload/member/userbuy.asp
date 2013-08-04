<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim userID,Page
userid=session("userID")
if isNul(userID) or not isNum(userID) then alertMsgAndGo "对不起,您的登录状态已经失效,请重新登录!","login.asp"
page=request.QueryString("page")
'if not isNul(page) and isNum(page) then page=clng(page) else echoMsgAndGo "页面不存在",3 end if

echo struserbuy(userid, Page)
If DebugMode Then echo timer - AppSpan
%>