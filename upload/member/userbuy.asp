<!--#include file="../inc/AspCms_SettingClass.asp" -->
<%
dim userID,Page
userid=session("userID")
if isNul(userID) or not isNum(userID) then alertMsgAndGo "�Բ���,���ĵ�¼״̬�Ѿ�ʧЧ,�����µ�¼!","login.asp"
page=request.QueryString("page")
'if not isNul(page) and isNum(page) then page=clng(page) else echoMsgAndGo "ҳ�治����",3 end if

echo struserbuy(userid, Page)
If DebugMode Then echo timer - AppSpan
%>