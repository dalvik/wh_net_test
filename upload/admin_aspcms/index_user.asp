<!--#include file="inc/AspCms_SettingClass.asp" -->
<%
if isnul(session("adminName")) then response.Redirect("login.asp"):response.End()

dim id, rsObj
id=getForm("id","get")
if not isnul(id) then
	set rsObj=conn.exec("select * from {prefix}Language where LanguageID="&id,"exe")
	if not rsObj.eof then 
		session("languageID")=id
		session("languagePath")=rsObj("LanguagePath")
		session("LanguageAlias")=rsObj("Alias")
	else		
		alertMsgAndGo "û���������","-1"
	end if
	rsObj.close : set rsObj=nothing
	response.Redirect("index_user.asp")
end if

dim SceneID
SceneID=getForm("SceneID","get")
if not isnul(SceneID) then
	set rsObj=conn.exec("select * from {prefix}Scene where SceneID="&SceneID,"exe")
	if not rsObj.eof then 
		'wCookie"SceneID",id
		session("SceneMenu")=rsObj("SceneMenu")
	else		
		alertMsgAndGo "û���������","-1"
	end if
	rsObj.close : set rsObj=nothing
	response.Redirect("index_user.asp")
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=setting.siteTitle%>-��վ���ݹ���ϵͳ- Powered by AspCms2.0</title>
</head>
	<frameset rows="63,*" cols="*" frameborder="no" border="0" framespacing="0" >
		<frame src="top_user.asp" name="topFrame" id="topFrame" scrolling="no" noresize />
		<frameset cols="176,*" name="bodyFrame" id="bodyFrame" frameborder="no" border="0" framespacing="0"  >
			<frame src="menu_user.asp?id=1&src=right_user.asp" name="menu" id="menu" scrolling="auto" noresize />
			<frame src="right_user.asp" name="main" id="main" scrolling="auto" noresize />
		</frameset>
	</frameset>
	<noframes>
		<body>����������֧�ֿ�ܣ�</body>
	</noframes>
</html>




      
<%
dim rs,i,sql,menuid,SceneStr,firstID
'die isnul(rCookie("SceneMenu")) 
if not isnul(session("SceneMenu")) then 
	SceneStr=" and MenuID in("&session("SceneMenu")&")"
end if

if session("GroupMenu")="all" then
	sql="select MenuID, MenuName, (select Count(*) from {prefix}Menu as b where MenuStatus=1 and b.ParentID=a.MenuID ) from {prefix}Menu as a where MenuStatus=1 and ParentID=0  "&SceneStr&" order by MenuOrder"
else
	sql="select MenuID, MenuName, (select Count(*) from {prefix}Menu as b where MenuStatus=1 and b.ParentID=a.MenuID ) from {prefix}Menu as a where MenuStatus=1 and ParentID=0 and MenuID in("&session("GroupMenu")&") "&SceneStr&" order by MenuOrder"
end if
set rs=conn.exec(sql, "r1")
i=1
do while not rs.eof 
	if i=1 then firstID=rs(0)
	'echo "<DIV class=menu id=menu"&i&"><a target=""mainframe"" href=""mainframe.asp?menuid="& rs(0)&""">"& rs(1)&"</a></DIV>"
	menuid=rs(0)
	echo "<DIV class=menu id=menu"&i&" onclick=""document.getElementById('mainframe').src='mainframe.asp?menuid="&rs(0)&"'"">"& rs(1)&"</DIV>"
	rs.moveNext
	i=i+1
loop
rs.close : set rs=nothing
%>

<DIV class=menu style="FLOAT: right" onClick="document.location.href ='login.asp?action=logout'">�˳�</DIV>
<DIV id="changepass" class=menu style="FLOAT: right" onClick="document.getElementById('mainframe').src='mainframe.asp?menuid=<%=menuid%>&src=editpass.asp'">�޸�����</DIV>
<DIV id="index"class=menu style="FLOAT: right" onClick="document.getElementById('mainframe').src='mainframe.asp?menuid=<%=menuid%>&src=right.asp'">��̨��ҳ</DIV>
<DIV class=menu style="FLOAT: right" onClick="window.open('../')">��վ��ҳ</DIV>
      <div class="lnglist">
      	�л����ԣ�
      <%
	  	dim languageName
		set rs=conn.exec("select LanguageID, LanguageName, Alias from {prefix}Language where LanguageStatus=1 order by LanguageOrder ,LanguageID", "exe")
		do while not rs.eof 
			if cint(session("languageID"))=cint(rs(0)) then 
				echo "<a href=""index.asp?id="&rs(0)&"""><strong>"&rs(1)&"</strong></a> "	
				languageName=rs(1)
			else
				echo "<a href=""index.asp?id="&rs(0)&""">"&rs(1)&"</a> "
			end if
			rs.moveNext
		loop
	  %>
      <a href="#" id="closelng">[�ر�]</a></div>
      
<div class="menu changelng" style="float:right">�л�����(<span style="color:red;"><%=languageName%></span>)</div>
      </DIV>
      </DIV></TD></TR>
  <TR vAlign=top>
    <TD>
    <iframe src="mainframe.asp?menuid=<%=firstID%>&src=right.asp"  id=mainframe name=mainframe marginHeight=0  frameBorder=0 width="100%" scrolling=none height="100%"></iframe>
   </TD></TR></TBODY></TABLE>   
