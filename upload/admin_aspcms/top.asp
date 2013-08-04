<!--#include file="inc/AspCms_SettingClass.asp" -->
<%CheckLogin()%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>Top</title>
<link href="css/css_top.css" rel="stylesheet" type="text/css" />
<style type="text/css">
.sitelink table{float:left}
label{ float:left; 
height:20px; line-height:20px;/* display:block;width:65px;*/
border:#1684bf 1px solid;
border-right:0px;
}
.lang{
	border:#1684bf 1px solid;
	border-left:0px;
	float:left;			
	height:22px;
	line-height:22px
	/*BEHAVIOR: url('/css/selectBox.htc');*/
}
</style>

</head>
<body>
<div class="topnav">
	<div class="sitenav">
		<div class="welcome">
			你好：<span style="color:#F00"><%=session("GroupName")%></span><span class="username"><%=session("adminName")%></span>
            <a href="editPass.asp" target="main">修改密码</a>
			<a href="login.asp?action=logout" target="_parent">注销登录</a>
		</div>
		<div class="resize">
		<a href="javascript:ChangeMenu(-1)"><img src='images/frame-l.gif' border='0' alt="减小左框架"></a>
    <a href="javascript:ChangeMenu(0)"><img src='images/frame_on.gif' border='0' alt="隐藏/显示左框架"></a>
    <a href="javascript:ChangeMenu(1)" title="增大左框架"><img src='images/frame-r.gif' border='0' alt="增大左框架"></a>
    </div>
		<div class="sitelink">
		
			<a href="right.asp" target="main">后台首页</a>
			<a href="http://b.aspcms.com/list/?24_1.html" target="_blank">帮助中心</a>
			<a href="../" target="_blank">网站首页</a>
			<a href="home.asp" target="main">后台导航</a>
		</div>
	</div>
	<div class="leftnav">
		<ul>
			<li class="navleft"></li>
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
	menuid=rs(0)
	if i=1 then 
		echo "<li id='d"&i&"' class=""thisclass""><a href=""javascript:OpenMenu("&rs(0)&",'','',"&i&")"">"& rs(1)&"</a></li>"&vbcrlf
	else
		echo "<li id='d"&i&"'><a href=""javascript:OpenMenu("&rs(0)&",'','',"&i&")"">"& rs(1)&"</a></li>"&vbcrlf
	end if
	rs.moveNext
	i=i+1
loop
rs.close : set rs=nothing
%>

			<li class="navright"></li>
		</ul>
	</div>
</div>
</body>
</html>
<script type="text/javascript" defer="defer" src="js/menu.js"></script>