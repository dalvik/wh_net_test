<!--#include file="inc/AspCms_SettingClass.asp" -->
<%CheckLogin()%>
<%
	dim sql,rs
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>menu</title>

<link href="css/css_menu.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/js/jquery.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>

<script language="javascript">


function getObject(objectId) {
 if(document.getElementById && document.getElementById(objectId)) {
 // W3C DOM
 return document.getElementById(objectId);
 }
 else if (document.all && document.all(objectId)) {
 // MSIE 4 DOM
 return document.all(objectId);
 }
 else if (document.layers && document.layers[objectId]) {
 // NN 4 DOM.. note: this won't find nested layers
 return document.layers[objectId];
 }
 else {
 return false;
 }
}

function showHide(objname){
    var obj = getObject(objname);
    if(obj.style.display == "none"){
		obj.style.display = "block";
	}else{
		obj.style.display = "none";
	}
}
</script>
<base target="main" />
<body>

<div class="infobox">
	<p>Powered by <A class=txt_C1  href="http://www.aspcms.com/" target=_blank>ASPCMS!</A></p>
	<p>&copy;2006-2011, <A class=txt_C1  href="http://www.4i.com.cn/" target=_blank>ChanCoo</A> Inc.</p>
</div>
<div class="menu">
<!-- Item 1 Strat -->
<dl>

<dt><a href="###" onclick="showHide('items0');" target="_self"><%=conn.exec("select MenuName from {prefix}Menu where MenuID="&getForm("id","get"),"r1")(0)%></a></dt>
    <dd id="items0" style="display:block;">
        <ul>
<%
dim menuid : menuid=getForm("id","get")
dim src : src=getForm("src","get")
dim i,firstLink,SceneStr
if not isnul(session("SceneMenu")) then 
	SceneStr=" and MenuID in("&session("SceneMenu")&")"
end if

if session("GroupMenu")="all" then
	sql="select MenuID, MenuName, MenuLink from {prefix}Menu where MenuStatus=1 and ParentID="&menuid&" "&SceneStr&"  order by MenuOrder"
else
	sql="select MenuID, MenuName, MenuLink from {prefix}Menu where MenuStatus=1 and ParentID="&menuid&" and MenuID in("&session("GroupMenu")&") "&SceneStr&"  order by MenuOrder"
end if

set rs=conn.exec(sql, "r1")

i=1
do while not rs.eof 
	if i=1 then firstLink=rs(2)
	echo "<li><a href='"&rs(2)&"' target='main'>"&rs(1)&"</a></li>"
	rs.moveNext
	i=i+1
loop
rs.close : set rs=nothing
if not isnul(src) then firstLink=src
%> 
        </ul>
    </dd>
</dl><!-- Item 1 End -->
<Script>top.document.getElementById("main").src = '<%=firstLink%>';</Script>
</div>

</body>
</html>
