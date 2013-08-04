<!--#include file="inc/AspCms_SettingClass.asp" -->
<%CheckLogin()%>
<%
sub menulist(ParentID)
	Dim rs :set rs =Conn.Exec ("select *,(select count(*) from {prefix}Sort where ParentID=t.SortID) as c from AspCms_Sort t where LanguageID="&session("languageID")&" and ParentID="&ParentID&" and sorttype<>7 and sortstatus=1 order by Sortorder,Sortorder ","r1")
	
	dim sortTypenames : sortTypenames=split(sortTypes,",")
	IF rs.eof or rs.bof Then
		echo "<tr bgcolor=""#ffffff"" align=""center""><td colspan=""8"">还没有记录</td></tr>"
	Else
		dim selecti:selecti=0
		Do While not rs.eof 
			selecti=selecti+1
			dim s,sstr
			if cint(rs("SortLevel"))=1 and rs("c")>0 then
				echo "<dl><dt><a href=""###"" onclick=""showHide('items"&selecti&"');"" target=""_self"">"&rs("SortName")&"</a></dt>"&vbcrlf
    			echo	"<dd id=""items"&selecti&""" style=""display:none;"">"&vbcrlf
        		echo	"<ul>"&vbcrlf
					if rs("c")>0 then menulist(rs("SortID"))
				echo    "</ul>"&vbcrlf
				echo		"</dd>"&vbcrlf
				echo		"</dl>"&vbcrlf
			elseif cint(rs("SortLevel"))=1 and rs("c")=0 and  rs("SortType")<>1 then
				echo "<dl><dt><a href=""###"" onclick=""showHide('items"&selecti&"');"" target=""_self"">"&rs("SortName")&"</a></dt>"&vbcrlf
    			echo	"<dd id=""items"&selecti&""" style=""display:none;"">"&vbcrlf
        		echo	"<ul>"&vbcrlf
				echo "<li><a href='_content/_Content/AspCms_ContentList_user.asp?sorttype="&rs("SortType")&"&SortID="&rs("SortID")&"' target='main'>"&rs("SortName")&"</a> | <a href='_content/_Content/AspCms_ContentAdd.asp?sortType="& rs("SortType")&"&sortid="& rs("SortID")&"' target='main'>添加</a></li>" & vbcr		
				echo    "</ul>"&vbcrlf
				echo		"</dd>"&vbcrlf
				echo		"</dl>"&vbcrlf	
			
			elseif cint(rs("SortLevel"))=1 and rs("SortType")=1  and rs("c")=0 then
echo "<dl><dt><a href=""###"" onclick=""showHide('items"&selecti&"');"" target=""_self"">"&rs("SortName")&"</a></dt>"&vbcrlf
    			echo	"<dd id=""items"&selecti&""" style=""display:none;"">"&vbcrlf
        		echo	"<ul>"&vbcrlf
				echo "<li><a href='_content/_About/AspCms_AboutEdit_user.asp?id="&rs("SortID")&"' target='main'>"&rs("SortName")&"</a></li>" & vbcr		
				echo    	"</ul>"&vbcrlf
				echo		"</dd>"&vbcrlf
				echo		"</dl>"&vbcrlf
			elseif cint(rs("SortLevel"))<>1 and rs("SortType")=1  and rs("c")>0 then
echo "<dl><dt><a href=""###"" onclick=""showHide('items"&selecti&"');"" target=""_self"">"&rs("SortName")&"</a></dt>"&vbcrlf
    			echo	"<dd id=""items"&selecti&""" style=""display:none;"">"&vbcrlf
        		echo	"<ul>"&vbcrlf
				if rs("c")>0 then menulist(rs("SortID"))
				echo    	"</ul>"&vbcrlf
				echo		"</dd>"&vbcrlf
				echo		"</dl>"&vbcrlf
			elseif cint(rs("SortLevel"))<>1 and rs("SortType")=1 and rs("c")=0 then
			echo "<li><a href='_content/_About/AspCms_AboutEdit_user.asp?id="&rs("SortID")&"' target='main'>"&rs("SortName")&"</a></li>" & vbcr	
			else
				echo "<li><a href='_content/_Content/AspCms_ContentList_user.asp?sorttype="&rs("SortType")&"&SortID="&rs("SortID")&"' target='main'>"&rs("SortName")&"</a> | <a href='_content/_Content/AspCms_ContentAdd.asp?sortType="& rs("SortType")&"&sortid="& rs("SortID")&"' target='main'>添加</a></li>" & vbcr		
			end if
			
			rs.MoveNext
		Loop
		rs.close : set rs = nothing
	End If
end sub
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>menu</title>

<link href="css/css_menu_user.css" rel="stylesheet" type="text/css" />
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



		<%menulist(0)%>

     <!-- Item 1 End -->
<!-- Item 2 Strat -->
<dl>

<dt><a href="###" onclick="showHide('itemsx');" target="_self">其它管理</a></dt>
    <dd id="itemsx" style="display:none;">
        <ul>
<li><a href='_content/_Comments/AspCms_Message.asp' target='main'>留言管理</a></li>
<li><a href='_content/_Order/AspCms_Order.asp' target='main'>订单管理</a></li>
<li><a href='_content/_Apply/AspCms_Apply.asp' target='main'>应聘管理</a></li>
<li><a href='_content/_Comments/AspCms_Comments.asp' target='main'>评论管理</a></li>
<li><a href='_adv/AspCms_Advlist.asp' target='main'>广告管理</a></li>
<li><a href='_seo/AspCms_MakeHtml.asp?actType=html' target='main'>静态生成</a></li>

        </ul>
    </dd>
</dl><!-- Item 2 End -->

<Script>top.document.getElementById("main").src = '<%=firstLink%>';</Script>
</div>

</body>
</html>
