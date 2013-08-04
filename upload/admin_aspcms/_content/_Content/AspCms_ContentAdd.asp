<!--#include file="AspCms_ContentFun.asp" -->
<%CheckAdmin("AspCms_ContentList.asp?sortType="&sortType)%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>
<script src="../../../js/jquery.js" type="text/javascript"></script>

<script language="javascript" type="text/javascript" src="../../js/getdate/WdatePicker.js"></script>
<script type="text/javascript" src="../../js/swfobject.js"></script>
<script src="../../js/iColorPicker.js" type="text/javascript"></script>
<script src="../../script/admin.js" type="text/javascript"></script>
<style type="text/css">
<!--
#pw {}
#pw .imgDiv { float:left; width:130px; height:110px; padding-top:10px; padding-left:5px; background:#ccc;}
#pw img{ border:0px; width:120px; height:90px}
-->
</style>
<script type="text/javascript" src="../../js/jquery.min.js"></script>
<script type="text/javascript" src="../../js/jquery.tinyTips.js"></script>
<script type="text/javascript">
		$(document).ready(function() {
			$('a.tTip').tinyTips('title');
			$('a.imgTip').tinyTips('title');
			$('img.tTip').tinyTips('title');
			$('h1.tagline').tinyTips('tinyTips are totally awesome!');
		});
</script>
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
		obj.style.display = "";
		
	}else{
		obj.style.display = "none";
	}
}
function imgnoneHide(objname){
    var obj = getObject(objname);
	
	if(obj.getAttribute("src",2)=='../../images/ico_4.gif')
	{	
		obj.src='../../images/ico_5.gif'
	}
	else
	{
		obj.src='../../images/ico_4.gif'
	}
}
</script>
<link rel="stylesheet" type="text/css" media="screen" href="../../css/tinyTips.css" />


</head>

<body>

<FORM name="form" action="?action=add<%="&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&"&id="&contentid&""%>" method="post" >
<DIV class=formzone>
<DIV class=namezone>添加<%=sortTypeName%></DIV>

<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>
    <TD align=middle width=153 height=30>分类</TD>    
    <TD><%LoadSelect "SortID","Aspcms_Sort","SortName","SortID",sortid,0," and( sortid in(select ParentID from {prefix}Sort where 1=1 and SortType="&sortType&" order by SortOrder) or SortType="&sortType&")","请选择分类"%></TD>
    <TD>标题颜色</TD>
    <TD><input id="TitleColor"  name="TitleColor" type="text" value="#000000" style="WIDTH: 60px"  class="iColorPicker input" /></TD>
  </tr>
  <TR>						
    <TD align=middle width=153 height=30>标题</TD>
    <TD colspan="3"><INPUT class="input" style="WIDTH: 350px" maxLength="200" name="Title"/> <FONT color=#ff0000>*</FONT></TD>
  </TR>
  <%getSpec(0)%>
  
 <%	if sortType="2" then%>
  <TR>
    <TD align=middle width=153 height=30>作者</TD>
    <TD width="231"><input class="input" maxlength="200" style="WIDTH: 160px" name="Author"/></TD>
    <TD width="78">来源 </TD>
    <TD width="395"><input class="input" maxlength="200" style="WIDTH: 160px" name="ContentSource" /></TD>
  </TR>
  <%end if%>
  <TR>
    <TD align=middle height=30>内容<BR><BR>
      <a href="#" onClick="SetEditorPage('Content', '{aspcms:page}')"> 插入分页<BR>
       {aspcms:page}</a></TD>
    <TD colspan="3">


<%Dim oFCKeditor:Set oFCKeditor = New FCKeditor:oFCKeditor.BasePath="../../editor/":oFCKeditor.ToolbarSet="AdminMode":oFCKeditor.Width="615":oFCKeditor.Height="300":oFCKeditor.Value=decodeHtml(Content):oFCKeditor.Create "Content"

'Default,AdminMode,Simple,UserMode,Basic
%>
</TD>
  </TR>
<%if sortType<>"4" then%> 
    <tr>
    <TD align=middle height=30 valign="top"><%=sortTypeName%>缩略图</td>
      <td colspan="3" ><input class="input" style="WIDTH: 450px" name="IndexImage" type="text" id="IndexImage" size="70" value="">
        <input type="hidden" name="ImagePath" id="ImagePath" value="">&nbsp;&nbsp;
  <!--      <input type="checkbox" name="AutoRemote" value='1'> 自动保存远程图片 -->
        <br>直接从上传图片中选择：   
        <select name="ImageFileList" id="ImageFileList" onChange="IndexImage.value=this.value;">
        <option value=''>不选择首页推荐图片</option>
        </select>
       <!-- <input type="button" name="selectpic" value="从已上传图片中选择" onClick="SelectPhoto()" class="button"> -->&nbsp;&nbsp;
       <br><div id="pw"></div>
<script type="text/javascript">
//parent.parent.FCKeditorAPI.GetInstance('content')
// 设置编辑器中内容
function SetEditorContents(EditorName, ContentStr) {
     var oEditor = FCKeditorAPI.GetInstance(EditorName) ;
	 oEditor.Focus();
	 //setTimeout(function() { oEditor.Focus(); }, 100);
     oEditor.InsertHtml("<img src="+"'"+ContentStr+"'"+"/>");	
}
function SetEditorPage(EditorName, ContentStr) {
     var oEditor = FCKeditorAPI.GetInstance(EditorName) ;
	 oEditor.Focus();
	 //setTimeout(function() { oEditor.Focus(); }, 100);
     oEditor.InsertHtml(ContentStr);	
}
</script>
<script type="text/javascript">
function dropThisDiv(t,p)
{
document.getElementById(t).style.display='none'
var str =document.getElementById("ImagePath").value;
var arr = str.split("|");
var nstr="";
for (var i=0; i<arr.length; i++)
{
	if(arr[i]!=p)
	{
		if (nstr!="")
		{
			nstr=nstr+"|";
		}		
		nstr=nstr+arr[i]
	}
}
document.getElementById("ImagePath").value=nstr;

doChange(document.getElementById("ImagePath"),document.getElementById("ImageFileList"))
}

function setIndexImage(t)
{
document.getElementById("IndexImage").value=t

doChange(document.getElementById("ImagePath"),document.getElementById("ImageFileList"))
}

function doChange(objText, objDrop){
if (!objDrop) return;
var str = objText.value;
var arr = str.split("|");
var nIndex = objDrop.selectedIndex;
objDrop.length=1;
for (var i=0; i<arr.length; i++){
objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
}
objDrop.selectedIndex = nIndex;
}
</script>
        </td>
    </tr>
    <tr>
    <TD align=middle height=30 valign="top">上传图片</td>
      <td colspan="3"><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../../editor/upload.asp?sortType=<%=sortType%>&stype=image&Tobj=content" allowTransparency="true"></iframe></td>
    </tr>
  <%end if%> 
  <%	if sortType="8" then%>
  <TR>
    <TD align=middle height=30> 观看权限 </TD>
    <TD colspan="3">
    <%=userGroupSelect("VideoGroupID", GroupID, 0)%>
    (只有满足权限的用户才能下载文件)
    </TD>
  </TR>	
  <TR>
    <TD align=middle height=30>来源</TD>
    <TD colspan="3"><input class="input" maxlength="200" style="WIDTH: 450px" name="ContentSource" id="videoPath"/></TD>
  </TR>
  <tr>
    <TD align=middle height=30 valign="top">上传视频</td>
      <td colspan="3"><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../../editor/upload.asp?sortType=<%=sortType%>&stype=video&Tobj=content" allowTransparency="true"></iframe></td>
    </tr>
  <%end if%>
 <TR>
    <TD align=middle height=30> 下载权限 </TD>
    <TD colspan="3">
    <%=userGroupSelect("DownGroupID", GroupID, 0)%>
    (只有满足权限的用户才能下载文件)
    </TD>
  </TR>	
  <TR>
    <TD align=middle width=153 height=30 valign="top">下载地址</TD>
    <TD colspan="3"><input class="input" maxlength="255" style="WIDTH: 400px" id="DownURL" name="DownURL"/>
</TD>
  </TR>
  <TR>
    <TD align=middle width=153 height=30 valign="top">上传文件</TD>
    <TD colspan="3">   <iframe name="image" frameborder="0" width='100%' height="40" scrolling="no"src="../../editor/upload.asp?sortType=<%=sortType%>&stype=file&Tobj=DownURL" allowTransparency="true"></iframe>
</TD>
  </TR>
  <TR>
    <TD align=center width=153 height=30 valign="top" colspan="2">
   <span onClick="showHide('moreform');imgnoneHide('himg');" style="cursor:pointer"><img id="himg" src="../../images/ico_5.gif"><strong>管理高级表单项目</strong></span>
</TD>
 <TD width="100%" colspan="2"></TD>
 </TR>
 
 
 
   <TR id="moreform" style="display:none;">
 <TD width="100%" colspan="4">
 
 <table cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <TR>
    <TD align=middle height=30 width="153"> 浏览权限 </TD>
    <TD colspan="3">
    <%=userGroupSelect("GroupID", 0, 0)%>
        <input type="radio" name="Exclusive" value=">=" checked="checked" />
        隶属
        <input type="radio" name="Exclusive" value="=" /> 
        专属（隶属：权限值≥可查看，专属：权限值＝可查看）    <img src="../../images/help.gif" class="tTip" title="该分类的访问权限的设定<br>隶属为权限大于该用户组都可以访问<br>
专属为只有设定的该组可以访问"></TD>
  </TR>	
  <TR>
    <TD align=middle height=30>星级 </TD>
    <TD>
    <select name="star">
			<option value="1">★</option>
			<option value="2">★★</option>
			<option value="3">★★★</option>
			<option value="4">★★★★</option>
			<option value="5">★★★★★</option>
			<option value="6">★★★★★★</option>
			<option value="7">★★★★★★★</option>
     </select>
    <img src="../../images/help.gif" class="tTip" title="星级分类，使用时请在模板中调用相关星级的分类标签内容">    </TD>
    <TD>浏览次数</TD>
    <TD><input class="input" maxlength="6" value="0" style="WIDTH: 60px" name="Visits"/></TD>
  </TR>
  <TR>
    <TD align=middle width=153 height=30>发布时间</TD>
    <TD><input class="input" value="<%=now%>" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" style="WIDTH: 120px" name="AddTime"/></TD>
    <TD>排序 </TD>
    <TD><input class="input" maxlength="6" value="0" style="WIDTH: 60px" name="ContentOrder"/>  <img src="../../images/help.gif" class="tTip" title="该排序需要模板中的新闻排序方法用order，按ID或其它方法排序该排序功能将失效"></TD>
  </TR>
  <TR>
    <TD align=middle height=30>定时发布</TD>
    <TD colspan="3"><INPUT type="checkbox" name="TimeStatus" value="1"/>
是否定时发布，定时发布时间 
  <input class="input" value="<%=now%>" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" style="WIDTH: 120px" name="Timeing"/></TD>
  </TR>
  <TR>
    <TD align=middle height=30>其它选项 </TD>
    <TD colspan="3"><INPUT type="checkbox" name="IsTop" value="1"/>
置顶
<INPUT type="checkbox" name="IsRecommend" value="1"/>
推荐 
<INPUT type="checkbox" name="IsFeatured" <% if IsFeatured then echo"checked=""checked"""%> value="1"/>
特别推荐 
<INPUT type="checkbox" name="IsImageNews" value="1"/>
图片新闻
<INPUT type="checkbox" name="IsNoComment" value="1"/>
禁止发表评论
<INPUT type="checkbox" checked="checked" name="ContentStatus" value="1"/>
发布  <img src="../../images/help.gif" class="tTip" title="该选项以模板中调用为准，如无相应调用部分功能将失效"></TD>
  </TR>
  <TR>
    <TD align=middle height=30>自定义文件名 </TD>
    <TD colspan="3"><input class="input" maxlength="200" style="WIDTH: 120px" name="PageFileName" />  
      <%=FileExt%>
      <INPUT type="hidden" name="IsGenerated" value="1"/>
      <img src="../../images/help.gif" class="tTip" title="静态模式下，当前文章的生成文件的命名规则<br>如果留空则以分类的统一命名规则为准"></TD>
  </TR>
  <TR>
    <TD align=middle height=30>TAG </TD>
    <TD colspan="3"> <input class="input" maxlength="200" style="WIDTH: 450px" name="ContentTag"/> 多个TAG请用逗号隔开!</TD>
  </TR>
  <TR>
    <TD align=middle height=30>页面关键词</TD>
    <TD colspan="3"> <input class="input" maxlength="200" style="WIDTH: 450px"  name="PageKeywords"/>    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>页面描述</TD>
    <TD colspan="3">  <TEXTAREA class="textarea" style="WIDTH: 450px" name="PageDesc" rows="3"></TEXTAREA>    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>转向链接地址  </TD>
    <TD colspan="3"><input class="input" maxlength="200" style="WIDTH: 200px" name="OutLink"  />
      <INPUT type="checkbox" name="IsOutLink" value="1" />
      转向链接</TD>
  </TR>
</table>
 
 </TD>
 </TR>

    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<input type="hidden" maxlength="200"  name="sortType" value="<%=sortType%>"/> 
<INPUT class="button" type="submit" value="添加" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 
<INPUT onClick="location.href='<%=getPageName()%>?sortType=<%=sortType%>'" type="button" value="刷新" class="button" /> 
</DIV></DIV></FORM>
</body>
</html>
