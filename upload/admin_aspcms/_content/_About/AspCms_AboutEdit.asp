<!--#include file="AspCms_AboutFun.asp" -->
<%getSort%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--
#pw {}
#pw .imgDiv { float:left; width:130px; height:110px; padding-top:10px; padding-left:5px; background:#ccc;}
#pw img{ border:0px; width:120px; height:90px}
-->
</style>
</HEAD>
<BODY>
<FORM name="form" action="?action=edit" method="post" >
<DIV class=formzone>
<DIV class=namezone>修改单篇文章</DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>						
    <TD align=middle width=100 height=30>分类名称</TD>
    <TD><INPUT class="input" style="WIDTH: 200px" maxLength="200" name="SortName" value="<%=SortName%>"/> <FONT color=#ff0000>*</FONT> </TD>
  </TR>
 
  <TR id="hid7">
    <TD align=middle height=30> 浏览权限 </TD>
    <TD>
    <%=userGroupSelect("GroupID", GroupID, 0)%>
        <input type="radio" name="Exclusive" value=">=" <% if Exclusive=">=" then echo "checked=""checked""" end if%> />
        隶属
        <input type="radio" name="Exclusive" value="=" <% if Exclusive="=" then echo "checked=""checked""" end if%> /> 
        专属（隶属：权限值≥可查看，专属：权限值＝可查看）    </TD>
  </TR>
  <TR id="hid8">
    <TD align="middle" height="30">
       	内容<BR><BR>
        <a href="#" onClick="SetEditorPage('Content', '{aspcms:page}')"> 插入分页<BR>
       {aspcms:page}</a></TD>
    <TD>
<%Dim oFCKeditor:Set oFCKeditor = New FCKeditor:oFCKeditor.BasePath="../../editor/":oFCKeditor.ToolbarSet="AdminMode":oFCKeditor.Width="615":oFCKeditor.Height="300":oFCKeditor.Value=decodeHtml(Content):oFCKeditor.Create "Content"
%>
<script type="text/javascript">

function SetEditorPage(EditorName, ContentStr) {
     var oEditor = FCKeditorAPI.GetInstance(EditorName) ;
	 oEditor.Focus();
	 //setTimeout(function() { oEditor.Focus(); }, 100);
     oEditor.InsertHtml(ContentStr);	
}
</script>
    </TD>
  </TR>
  <tr>
    <TD align=middle height=30 valign="top"><%=sortTypeName%>缩略图</td>
      <td colspan="3" ><input class="input" style="WIDTH: 450px" name="IndexImage" type="text" id="IndexImage" size="70"value="<%=IndexImage%>">
        <input type="hidden" name="ImagePath" id="ImagePath" value="<%=ImagePath%>">&nbsp;&nbsp;
        <!--<input type="checkbox" name="AutoRemote" value='1'> 自动保存远程图片 -->
        <br>直接从上传图片中选择：   
        <select name="ImageFileList" id="ImageFileList" onChange="IndexImage.value=this.value;">
        <option value=''>不选择首页推荐图片</option>
        </select>
       <!-- <input type="button" name="selectpic" value="从已上传图片中选择" onClick="SelectPhoto()" class="button"> -->&nbsp;&nbsp;
 <br><div id="pw"><%
if not isnul(ImagePath) then 
	dim j,images
	images=split(ImagePath,"|")
	for j=0 to ubound(images)
		echo "<div id=""img"&j&""" class=""imgDiv""><a href=""javascript:SetEditorContents('Content','"&images(j)&"')""><img border=""0"" src="""&images(j)&"""></a><br><input type=""radio"" value="""&images(j)&""" onclick=""setIndexImage('"&images(j)&"')"" name=""IndexImageradio"">设为缩略图 <a href=""javascript:dropThisDiv('img"&j&"','"&images(j)&"')"">删除</a></div>"
	next
end if
%></div>
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
doChange(document.getElementById("ImagePath"),document.getElementById("ImageFileList"))
</script>
        </td>
    </tr>
    <tr>
    <TD align=middle height=30 valign="top">上传图片</td>
      <td colspan="3"><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../../editor/upload.asp?sortType=10&stype=image&Tobj=content" allowTransparency="true"></iframe></td>
    </tr>
  <TR>
  <TR>
    <TD align=middle height=30>页面标题</TD>
    <TD> <input class="input" maxlength="200" style="WIDTH: 450px"  name="PageTitle" value="<%=PageTitle%>"/>  
    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>页面关键词</TD>
    <TD> <input class="input" maxlength="200" style="WIDTH: 450px"  name="PageKeywords" value="<%=PageKeywords%>"/>  
    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>页面描述</TD>
    <TD>  <TEXTAREA class="textarea" style="WIDTH: 450px" name="PageDesc" rows="3"><%=PageDesc%></TEXTAREA>
    </TD>
  </TR>

    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT type="hidden" name="SortID" value="<%=SortID%>" />
<INPUT class="button" type="submit" value="修改" />
<INPUT class="button" type="button" value="返回" onClick="history.go(-1)"/> 

<INPUT onClick="location.href='<%=getPageName()%>?id=<%=SortID%>'" type="button" value="刷新" class="button" /> 
</DIV></DIV></FORM>

</BODY></HTML>