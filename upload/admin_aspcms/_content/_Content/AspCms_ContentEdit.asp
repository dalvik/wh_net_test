<!--#include file="AspCms_ContentFun.asp" -->
<%CheckAdmin("AspCms_Contentlist.asp?sortType="&sortType)%>
<%getContent%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>


<script language="javascript" type="text/javascript" src="../../js/getdate/WdatePicker.js"></script>
<script type="text/javascript" src="../../js/swfobject.js"></script>
<script src="../../js/iColorPicker.js" type="text/javascript"></script>
<style type="text/css">
<!--
#pw {}
#pw .imgDiv { float:left; width:130px; height:110px; padding-top:10px; padding-left:5px; background:#ccc;}
#pw img{ border:0px; width:120px; height:90px}
-->
</style>

<script src="../../script/admin.js" type="text/javascript"></script>
<script type="text/javascript" src="../../js/jquery.min.js"></script>
<script type="text/javascript" src="../../js/jquery.tinyTips.js"></script>
<link rel="stylesheet" type="text/css" media="screen" href="../../css/tinyTips.css" />
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
</head>

<body>

<FORM name="form" action="?action=edit<%="&sortType="&sortType&"&sortid="&sortid&"&keyword="&keyword&"&page="&page&"&psize="&psize&"&order="&order&"&ordsc="&ordsc&"&id="&contentid&""%>" method="post" >
<DIV class=formzone>
<DIV class=namezone>�޸�<%=sortTypeName%></DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>
    <TD align=middle width=153 height=30>����</TD>    
    <TD><%LoadSelect "SortID","Aspcms_Sort","SortName","SortID",SortID, 0," and( sortid in(select ParentID from {prefix}Sort where 1=1 and SortType="&sortType&" order by SortOrder) or SortType="&sortType&")","��ѡ�����"%></TD>
    <TD>������ɫ</TD>
    <TD><input id="TitleColor"  name="TitleColor" type="text" value="<%=TitleColor%>" style="WIDTH: 60px"  class="iColorPicker input" /></TD>
  </tr>
  <TR>						
    <TD align=middle width=153 height=30>����</TD>
    <TD colspan="3"><INPUT class="input" style="WIDTH: 350px" maxLength="200" name="Title" value="<%=Title%>"/> <FONT color=#ff0000>*</FONT></TD>
  </TR> 
  <%getSpec(Contentid)%>
  
 <%	if sortType="2" then%>

  <TR>
    <TD align=middle width=153 height=30>����</TD>
    <TD width="231"><input class="input" maxlength="200" style="WIDTH: 160px" name="Author" value="<%=Author%>"/></TD>
    <TD width="78">��Դ </TD>
    <TD width="395"><input class="input" maxlength="200" style="WIDTH: 160px" name="ContentSource" value="<%=ContentSource%>"/></TD>
  </TR>  <%end if%>
  
  <TR>
    <TD align=middle height=30>����<BR><BR>
      <a href="#" onClick="SetEditorPage('Content', '{aspcms:page}')"> �����ҳ<BR>
       {aspcms:page}</a><BR>
        </TD>
    <TD colspan="3">

<%Dim oFCKeditor:Set oFCKeditor = New FCKeditor:oFCKeditor.BasePath="../../editor/":oFCKeditor.ToolbarSet="AdminMode":oFCKeditor.Width="615":oFCKeditor.Height="300":oFCKeditor.Value=decodeHtml(Content):oFCKeditor.Create "Content"

'Default,AdminMode,Simple,UserMode,Basic
%></TD>
  </TR>
<%	if sortType<>"4" then%> 
    <tr>
    <TD align=middle height=30 valign="top"><%=sortTypeName%>����ͼ</td>
      <td colspan="3" ><input class="input" style="WIDTH: 450px" name="IndexImage" type="text" id="IndexImage" size="70"value="<%=IndexImage%>">
        <input type="hidden" name="ImagePath" id="ImagePath" value="<%=ImagePath%>">&nbsp;&nbsp;
        <!--<input type="checkbox" name="AutoRemote" value='1'> �Զ�����Զ��ͼƬ -->
        <br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��   
        <select name="ImageFileList" id="ImageFileList" onChange="IndexImage.value=this.value;">
        <option value=''>��ѡ����ҳ�Ƽ�ͼƬ</option>
        </select>
       <!-- <input type="button" name="selectpic" value="�����ϴ�ͼƬ��ѡ��" onClick="SelectPhoto()" class="button"> -->&nbsp;&nbsp;
 <br><div id="pw"><%
if not isnul(ImagePath) then 
	dim j,images
	images=split(ImagePath,"|")
	for j=0 to ubound(images)
		echo "<div id=""img"&j&""" class=""imgDiv""><a href=""javascript:SetEditorContents('Content','"&images(j)&"')""><img border=""0"" src="""&images(j)&"""></a><br><input type=""radio"" value="""&images(j)&""" onclick=""setIndexImage('"&images(j)&"')"" name=""IndexImageradio"">��Ϊ����ͼ <a href=""javascript:dropThisDiv('img"&j&"','"&images(j)&"')"">ɾ��</a></div>"
	next
end if
%></div>
<script type="text/javascript">
//parent.parent.FCKeditorAPI.GetInstance('content')
// ���ñ༭��������
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
    <TD align=middle height=30 valign="top">�ϴ�ͼƬ</td>
      <td colspan="3"><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../../editor/upload.asp?sortType=<%=sortType%>&stype=image&Tobj=content" allowTransparency="true"></iframe></td>
    </tr>
  <%end if%> 
  <% if sortType="8" then%>
  <TR>
    <TD align=middle height=30> �ۿ�Ȩ�� </TD>
    <TD colspan="3">
    <%=userGroupSelect("VideoGroupID", GroupID, 0)%>
    (ֻ������Ȩ�޵��û����������ļ�)
    </TD>
  </TR>	
  <TR>
    <TD align=middle height=30 valign="top">��Ƶ��ַ </TD>
    <TD  colspan="3" ><input class="input" maxlength="200" style="WIDTH: 450px" name="ContentSource" value="<%=ContentSource%>" id="videoPath"/></TD>
  </TR>
  <tr>
    <TD align=middle height=30 valign="top">�ϴ���Ƶ</td>
      <td colspan="3"><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../../editor/upload.asp?sortType=<%=sortType%>&stype=video&Tobj=content" allowTransparency="true"></iframe></td>
    </tr>
  <%end if%>
  
 <TR>
    <TD align=middle height=30> ����Ȩ�� </TD>
    <TD colspan="3">
    <%=userGroupSelect("DownGroupID", GroupID, 0)%>
    (ֻ������Ȩ�޵��û����������ļ�)
    </TD>
  </TR>	
  <TR>
    <TD align=middle width=153 height=30>���ص�ַ</TD>
    <TD colspan="3"><input class="input" maxlength="255" style="WIDTH: 400px" id="DownURL"  name="DownURL" value="<%=DownURL%>"/> <br>
    <iframe src="../../editor/upload.asp?sortType=<%=sortType%>&stype=file&Tobj=DownURL" scrolling="no" topmargin="0" width="100%" height="24" marginwidth="0" marginheight="0" frameborder="0" align="left"></iframe>    </TD>
  </TR>
  <TR>
    <TD align=center width=153 height=30 valign="top" colspan="2">
   <span onClick="showHide('moreform');imgnoneHide('himg');" style="cursor:pointer"><img id="himg" src="../../images/ico_5.gif"><strong>����߼�����Ŀ</strong></span>
</TD>
 <TD width="100%" colspan="2">
 </TR>
   <TR id="moreform" style="display:none">
    <TD width="100%" colspan="4">
    <table cellSpacing=0 cellPadding=0 width="100%" align=center border=0> 
  <TR>
    <TD align=middle height=30> ���Ȩ�� </TD>
    <TD colspan="3">
    <%=userGroupSelect("GroupID", GroupID, 0)%>
        <input type="radio" name="Exclusive" value=">=" <% if Exclusive=">=" then echo "checked=""checked""" end if%> />
        ����
        <input type="radio" name="Exclusive" value="=" <% if Exclusive="=" then echo "checked=""checked""" end if%> /> 
        ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��   <img src="../../images/help.gif" class="tTip" title="�÷���ķ���Ȩ�޵��趨<br>����ΪȨ�޴��ڸ��û��鶼���Է���<br>
ר��Ϊֻ���趨�ĸ�����Է���"></TD>
  </TR>	
  <TR>
    <TD align=middle height=30>�����Ǽ� </TD>
    <TD>
    <select name="star">
        <option value="1"<%If Star= 1 Then echo" selected"%>>��</option>
        <option value="2"<%If Star= 2 Then echo" selected"%>>���</option>
        <option value="3"<%If Star= 3 Then echo" selected"%>>����</option>
        <option value="4"<%If Star= 4 Then echo" selected"%>>�����</option>
        <option value="5"<%If Star= 5 Then echo" selected"%>>������</option>
        <option value="6"<%If Star= 6 Then echo" selected"%>>�������</option>
        <option value="7"<%If Star= 7 Then echo" selected"%>>��������</option>
     </select> <img src="../../images/help.gif" class="tTip" title="�Ǽ����࣬ʹ��ʱ����ģ���е�������Ǽ��ķ����ǩ����">
    </TD>
    <TD>�������</TD>
    <TD><input class="input" maxlength="6" value="<%=Visits%>" style="WIDTH: 60px" name="Visits"/></TD>
  </TR>
  <TR>
    <TD align=middle width=153 height=30>����ʱ��</TD>
    <TD><input class="input" style="WIDTH: 120px"  onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})"  name="AddTime" value="<%=AddTime%>"/></TD>
    <TD>���� </TD>
    <TD><input class="input" maxlength="6" style="WIDTH: 60px" name="ContentOrder" value="<%=ContentOrder%>"/> <img src="../../images/help.gif" class="tTip" title="��������Ҫģ���е��������򷽷���order����ID��������������������ܽ�ʧЧ"></TD>
  </TR>
  <TR>
    <TD align=middle height=30>��ʱ����</TD>
    <TD colspan="3"><INPUT type="checkbox" name="TimeStatus" value="1" <% if TimeStatus then echo"checked=""checked"""%>/>
�Ƿ�ʱ��������ʱ����ʱ�� 
  <input class="input" value="<%=Timeing%>" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" style="WIDTH: 120px" name="Timeing"/></TD>
  </TR>
  <TR>
    <TD align=middle height=30>����ѡ�� </TD>
    <TD colspan="3"><INPUT type="checkbox" name="IsTop" <% if IsTop then echo"checked=""checked"""%> value="1"/>
�ö�
<INPUT type="checkbox" name="IsRecommend" <% if IsRecommend then echo"checked=""checked"""%> value="1"/>
�Ƽ� 
<INPUT type="checkbox" name="IsFeatured" <% if IsFeatured then echo"checked=""checked"""%> value="1"/>
�ر��Ƽ� 
<INPUT type="checkbox" name="IsImageNews" <% if IsImageNews then echo"checked=""checked"""%> value="1"/>
ͼƬ����
<INPUT type="checkbox" name="IsNoComment" <% if IsNoComment then echo"checked=""checked"""%> value="1"/>
��ֹ��������
<INPUT type="checkbox" checked="checked" name="ContentStatus" <% if ContentStatus then echo"checked=""checked"""%> value="1"/>
����  <img src="../../images/help.gif" class="tTip" title="��ѡ����ģ���е���Ϊ׼��������Ӧ���ò��ֹ��ܽ�ʧЧ">
</TD>
  </TR>
  <TR>
    <TD align=middle height=30>�Զ����ļ��� </TD>
    <TD colspan="3"><input class="input" maxlength="200" style="WIDTH: 120px" name="PageFileName" value="<%=PageFileName%>"/>  
      <%=FileExt%>
      <INPUT type="hidden" name="IsGenerated" value="1" /> <img src="../../images/help.gif" class="tTip" title="��̬ģʽ�£���ǰ���µ������ļ�����������<br>����������Է����ͳһ��������Ϊ׼"></TD>
  </TR>
  <TR>
    <TD align=middle height=30>TAG </TD>
    <TD colspan="3"> <input class="input" maxlength="200" style="WIDTH: 450px" name="ContentTag" value="<%=ContentTag%>"/>���TAG���ö��Ÿ���!    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>ҳ��ؼ���</TD>
    <TD colspan="3"> <input class="input" maxlength="200" style="WIDTH: 450px"  name="PageKeywords" value="<%=PageKeywords%>"/>    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>ҳ������</TD>
    <TD colspan="3">  <TEXTAREA class="textarea" style="WIDTH: 450px" name="PageDesc" rows="3"><%=PageDesc%></TEXTAREA>    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>ת�����ӵ�ַ  </TD>
    <TD colspan="3"><input class="input" maxlength="200" style="WIDTH: 200px" name="OutLink" value="<%=OutLink%>"/>
      <INPUT type="checkbox" name="IsOutLink" value="1"<% if IsOutLink then echo"checked=""checked"""%> />
      ת������</TD>
 
  </TR>
  </table>
</TD></TR>
  
    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<input type="hidden" maxlength="200"  name="ContentID" value="<%=ContentID%>"/> 
<input type="hidden" maxlength="200"  name="sortType" value="<%=sortType%>"/> 
<INPUT class="button" type="submit" value="����" />
<INPUT class="button" type="button" value="����" onClick="history.go(-1)"/> 
</DIV></DIV></FORM>
</body>
</html>