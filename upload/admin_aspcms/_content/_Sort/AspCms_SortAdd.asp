<!--#include file="AspCms_SortFun.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>
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
<link rel="stylesheet" type="text/css" media="screen" href="../../css/tinyTips.css" />
<style type="text/css">
<!--
#pw {}
#pw .imgDiv { float:left; width:130px; height:110px; padding-top:10px; padding-left:5px; background:#ccc;}
#pw img{ border:0px; width:120px; height:90px}
-->
</style>
</HEAD>
<BODY>
<FORM name="form" action="?action=add" method="post" >
<DIV class=formzone>
<DIV class=namezone>��ӷ���</DIV>
<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>						
    <TD align=middle width=100 height=30>��������</TD>
    <TD><INPUT class="input" style="WIDTH: 200px" maxLength="200" name="SortName"/> <FONT color=#ff0000>*</FONT></TD>
  </TR>
  <TR>
    <TD align=middle width=100 height=30>��������</TD>    
    <TD><%LoadSelect "ParentID","Aspcms_Sort","SortName","SortID",getForm("id","get"), 0,"order by SortOrder","������Ŀ"%></TD>
</tr>
  <TR>
    <TD align=middle width=100 height=30>��������</TD>    
    <TD><%=makeSortTypeSelect("sortType", 2, "onChange=""setInput(sortType.value)""")%><img src="../../images/help.gif" class="tTip" title="����÷���Ϊ����<br>��ѡ���븸����ͬ�ķ�������">
<Script type="text/javascript">
function setInput(t)
{
//'��ƪ1,����2,��Ʒ3,����4,��Ƹ5,���6,����7
//alert("AA");
 document.getElementById("hid1").style.display="";
 document.getElementById("hid2").style.display="";
 document.getElementById("hid3").style.display="";
 document.getElementById("hid4").style.display="";
 document.getElementById("hid5").style.display="";
 document.getElementById("hid6").style.display="";
 document.getElementById("hid7").style.display="";
 document.getElementById("hid8").style.display="";
 document.getElementById("hid9").style.display="";
switch(t)
   {
   case "1":
     document.getElementById("SortTemplate").value="about.html";
     document.getElementById("ContentTemplate").value="";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/about/";
     document.getElementById("ContentFolder").value="";
     document.getElementById("SortFileName").value="about-{sortid}";
     document.getElementById("ContentFileName").value="";
     document.getElementById("hid2").style.display="none";
     document.getElementById("hid4").style.display="none";
     document.getElementById("hid6").style.display="none";
	 document.getElementById("hid9").style.display="none";    
     break
   case "2":
     document.getElementById("SortTemplate").value="newslist.html";
     document.getElementById("ContentTemplate").value="news.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/newslist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/news/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";    
	 document.getElementById("hid8").style.display="none";  
	 document.getElementById("hid9").style.display="none";     
     break
   case "3":
     document.getElementById("SortTemplate").value="productlist.html";
     document.getElementById("ContentTemplate").value="product.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/productlist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/product/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";   
	 document.getElementById("hid8").style.display="none";     
	 document.getElementById("hid9").style.display="none";     
     break
   case "4":  
     document.getElementById("SortTemplate").value="downlist.html";
     document.getElementById("ContentTemplate").value="down.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/downlist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/down/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}"; 
	 document.getElementById("hid8").style.display="none";  
	 document.getElementById("hid9").style.display="none";      
     break
   case "5":
     document.getElementById("SortTemplate").value="joblist.html";
     document.getElementById("ContentTemplate").value="job.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/jobtlist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/job/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";  
	 document.getElementById("hid8").style.display="none";
	 document.getElementById("hid9").style.display="none";          
     break
   case "6":
     document.getElementById("SortTemplate").value="albumlist.html";
     document.getElementById("ContentTemplate").value="album.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/albumtlist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/album/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";
	 document.getElementById("hid8").style.display="none";
	 document.getElementById("hid9").style.display="none";            
     break
	case "8":
     document.getElementById("SortTemplate").value="videolist.html";
     document.getElementById("ContentTemplate").value="video.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/videolist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/video/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";    
	 document.getElementById("hid8").style.display="none";  
	 document.getElementById("hid9").style.display="none";      
     break
   default:
     document.getElementById("hid1").style.display="none";
     document.getElementById("hid2").style.display="none";
     document.getElementById("hid3").style.display="none";
     document.getElementById("hid4").style.display="none";
     document.getElementById("hid5").style.display="none";
     document.getElementById("hid6").style.display="none";
     document.getElementById("hid7").style.display="none";
     document.getElementById("hid8").style.display="none";
   }
}

</script>    </TD>
</tr>
  <TR id="hid9">
    <TD align=middle width=100 height=30>����</TD>
    <TD><input class="input" maxlength="200" style="WIDTH: 300px" name="SortURL"/> ע����������"/gbook/"</TD></TR>
  <TR>
    <TD align=middle width=100 height=30>����</TD>
    <TD><input class="input" maxlength="6" value="99" style="WIDTH: 60px" name="SortOrder"/> </TD></TR>
  <TR>
    <TD align=middle height=30>״̬</TD>
    <TD><INPUT class="checkbox" type="checkbox" checked="checked" name="SortStatus" value="1"/>    </TD>
  </TR>
  <TR id="hid1">
    <TD align=middle height=30>�б�ģ��</TD>
    <TD> <%tempList "SortTemplate","newslist.html" %>   <img src="../../images/help.gif" class="tTip" title="�÷���ʹ�õ��б��ģ����ļ���"></TD>
  </TR>
  <TR id="hid2">
    <TD align=middle height=30>����ҳģ��</TD>
    <TD> <%tempList "ContentTemplate","news.html" %>  <img src="../../images/help.gif" class="tTip" title="�÷����������ϸҳʹ�õ�ģ����ļ���">    </TD>
  </TR>
  <TR id="hid3">
    <TD align=middle height=30>�б�����Ŀ¼</TD>
    <TD> <input class="input" maxlength="200" style="WIDTH: 200px" name="SortFolder" id="SortFolder" value="{sitepath}<%=setting.languagePath&htmlDir%>/newslist/"/>
      <img src="../../images/help.gif" class="tTip" title="��̬ģʽ�£��÷������ɵľ�̬�б�ҳ���ŵ�λ��<br>{sitepath}Ϊ��վ��Ŀ¼"> </TD>
  </TR>
  <TR id="hid4">
    <TD align=middle height=30>��������Ŀ¼</TD>
    <TD> <input class="input" maxlength="200" style="WIDTH: 200px" name="ContentFolder" id="ContentFolder" value="{sitepath}<%=setting.languagePath&htmlDir%>/news/"/>
      <img src="../../images/help.gif" class="tTip" title="��̬ģʽ�£��÷������ɵľ�̬��ϸҳ���ŵ�λ��<br>{sitepath}Ϊ��վ��Ŀ¼"> </TD>
  </TR>
  <TR id="hid5">
    <TD align=middle height=30>�б��ļ�����</TD>
    <TD> <input class="input" maxlength="200" style="WIDTH: 200px" name="SortFileName" id="SortFileName" value="list-{sortid}-{page}"/>
    <%=FileExt%>     <img src="../../images/help.gif" class="tTip" title="��̬ģʽ�¸÷����б������ļ�����������<br>{sortid}Ϊ�÷����IDֵ<br>{page}Ϊҳ��
"> </TD>
  </TR>
  <TR  id="hid6">
    <TD align=middle height=30>����ҳ�ļ�����</TD>
    <TD> <input class="input" maxlength="200" style="WIDTH: 200px" name="ContentFileName" id="ContentFileName" value="{y}-{m}-{d}/{id}"/>
      <%=FileExt%>     <img src="../../images/help.gif" class="tTip" title="��̬ģʽ�¸÷�����ϸҳ�����ļ�����������<br>
{y}��<br>{m}��<br>{d}��<br>{id}���±��"> </TD>
  </TR>
  <TR id="hid7">
    <TD align=middle height=30> ���Ȩ�� </TD>
    <TD>
    <%=userGroupSelect("GroupID", 0, 0)%>
        <input type="radio" name="Exclusive" value=">=" checked="checked" />
        ����
        <input type="radio" name="Exclusive" value="=" /> 
        ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��    <img src="../../images/help.gif" class="tTip" title="�÷���ķ���Ȩ�޵��趨<br>����ΪȨ�޴��ڸ��û��鶼���Է���<br>
ר��Ϊֻ���趨�ĸ�����Է���"></TD>
  </TR>
  <TR id="hid8">
    <TD align="middle" height="30">
    ����<BR><BR>
       <a href="#" onClick="SetEditorPage('Content', '{aspcms:page}')"> �����ҳ<BR>
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
</script></TD>
  </TR>
    <tr>
    <TD align=middle height=30 valign="top">�ϴ�ͼƬ</td>
      <td colspan="3"><input class="input" maxlength="255" style="WIDTH: 400px" id="indeximage" name="indeximage" /> <img src="../../images/help.gif" class="tTip" title="��ĿͼƬ������ǰ̨����"><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../../editor/upload.asp?sortType=3&stype=file&Tobj=indeximage" allowTransparency="true"></iframe></td>
    </tr>
    	<tr>
    <TD align=middle height=30 valign="top">ICO����ͼ</td>
      <td colspan="3" ><input class="input" style="WIDTH: 450px" name="IcoImage" type="text" id="IcoImage" size="70"value="<%=IcoImage%>">
        <input type="hidden" name="IcoPath" id="ImagePath" value="<%=IcoPath%>">&nbsp;&nbsp;
        <!--<input type="checkbox" name="AutoRemote" value='1'> �Զ�����Զ��ͼƬ -->
        <br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��   
        <select name="ImageFileList" id="ImageFileList" onChange="IcoImage.value=this.value;">
        <option value=''>��ѡ����ҳ�Ƽ�ͼƬ</option>
        </select>
       <!-- <input type="button" name="selectpic" value="�����ϴ�ͼƬ��ѡ��" onClick="SelectPhoto()" class="button"> -->&nbsp;&nbsp;
 <br><div id="pw"><%
if not isnul(IcoPath) then 
	dim j,images
	images=split(IcoPath,"|")
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
var str =document.getElementById("IcoPath").value;
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
document.getElementById("IcoPath").value=nstr;

doChange(document.getElementById("IcoPath"),document.getElementById("ImageFileList"))
}

function setIndexImage(t)
{
document.getElementById("IcoImage").value=t

doChange(document.getElementById("IcoPath"),document.getElementById("ImageFileList"))
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
doChange(document.getElementById("IcoPath"),document.getElementById("ImageFileList"))
</script>
        </td>
    </tr>
    <tr>
    <TD align=middle height=30 valign="top">�ϴ�ͼƬ</td>
      <td colspan="3"><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../../editor/upload.asp?sortType=16&stype=image&Tobj=content" allowTransparency="true"></iframe></td>
    </tr>
  <TR>
    <TD align=middle height=30>ҳ�����</TD>
    <TD> <input class="input" maxlength="200" style="WIDTH: 450px"  name="PageTitle"/> </TD>
  </TR>
  <TR>
    <TD align=middle height=30>ҳ��ؼ���</TD>
    <TD> <input class="input" maxlength="200" style="WIDTH: 450px"  name="PageKeywords"/> </TD>
  </TR>
  <TR>
    <TD align=middle height=30>ҳ������</TD>
    <TD>  <TEXTAREA class="textarea" style="WIDTH: 450px" name="PageDesc" rows="3"></TEXTAREA>    </TD>
  </TR>
    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<INPUT class="button" type="submit" value="���" />
<INPUT class="button" type="button" value="����" onClick="history.go(-1)"/> 

<INPUT onClick="location.href='<%=getPageName()%>'" type="button" value="ˢ��" class="button" /> 
</DIV></DIV></FORM>

</BODY></HTML>