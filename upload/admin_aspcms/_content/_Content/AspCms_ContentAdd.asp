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
<DIV class=namezone>���<%=sortTypeName%></DIV>

<DIV class=tablezone>
<DIV class=noticediv id=notice></DIV>
<TABLE cellSpacing=0 cellPadding=2 width="100%" align=center border=0>
  <TBODY>
  <TR>
    <TD align=middle width=153 height=30>����</TD>    
    <TD><%LoadSelect "SortID","Aspcms_Sort","SortName","SortID",sortid,0," and( sortid in(select ParentID from {prefix}Sort where 1=1 and SortType="&sortType&" order by SortOrder) or SortType="&sortType&")","��ѡ�����"%></TD>
    <TD>������ɫ</TD>
    <TD><input id="TitleColor"  name="TitleColor" type="text" value="#000000" style="WIDTH: 60px"  class="iColorPicker input" /></TD>
  </tr>
  <TR>						
    <TD align=middle width=153 height=30>����</TD>
    <TD colspan="3"><INPUT class="input" style="WIDTH: 350px" maxLength="200" name="Title"/> <FONT color=#ff0000>*</FONT></TD>
  </TR>
  <%getSpec(0)%>
  
 <%	if sortType="2" then%>
  <TR>
    <TD align=middle width=153 height=30>����</TD>
    <TD width="231"><input class="input" maxlength="200" style="WIDTH: 160px" name="Author"/></TD>
    <TD width="78">��Դ </TD>
    <TD width="395"><input class="input" maxlength="200" style="WIDTH: 160px" name="ContentSource" /></TD>
  </TR>
  <%end if%>
  <TR>
    <TD align=middle height=30>����<BR><BR>
      <a href="#" onClick="SetEditorPage('Content', '{aspcms:page}')"> �����ҳ<BR>
       {aspcms:page}</a></TD>
    <TD colspan="3">


<%Dim oFCKeditor:Set oFCKeditor = New FCKeditor:oFCKeditor.BasePath="../../editor/":oFCKeditor.ToolbarSet="AdminMode":oFCKeditor.Width="615":oFCKeditor.Height="300":oFCKeditor.Value=decodeHtml(Content):oFCKeditor.Create "Content"

'Default,AdminMode,Simple,UserMode,Basic
%>
</TD>
  </TR>
<%if sortType<>"4" then%> 
    <tr>
    <TD align=middle height=30 valign="top"><%=sortTypeName%>����ͼ</td>
      <td colspan="3" ><input class="input" style="WIDTH: 450px" name="IndexImage" type="text" id="IndexImage" size="70" value="">
        <input type="hidden" name="ImagePath" id="ImagePath" value="">&nbsp;&nbsp;
  <!--      <input type="checkbox" name="AutoRemote" value='1'> �Զ�����Զ��ͼƬ -->
        <br>ֱ�Ӵ��ϴ�ͼƬ��ѡ��   
        <select name="ImageFileList" id="ImageFileList" onChange="IndexImage.value=this.value;">
        <option value=''>��ѡ����ҳ�Ƽ�ͼƬ</option>
        </select>
       <!-- <input type="button" name="selectpic" value="�����ϴ�ͼƬ��ѡ��" onClick="SelectPhoto()" class="button"> -->&nbsp;&nbsp;
       <br><div id="pw"></div>
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
</script>
        </td>
    </tr>
    <tr>
    <TD align=middle height=30 valign="top">�ϴ�ͼƬ</td>
      <td colspan="3"><iframe name="image" frameborder="0" width='100%' height="40" scrolling="no" src="../../editor/upload.asp?sortType=<%=sortType%>&stype=image&Tobj=content" allowTransparency="true"></iframe></td>
    </tr>
  <%end if%> 
  <%	if sortType="8" then%>
  <TR>
    <TD align=middle height=30> �ۿ�Ȩ�� </TD>
    <TD colspan="3">
    <%=userGroupSelect("VideoGroupID", GroupID, 0)%>
    (ֻ������Ȩ�޵��û����������ļ�)
    </TD>
  </TR>	
  <TR>
    <TD align=middle height=30>��Դ</TD>
    <TD colspan="3"><input class="input" maxlength="200" style="WIDTH: 450px" name="ContentSource" id="videoPath"/></TD>
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
    <TD align=middle width=153 height=30 valign="top">���ص�ַ</TD>
    <TD colspan="3"><input class="input" maxlength="255" style="WIDTH: 400px" id="DownURL" name="DownURL"/>
</TD>
  </TR>
  <TR>
    <TD align=middle width=153 height=30 valign="top">�ϴ��ļ�</TD>
    <TD colspan="3">   <iframe name="image" frameborder="0" width='100%' height="40" scrolling="no"src="../../editor/upload.asp?sortType=<%=sortType%>&stype=file&Tobj=DownURL" allowTransparency="true"></iframe>
</TD>
  </TR>
  <TR>
    <TD align=center width=153 height=30 valign="top" colspan="2">
   <span onClick="showHide('moreform');imgnoneHide('himg');" style="cursor:pointer"><img id="himg" src="../../images/ico_5.gif"><strong>����߼�����Ŀ</strong></span>
</TD>
 <TD width="100%" colspan="2"></TD>
 </TR>
 
 
 
   <TR id="moreform" style="display:none;">
 <TD width="100%" colspan="4">
 
 <table cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <TR>
    <TD align=middle height=30 width="153"> ���Ȩ�� </TD>
    <TD colspan="3">
    <%=userGroupSelect("GroupID", 0, 0)%>
        <input type="radio" name="Exclusive" value=">=" checked="checked" />
        ����
        <input type="radio" name="Exclusive" value="=" /> 
        ר����������Ȩ��ֵ�ݿɲ鿴��ר����Ȩ��ֵ���ɲ鿴��    <img src="../../images/help.gif" class="tTip" title="�÷���ķ���Ȩ�޵��趨<br>����ΪȨ�޴��ڸ��û��鶼���Է���<br>
ר��Ϊֻ���趨�ĸ�����Է���"></TD>
  </TR>	
  <TR>
    <TD align=middle height=30>�Ǽ� </TD>
    <TD>
    <select name="star">
			<option value="1">��</option>
			<option value="2">���</option>
			<option value="3">����</option>
			<option value="4">�����</option>
			<option value="5">������</option>
			<option value="6">�������</option>
			<option value="7">��������</option>
     </select>
    <img src="../../images/help.gif" class="tTip" title="�Ǽ����࣬ʹ��ʱ����ģ���е�������Ǽ��ķ����ǩ����">    </TD>
    <TD>�������</TD>
    <TD><input class="input" maxlength="6" value="0" style="WIDTH: 60px" name="Visits"/></TD>
  </TR>
  <TR>
    <TD align=middle width=153 height=30>����ʱ��</TD>
    <TD><input class="input" value="<%=now%>" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" style="WIDTH: 120px" name="AddTime"/></TD>
    <TD>���� </TD>
    <TD><input class="input" maxlength="6" value="0" style="WIDTH: 60px" name="ContentOrder"/>  <img src="../../images/help.gif" class="tTip" title="��������Ҫģ���е��������򷽷���order����ID��������������������ܽ�ʧЧ"></TD>
  </TR>
  <TR>
    <TD align=middle height=30>��ʱ����</TD>
    <TD colspan="3"><INPUT type="checkbox" name="TimeStatus" value="1"/>
�Ƿ�ʱ��������ʱ����ʱ�� 
  <input class="input" value="<%=now%>" onClick="WdatePicker({dateFmt:'yyyy-MM-dd HH:mm:ss'})" style="WIDTH: 120px" name="Timeing"/></TD>
  </TR>
  <TR>
    <TD align=middle height=30>����ѡ�� </TD>
    <TD colspan="3"><INPUT type="checkbox" name="IsTop" value="1"/>
�ö�
<INPUT type="checkbox" name="IsRecommend" value="1"/>
�Ƽ� 
<INPUT type="checkbox" name="IsFeatured" <% if IsFeatured then echo"checked=""checked"""%> value="1"/>
�ر��Ƽ� 
<INPUT type="checkbox" name="IsImageNews" value="1"/>
ͼƬ����
<INPUT type="checkbox" name="IsNoComment" value="1"/>
��ֹ��������
<INPUT type="checkbox" checked="checked" name="ContentStatus" value="1"/>
����  <img src="../../images/help.gif" class="tTip" title="��ѡ����ģ���е���Ϊ׼��������Ӧ���ò��ֹ��ܽ�ʧЧ"></TD>
  </TR>
  <TR>
    <TD align=middle height=30>�Զ����ļ��� </TD>
    <TD colspan="3"><input class="input" maxlength="200" style="WIDTH: 120px" name="PageFileName" />  
      <%=FileExt%>
      <INPUT type="hidden" name="IsGenerated" value="1"/>
      <img src="../../images/help.gif" class="tTip" title="��̬ģʽ�£���ǰ���µ������ļ�����������<br>����������Է����ͳһ��������Ϊ׼"></TD>
  </TR>
  <TR>
    <TD align=middle height=30>TAG </TD>
    <TD colspan="3"> <input class="input" maxlength="200" style="WIDTH: 450px" name="ContentTag"/> ���TAG���ö��Ÿ���!</TD>
  </TR>
  <TR>
    <TD align=middle height=30>ҳ��ؼ���</TD>
    <TD colspan="3"> <input class="input" maxlength="200" style="WIDTH: 450px"  name="PageKeywords"/>    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>ҳ������</TD>
    <TD colspan="3">  <TEXTAREA class="textarea" style="WIDTH: 450px" name="PageDesc" rows="3"></TEXTAREA>    </TD>
  </TR>
  <TR>
    <TD align=middle height=30>ת�����ӵ�ַ  </TD>
    <TD colspan="3"><input class="input" maxlength="200" style="WIDTH: 200px" name="OutLink"  />
      <INPUT type="checkbox" name="IsOutLink" value="1" />
      ת������</TD>
  </TR>
</table>
 
 </TD>
 </TR>

    </TBODY></TABLE>
</DIV>
<DIV class=adminsubmit>
<input type="hidden" maxlength="200"  name="sortType" value="<%=sortType%>"/> 
<INPUT class="button" type="submit" value="���" />
<INPUT class="button" type="button" value="����" onClick="history.go(-1)"/> 
<INPUT onClick="location.href='<%=getPageName()%>?sortType=<%=sortType%>'" type="button" value="ˢ��" class="button" /> 
</DIV></DIV></FORM>
</body>
</html>
