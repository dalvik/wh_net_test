<!--#include file="AspCms_smsSortFun.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="images/style.css" type=text/css rel=stylesheet>
</HEAD>
<SCRIPT>
function SelAll(theForm){
		for ( i = 0 ; i < theForm.elements.length ; i ++ )
		{
			if ( theForm.elements[i].type == "checkbox" && theForm.elements[i].name != "SELALL" )
			{
				theForm.elements[i].checked = ! theForm.elements[i].checked ;
			}
		}
}
</SCRIPT>

<Script type="text/javascript">
function setInput(t)
{
//'单篇1,文章2,产品3,下载4,招聘5,相册6,链接7
//alert("AA");

switch(t)
   {
   case "1":
     document.getElementById("SortTemplate").value="about.html";
     document.getElementById("ContentTemplate").value="";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/about/";
     document.getElementById("ContentFolder").value="";
     document.getElementById("SortFileName").value="about-{sortid}";
     document.getElementById("ContentFileName").value="";  
     break
   case "2":
     document.getElementById("SortTemplate").value="newslist.html";
     document.getElementById("ContentTemplate").value="news.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/newslist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/news/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";    
 
     break
   case "3":
     document.getElementById("SortTemplate").value="teacherlist.html";
     document.getElementById("ContentTemplate").value="teacher.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/teacherlist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/teacher/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";   
  
     break
   case "4":  
     document.getElementById("SortTemplate").value="downlist.html";
     document.getElementById("ContentTemplate").value="down.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/downlist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/down/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}"; 

     break
   case "5":
     document.getElementById("SortTemplate").value="joblist.html";
     document.getElementById("ContentTemplate").value="job.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/jobtlist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/job/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";  
     
     break
   case "6":
     document.getElementById("SortTemplate").value="albumlist.html";
     document.getElementById("ContentTemplate").value="album.html";
     document.getElementById("SortFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/albumtlist/";
     document.getElementById("ContentFolder").value="{sitepath}<%=setting.languagePath&htmlDir%>/album/";
     document.getElementById("SortFileName").value="list-{sortid}-{page}";
     document.getElementById("ContentFileName").value="{y}-{m}-{d}/{id}";
      
     break
   default:
 
   }
}

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

function showHide(objname,objimgname){
    var obj = getObject(objname);
	var objimg =getObject(objimgname);
    if(obj.style.display == "none"){
		obj.style.display = "block";
		objimg.src="../../images/ico_4.gif"
	}else{
		obj.style.display = "none";
		objimg.src="../../images/ico_5.gif"
	}
}

function blockHide(objname){
    var obj = getObject(objname);

		obj.style.display = "block";
		

}

function noneHide(objname){
    var obj = getObject(objname);
		obj.style.display = "none";
	
	

}

function imgnoneHide(objname){
    var obj = getObject(objname);
	
	if(obj.getAttribute("src",2)=='../../images/ico_4.gif')
	{	
		obj.src='../../images/ico_5.gif'
	}
}

function imgblockHide(objname){
    var obj = getObject(objname);
	if(obj.getAttribute("src",2)=='../../images/ico_5.gif')
	{
		obj.src='../../images/ico_4.gif'
	}
}

function allshowHide(styletype)
{
	 
	var TR_data=document.getElementById("listtable").getElementsByTagName("table");
	var img_data=document.getElementById("listtable").getElementsByTagName("img");
	
	for(var i=0;i<img_data.length;i++)
	{	
		if(img_data[i].id!='')
		{
			if(styletype=='1')
			{	
				
				imgnoneHide(img_data[i].id);
			}
			else
			{
				
				imgblockHide(img_data[i].id);
			}
		}
	}
	
	for(var i=0;i<TR_data.length;i++)
	{
		if(TR_data[i].id!='' && TR_data[i].id!='tab0')
		{
			if(styletype=='1')
			{
			noneHide(TR_data[i].id)
			}
			else
			{
			blockHide(TR_data[i].id)
			}
		
		}
	}
}
</script>
<BODY>
<!--#include file="AspCms_smsHead.asp" -->
<DIV class=searchzone>
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD height=30>&nbsp;</TD>
    <TD align=right colSpan=2>&nbsp;</TD>
  </TR></TBODY></TABLE></DIV>
<FORM name="form" action="?action=add" method="post" >
<DIV class=searchzone>
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
   
    <TD width="80%"> 联系人分类名称
      <INPUT class="input" style="WIDTH:120px" maxLength="200" name="smsSortName"/>
      <INPUT class="button" type="submit" value="添加" /></TD>
  
  </TR></TBODY></TABLE>
</DIV>
</FORM>
<FORM name="" action="" method="post">
<DIV class=listzone>
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 id="listtable">
  <TBODY>
  <TR class=list>
    <TD width=48 align="center" class=biaoti>选择</TD>
    <TD width=144 align="center" class=biaoti>编号</TD>
    <TD width="1057" height=28 align="center" class=biaoti><span class="searchzone">分类名称</span></TD>
    <TD width=606 align="center" class=biaoti><span class="searchzone">操作</span></TD>
    </TR>
	<%smssortList()%>
    </TBODY></TABLE>
</DIV>
<DIV class=piliang>
<INPUT onClick="SelAll(this.form)" type="checkbox" value="1" name="SELALL"> 全选&nbsp;
<INPUT class="button" type="submit" value="删除" onClick="if(confirm('确定要删除吗')){form.action='?action=del';}else{return false};"/> <INPUT class="button" type="submit" value="保存" onClick="form.action='?action=saveall';"/>
</DIV>
</FORM>
</BODY></HTML>
