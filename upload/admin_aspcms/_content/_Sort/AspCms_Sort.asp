<!--#include file="AspCms_SortFun.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gbk">
<LINK href="../../images/style.css" type=text/css rel=stylesheet>
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
<DIV class=searchzone>
<TABLE height=30 cellSpacing=0 cellPadding=0 width="100%" border=0>
  <TBODY>
  <TR>
    <TD height=30>&nbsp;</TD>
    <TD align=right colSpan=2><a href="###" onClick="allshowHide('1');" target="_self">全部收起</a> <a href="###" onClick="allshowHide('0');" target="_self">全部展开</a>
    <INPUT onClick="location.href='AspCms_QuickAdd.asp'" type="button" value="批量添加分类" class="button" >
    <INPUT onClick="location.href='AspCms_SortAdd.asp?id=0'" type="button" value="添加分类" class="button" > 
    <INPUT onClick="location.href='<%=getPageName()%>'" type="button" value="刷新" class="button" /> 
</TD></TR></TBODY></TABLE></DIV>
<FORM name="" action="" method="post">
<DIV class=listzone>
<TABLE cellSpacing=0 cellPadding=3 width="100%" align=center border=0 id="listtable">
  <TBODY>
  <TR class=list>
    <TD width=47 align="center" class=biaoti>选择</TD>
    <TD width=46 align="center" class=biaoti>编号</TD>
    <TD width="227" height=28 align="center" class=biaoti><span class="searchzone">分类名称</span></TD>
    <TD width=97 align="center" class=biaoti><span class="searchzone">类型</span></TD>
    <TD width=166 align="center" class=biaoti><span class="searchzone">链接</span></TD>
    <TD width=78 align="center" class=biaoti><span class="searchzone">排序</span></TD>
    <TD width=79 align="center" class=biaoti><span class="searchzone">状态</span></TD>
    <TD width=185 align="center" class=biaoti><span class="searchzone">操作</span></TD>
    </TR>
	<%sortList(0)%>
    </TBODY></TABLE>
</DIV>
<DIV class=piliang>
<INPUT onClick="SelAll(this.form)" type="checkbox" value="1" name="SELALL"> 全选&nbsp;
<INPUT class="button" type="submit" value="删除" onClick="if(confirm('确定要删除吗')){form.action='?action=del';}else{return false};"/>  
<INPUT class="button" type="submit" value="禁用" onClick="form.action='?action=off';" />
<INPUT class="button" type="submit" value="启用" onClick="form.action='?action=on';"/>
<INPUT class="button" type="submit" value="保存" onClick="form.action='?action=saveall';"/>
移动到:<%LoadSelect "moveSortID","Aspcms_Sort","SortName","SortID",getForm("id","get"), 0," order by SortOrder","顶级分类"%>
<INPUT class="button" type="submit" value="移动" onClick="form.action='?action=move';"/>

</DIV>
</FORM>
</BODY></HTML>
