
function AddClass(){
	var thissort='目标分类名称|当前分类名称'; 
	var sort=prompt('请输入目标分类名称和当前分类名称，中间用“|”隔开：',thissort);
	if(sort!=null&&sort!=''){document.myform.RetuneClass.options[document.myform.RetuneClass.length]=new Option(sort,sort);}
}
function ModifyClass(){
	if(document.myform.RetuneClass.length==0) return false;
	var thissort=document.myform.RetuneClass.value; 
	if (thissort=='') {alert('请先选择一个分类，再点修改按钮！');return false;}
	var sort=prompt('请输入目标分类名称和当前分类名称，中间用“|”隔开：',thissort);
	if(sort!=thissort&&sort!=null&&sort!=''){document.myform.RetuneClass.options[document.myform.RetuneClass.selectedIndex]=new Option(sort,sort);}
}
function DelClass(){
	if(document.myform.RetuneClass.length==0) return false;
	var thissort=document.myform.RetuneClass.value; 
	if (thissort=='') {alert('请先选择一个分类，再点删除按钮！');return false;}
	document.myform.RetuneClass.options[document.myform.RetuneClass.selectedIndex]=null;
}

function AddReplace(){
	var strthis='替换前的字符串|替换后的字符串'; 
	var str=prompt('请输入替换前的字符串和替换后的字符串，中间用“|”隔开：',strthis);
	if(str!=null&&str!=''){document.myform.strReplace.options[document.myform.strReplace.length]=new Option(str,str);}
}
function ModifyReplace(){
	if(document.myform.strReplace.length==0) return false;
	var strthis=document.myform.strReplace.value; 
	if (strthis=='') {alert('请先选择一个字符串，再点修改按钮！');return false;}
	var str=prompt('请输入替换前的字符串和替换后的字符串，中间用“|”隔开：',strthis);
	if(str!=strthis&&str!=null&&str!=''){document.myform.strReplace.options[document.myform.strReplace.selectedIndex]=new Option(str,str);}
}
function DelReplace(){
	if(document.myform.strReplace.length==0) return false;
	var strthis=document.myform.strReplace.value; 
	if (strthis=='') {alert('请先选择一个字符串，再点删除按钮！');return false;}
	document.myform.strReplace.options[document.myform.strReplace.selectedIndex]=null;
}

function CheckForm(){
	if (document.myform.ItemName.value==''){
		alert('项目名称不能为空！');
		return false;
	}
	if (document.myform.SiteUrl.value==''){
		alert('请输入目标站点URL！');
		return false;
	}
	if (document.myform.classid.value==''){
		alert('该一级分类已经有下属分类，请选择其下属分类！');
		return false;
	}
	if (document.myform.classid.value=='0'){
		alert('该分类是外部连接，不能添加内容！');
		return false;
	}
	if (document.myform.action.value=='step2'){
		for(var i=0;i<document.myform.RetuneClass.length;i++){
			if (document.myform.ClassList.value=='') document.myform.ClassList.value=document.myform.RetuneClass.options[i].value;
			else document.myform.ClassList.value+='$$$'+document.myform.RetuneClass.options[i].value;
		}
		for(var n=0;n<document.myform.strReplace.length;n++){
			if (document.myform.ReplaceList.value=='') document.myform.ReplaceList.value=document.myform.strReplace.options[n].value;
			else document.myform.ReplaceList.value+='$$$'+document.myform.strReplace.options[n].value;
		}
	}
}