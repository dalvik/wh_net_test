function checkPost(){
	var form1=document.forms[0];
	try{
		if (GetContentLength()==0){
			alert("简介不能为空!");
			return false;
		}
	}
	catch(e){
		if (Trim(form1.content.value)=="") {
			alert("简介不能为空!");
			form1.content.focus();
			return false;
		}
	}

	if (form1.title.value==""){
		alert("标题不能为空！");
		form1.titlee.focus();
		return false;
	}
	if (form1.classid.value==""){
		alert("该一级分类已经有下属分类，请选择其下属分类！");
		form1.classid.focus();
		return false;
	}
	if (form1.classid.value=="0"){
		alert("该分类是外部连接，不能添加内容！");
		form1.classid.focus();
		return false;
	}
	if (form1.filesize.value==""){
		alert("动画大小不能为空！");
		form1.filesize.focus();
		return false;
	}
	if (form1.AllHits.value==""){
		alert("初始点击数不能为空！");
		form1.AllHits.focus();
		return false;
	}

}