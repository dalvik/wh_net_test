function ToRunsystem(addTitle) {
	var revisedTitle;
	var currentTitle;
	currentTitle = document.myform.RunSystem.value;
	revisedTitle = currentTitle+addTitle;
	document.myform.RunSystem.value=revisedTitle;
	document.myform.RunSystem.focus();
	return; 
}

function doSubmit(){
	if (checkPost()!=false){
		document.myform.submit();
	}
}

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

	if (form1.SoftName.value==""){
		alert("软件名称不能为空！");
		form1.SoftName.focus();
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
	if (form1.SoftSize.value==""){
		alert("软件大小还没有填！");
		form1.SoftSize.focus();
		return false;
	}
	if (form1.RunSystem.value==""){
		alert("软件运行环境不能为空！");
		form1.RunSystem.focus();
		return false;
	}
	if (form1.SoftType.value==""){
		alert("软件类型不能为空！");
		form1.SoftType.focus();
		return false;
	}
	if (form1.AllHits.value==""){
		alert("初始点击数不能为空！");
		form1.AllHits.focus();
		return false;
	}
	return true;
}