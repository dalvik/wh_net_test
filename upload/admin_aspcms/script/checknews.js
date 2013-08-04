function doChange(objText, objDrop){
	if (!objDrop) return;
	if(document.myform.BriefTopic.selectedIndex<2){
		document.myform.BriefTopic.selectedIndex+=1;
	}
	var str = objText.value;
	var arr = str.split("|");
	var nIndex = objDrop.selectedIndex;
	objDrop.length=1;
	for (var i=0; i<arr.length; i++){
		objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
	}
	objDrop.selectedIndex = nIndex;
}
function doSubmit(){
	if (checkPost()!=false){
		document.myform.submit();
	}
}
function checkPost(){
	var form1=document.myform;
	try{
		if (GetContentLength()==0){
			alert("内容不能为空!");
			return false;
		}
	}
	catch(e){
		if (Trim(form1.content.value)=="") {
			alert("内容不能为空!");
			form1.content.focus();
			return false;
		}
	}

	if (form1.title.value==""){
		alert("标题不能为空！");
		form1.title.focus();
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
	if (form1.Author.value==""){
		alert("作者不能为空！");
		form1.Author.focus();
		return false;
	}
	if (form1.ComeFrom.value==""){
		alert("来源不能为空！");
		form1.ComeFrom.focus();
		return false;
	}
	if (form1.AllHits.value==""){
		alert("初始点击数不能为空！");
		form1.AllHits.focus();
		return false;
	}
	return true;
}