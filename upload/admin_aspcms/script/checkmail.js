function checkPost(){
	try{
		if (GetContentLength()==0){
			alert("邮件内容不能为空!");
			return false;
		}
	}
	catch(e){
		if (Trim(document.forms[0].content.value)=="") {
			alert("邮件内容不能为空!");
			document.forms[0].content.focus();
			return false
		}
	}
	try{
		if (document.forms[0].useremail.value==""){
			alert("E-mail地址不能为空！");
			document.forms[0].useremail.focus();
			return false;
		}
	}
	catch(e){}
	if (document.forms[0].topic.value==""){
		alert("邮件标题不能为空！");
		document.forms[0].topic.focus();
		return false;
	}
}