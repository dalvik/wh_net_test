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
			alert("��鲻��Ϊ��!");
			return false;
		}
	}
	catch(e){
		if (Trim(form1.content.value)=="") {
			alert("��鲻��Ϊ��!");
			form1.content.focus();
			return false;
		}
	}

	if (form1.SoftName.value==""){
		alert("������Ʋ���Ϊ�գ�");
		form1.SoftName.focus();
		return false;
	}
	if (form1.classid.value==""){
		alert("��һ�������Ѿ����������࣬��ѡ�����������࣡");
		form1.classid.focus();
		return false;
	}
	if (form1.classid.value=="0"){
		alert("�÷������ⲿ���ӣ�����������ݣ�");
		form1.classid.focus();
		return false;
	}
	if (form1.SoftSize.value==""){
		alert("�����С��û���");
		form1.SoftSize.focus();
		return false;
	}
	if (form1.RunSystem.value==""){
		alert("������л�������Ϊ�գ�");
		form1.RunSystem.focus();
		return false;
	}
	if (form1.SoftType.value==""){
		alert("������Ͳ���Ϊ�գ�");
		form1.SoftType.focus();
		return false;
	}
	if (form1.AllHits.value==""){
		alert("��ʼ���������Ϊ�գ�");
		form1.AllHits.focus();
		return false;
	}
	return true;
}