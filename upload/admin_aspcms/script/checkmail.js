function checkPost(){
	try{
		if (GetContentLength()==0){
			alert("�ʼ����ݲ���Ϊ��!");
			return false;
		}
	}
	catch(e){
		if (Trim(document.forms[0].content.value)=="") {
			alert("�ʼ����ݲ���Ϊ��!");
			document.forms[0].content.focus();
			return false
		}
	}
	try{
		if (document.forms[0].useremail.value==""){
			alert("E-mail��ַ����Ϊ�գ�");
			document.forms[0].useremail.focus();
			return false;
		}
	}
	catch(e){}
	if (document.forms[0].topic.value==""){
		alert("�ʼ����ⲻ��Ϊ�գ�");
		document.forms[0].topic.focus();
		return false;
	}
}