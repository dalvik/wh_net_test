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

	if (form1.title.value==""){
		alert("���ⲻ��Ϊ�գ�");
		form1.titlee.focus();
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
	if (form1.filesize.value==""){
		alert("������С����Ϊ�գ�");
		form1.filesize.focus();
		return false;
	}
	if (form1.AllHits.value==""){
		alert("��ʼ���������Ϊ�գ�");
		form1.AllHits.focus();
		return false;
	}

}