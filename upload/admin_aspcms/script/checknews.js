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
			alert("���ݲ���Ϊ��!");
			return false;
		}
	}
	catch(e){
		if (Trim(form1.content.value)=="") {
			alert("���ݲ���Ϊ��!");
			form1.content.focus();
			return false;
		}
	}

	if (form1.title.value==""){
		alert("���ⲻ��Ϊ�գ�");
		form1.title.focus();
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
	if (form1.Author.value==""){
		alert("���߲���Ϊ�գ�");
		form1.Author.focus();
		return false;
	}
	if (form1.ComeFrom.value==""){
		alert("��Դ����Ϊ�գ�");
		form1.ComeFrom.focus();
		return false;
	}
	if (form1.AllHits.value==""){
		alert("��ʼ���������Ϊ�գ�");
		form1.AllHits.focus();
		return false;
	}
	return true;
}