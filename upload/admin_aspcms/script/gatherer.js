
function AddClass(){
	var thissort='Ŀ���������|��ǰ��������'; 
	var sort=prompt('������Ŀ��������ƺ͵�ǰ�������ƣ��м��á�|��������',thissort);
	if(sort!=null&&sort!=''){document.myform.RetuneClass.options[document.myform.RetuneClass.length]=new Option(sort,sort);}
}
function ModifyClass(){
	if(document.myform.RetuneClass.length==0) return false;
	var thissort=document.myform.RetuneClass.value; 
	if (thissort=='') {alert('����ѡ��һ�����࣬�ٵ��޸İ�ť��');return false;}
	var sort=prompt('������Ŀ��������ƺ͵�ǰ�������ƣ��м��á�|��������',thissort);
	if(sort!=thissort&&sort!=null&&sort!=''){document.myform.RetuneClass.options[document.myform.RetuneClass.selectedIndex]=new Option(sort,sort);}
}
function DelClass(){
	if(document.myform.RetuneClass.length==0) return false;
	var thissort=document.myform.RetuneClass.value; 
	if (thissort=='') {alert('����ѡ��һ�����࣬�ٵ�ɾ����ť��');return false;}
	document.myform.RetuneClass.options[document.myform.RetuneClass.selectedIndex]=null;
}

function AddReplace(){
	var strthis='�滻ǰ���ַ���|�滻����ַ���'; 
	var str=prompt('�������滻ǰ���ַ������滻����ַ������м��á�|��������',strthis);
	if(str!=null&&str!=''){document.myform.strReplace.options[document.myform.strReplace.length]=new Option(str,str);}
}
function ModifyReplace(){
	if(document.myform.strReplace.length==0) return false;
	var strthis=document.myform.strReplace.value; 
	if (strthis=='') {alert('����ѡ��һ���ַ������ٵ��޸İ�ť��');return false;}
	var str=prompt('�������滻ǰ���ַ������滻����ַ������м��á�|��������',strthis);
	if(str!=strthis&&str!=null&&str!=''){document.myform.strReplace.options[document.myform.strReplace.selectedIndex]=new Option(str,str);}
}
function DelReplace(){
	if(document.myform.strReplace.length==0) return false;
	var strthis=document.myform.strReplace.value; 
	if (strthis=='') {alert('����ѡ��һ���ַ������ٵ�ɾ����ť��');return false;}
	document.myform.strReplace.options[document.myform.strReplace.selectedIndex]=null;
}

function CheckForm(){
	if (document.myform.ItemName.value==''){
		alert('��Ŀ���Ʋ���Ϊ�գ�');
		return false;
	}
	if (document.myform.SiteUrl.value==''){
		alert('������Ŀ��վ��URL��');
		return false;
	}
	if (document.myform.classid.value==''){
		alert('��һ�������Ѿ����������࣬��ѡ�����������࣡');
		return false;
	}
	if (document.myform.classid.value=='0'){
		alert('�÷������ⲿ���ӣ�����������ݣ�');
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