var selects = document.getElementsByTagName('select');
var isIE = (document.all && window.ActiveXObject && !window.opera) ? true : false;

function stopBubbling (ev) {	
	ev.stopPropagation();
}

function rSelects() {
	for (i=0;i<selects.length;i++){
		if(selects[i].id!="txtInfoAreaID"  && selects[i].id!="infosearchClassID" && selects[i].id!="txtInfoYouxiao"  && selects[i].id!="cYear" && selects[i].id!="calendarYear"  && selects[i].id!="cMonth" && selects[i].id.indexOf("AttrID")==-1 )
		{
			selects[i].style.display = 'none';
			select_tag = document.createElement('div');
			select_tag.id = 'select_' + selects[i].name;
			select_tag.className = 'select_box';
			selects[i].parentNode.insertBefore(select_tag,selects[i]);
	
			select_info = document.createElement('div');	
			select_info.id = 'select_info_' + selects[i].name;
			select_info.className='tag_select';
			select_info.style.cursor='pointer';
			select_tag.appendChild(select_info);
	
			select_ul = document.createElement('ul');	
			select_ul.id = 'options_' + selects[i].name;
			select_ul.className = 'tag_options';
			select_ul.style.position='absolute';
			select_ul.style.display='none';
			select_ul.style.zIndex='999';
			select_tag.appendChild(select_ul);
	
			rOptions(i,selects[i].name);
			
			mouseSelects(selects[i].name);
	
			if (isIE){
				selects[i].onclick = new Function("clickLabels3('"+selects[i].name+"');window.event.cancelBubble = true;");
			}
			else if(!isIE){
				selects[i].onclick = new Function("clickLabels3('"+selects[i].name+"')");
				selects[i].addEventListener("click", stopBubbling, false);
			}
		}
	}
}


function rOptions(i, name) {
	var options = selects[i].getElementsByTagName('option');
	var options_ul = 'options_' + name;
	for (n=0;n<selects[i].options.length;n++){	
		option_li = document.createElement('li');
		option_li.style.cursor='pointer';
		option_li.className='open';
		$_D(options_ul).appendChild(option_li);

		option_text = document.createTextNode(selects[i].options[n].text);
		option_li.appendChild(option_text);

		option_selected = selects[i].options[n].selected;

		if(option_selected){
			option_li.className='open_selected';
			option_li.id='selected_' + name;
			$_D('select_info_' + name).appendChild(document.createTextNode(option_li.innerHTML));
		}
		
		option_li.onmouseover = function(){	this.className='open_hover';}
		
		option_li.onmouseout = function()
		{
			if(this.id=='selected_' + name){
					this.className='open_selected';
			}
			else {
					this.className='open';
			}
		} 
		option_li.onclick = new Function("clickOptions("+i+","+n+",'"+selects[i].name+"')");
	}
}

function mouseSelects(name){
	var sincn = 'select_info_' + name;

	$_D(sincn).onmouseover = function(){ if(this.className=='tag_select') this.className='tag_select_hover'; }
	$_D(sincn).onmouseout = function(){ if(this.className=='tag_select_hover') this.className='tag_select'; }

	if (isIE){
		$_D(sincn).onclick = new Function("clickSelects('"+name+"');window.event.cancelBubble = true;");
	}
	else if(!isIE){
		$_D(sincn).onclick = new Function("clickSelects('"+name+"');");
		$_D('select_info_' +name).addEventListener("click", stopBubbling, false);
	}

}

function clickSelects(name){
	var sincn = 'select_info_' + name;
	var sinul = 'options_' + name;	
	for (i=0;i<selects.length;i++){	
		if(selects[i].id!="txtInfoAreaID" && selects[i].id!="infosearchClassID"  && selects[i].id!="txtInfoYouxiao"  && selects[i].id!="cYear" && selects[i].id!="calendarYear"  && selects[i].id!="cMonth"  && selects[i].id.indexOf("AttrID")==-1 )
		{
			if(selects[i].name == name){				
				if( $_D(sincn).className =='tag_select_hover'){
					$_D(sincn).className ='tag_select_open';
					$_D(sinul).style.display = '';
				}
				else if( $_D(sincn).className =='tag_select_open'){
					$_D(sincn).className = 'tag_select_hover';
					$_D(sinul).style.display = 'none';
				}
			}
			else{
				$_D('select_info_' + selects[i].name).className = 'tag_select';
				$_D('options_' + selects[i].name).style.display = 'none';
			}
		}
	}

}

function clickOptions(i, n, name){		
	var li = $_D('options_' + name).getElementsByTagName('li');

	$_D('selected_' + name).className='open';
	$_D('selected_' + name).id='';
	li[n].id='selected_' + name;
	li[n].className='open_hover';
	$_D('select_' + name).removeChild($_D('select_info_' + name));

	select_info = document.createElement('div');
	select_info.id = 'select_info_' + name;
	select_info.className='tag_select';
	select_info.style.cursor='pointer';
	$_D('options_' + name).parentNode.insertBefore(select_info,$_D('options_' + name));

	mouseSelects(name);

	$_D('select_info_' + name).appendChild(document.createTextNode(li[n].innerHTML));
	$_D( 'options_' + name ).style.display = 'none' ;
	$_D( 'select_info_' + name ).className = 'tag_select';
	selects[i].options[n].selected = 'selected';
	if(name=="HouseTopSelected")
	{
		if(li[n].innerHTML=="�Ҷ��ַ�"){
			$("#HouseTopSelectedA").show();
			$("#HouseTopSelectedB").hide();
			$("#HouseTopSelectedC").hide();
			$("#HouseTopSelectedD").hide();
		}
		else if(li[n].innerHTML=="���·�"){
			$("#HouseTopSelectedA").hide();
			$("#HouseTopSelectedC").hide();
			$("#HouseTopSelectedD").hide();
			$("#HouseTopSelectedB").show();
		}
		else if(li[n].innerHTML=="���ⷿ"){
			$("#HouseTopSelectedA").hide();
			$("#HouseTopSelectedB").hide();
			$("#HouseTopSelectedD").hide();
			$("#HouseTopSelectedC").show();
		}
		else if(li[n].innerHTML=="��С��"){
			$("#HouseTopSelectedA").hide();
			$("#HouseTopSelectedB").hide();
			$("#HouseTopSelectedC").hide();
			$("#HouseTopSelectedD").show();
		}
	}
}

window.onload = function(e) {
	/*bodyclick = document.getElementsByTagName('body').item(0);
	rSelects();
	bodyclick.onclick = function(){
		for (i=0;i<selects.length;i++){	
			$_D('select_info_' + selects[i].name).className = 'tag_select';
			$_D('options_' + selects[i].name).style.display = 'none';
		}
	}*/
}