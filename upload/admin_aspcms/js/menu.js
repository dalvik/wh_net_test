function $Nav(){
	if(window.navigator.userAgent.indexOf("MSIE")>=1) return 'IE';
  else if(window.navigator.userAgent.indexOf("Firefox")>=1) return 'FF';
  else return "OT";
}

var preID = 1;

function OpenMenu(cid,lurl,rurl,bid){
   if($Nav()=='IE'){
   	 //if(top.document.frames.menu.document.body.outerHTML.indexOf("Powered by AspCms")==-1)
	 //{
	 	//alert("Powered by AspCms!");
		//top.document.frames.main.location = 'http://www.aspcms.com';
	 //}
	 //else{
     if(rurl!='') top.document.frames.main.location = rurl;
     if(cid > -1) top.document.frames.menu.location = 'menu.asp?id='+cid;
     else if(lurl!='') top.document.frames.menu.location = lurl;
     if(bid>0) document.getElementById("d"+bid).className = 'thisclass';
     if(preID>0 && preID!=bid) document.getElementById("d"+preID).className = '';
     preID = bid;
	 //}
   }else{
	 //if(top.document.getElementById("menu").contentWindow.document.body.innerHTML.indexOf("www.benkin.net")==-1)
	 //{
	 	//alert("Powered by AspCms!");
		//top.document.getElementById("main").src = 'http://www.aspcms.com'; 
	 //}
	 //else{
     if(rurl!='') top.document.getElementById("main").src = rurl;
     if(cid > -1) top.document.getElementById("menu").src = 'menu.asp?id='+cid;
     else if(lurl!='') top.document.getElementById("menu").src = lurl;
     if(bid>0) document.getElementById("d"+bid).className = 'thisclass';
     if(preID>0 && preID!=bid) document.getElementById("d"+preID).className = '';
     preID = bid;
	 //}
   }
}

var preFrameW = '160,*';
var FrameHide = 0;

function ChangeMenu(way){
	var addwidth = 10;
	var fcol = top.document.all.bodyFrame.cols;
	if(way==1) addwidth = 10;
	else if(way==-1) addwidth = -10;
	else if(way==0){
		if(FrameHide == 0){
			preFrameW = top.document.all.bodyFrame.cols;
			top.document.all.bodyFrame.cols = '0,*';
			FrameHide = 1;
			return;
		}else{
			top.document.all.bodyFrame.cols = preFrameW;
			FrameHide = 0;
			return;
		}
	}
	fcols = fcol.split(',');
	fcols[0] = parseInt(fcols[0]) + addwidth;
	top.document.all.bodyFrame.cols = fcols[0]+',*';
}

function resetBT(){
	if(preID>0) document.getElementById("d"+preID).className = 'bdd';
	preID = 0;
}
function changeLang(sel){
	window.parent.location.href = "index.asp?id="+sel.options[sel.selectedIndex].value;
}

