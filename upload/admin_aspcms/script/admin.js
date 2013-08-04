// 修改编辑栏高度
function admin_Size(num,objname) {
	var obj=document.getElementById(objname)
	if (parseInt(obj.rows)+num>=3) {
		obj.rows = parseInt(obj.rows) + num;	
	}
	if (num>0) {
		obj.width="90%";
	}
}

function CtrlEnter() {
	if(event.ctrlKey && window.event.keyCode==13) {
		this.document.myform.submit();
	}	
}
function CheckAll(form) {  
	for (var i=0;i<form.elements.length;i++)  {  
		var e = form.elements[i];  
		if (e.name != 'chkall')  
		e.checked = true // form.chkall.checked;  
	}  
} 
 
function ContraSel(form) {
	for (var i=0;i<form.elements.length;i++){
		var e = form.elements[i];
		if (e.name != 'chkall')
		e.checked=!e.checked;
	}
}
function detectMacXFF() {
  var userAgent = navigator.userAgent.toLowerCase();
  if (userAgent.indexOf('mac') != -1 || userAgent.indexOf('firefox')!=-1 || userAgent.indexOf('opera')!=-1) {
	return true;
  }else{
	return false;
  }
}
var ColorImg;
var ColorValue;
function hideColourPallete() {
	document.getElementById("colourPalette").style.visibility="hidden";
}
function Getcolor(img_val,input_val){
	var obj = document.getElementById("colourPalette");
	ColorImg = img_val;
	ColorValue = document.getElementById(input_val);
	if (obj){
	obj.style.left = getOffsetLeft(ColorImg) + "px";
	obj.style.top = (getOffsetTop(ColorImg) + ColorImg.offsetHeight) + "px";
	if (obj.style.visibility=="hidden")
	{
	obj.style.visibility="visible";
	}else {
	obj.style.visibility="hidden";
	}
	}
}
//Colour pallete top offset
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}

//Colour pallete left offset
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
function setColor(color)
{
	if (ColorValue){ColorValue.value = color;}
	if (ColorImg){ColorImg.style.backgroundColor = color;}
	document.getElementById("colourPalette").style.visibility="hidden";
}

function bbimg(o){
	var zoom=parseInt(o.style.zoom, 10)||100;zoom+=event.wheelDelta/12;if (zoom>0) o.style.zoom=zoom+'%';
	return false;
}
function imgzoom(img,maxsize){
	var a=new Image();
	a.src=img.src
	if(a.width > maxsize * 4)
	{
		img.style.width=maxsize;
	}
	else if(a.width >= maxsize)
	{
		img.style.width=Math.round(a.width * Math.floor(4 * maxsize / a.width) / 4);
	}
	return false;
}
//图片自动调整的模式，1为按比例调整 ，2 按大小调整。
var resizemode=1
function imgresize(o){
	 if (resizemode==2 || o.onmousewheel){
	 	if(o.width > 500 ){
				o.style.width='500px';
			}
			if(o.height > 800){
				o.style.height='800px';
			}
		}else{
		var parentNode=o.parentNode.parentNode
		if (parentNode){
		if (o.offsetWidth>=parentNode.offsetWidth) o.style.width='98%';
		}else{
		var parentNode=o.parentNode
		if (parentNode){
			if (o.offsetWidth>=parentNode.offsetWidth) o.style.width='98%';
			}
		}
	}
}

//运行代码
function runEx(cod1)  {
	 cod=document.getElementById(cod1)
	  var code=cod.value;
	  if (code!=""){
		  var newwin=window.open('','','');  
		  newwin.opener = null 
		  newwin.document.write(code);  
		  newwin.document.close();
	}
}
//复制代码
function doCopy(ID) { 
	if (document.all){
		 textRange = document.getElementById(ID).createTextRange(); 
		 textRange.execCommand("Copy"); 
	}
	else{
		 //alert("此功能只能在IE上有效");
		 copyToClipboard(document.getElementById(ID).value);
	}
}
//另存为代码
function saveCode(cod1) {
	cod=document.getElementById(cod1)
	var code=cod.value;
	if (code!=""){
        var winname = window.open('', '_blank', 'top=10000');
        winname.document.open('text/html', 'replace');
        winname.document.write(code);
        winname.document.execCommand('saveas','','code.htm');
        winname.close();
	}
}
function copyToClipboard(txt) {
     if(window.clipboardData) {
             window.clipboardData.clearData();
             window.clipboardData.setData("Text", txt);
     } else if(navigator.userAgent.indexOf("Opera") != -1) {
          window.location = txt;
     } else if (window.netscape) {
          try {
               netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
          } catch (e) {
               alert("被浏览器拒绝！\n请在浏览器地址栏输入'about:config'并回车\n然后将'signed.applets.codebase_principal_support'设置为'true'");
          }
          var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
          if (!clip)
               return;
          var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);
          if (!trans)
               return;
          trans.addDataFlavor('text/unicode');
          var str = new Object();
          var len = new Object();
          var str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
          var copytext = txt;
          str.data = copytext;
          trans.setTransferData("text/unicode",str,copytext.length*2);
          var clipid = Components.interfaces.nsIClipboard;
          if (!clip)
               return false;
          clip.setData(trans,null,clipid.kGlobalClipboard);
     }
}
function GetContentLength(){
	var oEditor = FCKeditorAPI.GetInstance('content') ;
	var oDOM = oEditor.EditorDocument ;
	var iLength ;

	if ( document.all )
	{
		iLength = oDOM.body.innerText.length ;
	}
	else
	{
		var r = oDOM.createRange() ;
		r.selectNodeContents( oDOM.body ) ;
		iLength = r.toString().length ;
	}
	if (iLength > 65534)
	{
		alert ("最大长度为 65534 字节，当前长度为" + iLength + "字节！");
		return false;
	}else{
		return iLength
	}
}

function Trim(TRIM_VALUE){
	if(TRIM_VALUE.length < 1){
		return"";
	}
	TRIM_VALUE = RTrim(TRIM_VALUE);
	TRIM_VALUE = LTrim(TRIM_VALUE);
	if(TRIM_VALUE==""){
		return "";
	}
	else{
		return TRIM_VALUE;
	}
}

function RTrim(VALUE){
	var w_space = String.fromCharCode(32);
	var v_length = VALUE.length;
	var strTemp = "";
	if(v_length < 0){
		return"";
	}
	var iTemp = v_length -1;

	while(iTemp > -1){
		if(VALUE.charAt(iTemp) == w_space){
		}
		else{
			strTemp = VALUE.substring(0,iTemp +1);
			break;
		}
		iTemp = iTemp-1;

	}
	return strTemp;

}

function LTrim(VALUE){
	var w_space = String.fromCharCode(32);
	if(v_length < 1){
		return"";
	}
	var v_length = VALUE.length;
	var strTemp = "";
	var iTemp = 0;

	while(iTemp < v_length){
		if(VALUE.charAt(iTemp) == w_space){
			}
			else{
				strTemp = VALUE.substring(iTemp,v_length);
				break;
		}
		iTemp = iTemp + 1;
	}
	return strTemp;
}

function CheckDate(str){
	var reg = /^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2}) (\d{1,2}):(\d{1,2}):(\d{1,2})$/; 
	var r = str.match(reg); 
	if(r==null)return false; 
	var d= new Date(r[1],r[3]-1,r[4],r[5],r[6],r[7]); 
	return (d.getFullYear()+"-"+(d.getMonth()+1)+"-"+d.getDate()+" "+d.getHours()+":"+d.getMinutes()+":"+d.getSeconds());
}

var MediaTemp=new Array()
function MediaShow(strType,strID,strURL,intWidth,intHeight,strPath)
{
	var tmpstr
	if (MediaTemp[strID]==undefined) MediaTemp[strID]=false; else MediaTemp[strID]=!MediaTemp[strID];
	if(MediaTemp[strID]){
			if ( document.all )	{
	         	document.getElementById(strID).outerHTML = '<div id="'+strID+'"></div>'
			}
			else
			{
	         	document.getElementById(strID).innerHTML = ''
			}

		document.images[strID+"_img"].src=strPath+"mm_snd.gif" 		
		document.getElementById(strID+"_text").innerHTML="在线播放"	
	}else{
		document.images[strID+"_img"].src=strPath+"mm_snd_stop.gif" 		
		document.getElementById(strID+"_text").innerHTML="关闭在线播放"
		switch(strType){
			case "flash":
				tmpstr='<div style="height:6px;overflow:hidden"></div><object codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="'+intWidth+'" height="'+intHeight+'"><param name="movie" value="'+strURL+'" /><param name="quality" value="high" /><param name="AllowScriptAccess" value="never" /><embed src="'+strURL+'" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="'+intWidth+'" height="'+intHeight+'" /></object>';
				break;
			case "wma":
				tmpstr='<div style="height:6px;overflow:hidden"></div><object classid="CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95"  id="MediaPlayer" width="450" height="70"><param name=""howStatusBar" value="-1"><param name="AutoStart" value="True"><param name="Filename" value="'+strURL+'"></object>';
				break;
			case "wmv":
				tmpstr='<div style="height:6px;overflow:hidden"></div><object classid="clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95" codebase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=6,0,02,902" type="application/x-oleobject" standby="Loading..." width="'+intWidth+'" height="'+intHeight+'"><param name="FileName" VALUE="'+strURL+'" /><param name="ShowStatusBar" value="-1" /><param name="AutoStart" value="true" /><embed type="application/x-mplayer2" pluginspage="http://www.microsoft.com/Windows/MediaPlayer/" src="'+strURL+'" autostart="true" width="'+intWidth+'" height="'+intHeight+'" /></object>';
				break;
			case "rm":
				tmpstr='<div style="height:6px;overflow:hidden"></div><object classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" width="'+intWidth+'" height="'+intHeight+'"><param name="SRC" value="'+strURL+'" /><param name="CONTROLS" VALUE="ImageWindow" /><param name="CONSOLE" value="one" /><param name="AUTOSTART" value="true" /><embed src="'+strURL+'" nojava="true" controls="ImageWindow" console="one" width="'+intWidth+'" height="'+intHeight+'"></object>'+
                '<br/><object classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" width="'+intWidth+'" height="32" /><param name="CONTROLS" value="StatusBar" /><param name="AUTOSTART" value="true" /><param name="CONSOLE" value="one" /><embed src="'+strURL+'" nojava="true" controls="StatusBar" console="one" width="'+intWidth+'" height="24" /></object>'+'<br /><object classid="clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA" width="'+intWidth+'" height="32" /><param name="CONTROLS" value="ControlPanel" /><param name="AUTOSTART" value="true" /><param name="CONSOLE" value="one" /><embed src="'+strURL+'" nojava="true" controls="ControlPanel" console="one" width="'+intWidth+'" height="24" autostart="true" loop="false" /></object>';
				break;
			//case "rm":
				//tmpstr='<div style="height:6px;overflow:hidden"></div><object classid=""clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA"" class=""OBJECT"" id=""RAOCX"" width=""'+intWidth+'"" height=""'+intHeight+'""><param name=""src"" value=""'+strURL+'""/><param name=""console"" value=""Clip1""/><param name=""controls"" value=""imagewindow""/><param name=""autostart"" value=""true""/></object><br/><object classid=""CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA"" height=""32"" id=""video2"" width=""'+intWidth+'""><param name=src value=""'+strURL+'""/><param name=""autostart"" value=""-1""/><param name=""controls"" value=""controlpanel""/><param name=""console"" value=""Clip1""/></object>';
				//break;
			case "ra":
				tmpstr='<div style="height:6px;overflow:hidden"></div><object classid="clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA" id="RAOCX" width="450" height="60"><param name="_ExtentX" value="6694"><param name="_ExtentY" value="1588"><param name="AUTOSTART" value="true"><param name="SHUFFLE" value="0"><param name="PREFETCH" value="0"><param name="NOLABELS" value="0"><param name="SRC" value="'+strURL+'"><param name="CONTROLS" value="StatusBar,ControlPanel"><param name="LOOP" value="0"><param name="NUMLOOP" value="0"><param name="CENTER" value="0"><param name="MAINTAINASPECT" value="0"><param name="BACKGROUNDCOLOR" value="#000000"><embed src="'+strURL+'" width="450" autostart="true" height="60"></embed></object>';
				break;
			case "qt":
				tmpstr='<div style="height:6px;overflow:hidden"></div><embed src="'+strURL+'" autoplay="true" loop="false" controller="true" playeveryframe="false" cache="false" scale="TOFIT" bgcolor="#000000" kioskmode="false" targetcache="false" pluginspage="http://www.apple.com/quicktime/" />';
		}
		document.getElementById(strID).innerHTML = tmpstr;
	}
		document.getElementById(strID+"_href").blur()
}
function mover(src,clrOver){ 
	src.bgColor = clrOver; 
}
function mout(src,clrout)  { 
	src.bgColor = clrout; 
}