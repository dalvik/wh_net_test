
//***********************************��Ϊ��ҳ���ղؿ�ʼ*********************************************
function AddToFavorite()  
{  
    if (document.all){  
       window.external.addFavorite(document.URL,document.title);  
    }else if (window.sidebar){  
       window.sidebar.addPanel(document.title, document.URL, "");  
    }  
}  
  
//��Ϊ��ҳ  
function setHomepage(){  
    if (document.all){  
    document.body.style.behavior='url(#default#homepage)';  
    document.body.setHomePage(document.URL);  
  }else if (window.sidebar){  
        if(window.netscape){  
       try{   
          netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");   
       }catch (e){   
                    alert( "�ò�����������ܾ�����������øù��ܣ����ڵ�ַ�������� about:config,Ȼ���� signed.applets.codebase_principal_support ֵ��Ϊtrue" );   
       }  
        }   
    var prefs = Components.classes['@mozilla.org/preferences-service;1'].getService(Components. interfaces.nsIPrefBranch);  
    prefs.setCharPref('browser.startup.homepage',document.URL);  
    }  
} 
//***********************************��Ϊ��ҳ���ղؽ���*********************************************


<!-- Beginning of JavaScript Code -------------------

nereidFadeObjects = new Object();
nereidFadeTimers = new Object();
function nereidFade(object, destOp, rate, delta){
if (!document.all)
return
    if (object != "[object]"){ setTimeout("nereidFade("+object+","+destOp+","+rate+","+delta+")",0);
        return;
    }
    clearTimeout(nereidFadeTimers[object.sourceIndex]);
    diff = destOp-object.filters.alpha.opacity;
    direction = 1;
    if (object.filters.alpha.opacity > destOp){
        direction = -1;
    }
    delta=Math.min(direction*diff,delta);
    object.filters.alpha.opacity+=direction*delta;
    if (object.filters.alpha.opacity != destOp){
        nereidFadeObjects[object.sourceIndex]=object;
        nereidFadeTimers[object.sourceIndex]=setTimeout("nereidFade(nereidFadeObjects["+object.sourceIndex+"],"+destOp+","+rate+","+delta+")",rate);
    }
}
// -- End of JavaScript Code -->

<!--
//ͼƬ����������
var flag=false;
function DrawImage(ImgD){
var image=new Image();
var iwidth = 180; //��������ͼƬ���
var iheight = 160; //��������ͼƬ�߶�
image.src=ImgD.src;
if(image.width>0 && image.height>0){
flag=true;
if(image.width/image.height>= iwidth/iheight){
if(image.width>iwidth){ 
ImgD.width=iwidth;
ImgD.height=(image.height*iwidth)/image.width;
}else{
ImgD.width=image.width; 
ImgD.height=image.height;
}

ImgD.alt=image.width+"��"+image.height;
}
else{
if(image.height>iheight){ 
ImgD.height=iheight;
ImgD.width=(image.width*iheight)/image.height; 
}else{
ImgD.width=image.width; 
ImgD.height=image.height;
}
ImgD.alt=image.width+"��"+image.height;
}
}
} 
//���ã�<img src="ͼƬ" onload="javascript:DrawImage(this)">
//-->

<!--ff֧��outerHTML-->
 if(typeof(HTMLElement)!="undefined" && !window.opera)
 {
  HTMLElement.prototype.__defineGetter__("outerHTML",function()
  {
   var a=this.attributes, str="<"+this.tagName, i=0;for(;i<a.length;i++)
   if(a[i].specified)
    str+=" "+a[i].name+'="'+a[i].value+'"';
   if(!this.canHaveChildren)
    return str+" />";
   return str+">"+this.innerHTML+"</"+this.tagName+">";
  });
  HTMLElement.prototype.__defineSetter__("outerHTML",function(s)
  {
   var r = this.ownerDocument.createRange();
   r.setStartBefore(this);
   var df = r.createContextualFragment(s);
   this.parentNode.replaceChild(df, this);
   return s;
  });
  HTMLElement.prototype.__defineGetter__("canHaveChildren",function()
  {
   return !/^(area|base|basefont|col|frame|hr|img|br|input|isindex|link|meta|param)$/.test(this.tagName.toLowerCase());
  });
 }
 <!--ff֧��outerHTML-->

  // ͼƬ��С����
function ReImgSize(){
	var obj =document.getElementById("content").getElementsByTagName("img");
	if(!obj) obj = document.images;
	for (i=0; i<obj.length; i++){
		if (obj){
			if (obj[i].width>450){
				obj[i].width = 450;
				obj[i].style.width = 450;
				obj[i].title = "���´��ڴ�ͼƬ!";
				obj[i].style.cursor = "pointer";
				obj[i].border = 0;
				obj[i].outerHTML='<a href="'+obj[i].src+'" rel="lightbox[loval]"  target="_blank" title="[pic:title]">'+obj[i].outerHTML+'</a>'
			}
		}
	}
}
