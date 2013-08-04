
var yuxingbrOK=false;
var mie=false;
var aver=parseInt(navigator.appVersion.substring(0,1));
var aname=navigator.appName;
var yuxingmystop=0;

function yuxingcheckbrOK()
{if(aname.indexOf("Internet Explorer")!=-1)
{if(aver>=4) yuxingbrOK=navigator.javaEnabled();
mie=true;
}
if(aname.indexOf("Netscape")!=-1)  
{if(aver>=4) yuxingbrOK=navigator.javaEnabled();}
}
var vmin=2;
var vmax=5;
var vr=2;
var timer1;

function yuxingchip(chipname,width,height)
{
this.named=chipname;
this.vx=vmin+vmax*Math.random();
this.vy=vmin+vmax*Math.random();
this.w=width;
this.h=height;
this.xx=0;
this.yy=0;
this.timer1=null;
}

function yuxingmovechip(chipname)
{
if(yuxingbrOK && yuxingmystop==0)
{eval("yuxingchip="+chipname);
if(!mie)
{pageX=window.pageXOffset;
pageW=window.innerWidth;
pageY=window.pageYOffset;
pageH=window.innerHeight;
}
else
{pageX=window.document.body.scrollLeft;
pageW=window.document.body.offsetWidth-8;
pageY=window.document.documentElement.scrollTop;
pageH=window.document.documentElement.offsetHeight;
} 
yuxingchip.xx=yuxingchip.xx+yuxingchip.vx;
yuxingchip.yy=yuxingchip.yy+yuxingchip.vy;
yuxingchip.vx+=vr*(Math.random()-0.5);
yuxingchip.vy+=vr*(Math.random()-0.5);
if(yuxingchip.vx>(vmax+vmin))  yuxingchip.vx=(vmax+vmin)*2-yuxingchip.vx;
if(yuxingchip.vx<(-vmax-vmin)) yuxingchip.vx=(-vmax-vmin)*2-yuxingchip.vx;
if(yuxingchip.vy>(vmax+vmin))  yuxingchip.vy=(vmax+vmin)*2-yuxingchip.vy;
if(yuxingchip.vy<(-vmax-vmin)) yuxingchip.vy=(-vmax-vmin)*2-yuxingchip.vy;
if(yuxingchip.xx<=pageX)
{yuxingchip.xx=pageX;
yuxingchip.vx=vmin+vmax*Math.random();
}
if(yuxingchip.xx>=pageX+pageW-yuxingchip.w)
{yuxingchip.xx=pageX+pageW-yuxingchip.w;
yuxingchip.vx=-vmin-vmax*Math.random();
}
if(yuxingchip.yy<=pageY)
{yuxingchip.yy=pageY;
yuxingchip.vy=vmin+vmax*Math.random();
}
if(yuxingchip.yy>=pageY+pageH-yuxingchip.h)
{yuxingchip.yy=pageY+pageH-yuxingchip.h;
yuxingchip.vy=-vmin-vmax*Math.random();
}
if(!mie)
{eval('document.'+yuxingchip.named+'.top ='+yuxingchip.yy);
eval('document.'+yuxingchip.named+'.left='+yuxingchip.xx);
} 
else
{eval('document.all.'+yuxingchip.named+'.style.pixelLeft='+yuxingchip.xx);
eval('document.all.'+yuxingchip.named+'.style.pixelTop ='+yuxingchip.yy); 
}
	yuxingchip.timer1=setTimeout("yuxingmovechip('"+yuxingchip.named+"')",100);
}
}
function yuxingstop(x)
{
yuxingbrOK=true;
yuxingmystop=x;
yuxingmovechip("yuxingtome");
}
var yuxingtome;
var yuxingchip;
function yuxingtome()
{yuxingcheckbrOK(); 
yuxingtome=new yuxingchip("yuxingtome",80,80);
if(yuxingbrOK) 
{ yuxingmovechip("yuxingtome");
}
}

ns4=(document.layers)?true:false;
ie4=(document.all)?true:false;

function yuxingover()
{
if(ns4){
	//document.yuxing.left=window.innerWidth/2-400;
	eval('document.yuxing.left=document.'+yuxingchip.named+'.left');
	eval('document.yuxing.top=document.'+yuxingchip.named+'.top-15');
	document.yuxing.visibility="show";
	}else if(ie4) 
	{
	//document.all.yuxing.style.left=window.document.body.offsetWidth/2-400;
	document.all.yuxing.style.left=parseInt(document.all.yuxingtome.style.left);
	document.all.yuxing.style.top=parseInt(document.all.yuxingtome.style.top)+80;
	document.all.yuxing.style.visibility="visible";
	yuxingstop(1);
	}
}

function yuxingout()
{
if(ns4){
	document.yuxing.visibility="hide";
	}else if(ie4) 
	{
	document.all.yuxing.style.visibility="hidden";
	yuxingstop(0);
	}
}
document.writeln("<div id='yuxingtome' style='position:absolute;'>{$FloatCode}</div>");
yuxingtome()
