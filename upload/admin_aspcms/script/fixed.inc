
lastScrollY=0; 

function heartBeat(){ 
	diffY=document.body.scrollTop; 
	percent=.1*(diffY-lastScrollY); 
	if(percent>0)percent=Math.ceil(percent); 
	else percent=Math.floor(percent); 
	{$Scroll} 
	lastScrollY=lastScrollY+percent; 
} 

{$SuspendCode}
window.setInterval("heartBeat()",1); 
