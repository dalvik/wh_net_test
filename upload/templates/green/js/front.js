/*tab切换*/
tabchange = function(btnId,showId){
	$("#"+btnId).find("dt").mouseover(function(){
		$(this).addClass("onfocus").siblings().removeClass("onfocus");					
		var index = $("#"+btnId).find("dt").index(this);		
		$("#"+showId).find("ul").eq(index).removeClass("none").siblings().addClass("none");
	});	
}




//加载
$(function(){
		   //tab切换
		   tabchange("news_t","news_c");
		   })	