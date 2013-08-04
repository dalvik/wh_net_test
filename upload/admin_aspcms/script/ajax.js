var ajax_request_type = "GET";
var ajax_debug_mode = false;
function ajaxDebugPrint(text) {
	if (ajax_debug_mode)
	alert("RSD: " + text);
}
function initAjaxObject() {
	ajaxDebugPrint("initAjaxObject() called..");	
	var RetValue;
	try {
			RetValue = new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
		try {
		RetValue = new ActiveXObject("Microsoft.XMLHTTP");
		} catch (oc) {
		RetValue = null;
		}
	}
	if(!RetValue && typeof XMLHttpRequest != "undefined")
		RetValue = new XMLHttpRequest();
		if (!RetValue)
			ajaxDebugPrint("Could not create connection object.");
		return RetValue;
}

function getById(id) {
	itm = null;
	if (document.getElementById) {
		itm = document.getElementById(id);
	} else if (document.all)	{
		itm = document.all[id];
	} else if (document.layers) {
		itm = document.layers[id];
	}
	return itm;
}

function urlencode(text){
	text = text.toString();
	var matches = text.match(/[\x90-\xFF]/g);
	if (matches)
	{
		for (var matchid = 0; matchid < matches.length; matchid++)
		{
			var char_code = matches[matchid].charCodeAt(0);
			text = text.replace(matches[matchid], '%u00' + (char_code & 0xFF).toString(16).toUpperCase());
		}
	}
	return escape(text).replace(/\+/g, "%2B");
}

function ajaxcheckcode(stype,objid,sid){
    var objname = getById(objid).value;
	if (!sid) sid = "isok_checkcode";
	if (objname){
		if (objname.length>=4)
		{
			ajaxcheckdata(stype,objname,sid);
		}			
	}
}

function ajaxcheckdata(x_stype,x_objname,x_sid) {
	var xmlhttp = initAjaxObject();
	var post_data = null;
	var send_url = "../common/ajaxcheck.asp";
	try{
		if (ajax_request_type == "GET") {
			send_url = send_url + "?rs=" + x_stype + "&value="+ x_objname;
			post_data = null;
		}else{
			post_data = "rs=" + x_stype + "&value="+ x_objname;
		}
		xmlhttp.open(ajax_request_type, send_url.replace("&&","&"), true);
		if (ajax_request_type == "POST") {
			xmlhttp.setRequestHeader("Method", "POST " + send_url + " HTTP/1.1");
			xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		}
		xmlhttp.onreadystatechange = function() {
			if (xmlhttp.readyState != 4 || xmlhttp.status != 200) 
				return;
			var response = xmlhttp.responseText;
			ajaxDebugPrint("received " + response.substring(0));
			var innerEl = getById(x_sid);
			if(typeof(innerEl)=='object')
			innerEl.innerHTML = response;
		}
		xmlhttp.send(post_data);
		ajaxDebugPrint(x_stype + " url = " + send_url + "/post = " + post_data);
	}catch(e){}
	delete xmlhttp;
}