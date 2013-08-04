var completeKeydownSubmit = false;	/***自动完成查询列表后，键盘回车是否直接搜索，否则直接转到相关页面***/
var completeQuerySubmit = false;	/***自动完成查询点击后直接提交到搜索页面，否则直接转到相关页面***/
(function() {
	function $G(id) {return document.getElementById(id);}
	function $C(tg) {return document.createElement(tg);}
	var sugComplete = $G("sugAutoComplete"),
	oInputQuery=oQueryKeyword,
	completeData,
	createScriptObj = null,
	completeTable = null,
	completeStringIndex = -1,
	completeKeyed = false,
	mouseOnSug = false,
	completeMouse = false;
	var isIE = navigator.userAgent.indexOf("MSIE") != -1 && !window.opera;
	var isWebKit = (navigator.userAgent.indexOf("AppleWebKit/") > -1);
	var isGecko = (navigator.userAgent.indexOf("Gecko") > -1) && (navigator.userAgent.indexOf("KHTML") == -1);
	var timer3 = 0;
	function hiddenCompleteIFrame() {
		if (isIE) {
			$G("sugCompleteIFrame").style.display = "none"
		}
		sugComplete.style.display = "none";
		timer3 = 0
	}
	function createTable() {
		var completeRows = completeTable.rows;
		for (var i = 0; i < completeRows.length; i++) {
			completeRows[i].className = "trmouseout"
		}
	}
	function clickTableRow(j) {
		return function() {
			oInputQuery.blur();
			hiddenCompleteIFrame();
			clearTimeout(timer1);
			timer1 = 0;
			clearTimeout(timer2);
			timer2 = 0;
			if (!completeQuerySubmit) {
				window.location.href=completeData.a[j];
			} else {
				submitQueryResult(this.cells[0].innerHTML)
			}
		}
	}
	function submitQueryResult(v) {
		oInputQuery.value = v;
		oSearchForm.submit()
	}
	function closeCompleteIFrame() {
		hiddenCompleteIFrame();
		clearInterval(timer1);
		oInputQuery.blur();
		oInputQuery.setAttribute("autocomplete", "on");
		oInputQuery.focus();
		if (navigator.cookieEnabled && !/sugComplete=0/.test(document.cookie)) {
			document.cookie = "sugComplete=0;domain="+document.domain+";path=/";
		}
	}
	function setSug() {
		if (typeof(completeData) != "object" || typeof(completeData.s) == "undefined") {
			return
		}
		var autoCompTable= $C("table");
		with(autoCompTable) {
			id = "autoCompleteTable";
			style.width = "100%";
			style.backgroundColor = "#fff";
			cellSpacing = 0;
			cellPadding = 2;
			style.cursor = "default"
		}
		var tb = $C("tbody");
		autoCompTable.appendChild(tb);
		for (var i = 0; i < completeData.s.length; i++) {
			var tr = tb.insertRow( - 1);
			tr.onmouseover = function() {
				createTable();
				this.className = "trmouseon";
				mouseOnSug = true
			};
			tr.onmouseout = createTable;
			tr.onmousedown = function(e) {
				completeMouse = true;
				if (!isIE) {
					e.stopPropagation();
					return false
				}
			};
			tr.onclick = clickTableRow(i);
			var td = tr.insertCell( - 1);
			td.innerHTML = completeData.s[i]
		}
		var th = tb.insertRow( - 1);
		var td = th.insertCell( - 1);
		td.style.textAlign = "right";
		var oHideCompFrame = $C("A");
		oHideCompFrame.href = "javascript:void(0)";
		oHideCompFrame.innerHTML = "关闭";
		oHideCompFrame.onclick = closeCompleteIFrame;
		td.appendChild(oHideCompFrame);
		sugComplete.innerHTML = "";
		sugComplete.appendChild(autoCompTable);
		sugComplete.style.width = (isIE ? oInputQuery.offsetWidth: oInputQuery.offsetWidth - 2) + "px";
		sugComplete.style.top = (isGecko ? oInputQuery.offsetHeight - 1 : oInputQuery.offsetHeight) + "px";
		sugComplete.style.display = "block";
		if (isIE) {
			var sugCompIFrame = $G("sugCompleteIFrame");
			with(sugCompIFrame.style) {
				display = "";
				position = "absolute";
				top = oInputQuery.offsetHeight + "px";
				left = "0";
				width = sugComplete.offsetWidth + "px";
				height = autoCompTable.offsetHeight + "px"
			}
		}
		completeTable = $G("autoCompleteTable");
		completeStringIndex = -1;
		s3 = ""
	}
	function documentOnkeydown(e) {
		e = e || window.event;
		completeKeyed = false;
		var ctr;
		if (e.keyCode == 13) {
			if (isIE) {
				e.returnValue = false
			} else {
				e.preventDefault()
			}
			if (completeKeydownSubmit) {
				oSearchForm.submit();
			} else {
				if (completeStringIndex == -1){
					oSearchForm.submit();
				} else {
					window.location.href=completeData.a[completeStringIndex];
				}
			}
			return
		}
		if (e.keyCode == 38 || e.keyCode == 40) {
			mouseOnSug = false;
			if (sugComplete.style.display != "none") {
				var completeRows = completeTable.rows;
				var l = completeRows.length - 1;
				for (var i = 0; i < l; i++) {
					if (completeRows[i].className == "trmouseon") {
						completeStringIndex = i;
						break
					}
				}
				createTable();
				var ctr;
				if (e.keyCode == 38) {
					if (completeStringIndex == 0) {
						oInputQuery.value = completeData.q;
						completeStringIndex = -1;
						completeKeyed = true
					} else {
						if (completeStringIndex == -1) {
							completeStringIndex = l
						}
						ctr = completeRows[--completeStringIndex];
						ctr.className = "trmouseon";
						oInputQuery.value = ctr.cells[0].innerHTML
					}
				}
				if (e.keyCode == 40) {
					if (completeStringIndex == l - 1) {
						oInputQuery.value = completeData.q;
						completeStringIndex = -1;
						completeKeyed = true
					} else {
						ctr = completeRows[++completeStringIndex];
						ctr.className = "trmouseon";
						oInputQuery.value = ctr.cells[0].innerHTML
					}
				}
				if (!isIE) {
					e.preventDefault()
				}
			}
		}
	}
	window.newasp = {
		sug: function(data) {
			if (typeof(data) == "object" && typeof(data.s) != "undefined" && typeof(data.s[0]) != "undefined") {
				completeData = data;
				setSug()
			} else {
				hiddenCompleteIFrame();
				completeData = {}
			}
		},
		init: function() {
			s3 = oInputQuery.value;
			timer1 = setInterval(formInputQuery, 10)
		}
	};
	var s1 = "";
	var s3;
	var timer2 = 0;
	function completeRequest() {
		var completeInputQuery = true;
		var completeQueryValue = oInputQuery.value;
		if (typeof(completeTable) != "undefined" && completeTable != null) {
			var completeRows = completeTable.rows;
			for (var i = 0; i < completeRows.length; i++) {
				if (completeRows[i].className == "trmouseon") {
					if (completeQueryValue == completeRows[i].cells[0].innerHTML && !mouseOnSug) {
						completeInputQuery = false
					}
				}
			}
		}
		if (completeInputQuery && !completeKeyed) {
			if (createScriptObj) {
				document.body.removeChild(createScriptObj)
			}
			createScriptObj = $C("script");
			createScriptObj.src = "../../common/query.asp?word=" + urlencode(oInputQuery.value) + dataQueryParam + "&d=" + (new Date()).getTime();
			document.body.appendChild(createScriptObj)
		}
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
	function formInputQuery() {
		var s2 = oInputQuery.value;
		if (s2 == s1 && s2 != "" && s2 != s3) {
			if (timer2 == 0) {
				timer2 = setTimeout(completeRequest, 100)
			}
		} else {
			clearTimeout(timer2);
			timer2 = 0;
			s1 = s2;
			if (s2 == "") {
				hiddenCompleteIFrame()
			}
			if (s3 != oInputQuery.value) {
				s3 = ""
			}
		}
	}
	var timer1;
	oInputQuery.onkeydown = documentOnkeydown;
	var isClkSug = false;
	oInputQuery.onblur = function(e) {
		if (!isClkSug) {
			if (timer3 == 0) {
				timer3 = setTimeout(hiddenCompleteIFrame, 200)
			}
		}
		isClkSug = false
	};
	document.onmousedown = function(e) {
		e = e || window.event;
		var elm = e.target || e.srcElement;
		if (elm == oInputQuery) {
			return
		}
		while (elm = elm.parentNode) {
			if (elm == sugComplete || elm == oInputQuery) {
				isClkSug = true;
				return
			}
		}
		if (timer3 == 0) {
			timer3 = setTimeout(hiddenCompleteIFrame, 200)
		}
	};
	window.onresize = function() {
		if (typeof(timer3) != "undefined" && timer3 != 0) {
			clearTimeout(timer3)
		}
		resetSuggestion()
	};
	sugComplete.style.zIndex = 200;
	if (isIE) {
		var sugCompIFrame = $C("IFRAME");
		sugCompIFrame.src = "javascript:void(0)";
		sugCompIFrame.id = "sugCompleteIFrame";
		with(sugCompIFrame.style) {
			display = "none";
			position = "absolute"
		}
		sugComplete.parentNode.insertBefore(sugCompIFrame, sugComplete)
	}
	function resetSuggestion() {
		if (sugComplete.style.display != "none") {
			setTimeout(function() {
				hiddenCompleteIFrame();
				if (completeData != null) {
					setSug(completeData);
					oInputQuery.focus()
				}
			},
			100)
		}
	}
	document.onkeydown = function(e) {
		if (isGecko) {
			e = e || window.event;
			if (e.ctrlKey) {
				if (e.keyCode == 61 || e.keyCode == 107 || e.keyCode == 109 || e.keyCode == 96 || e.keyCode == 48) {
					resetSuggestion()
				}
			}
		}
	};
	oInputQuery.onbeforedeattivate = function() {
		if (completeMouse) {
			window.event.cancelBubble = true;
			window.event.returnValue = false;
			completeMouse = false
		}
	};
	function addRule(className, rule) {
		var sheet = document.styleSheets[0];
		if (isIE) {
			sheet.addRule(className, rule)
		} else {
			var oneCssRule = className + "{" + rule + "}";
			sheet.insertRule(oneCssRule, sheet.cssRules.length)
		}
	}
	addRule("#sugAutoComplete", "border:1px solid #817F82;display:none;position:absolute;top:28px;left:0;-moz-user-select:none;");
	addRule("#sugAutoComplete td", "font:12px verdana;line-height:21px;color:#333;");
	addRule("#sugAutoComplete td a", "font:12px verdana;color:#333;");
	addRule(".trmouseon", "background-color:#316ac5;color:#fff;");
	addRule(".trmouseout", "background-color:#fff;color:#333;");
	addRule("#sugAutoComplete .trmouseon td", "color:#fff;");
	oInputQuery.onbeforedeactivate = function() {
		if (completeMouse) {
			window.event.cancelBubble = true;
			window.event.returnValue = false;
			completeMouse = false
		}
	};
	oInputQuery.setAttribute("autocomplete", "off")
})();