<%
Dim errno,errmsg
errno=Array("1","-2","-4","-5","-6","-7","-8","-9","-10","-11","-12")
errmsg=Array("没有需要取得的数据","帐号/密码不正确","余额不足","数据格式错误","参数有误","权限受限","流量控制错误","扩展码权限错误","内容长度过长","内部数据库错误","序列号状态错误")
'--------------------------------------
'注册激活
'--------------------------------------
function Register()
	Dim SoapRequest,xmlhttp,xmlDOC
	SoapRequest="<?xml version="&CHR(34)&"1.0"&CHR(34)&" encoding="&CHR(34)&"utf-8"&CHR(34)&"?>"& _
	"<soap:Envelope xmlns:xsi="&CHR(34)&"http://www.w3.org/2001/XMLSchema-instance"&CHR(34)&" "& _
	"xmlns:xsd="&CHR(34)&"http://www.w3.org/2001/XMLSchema"&CHR(34)&" "& _
	"xmlns:soap="&CHR(34)&"http://schemas.xmlsoap.org/soap/envelope/"&CHR(34)&">"& _
	"<soap:Body>"& _
	"<Register xmlns="&CHR(34)&"http://tempuri.org/"&CHR(34)&">"& _
	"<sn>"&smsSn&"</sn>"& _
	"<pwd>"&smsPwd&"</pwd>"& _
	"<province>"&smsProvince&"</province>"& _
	"<city>"&smsCity&"</city>"& _
	"<trade>"&smsTrade&"</trade>"& _
	"<entname>"&smsEntname&"</entname>"& _
	"<linkman>"&smsLinkman&"</linkman>"& _
	"<phone>"&smsPhone&"</phone>"& _
	"<mobile>"&smsMobile&"</mobile>"& _
	"<email>"&smsEmail&"</email>"& _
	"<fax>"&smsFax&"</fax>"& _
	"<address>"&smsAddress&"</address>"& _
	"<postcode>"&smsPostcode&"</postcode>"& _
	"<sign>"&smsSign&"</sign>"& _
	"</Register>"& _
	"</soap:Body>"& _
	"</soap:Envelope>"
	Set xmlhttp = server.CreateObject("Msxml2.XMLHTTP")
	xmlhttp.Open "POST",smsUrl,false
	xmlhttp.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
	xmlhttp.setRequestHeader "HOST",smsHost
	xmlhttp.setRequestHeader "Content-Length",LEN(SoapRequest)
	xmlhttp.setRequestHeader "SOAPAction", "http://tempuri.org/Register" 
	xmlhttp.Send(SoapRequest)	
	If xmlhttp.Status = 200 Then
		Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
		xmlDOC.load(xmlhttp.responseXML)
		Register=xmlDOC.documentElement.selectNodes("//RegisterResult")(0).text
		Set xmlDOC = nothing
	Else
		Register=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function
function SendSMS(mobile,content)
	Dim SoapRequest,xmlhttp,xmlDOC
	SoapRequest="<?xml version="&CHR(34)&"1.0"&CHR(34)&" encoding="&CHR(34)&"utf-8"&CHR(34)&"?>"& _
	"<soap:Envelope xmlns:xsi="&CHR(34)&"http://www.w3.org/2001/XMLSchema-instance"&CHR(34)&" "& _
	"xmlns:xsd="&CHR(34)&"http://www.w3.org/2001/XMLSchema"&CHR(34)&" "& _
	"xmlns:soap="&CHR(34)&"http://schemas.xmlsoap.org/soap/envelope/"&CHR(34)&">"& _
	"<soap:Body>"& _
	"<SendSMS xmlns="&CHR(34)&"http://tempuri.org/"&CHR(34)&">"& _
	"<sn>"&smsSn&"</sn>"& _
	"<pwd>"&smsPwd&"</pwd>"& _
	"<mobile>"&mobile&"</mobile>"& _
	"<content>"&content&"</content>"& _
	"</SendSMS>"& _
	"</soap:Body>"& _
	"</soap:Envelope>"
	Set xmlhttp = server.CreateObject("Msxml2.XMLHTTP")
	xmlhttp.Open "POST",smsUrl,false
	xmlhttp.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
	xmlhttp.setRequestHeader "HOST",smsHost
	xmlhttp.setRequestHeader "Content-Length",LEN(SoapRequest)
	xmlhttp.setRequestHeader "SOAPAction", "http://tempuri.org/SendSMS" 
	xmlhttp.Send(SoapRequest)	
	If xmlhttp.Status = 200 Then
	
	Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
	xmlDOC.load(xmlhttp.responseXML)
	SendSMS=xmlDOC.documentElement.selectNodes("//SendSMSResult")(0).text
	Set xmlDOC = nothing
	Else
	SendSMS=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
END function
'--------------------------------------
'获取余额
'--------------------------------------
function getBalance()
	Dim SoapRequest,xmlhttp,xmlDOC
	SoapRequest="<?xml version="&CHR(34)&"1.0"&CHR(34)&" encoding="&CHR(34)&"utf-8"&CHR(34)&"?>"& _
	"<soap:Envelope xmlns:xsi="&CHR(34)&"http://www.w3.org/2001/XMLSchema-instance"&CHR(34)&" "& _
	"xmlns:xsd="&CHR(34)&"http://www.w3.org/2001/XMLSchema"&CHR(34)&" "& _
	"xmlns:soap="&CHR(34)&"http://schemas.xmlsoap.org/soap/envelope/"&CHR(34)&">"& _
	"<soap:Body>"& _
	"<balance xmlns="&CHR(34)&"http://tempuri.org/"&CHR(34)&">"& _
	"<sn>"&smsSn&"</sn>"& _
	"<pwd>"&smsMd5&"</pwd>"& _
	"</balance>"& _
	"</soap:Body>"& _
	"</soap:Envelope>"
	Set xmlhttp = server.CreateObject("Msxml2.XMLHTTP")
	xmlhttp.Open "POST",smsUrl,false
	xmlhttp.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
	xmlhttp.setRequestHeader "HOST",smsHost
	xmlhttp.setRequestHeader "Content-Length",LEN(SoapRequest)
	xmlhttp.setRequestHeader "SOAPAction", "http://tempuri.org/balance"
	xmlhttp.Send(SoapRequest)
	If xmlhttp.Status = 200 Then
		Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
		xmlDOC.load(xmlhttp.responseXML)
		getBalance=xmlDOC.documentElement.selectNodes("//balanceResult")(0).text
		Set xmlDOC = nothing
	Else
		getBalance=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function
'--------------------------------------
'充值
'--------------------------------------
function ChargUp(cardno,cardpwd)
	Dim SoapRequest,xmlhttp,xmlDOC
	SoapRequest="<?xml version="&CHR(34)&"1.0"&CHR(34)&" encoding="&CHR(34)&"utf-8"&CHR(34)&"?>"& _
	"<soap:Envelope xmlns:xsi="&CHR(34)&"http://www.w3.org/2001/XMLSchema-instance"&CHR(34)&" "& _
	"xmlns:xsd="&CHR(34)&"http://www.w3.org/2001/XMLSchema"&CHR(34)&" "& _
	"xmlns:soap="&CHR(34)&"http://schemas.xmlsoap.org/soap/envelope/"&CHR(34)&">"& _
	"<soap:Body>"& _
	"<ChargUp xmlns="&CHR(34)&"http://tempuri.org/"&CHR(34)&">"& _
	"<sn>"&smsSn&"</sn>"& _
	"<pwd>"&smsMd5&"</pwd>"& _
	"</ChargUp>"& _
	"</soap:Body>"& _
	"</soap:Envelope>"
	Set xmlhttp = server.CreateObject("Msxml2.XMLHTTP")
	xmlhttp.Open "POST",smsUrl,false
	xmlhttp.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
	xmlhttp.setRequestHeader "HOST",smsHost
	xmlhttp.setRequestHeader "Content-Length",LEN(SoapRequest)
	xmlhttp.setRequestHeader "SOAPAction", "http://tempuri.org/ChargUp"
	xmlhttp.Send(SoapRequest)
	If xmlhttp.Status = 200 Then
		Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
		xmlDOC.load(xmlhttp.responseXML)
		ChargUp=xmlDOC.documentElement.selectNodes("//ChargUpResult")(0).text
		Set xmlDOC = nothing
	Else
		ChargUp=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function

'--------------------------------------
'发送短信
'参数：mobile 手机，content 内容，ext 扩展码，stime 定时时间，rrid 唯一标识如为空系统返回
'--------------------------------------
function mt(mobile,content,ext,stime,rrid)
	Dim SoapRequest,xmlhttp,xmlDOC
	SoapRequest="<?xml version="&CHR(34)&"1.0"&CHR(34)&" encoding="&CHR(34)&"utf-8"&CHR(34)&"?>"& _
	"<soap:Envelope xmlns:xsi="&CHR(34)&"http://www.w3.org/2001/XMLSchema-instance"&CHR(34)&" "& _
	"xmlns:xsd="&CHR(34)&"http://www.w3.org/2001/XMLSchema"&CHR(34)&" "& _
	"xmlns:soap="&CHR(34)&"http://schemas.xmlsoap.org/soap/envelope/"&CHR(34)&">"& _
	"<soap:Body>"& _
	"<mt xmlns="&CHR(34)&"http://tempuri.org/"&CHR(34)&">"& _
	"<sn>"&smsSn&"</sn>"& _
	"<pwd>"&smsMd5&"</pwd>"& _
	"<mobile>"&mobile&"</mobile>"& _
	"<content>"&content&"</content>"& _
	"<ext>"&ext&"</ext>"& _
	"<stime>"&stime&"</stime>"& _
	"<rrid>"&rrid&"</rrid>"& _
	"</mt>"& _
	"</soap:Body>"& _
	"</soap:Envelope>"
	Set xmlhttp = server.CreateObject("Msxml2.XMLHTTP")
	xmlhttp.Open "POST",smsUrl,false
	xmlhttp.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
	xmlhttp.setRequestHeader "HOST",smsHost
	xmlhttp.setRequestHeader "Content-Length",LEN(SoapRequest)
	xmlhttp.setRequestHeader "SOAPAction", "http://tempuri.org/mt"
	xmlhttp.Send(SoapRequest)
	If xmlhttp.Status = 200 Then
		Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
		xmlDOC.load(xmlhttp.responseXML)
		mt=xmlDOC.documentElement.selectNodes("//mtResult")(0).text
		Set xmlDOC = nothing
	Else
		mt=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function
'--------------------------------------
'接收短信
'--------------------------------------
function mo()
	Dim SoapRequest,xmlhttp,xmlDOC
	SoapRequest="<?xml version="&CHR(34)&"1.0"&CHR(34)&" encoding="&CHR(34)&"utf-8"&CHR(34)&"?>"& _
	"<soap:Envelope xmlns:xsi="&CHR(34)&"http://www.w3.org/2001/XMLSchema-instance"&CHR(34)&" "& _
	"xmlns:xsd="&CHR(34)&"http://www.w3.org/2001/XMLSchema"&CHR(34)&" "& _
	"xmlns:soap="&CHR(34)&"http://schemas.xmlsoap.org/soap/envelope/"&CHR(34)&">"& _
	"<soap:Body>"& _
	"<mo xmlns="&CHR(34)&"http://tempuri.org/"&CHR(34)&">"& _
	"<sn>"&smsSn&"</sn>"& _
	"<pwd>"&smsMd5&"</pwd>"& _
	"</mo>"& _
	"</soap:Body>"& _
	"</soap:Envelope>"
	Set xmlhttp = server.CreateObject("Msxml2.XMLHTTP")
	xmlhttp.Open "POST",smsUrl,false
	xmlhttp.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
	xmlhttp.setRequestHeader "HOST",smsHost
	xmlhttp.setRequestHeader "Content-Length",LEN(SoapRequest)
	xmlhttp.setRequestHeader "SOAPAction", "http://tempuri.org/mo"
	xmlhttp.Send(SoapRequest)
	If xmlhttp.Status = 200 Then
		Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
		xmlDOC.load(xmlhttp.responseXML)
		mo=xmlDOC.documentElement.selectNodes("//moResult")(0).text
		Set xmlDOC = nothing
	Else
		mo=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function

function report(maxid)
	Dim SoapRequest,xmlhttp,xmlDOC
	SoapRequest="<?xml version="&CHR(34)&"1.0"&CHR(34)&" encoding="&CHR(34)&"utf-8"&CHR(34)&"?>"& _
	"<soap:Envelope xmlns:xsi="&CHR(34)&"http://www.w3.org/2001/XMLSchema-instance"&CHR(34)&" "& _
	"xmlns:xsd="&CHR(34)&"http://www.w3.org/2001/XMLSchema"&CHR(34)&" "& _
	"xmlns:soap="&CHR(34)&"http://schemas.xmlsoap.org/soap/envelope/"&CHR(34)&">"& _
	"<soap:Body>"& _
	"<report xmlns="&CHR(34)&"http://tempuri.org/"&CHR(34)&">"& _
	"<sn>"&smsSn&"</sn>"& _
	"<pwd>"&smsPwd&"</pwd>"& _
	"<maxid>"&maxid&"</maxid>"& _
	"</report>"& _
	"</soap:Body>"& _
	"</soap:Envelope>"
	Set xmlhttp = server.CreateObject("Msxml2.XMLHTTP")
	xmlhttp.Open "POST",smsUrl,false
	xmlhttp.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
	xmlhttp.setRequestHeader "HOST",smsHost
	xmlhttp.setRequestHeader "Content-Length",LEN(SoapRequest)
	xmlhttp.setRequestHeader "SOAPAction", "http://tempuri.org/report"
	xmlhttp.Send(SoapRequest)
	If xmlhttp.Status = 200 Then
		Set xmlDOC = server.CreateObject("MSXML.DOMDocument")
		xmlDOC.load(xmlhttp.responseXML)
		report=xmlDOC.documentElement.selectNodes("//reportResult")(0).text
		Set xmlDOC = nothing
	Else
		report=xmlhttp.Status&"&nbsp;"&xmlhttp.StatusText
	End if
	Set xmlhttp = Nothing
end function

'功能:判断一个值是否存在于数组
Function InArray( sValue, aArray )
    Dim x
    InArray = False
    For Each x In aArray
        If x = sValue Then
            InArray = True
            Exit For
        End If
    Next
End Function
%>