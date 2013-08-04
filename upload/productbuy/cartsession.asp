<%

'response.Write action
'response.end
'加入购物车
Sub Add2Session	 

	Dim dicCount,dicPrice,pid,count,price

	pid = Request.QueryString("proid")
	count = Request.QueryString("count")
	price = Request.QueryString("price")
	
	if pid="" then exit sub
	If IsEmpty(Session("Cart")) Then
		Set dicCount = Server.CreateObject("Scripting.Dictionary")
		Set dicPrice = Server.CreateObject("Scripting.Dictionary")
	Else
		Set dicCount = Session("Cart")
		Set dicPrice = Session("Cart.Price")
	End If
	
	if dicCount.Exists(pid) then
		dicCount(pid) = count
		dicPrice(pid) = price
	else
		dicCount.add pid,count
		dicPrice.add pid,price
	end if
	'echo dicCount(pid)
	Set Session("Cart") = dicCount
	Set Session("Cart.Price") = dicPrice
End Sub



Sub DeleteFromSession()
Dim dicCount,dicPrice
dim proid:proid=Request.QueryString("proid")
echo "进入1"
	If Not IsEmpty(Session("Cart")) Then
		Set dicCount = Session("Cart")
		Set dicPrice = Session("Cart.Price")
		echo "进入2"
		If dicCount.exists(proid) Then dicCount.Remove(proid)
		If dicPrice.exists(proid) Then dicPrice.Remove(proid)
	End If	
	Set Session("Cart") = dicCount
	Set Session("Cart.Price") = dicPrice	
End Sub


Sub Echo(s)
	Response.Write s	
End Sub
Sub Die(s)
	Echo s
	Response.End
End Sub




dim action:action=Request.QueryString("action")




select case LCase(action)
	case "add2cart"
		Call Add2Session()
	case "deleteall"
		Session("Cart")="":Session("Cart.Price")=""
	case "delete"
		DeleteFromSession
end select
%>