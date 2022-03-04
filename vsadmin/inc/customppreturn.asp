<%
' This is an example of how you would go about setting up a custom payment 
' provider for the Ecommerce Plus template range. More information can be found
' at http://www.ecommercetemplates.com
' Here we have used the 2Checkout.com system as an example of how a common payment
' processor works. You can edit this file to match the details of your particular payment system

' Payment systems will normally pass back 3 different pieces of information. One will be the order
' id that we sent with the transaction. The second will be the authorization code and sometimes a
' variable will be passed back indicating the success of the transaction. You can use these to
' check that the order did indeed come from the payment system you are implementing
theorderid=Trim(Request.Form("cart_order_id"))
theauthcode=Trim(Request.Form("order_number"))
thesuccess=Trim(Request.Form("credit_card_processed"))
if theorderid<>"" AND theauthcode<>"" AND thesuccess="Y" then
	' You should not normally need to change the code below
	do_stock_management(trim(request.form("cart_order_id")))
	sSQL="UPDATE cart SET cartCompleted=1 WHERE cartOrderID="&theorderid
	cnn.Execute(sSQL)
	sSQL="UPDATE orders SET ordStatus=3,ordAuthNumber='"&theauthcode&"' WHERE ordID="&theorderid
	cnn.Execute(sSQL)
	Call order_success(theorderid,emailAddr,sendEmail)
else
	' Make sure you leave this condition here. It calls a failure routine if no match is found for any payment system.
	Call order_failed
end if
%>