<%
' This is an example of how you would go about setting up a custom payment 
' provider for the Ecommerce Plus template range. More information can be found
' at http://www.ecommercetemplates.com
' Here we have used the 2Checkout.com system as an example of how a common payment
' processor works. You can edit this file to match the details of your particular payment system
   ' Firstly you will need to set the URL to pass payment variables below in the FORM action %>
	<form method="POST" action="https://www.2checkout.com/cgi-bin/sbuyers/cartpurchase.2c">
<% ' A unique id is assigned to each order so that we can track the order. This is available as the orderid. Edit the name cart_order_id to that which is used by your payment system. %>
	  <input type="hidden" name="cart_order_id" value="<%=orderid%>">
<% ' In the Ecommerce Templates admin section for the Custom Payment System, up to 2 pieces of data can be entered %>
<% ' to configure a payment system. These are Data 1 and Data 2 and are available in the variables data1 and data2 %>
	  <input type="hidden" name="sid" value="<%=data1%>">
<% ' Our example of 2Checkout.com does not require a return URL, but I´ve included one below as an example if needed %>
	  <input type="hidden" name="returnurl" VALUE="<%=storeurl%>thanks.asp">
<% ' The variable ppmethod is available if needed to choose between authorize only and authorize capture payments. If this does not apply to your payment system just delete the line below %>
	  <input type="hidden" name="paymenttype" value="<% if ppmethod=1 then response.write "1" else response.write "0" %>">
<% ' The following should be quite self explanatory %>
	  <input type="hidden" name="total" value="<%=grandtotal%>">
	  <input type="hidden" name="card_holder_name" value="<%=Request.form("name")%>">
	  <input type="hidden" name="street_address" value="<%=Request.form("address")%>">
	  <input type="hidden" name="city" value="<%=Request.form("city")%>">
	  <input type="hidden" name="state" value="<%if Trim(Request.form("state"))<>"" then response.write Trim(Request.form("state")) else response.write Trim(Request.form("state2"))%>">
	  <input type="hidden" name="zip" value="<%=Request.form("zip")%>">
	  <input type="hidden" name="country" value="<%=Request.form("country")%>">
	  <input type="hidden" name="email" value="<%=Request.form("email")%>">
	  <input type="hidden" name="phone" value="<%=Request.form("phone")%>">
	  <%	if trim(Request.form("sname")) <> "" OR trim(Request.form("saddress")) <> "" then %>
	  <input type="hidden" name="ship_name" value="<%=Request.form("sname")%>">
	  <input type="hidden" name="ship_street_address" value="<%=Request.form("saddress")%>">
	  <input type="hidden" name="ship_city" value="<%=Request.form("scity")%>">
	  <input type="hidden" name="ship_state" value="<%if Trim(Request.form("sstate"))<>"" then response.write Trim(Request.form("sstate")) else response.write Trim(Request.form("sstate2"))%>">
	  <input type="hidden" name="ship_zip" value="<%=Request.form("szip")%>">
	  <input type="hidden" name="ship_country" value="<%=Request.form("scountry")%>">
	  <%	end if %>
<% ' A variable "demomode" is made available to the admin section that, if provided by the payment system will turn on a demo transaction mode %>
<% if demomode then Response.write "<input type='hidden' name='demo' value='Y'>" %>
<% ' IMPORTANT NOTE ! You may notice there is not closing <FORM> tag. This is intentional. %>