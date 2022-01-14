<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com

version="20210621"

inlinestyle=""
if isempty(ectqbuystyle) then ectqbuystyle=1
if isempty(ectproductstyle) then ectproductstyle=1
if isempty(ectdetailstyle) then ectdetailstyle=1
if categorycolumns="" then categorycolumns=2

if ectqbuystyle="" then ectqbuystyle=0
if ectproductstyle="" then ectproductstyle=0
if ectdetailstyle="" then ectdetailstyle=0

if productcolumns="" then
	if ectproductstyle=2 then productcolumns=3 else productcolumns=2
end if

if quickbuylayout="" then
	if ectqbuystyle=1 then
		quickbuylayout="productimage,productname,discounts,reviewstars,instock,productid,sku,description,options,listprice,price,currency,detaillink,addtocartquant"
	end if
	if ectqbuystyle=2 then
		quickbuylayout="productimage,productname,discounts,reviewstars,instock,productid,manufacturer,description,options,listprice,price,currency,detaillink,addtocartquant"
	end if
	if ectqbuystyle=3 then
		quickbuylayout="productimage,productname,discounts,reviewstars,instock,productid,manufacturer,description,options,listprice,price,currency,detaillink,addtocartquant"
	end if
end if
if productpagelayout="" then
	if ectproductstyle=1 then
		productpagelayout="productimage,productname,discounts,reviewstars,instock,listprice,price,detaillink,quickbuy"
	end if
	if ectproductstyle=2 then
		productpagelayout="productimage,productname,discounts,reviewstars,instock,description,listprice,price,currency,detaillink,quickbuy"
	end if
	if ectproductstyle=3 then
		productpagelayout="productimage,productname,discounts,reviewstars,options,listprice,price,instock,detaillink,addtocart"
	end if
end if
if detailpagelayout="" then
	if ectdetailstyle=1 then
		detailpagelayout="navigation,productimage,productname,discounts,listprice,price,currency,instock,reviewstarslarge,shortdescription,options,addtocartquant,productid,manufacturer,sku,dateadded,description,previousnext,searchwords,socialmedia,reviews"
	end if
	if ectdetailstyle=2 then
		detailpagelayout="navigation,productimage,reviewstarslarge,productname,discounts,shortdescription,listprice,price,currency,instock,productid,manufacturer,sku,dateadded,options,addtocartquant,description,previousnext,searchwords,socialmedia,reviews"
	end if
	if ectdetailstyle=3 then
		detailpagelayout="navigation,productname,discounts,productimage,instock,reviewstarslarge,shortdescription,productid,manufacturer,sku,dateadded,listprice,price,currency,options,addtocartquant,description,previousnext,searchwords,socialmedia,reviews"
	end if
end if

if categorycolumns>1 then
	inlinestyle=inlinestyle&"div.category{width:" & ((100/categorycolumns)-1) & "%;}" & vbLf
	inlinestyle=inlinestyle&"@media screen and (max-width: 800px) {div.category{width:" & ((100/(categorycolumns-1))-1) & "%;}}" & vbLf
	if categorycolumns>2 then
		inlinestyle=inlinestyle&"@media screen and (max-width: 480px) {div.category{width:99%;}}" & vbLf
	end if
end if
if productcolumns>1 then
	inlinestyle=inlinestyle&"div.product{width:" & ((100/productcolumns)-1) & "%;}" & vbLf
	inlinestyle=inlinestyle&"@media screen and (max-width: 800px) {div.product{width:" & ((100/(productcolumns-1))-1) & "%;}}" & vbLf
	if productcolumns>2 then
		inlinestyle=inlinestyle&"@media screen and (max-width: 480px) {div.product{width:99%;}}" & vbLf
	end if
end if

	if NOT noectcartfiles then %>
<link href="css/ectcart.css?ver=<%=version%>" rel="stylesheet" type="text/css" />
<script src="js/ectcart.js?ver=<%=version%>"></script>
<%	end if
	if ectproductstyle<>0 OR ectqbuystyle<>0 OR ectdetailstyle<>0 then %>
<link href="css/ectstylebase.css?ver=<%=version%>" rel="stylesheet" type="text/css" />
<%	else
		inlinestyle=""
	end if
	if ectproductstyle<>0 then %>
<link href="css/ectstyleproduct<%=ectproductstyle%>.css?ver=<%=version%>" rel="stylesheet" type="text/css" />
<%	end if
	if ectqbuystyle<>0 then %>
<link href="css/ectstyleqbuy<%=ectqbuystyle%>.css?ver=<%=version%>" rel="stylesheet" type="text/css" />
<%	end if
	if ectdetailstyle<>0 then %>
<link href="css/ectstyledetails<%=ectdetailstyle%>.css?ver=<%=version%>" rel="stylesheet" type="text/css" />
<%	end if
if inlinestyle<>"" then
	print "<style>" & inlinestyle & "</style>"
end if
%>
