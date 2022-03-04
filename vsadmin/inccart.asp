<!--#include file="incemail.asp"-->
<!--#include file="md5.asp"-->
<%
'This code is copyright (c) ViciSoft SL, all rights reserved.
'The contents of this file are protected under law as the intellectual property
'of ViciSoft SL. Any use, reproduction, disclosure or copying
'of any kind without the express and written permission of ViciSoft SL is forbidden.
'Author: Vince Reid, vincereid@gmail.com
if sDSN="" then response.write "Database connection not set" : response.end
Dim sSQL,rs,alldata,quantity,grandtotal,netnav,cartID,cartEmpty,index,index2,rowcounter,objItem,totShipOptions,totaldiscounts,blockmultipurchase,multipurchaseblockmessage
Dim demomode,data1,data2,success,errormsg,shipping,totalgoods,handlingeligablegoods,orderid,sXML,destZip,allzones,stateTax,stateTaxRate,countryID,somethingToShip,taxfreegoods,uspsmethods,freeshipamnt
Dim iTotItems,international,multipleoptions,shipMethod,shipArr,shipcountry,intShipping(8,40),outofstockarr(),havematch,dHighest(10),dHighWeight,dTotalWeight,dTotalWeightOz,thePQuantity,thepname,thepprice,thepweight,xmlfnheaders
if xxHasDel="" then xxHasDel="Some of the items in your cart have been deleted and must be removed before you can continue."
if dateadjust="" then dateadjust=0
allfreeshipexempt=TRUE : success=TRUE : shiphomecountry=TRUE
ordGrandTotal=0 : ordTotal=0 : ordStateTax=0 : ordHSTTax=0 : ordCountryTax=0 : ordShipping=0 : ordHandling=0 : ordDiscount=0 : packnumber=1
affilID="" : ordState="" : ordCountry="" : ordDiscountText=""
termscontentregion=FALSE : fromshipselector=FALSE : nodiscounts=FALSE : usehst=FALSE : multipleoptions=FALSE : stockwarning=FALSE : backorder=FALSE : cartEmpty=FALSE : handlingeligableitem=FALSE : noshowcart=FALSE : isavsmismatch=FALSE
willpickup_=FALSE : insidedelivery_=FALSE : commercialloc_=FALSE : wantinsurance_=FALSE : saturdaydelivery_=FALSE : signaturerelease_=FALSE : hasstates=FALSE : returntocustomerdetails=FALSE : minquantityerror=FALSE : deleteditemerror=FALSE
shipping=0 : iTotItems=0 : stateTaxRate=0 : countryTax=0 : stateTax=0 : outofstockcnt=0 : ordComLoc=0
alldata="" : shipMethod="" : WSP="" : OWSP="" : appliedcouponname="" : ordAVS="" : ordCVV="" : stateAbbrev="" : international="" : cpnmessage="" : cpnerror="" : shipselectoraction="" : altrate="" : shipstate="" : shipstateid="" : backorderlist=""
countrytaxthreshold=0 : appliedcouponamount=0 : totalquantity=0 : statetaxfree=0 : countrytaxfree=0 : shipfreegoods=0 : : shipdiscountexempt=0 : numshipdiscountexempt=0 : totalgoods=0 : handlingeligablegoods=0 : shippingtax=0
freeshippingincludeshandling=FALSE : freeshippingincludesservices=FALSE : somethingToShip=FALSE : freeshippingapplied=FALSE : warncheckspamfolder=FALSE : homecountry=FALSE : gotcpncode=FALSE : freeshipmethodexists=FALSE
errordname=FALSE : errordaddress=FALSE : errordcity=FALSE : errordstate=FALSE : errordshipstate=FALSE : errordcountry=FALSE : errordzip=FALSE : errordphone=FALSE : errordemail=FALSE : errordemailv=FALSE : errtermsandconditions=FALSE : errordshipaddress=FALSE : errordshipcountry=FALSE
selectedshiptype=0 : numshipoptions=0 : freeshipamnt=0 : rowcounter=0 : stockrelitems=0 : thePQuantity=0 : thepweight=0 : grandtotal=0 : totaldiscounts=0 : giftcertsamount=0 : loyaltypointdiscount=0
analyticsoutput="" : ordShipName="":ordShipLastName="":ordShipAddress="":ordShipAddress2="":ordShipCity="":ordShipState="":ordShipZip="":ordShipCountry="":ordShipPhone=""
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs2=Server.CreateObject("ADODB.RecordSet")
Set rs3=Server.CreateObject("ADODB.RecordSet")
Set cnn=Server.CreateObject("ADODB.Connection")
cnn.open sDSN
alreadygotadmin=getadminsettings()
if nopriceanywhere then adminAltRates=0 : estimateshipping=FALSE
if adminAltRates>0 then
	rs.open "SELECT altrateid FROM alternaterates WHERE usealtmethod<>0 ORDER BY altrateorder",cnn,0,1
	if NOT rs.EOF then shipType=rs("altrateid")
	rs.close
	rs.open "SELECT altrateid FROM alternaterates WHERE usealtmethodintl<>0 ORDER BY altrateorderintl",cnn,0,1
	if NOT rs.EOF then adminIntShipping=rs("altrateid")
	rs.close
end if
homeCountryTaxRate=countryTaxRate
if zipposition="" then
	zipposition=1
	if origCountryID=65 then zipposition=2
	if origCountryID=133 then zipposition=4
end if
if xxAuNetR="" then xxAuNetR="Thank you! Your order has been received and for security reasons is currently being reviewed. We will be in touch as soon as possible!"
if xxChoIns="" then xxChoIns="Please choose if you would like to add shipping insurance"
if cartisincluded<>TRUE then
	if request.totalbytes>100000 then response.end
	rgcpncode=trim(replace(strip_tags2(request("cpncode")),"'",""))
	if instr(1, SESSION("cpncode"), rgcpncode & " ", 1)>0 OR instr(1, SESSION("giftcerts"), rgcpncode & " ", 1)>0 then rgcpncode=""
	if rgcpncode<>"" then ' Check for gift certs
		sSQL="SELECT gcID FROM giftcertificate WHERE gcRemaining>0 AND gcAuthorized<>0 AND gcID='" & trim(replace(rgcpncode,"'","")) & "'"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			if instr(SESSION("giftcerts"), rs("gcID") & " ")=0 then SESSION("giftcerts")=SESSION("giftcerts") & rs("gcID") & " "
			rgcpncode=""
			rs.movenext
		loop
		rs.close
	end if
	if rgcpncode<>"" then
		if trim(SESSION("cpncode"))<>"" then cpnerror=xxCanApp & " " & rgcpncode & ". " & xxOnOnCp & "<br>" else SESSION("cpncode")=trim(rgcpncode) & " "
	end if
	rgcpncode=trim(SESSION("cpncode"))
	payerid=getpost("payerid")
	token=request("token")
	if replace(getpost("sessionid"),"'","")<>"" then thesessionid=replace(getpost("sessionid"),"'","") else thesessionid=getsessionid()
	theid=replace(getpost("id"),"'","")
	if getget("mode")="checkout" then checkoutmode=getget("mode") else checkoutmode=getpost("mode")
	commercialloc_=(getpost("commercialloc")="Y") : SESSION("commercialloc_")=commercialloc_
	wantinsurance_=(getpost("wantinsurance")="Y") : SESSION("wantinsurance_")=wantinsurance_
	saturdaydelivery_=(getpost("saturdaydelivery")="Y") : SESSION("saturdaydelivery_")=saturdaydelivery_
	signaturerelease_=(getpost("signaturerelease")="Y") : SESSION("signaturerelease_")=signaturerelease_
	insidedelivery_=(getpost("insidedelivery")="Y")
	willpickup_=(getpost("willpickup")="Y") : SESSION("willpickup_")=willpickup_
	homedelivery_=getpost("homedelivery")
	ordPayProvider=getpost("payprovider")
	if NOT is_numeric(ordPayProvider) then ordPayProvider=""
	if getget("token")<>"" AND ordPayProvider="" then ordPayProvider=19
	if getget("action")="paypalcancel" then checkoutmode="paypalcancel"
	shipselectoraction=getpost("shipselectoraction")
	if getpost("shipselectoraction")="selector" then fromshipselector=TRUE
	if getpost("noredeempoints")="1" then SESSION("noredeempoints")=TRUE
	if is_numeric(getpost("altrates")) then altrate=int(getpost("altrates"))

	sSQL="SELECT contentData FROM contentregions WHERE contentName='termsandconditions'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		termscontentregion=TRUE
		termsandconditions=TRUE
	end if
	rs.close
end if
if getget("sharecart")<>"" AND getget("key")<>"" AND getget("list")<>"" then
	mykey=b64_hmac_sha256(adminSecret,getget("sharecart") & "this is a saved cart:" & getget("list"))
	if mykey=getget("key") then
		if getget("sharecart")=thesessionid then
			print "<div style=""text-align:center;padding:10px;margin-bottom:10px;border:1px solid red"">The source and destination carts are the same</div>"
		else
			sSQL="UPDATE cart SET cartSessionID='" & escape_string(thesessionid) & "', cartClientID=" & IIfVr(SESSION("clientID")<>"", escape_string(SESSION("clientID")), 0) & " WHERE cartID IN (" & escape_string(getget("list")) & ") AND cartSessionID='" & escape_string(getget("sharecart")) & "' AND cartCompleted=0 AND cartDateAdded>=" & vsusdate(date()-3)
			ect_query(sSQL)
			print "<div style=""text-align:center;padding:10px;margin-bottom:10px;border:1px solid grey"">The cart was copied successfully</div>"
		end if
	end if
end if
if is_numeric(getget("acartid")) AND getget("acarthash")<>"" then
	sSQL="SELECT aceID FROM abandonedcartemail WHERE aceOrderID=" & getget("acartid") & " AND aceKey='" & escape_string(getget("acarthash")) & "'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		cnn.execute("UPDATE cart SET cartSessionID='" & thesessionid & "' WHERE cartOrderID=" & getget("acartid") & " AND cartCompleted=0")
		cnn.execute("UPDATE orders SET ordSessionID='" & thesessionid & "' WHERE ordID=" & getget("acartid") & " AND ordStatus=2 AND ordAuthNumber=''")
		call setacookie("ectordid",getget("acartid"),0)
		call setacookie("ectsessid",thesessionid,0)
		call setacookie("ecthash",sha256(getget("acartid")&thesessionid&adminSecret),0)
	end if
	rs.close
end if
amazonpayment=FALSE : paypalexpress=FALSE
thefrompage=strip_tags2(IIfVr(getget("rp")<>"", getget("rp"), request.servervariables("HTTP_REFERER")))
if getget("rp")="" then
	if instr(1, storeurl, replace(parse_url(thefrompage,2),"www.",""), 1)=0 then thefrompage=""
end if
if instr(1,thefrompage,"javascript:",1)>0 OR instr(1,thefrompage,"cart"&extension,1)>0 OR instr(1,thefrompage,"thanks"&extension,1)>0 then thefrompage=""
if SESSION("clientID")<>"" AND SESSION("clientLoginLevel")<>"" then minloglevel=SESSION("clientLoginLevel") else minloglevel=0
countryTax=0 ' At present both countryTaxRate and countryTax are set in incfunctions
origShipType=shipType
orighandling=handling
orighandlingpercent=handlingchargepercent
function getcartforganalytics(gcartid)
	ao=""
	index=1
	sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pManufacturer,sectionName FROM (cart INNER JOIN products ON cart.cartProdID=products.pID) INNER JOIN sections ON products.pSection=sections.sectionID WHERE " & IIfVs(gcartid<>"","cartID=" & gcartid & " AND ") & "cartCompleted=0 AND " & getsessionsql() & " ORDER BY cartDateAdded"
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		sSQL="SELECT scName FROM searchcriteria WHERE scID=" & rs("pManufacturer")
		rs3.open sSQL,cnn,0,1
		if NOT rs3.EOF then brand=rs3("scName") else brand=""
		rs3.close
		if ao<>"" then ao=ao & ","
		ao=ao & "{item_id:'" & jsescapel(rs("cartProdID")) & "',item_name:'" & jsescapel(rs("cartProdName")) & "',index:" & index & ","
		if brand<>"" then ao=ao & "item_brand:'" & jsescapel(brand) & "',"
		ao=ao & "item_category:'" & jsescapel(rs("sectionName")) & "',price:" & jsescapel(rs("cartProdPrice")) & "}"
		index=index+1
		rs.movenext
	loop
	rs.close
	getcartforganalytics=ao
end function
function getstatefromid(tstate)
	getstatefromid=tstate
	if is_numeric(tstate) then
		sSQL="SELECT stateName,stateAbbrev,stateCountryID FROM states WHERE stateID=" & tstate
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then getstatefromid=IIfVr(usestateabbrev AND (rs("stateCountryID")=1 OR rs("stateCountryID")=2), rs("stateAbbrev"), rs("stateName"))
		rs.close
	end if
end function
function getidfromstate(tstate,statecountry)
	getidfromstate=""
	sSQL="SELECT stateID FROM states WHERE stateEnabled<>0 AND stateCountryID=" & IIfVr(is_numeric(statecountry),statecountry,0) & " AND (stateName='" & escape_string(tstate) & "'" & IIfVr((statecountry=1 OR statecountry=2), " OR stateAbbrev='" & escape_string(tstate) & "')", ")")
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then getidfromstate=rs("stateID")
	rs.close
end function
sub gettermsjsfunction()
		if termscontentregion then %>
function closetandc(){
	document.getElementById('ecttncdiv').parentNode.removeChild(document.getElementById('ecttncdiv'));
	return false;
}
function doprintterms(){
	var stylesheetlist='';
	if(document.styleSheets){
		for(var dpci=0;dpci<document.styleSheets.length;dpci++){
			if(document.styleSheets[dpci].href){
				stylesheetlist+='<link rel="stylesheet" type="text/css" href="'+document.styleSheets[dpci].href+'" />\n';
			}
		}
	}
	var prnttext='<html><head>'+stylesheetlist+'</head><body onload="window.print()">\n';
	prnttext+=document.getElementById('ecttermsandconds').outerHTML+"\n";
	prnttext+='</body></'+'html>\n';
	var newwin=window.open("","printit",'menubar=no, scrollbars=yes, width=600, height=450, directories=no,location=no,resizable=yes,status=no,toolbar=no');
	newwin.document.open();
	newwin.document.write(prnttext);
	newwin.document.close();
}
<%		end if
		print "function showtermsandconds(){" & vbCrLf
		if termscontentregion then %>
ecttncdiv=document.createElement('div');
ecttncdiv.setAttribute('id','ecttncdiv');
ecttncdiv.style.zIndex=1000;
ecttncdiv.style.position='fixed';
ecttncdiv.style.width='100%';
ecttncdiv.style.height='100%';
ecttncdiv.style.top='0px';
ecttncdiv.style.left='0px';
ecttncdiv.style.backgroundColor='rgba(140,140,150,0.5)';
document.body.appendChild(ecttncdiv);
ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
ajaxobj.open("GET",'vsadmin/ajaxservice.asp?action=termsandconditions',false);
ajaxobj.send(null);
ecttncdiv.innerHTML=ajaxobj.responseText;
<% 		else %>
newwin=window.open("termsandconditions.asp","Terms","menubar=no, scrollbars=yes, width=420, height=380, directories=no,location=no,resizable=yes,status=no,toolbar=no");
<%		end if
	print "}" & vbCrLf
end sub
function getamazonjsurl(isdemomode)
	scripturl="static-eu.payments-amazon.com/OffAmazonPayments/" & IIfVr(origCountryCode="DE","de/","uk/") & IIfVs(isdemomode,"sandbox/") & "lpa/"
	if origCountryCode="US" then scripturl="static-na.payments-amazon.com/OffAmazonPayments/us/" & IIfVs(isdemomode,"sandbox/")
	if origCountryCode="JP" then scripturl="origin-na.ssl-images-amazon.com/images/G/09/EP/offAmazonPayments/sandbox/prod/lpa/"
	getamazonjsurl="https://" & scripturl & "js/Widgets.js"
end function
function calculateStringToSignV2()
	calculateStringToSignV2="POST" & vbLf & scripturl & vbLf & endpointpath & vbLf & amazonstr
end function
function amazonparam2(nam, val)
	amazonstr=amazonstr & IIfVs(amazonstr<>"","&") & nam & "=" & replace(rawurlencode(replaceaccents(val)),"%7E","~")
end function
function iseuropean(cntryid)
	iseuropean=cntryid="BE" OR cntryid="BG" OR cntryid="CZ" OR cntryid="DK" OR cntryid="DE" OR cntryid="EE" OR cntryid="IE" OR cntryid="EL" OR cntryid="ES" OR cntryid="FR" OR cntryid="GB" OR cntryid="HR" OR cntryid="IT" OR cntryid="CY" OR cntryid="LV" OR cntryid="LT" OR cntryid="LU" OR cntryid="HU" OR cntryid="MT" OR cntryid="NL" OR cntryid="AT" OR cntryid="PL" OR cntryid="PT" OR cntryid="RO" OR cntryid="SI" OR cntryid="SK" OR cntryid="FI" OR cntryid="SE" OR cntryid="UK"
end function
function getstateabbrev(statename)
	getstateabbrev=""
	sSQL="SELECT stateAbbrev FROM states WHERE (stateCountryID=1 OR stateCountryID=2) AND " & IIfVr(is_numeric(statename), "stateID=" & statename, "(stateName='" & escape_string(statename) & "' OR stateAbbrev='" & escape_string(statename) & "')")
	rs2.Open sSQL,cnn,0,1
	if NOT rs2.EOF then getstateabbrev=rs2("stateAbbrev")
	rs2.Close
end function
function zipisoptional(sci)
	zipisoptional=FALSE
	for zoi=0 to UBOUND(zipoptional)
		if int(sci)=zipoptional(zoi) then zipisoptional=TRUE
	next
end function
function getDPs(currcode)
	getDPs=checkDPs(currcode)
end function
sub createdynamicstates(sSQL) %>
	function getziptext(cntid){
		if(cntid==1) return("<%=jsescape(xxZip)%>"); else return("<%=jsescape(xxPostco)%>");
	}
	function dynamiccountries(citem,stateid){
		if(citem!=null){
			var st,smen,cntid=citem[citem.selectedIndex].value;
			if(st=document.getElementById(stateid+'statetxt')){
				if(cntid==1) st.innerHTML='<%=jsescape(xxStateD)%>';
				else if(cntid==2||cntid==175) st.innerHTML='<%=jsescape(xxProvin)%>';
				else if(cntid==142||cntid==201) st.innerHTML='<%=jsescape(xxCounty)%>';
				else st.innerHTML='<%=jsescape(xxStaPro)%>';
				if(st2=document.getElementById(stateid+'state2')) st2.placeholder=st.innerHTML;
			}
			if(st=document.getElementById(stateid+'ziptxt')) st.innerHTML=getziptext(cntid);
			if(st=document.getElementById(stateid+'zip')) st.placeholder=getziptext(cntid);
			if(smen=document.getElementById(stateid+'state')){
				smen.disabled=false;
				if(countryhasstates[cntid]){
					smen.options[0].value='';
					if(cntid==1) smen.options[0].innerHTML='<%=jsescape(xxPSelUS)%>';
					else if(cntid==2) smen.options[0].innerHTML='<%=jsescape(xxPSelCA)%>';
					else if(cntid==201) smen.options[0].innerHTML='<%=jsescape(xxPSelUK)%>';
					else smen.options[0].innerHTML='<%=jsescape(xxPlsSel)%>';
					for(var cind=0;cind<dynst[cntid].length;cind++){
						if(cind>=smen.length-1)
							smen.options[cind+1]=new Option();
						smen.options[cind+1].value=dynst[cntid][cind][2];
						smen.options[cind+1].innerHTML=dynst[cntid][cind][0];
					}
					smen.length=cind+1;
					stateselectordisabled[stateid=='s'?1:0]=false;
				}else{
					smen.options[0].innerHTML='<%=jsescape(xxOutsid&" "&origCountryCode)%>';
					smen.disabled=true;
					stateselectordisabled[stateid=='s'?1:0]=true;
				}
				smen.selectedIndex=0;
			}
		}
	}
	function setinitialstate(isshp){<%
sistate=IIfVr(getpost("state")<>"",getpost("state"),ordState)
sisstate=IIfVr(getpost("sstate")<>"",getpost("sstate"),ordShipState)
sistatename=IIfVr(NOT is_numeric(sistate),sistate,"")
sisstatename=IIfVr(NOT is_numeric(sisstate),sisstate,"")
if sistate<>"" then
	sSQL2="SELECT stateID FROM states WHERE " & IIfVr(is_numeric(sistate), "stateID=" & sistate, "statename='" & escape_string(sistate) & "' OR stateAbbrev='" & escape_string(sistate) & "'")
	rs.open sSQL2,cnn,0,1
	if NOT rs.EOF then sistate=rs("stateID")
	rs.close
end if
if sisstate<>"" then
	sSQL2="SELECT stateID FROM states WHERE " & IIfVr(is_numeric(sisstate), "stateID=" & sisstate, "statename='" & escape_string(sisstate) & "' OR stateAbbrev='" & escape_string(sisstate) & "'")
	rs.open sSQL2,cnn,0,1
	if NOT rs.EOF then sisstate=rs("stateID")
	rs.close
end if
%>		var initstate=['<%=jsescape(sistate)%>','<%=jsescape(sisstate)%>'];
		var gotstate=false;
		if(document.getElementById(isshp+"state")){
			var smen=document.getElementById(isshp+"state");
			for(var cind=0; cind<smen.length; cind++){
				if(smen.options[cind].value==initstate[isshp=='s'?1:0]){
					smen.selectedIndex=cind;
					gotstate=true;
					break;
				}
			}
		}
		if(document.getElementById(isshp+"state2"))
			document.getElementById(isshp+"state2").value=(gotstate?'':(isshp=='s'?'<%=jsescape(sisstatename)%>':'<%=jsescape(sistatename)%>'));
	}
	var stateselectordisabled=[true,true];
	var dynst=[];var countryhasstates=[];
	var savstates=[];var savstatab=[];
<%	currcountry=0
	if sSQL<>"" then
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			do while NOT rs.EOF
				if currcountry<>rs("stateCountryID") then
					if currcountry<>0 then print "];" & vbLf
					currcountry=rs("stateCountryID")
					print "countryhasstates[" & currcountry & "]=" & currcountry & ";" & vbLf
					print "dynst[" & currcountry & "]=["
				else
					print "," & vbLf
				end if
				print "['" & jsescape(rs(getlangid("stateName",1048576))) & "','" & IIfVs(currcountry=1 OR currcountry=2,jsescape(rs("stateAbbrev"))) & "'," & rs("stateID") & "]"
				rs.movenext
			loop
			print "];" & vbLf
		end if
		rs.close
	end if
end sub
sub updategiftwrap()
	quantity=0
	currquant=0
	theid=giftwrappingid
	sSQL="SELECT SUM(cartQuantity) AS cartquant FROM cart WHERE cartGiftWrap<>0 AND cartCompleted=0 AND " & getsessionsql()
	rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT isnull(rs("cartquant")) then quantity=cint(rs("cartquant"))
		end if
	rs.close
	sSQL="SELECT cartQuantity FROM cart WHERE cartProdID='"&giftwrappingid&"' AND cartCompleted=0 AND " & getsessionsql()
	rs.open sSQL,cnn,0,1
		if NOT rs.EOF then currquant=rs("cartQuantity") else currquant=-1
	rs.close
	if quantity<>currquant then
		if currquant=-1 then
			if is_numeric(giftwrappingcost) AND giftwrappingcost<>0 AND quantity>0 then call additemtocart(xxGifPro,giftwrappingcost)
		elseif quantity=0 OR NOT is_numeric(giftwrappingcost) then
			ect_query("DELETE FROM cart WHERE cartProdID='"&giftwrappingid&"' AND cartCompleted=0 AND " & getsessionsql())
		else
			ect_query("UPDATE cart SET cartQuantity=" & quantity & ",cartProdPrice=" & giftwrappingcost & " WHERE cartProdID='"&giftwrappingid&"' AND cartCompleted=0 AND " & getsessionsql())
		end if
	end if
end sub
function getshiplogo(stype)
	if stype=3 then
		getshiplogo="<img src=""images/usps_logo.png"" alt=""USPS Logo"" />"
	elseif stype=4 then
		getshiplogo="<img src=""images/upslogo.png"" alt=""UPS Logo"" />"
	elseif stype=6 then
		getshiplogo="<img src=""images/canadapost.gif"" alt=""CanadaPost Logo"" />"
	elseif stype=7 OR stype=8 then
		getshiplogo="<img src=""images/fedexlogo.png"" alt=""FedEx Logo"" />"
	elseif stype=9 then
		getshiplogo="<img src=""images/dhllogo.gif"" alt=""DHL Logo"" />"
	else
		getshiplogo="<img src="""&IIfVr(shippinglogo<>"",shippinglogo,"images/defaultshiplogo.png")&""" alt=""Logo"" />"
	end if
end function
sub writealtshipline(altmethod,altid,pretext,defpretext,isestimator)
	if NOT isestimator then
		print "<div class=""altshippingselector"">"
			print "<div onclick=""selaltrate("&altid&")"" style=""cursor:pointer"">" & getshiplogo(altid) & "</div>"
			print "<div><a href=""#"" onclick=""selaltrate("&altid&")"" class=""ectlink"">" & altmethod & "</a></div>"
		print "</div>" & vbLf
	else
		if shippingoptionsasradios=TRUE then
			if altmethod<>"" OR origShipType=altid then print "<div class=""shipline""" & IIfVs(shipType=altid," style=""font-weight:bold""") & "><label class=""ectlabel shipradio""><input type=""radio"" class=""ectradio shipradio"" value="""&altid&""""&IIfVr(shipType=altid," checked=""checked""","")&" onclick=""selaltrate("&altid&")"" /><span class=""shiplinetext"">" & IIfVr(shipType=altid,defpretext,pretext) & altmethod & "</span></label></div>"
		else
			if altmethod<>"" OR origShipType=altid then print "<option value="""&altid&""""&IIfVr(shipType=altid," selected=""selected""","")&">"&IIfVr(shipType=altid,defpretext,pretext)&altmethod&"</option>"
		end if
	end if
end sub
sub retrieveorderdetails(ordid, sessid)
	if is_numeric(ordid) AND len(sessid)<100 then
		sSQL="SELECT ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordComLoc,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordAffiliate,ordAVS,ordCVV,ordAddInfo FROM orders WHERE ordID="&replace(ordid,"'","")&" AND ordSessionID='"&escape_string(sessid)&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			ordName=trim(rs("ordName")&"")
			ordLastName=trim(rs("ordLastName")&"")
			ordAddress=rs("ordAddress")
			ordAddress2=rs("ordAddress2")
			ordCity=rs("ordCity")
			ordState=rs("ordState")
			ordZip=rs("ordZip")
			ordCountry=rs("ordCountry")
			ordEmail=rs("ordEmail")
			ordEmail2=ordEmail
			ordPhone=rs("ordPhone")
			ordShipName=trim(rs("ordShipName")&"")
			ordShipLastName=trim(rs("ordShipLastName")&"")
			ordShipAddress=rs("ordShipAddress")
			ordShipAddress2=rs("ordShipAddress2")
			ordShipCity=rs("ordShipCity")
			ordShipState=rs("ordShipState")
			ordShipZip=rs("ordShipZip")
			ordShipCountry=rs("ordShipCountry")
			ordShipPhone=rs("ordShipPhone")
			ordPayProvider=rs("ordPayProvider")
			ordComLoc=rs("ordComLoc")
			ordExtra1=rs("ordExtra1")
			ordExtra2=rs("ordExtra2")
			ordShipExtra1=rs("ordShipExtra1")
			ordShipExtra2=rs("ordShipExtra2")
			ordCheckoutExtra1=rs("ordCheckoutExtra1")
			ordCheckoutExtra2=rs("ordCheckoutExtra2")
			ordAffiliate=rs("ordAffiliate")
			ordAVS=rs("ordAVS")
			ordCVV=rs("ordCVV")
			ordAddInfo=""
			if is_numeric(left(getpost("changeaction"),1)) AND is_numeric(ordid) then
				thebit=2^int(left(getpost("changeaction"),1))
				if mid(getpost("changeaction"),2,1)="y" then
					if (ordComLoc AND thebit)<>thebit then ordComLoc=ordComLoc+thebit
				else
					if (ordComLoc AND thebit)=thebit then ordComLoc=ordComLoc-thebit
				end if
				cnn.execute("UPDATE orders SET ordComLoc=" & ordComLoc & " WHERE ordStatus=2 AND ordID="&replace(ordid,"'","")&" AND ordSessionID='"&escape_string(sessid)&"'")
			end if
			if (ordComLoc AND 1)=1 then commercialloc_=TRUE
			if (ordComLoc AND 2)=2 OR abs(addshippinginsurance)=1 then wantinsurance_=TRUE
			if (ordComLoc AND 4)=4 then saturdaydelivery_=TRUE
			if (ordComLoc AND 8)=8 then signaturerelease_=TRUE
			if (ordComLoc AND 16)=16 then insidedelivery_=TRUE
		end if
		rs.close
	end if
end sub
sub getpayprovhandling()
	if is_numeric(ordPayProvider) then
		rs.open "SELECT ppHandlingCharge,ppHandlingPercent FROM payprovider WHERE payProvID="&ordPayProvider,cnn,0,1
		if NOT rs.EOF then
			handling=handling + rs("ppHandlingCharge")
			handlingchargepercent=handlingchargepercent + rs("ppHandlingPercent")
		end if
		rs.close
	end if
	orighandling=handling
	orighandlingpercent=handlingchargepercent
end sub
get_wholesaleprice_sql()
function getcctypefromnum(thecardnum)
	getcctypefromnum="Visa"
	if left(thecardnum, 1)="5" then
		getcctypefromnum="MasterCard"
	elseif left(thecardnum, 1)="6" then
		getcctypefromnum="Discover"
	elseif left(thecardnum, 1)="3" then
		getcctypefromnum="Amex"
	end if
end function
function show_states(tstate)
	show_states=FALSE
	print "<option value=''>"&xxOutsid&" "&origCountryCode&"</option>"
end function
function getcountryfromid(cntryid)
	getcountryfromid=""
	if is_numeric(cntryid) OR len(cntryid)=2 then
		sSQL="SELECT countryName FROM countries WHERE "
		if cntryid="GB" then
			sSQL=sSQL&"countryID=201"
		elseif cntryid="FR" then
			sSQL=sSQL&"countryID=65"
		elseif cntryid="PT" then
			sSQL=sSQL&"countryID=153"
		elseif cntryid="ES" then
			sSQL=sSQL&"countryID=175"
		elseif is_numeric(cntryid) then
			sSQL=sSQL&"countryID=" & escape_string(cntryid)
		else
			sSQL=sSQL&"countryCode='" & escape_string(cntryid) & "'"
		end if
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then getcountryfromid=rs("countryName")
		rs.close
	end if
end function
function getidfromcountry(cntry)
	getidfromcountry=1
	sSQL="SELECT countryID FROM countries WHERE countryName='" & escape_string(cntry) & "'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then getidfromcountry=rs("countryID")
	rs.close
end function
sub show_countries(tcountry,showplssel)
	if IsArray(allcountries) then
		if UBOUND(allcountries,2)>0 AND showplssel then print "<option value="""">"&xxPlsSel&"</option>"
		for rowcounter=0 to UBOUND(allcountries,2)
			print "<option value=""" & allcountries(3,rowcounter) & """"
			if tcountry=allcountries(0,rowcounter) then print " selected=""selected"""
			print ">"&allcountries(2,rowcounter)&"</option>"&vbCrLf
		next
	end if
end sub
function checkuserblock(thepayprov)
	multipurchaseblocked=FALSE
	if multipurchaseblockmessage="" then multipurchaseblockmessage="I'm sorry. We are experiencing temporary difficulties at the moment. Please try your purchase again later."
	if thepayprov<>"7" AND thepayprov<>"13" AND len(REMOTE_ADDR)<=15 then
		theip=REMOTE_ADDR
		if theip="" then theip="none"
		if blockmultipurchase>0 AND shipselectoraction="" then
			ect_query("DELETE FROM multibuyblock WHERE lastaccess<" & vsusdatetime(Now()-1))
			sSQL="SELECT ssdenyid,sstimesaccess FROM multibuyblock WHERE ssdenyip='" & theip & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				ect_query("UPDATE multibuyblock SET sstimesaccess=sstimesaccess+1,lastaccess=" & vsusdatetime(Now()) & " WHERE ssdenyid=" & rs("ssdenyid"))
				if rs("sstimesaccess")>=blockmultipurchase then multipurchaseblocked=TRUE
			else
				ect_query("INSERT INTO multibuyblock (ssdenyip,lastaccess) VALUES ('" & theip & "'," & vsusdatetime(Now()) & ")")
			end if
			rs.close
		end if
		if theip="none" then
			sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1")&" dcid FROM ipblocking"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
		else
			sSQL="SELECT dcid FROM ipblocking WHERE (dcip1=" & ip2long(theip) & " AND dcip2=0) OR (dcip1<=" & ip2long(theip) & " AND " & ip2long(theip) & "<=dcip2 AND dcip2<>0)"
		end if
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then multipurchaseblocked=TRUE
		rs.close
	end if
	checkuserblock=multipurchaseblocked
end function
function multShipWeight(theweight, themul)
	multShipWeight=(theweight*themul)/100.0
end function
sub subtaxesfordiscounts(theExemptions, discAmount)
	if (theExemptions AND 1)=1 then statetaxfree=statetaxfree - discAmount
	if (theExemptions AND 2)=2 then countrytaxfree=countrytaxfree - discAmount
	if (theExemptions AND 4)=4 then shipfreegoods=shipfreegoods - discAmount
end sub
sub addadiscount(resset, groupdiscount, dscamount, subcpns, cdcpncode, statetaxhandback, countrytaxhandback, theexemptions, thetax)
	totaldiscounts=totaldiscounts + dscamount
	if groupdiscount then
		statetaxfree=statetaxfree - (dscamount * statetaxhandback)
		countrytaxfree=countrytaxfree - (dscamount * countrytaxhandback)
	else
		call subtaxesfordiscounts(theexemptions, dscamount)
		if perproducttaxrate then countryTax=countryTax - ((dscamount * thetax) / 100.0)
	end if
	if InStr(cpnmessage,"<br>" & resset("cpnName") & "<br>")=0 then cpnmessage=cpnmessage & resset("cpnName") & "<br>"
	if subcpns then
		Set theres=cnn.execute("SELECT cpnID FROM coupons WHERE cpnNumAvail>0 AND cpnNumAvail<30000000 AND cpnID=" & resset("cpnID") & " AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))")
		if NOT theres.EOF then SESSION("couponapply")=SESSION("couponapply") & "," & resset("cpnID")
		ect_query("UPDATE coupons SET cpnNumAvail=cpnNumAvail-1 WHERE cpnNumAvail>0 AND cpnNumAvail<30000000 AND cpnID=" & resset("cpnID") & " AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))")
	end if
	if cdcpncode<>"" AND LCase(trim(resset("cpnNumber")))=LCase(cdcpncode) then gotcpncode=TRUE : appliedcouponname=resset("cpnName") : appliedcouponamount=dscamount
end sub
function timesapply(taquant,tathresh,tamaxquant,tamaxthresh,taquantrepeat,tathreshrepeat)
	if tamaxquant=0 then taquantrepeat=0
	if tamaxthresh=0 then tathreshrepeat=0
	if taquantrepeat=0 AND tathreshrepeat=0 then
		tatimesapply=1.0
	elseif tamaxquant=0 OR taquantrepeat=0 then
		tatimesapply=int((tathresh-tamaxthresh) / tathreshrepeat)+1
	elseif tamaxthresh=0 OR tathreshrepeat=0 then
		tatimesapply=int((taquant-tamaxquant) / taquantrepeat)+1
	else
		ta1=int((taquant-tamaxquant) / taquantrepeat)+1
		ta2=int((tathresh-tamaxthresh) / tathreshrepeat)+1
		if ta2 < ta1 then tatimesapply=ta2 else tatimesapply=ta1
	end if
	timesapply=tatimesapply
end function
function jschk(str)
	jschk=replace(trim(str&""),"\","\\")
	jschk=replace(jschk,"'","\'")
	jschk=replace(jschk,"<","\<")
	jschk=replace(jschk,">","\>")
end function
sub calculatediscounts(cdgndtot, subcpns, cdcpncode)
	totaldiscounts=0
	cdtotquant=0 : cdtotprice=0
	cpnmessage="<br>"
	if cdgndtot=0 then
		statetaxhandback=0.0
		countrytaxhandback=0.0
	else
		statetaxhandback=1.0 - ((cdgndtot - statetaxfree) / cdgndtot)
		countrytaxhandback=1.0 - ((cdgndtot - countrytaxfree) / cdgndtot)
	end if
	if NOT nodiscounts then
		Session.LCID=1033
		cdalldata=""
		sSQL="SELECT cartProdID,SUM(cartProdPrice*cartQuantity),SUM(cartQuantity),pSection,COUNT(cartProdID),pExemptions,pTax FROM products INNER JOIN cart ON cart.cartProdID=products.pID WHERE " & getsessionsql() & " AND cartProdID<>'"&giftcertificateid&"' AND cartProdID<>'"&donationid&"' AND cartProdID<>'"&giftwrappingid&"' AND cartCompleted=0 AND pExemptions<64 GROUP BY cartProdID,pSection,pExemptions,pTax"
		rs2.Open sSQL,cnn,0,1
		if NOT (rs2.EOF OR rs2.BOF) then cdalldata=rs2.getrows
		rs2.Close
		if IsArray(cdalldata) then
			for index=0 to UBOUND(cdalldata,2)
				' if (alldata(0,index)=giftcertificateid OR alldata(0,index)=donationid OR alldata(0,index)=giftwrappingid) AND isnull(alldata(8,index)) then alldata(8,index)=15
				sSQL="SELECT SUM(coPriceDiff*cartQuantity) AS totOpts FROM cart INNER JOIN cartoptions ON cart.cartID=cartoptions.coCartID WHERE cartCompleted=0 AND " & getsessionsql() & " AND cartProdID='" & escape_string(cdalldata(0,index)) & "'"
				rs2.Open sSQL,cnn,0,1
				if NOT isnull(rs2("totOpts")) then cdalldata(1,index)=cdalldata(1,index) + rs2("totOpts")
				rs2.Close
				cdalldata(2,index)=clng(cdalldata(2,index))
				cdtotquant=cdtotquant + cdalldata(2,index)
				cdtotprice=cdtotprice + cdalldata(1,index)
				topcpnids=cdalldata(3,index)
				thetopts=cdalldata(3,index)
				if isnull(cdalldata(6,index)) then cdalldata(6,index)=countryTaxRate
				if NOT isnull(thetopts) then
					for cpnindex=0 to 10
						if thetopts=0 then
							exit for
						else
							sSQL="SELECT topSection FROM sections WHERE sectionID=" & thetopts
							rs.open sSQL,cnn,0,1
							if NOT rs.EOF then
								thetopts=rs("topSection")
								topcpnids=topcpnids & "," & thetopts
							else
								rs.close
								exit for
							end if
							rs.close
						end if
					next
				end if
				attributelist=""
				sSQL="SELECT mSCscID FROM multisearchcriteria WHERE mSCpID='"&escape_string(cdalldata(0,index))&"'"
				rs.open sSQL,cnn,0,1
				do while NOT rs.EOF
					attributelist=attributelist&rs("mSCscID")&" "
					rs.movenext
				loop
				rs.close
				sSQL="SELECT DISTINCT cpnID,cpnDiscount,cpnType,cpnNumber,cpnName,cpnThreshold,cpnQuantity,cpnThresholdRepeat,cpnQuantityRepeat FROM coupons LEFT OUTER JOIN cpnassign ON coupons.cpnID=cpnassign.cpaCpnID WHERE cpnNumAvail>0 AND cpnStartDate<=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND cpnEndDate>=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND (cpnIsCoupon=0"
				if cdcpncode<>"" then sSQL=sSQL & " OR (cpnIsCoupon=1 AND cpnNumber='"&cdcpncode&"')"
				sSQL=sSQL & ") AND cpnThreshold<="&cdalldata(1,index)&" AND (cpnThresholdMax>"&cdalldata(1,index)&" OR cpnThresholdMax=0) AND cpnQuantity<="&cdalldata(2,index)&" AND (cpnQuantityMax>"&cdalldata(2,index)&" OR cpnQuantityMax=0) AND (cpnSitewide=0 OR cpnSitewide=2) AND " & _
					"(cpnSitewide=2 OR (cpaType=2 AND cpaAssignment='"&cdalldata(0,index)&"') "
				if attributelist<>"" then sSQL=sSQL&"OR (cpaType=3 AND cpaAssignment IN ('"&replace(trim(attributelist)," ","','")&"')) "
				sSQL=sSQL & "OR (cpaType=1 AND cpaAssignment IN ('"&replace(topcpnids,",","','")&"')))" & _
					" AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"
				rs2.Open sSQL,cnn,0,1
				do while NOT rs2.EOF
					if rs2("cpnType")=1 then ' Flat Rate Discount
						thedisc=cdbl(rs2("cpnDiscount")) * timesapply(cdalldata(2,index),cdalldata(1,index),rs2("cpnQuantity"),rs2("cpnThreshold"),rs2("cpnQuantityRepeat"),rs2("cpnThresholdRepeat"))
						if cdalldata(1,index) < thedisc then thedisc=cdalldata(1,index)
						call addadiscount(rs2, FALSE, thedisc, subcpns, cdcpncode, statetaxhandback, countrytaxhandback, cdalldata(5,index), cdalldata(6,index))
					elseif rs2("cpnType")=2 then ' Percentage Discount
						call addadiscount(rs2, FALSE, ((cdbl(rs2("cpnDiscount")) * cdbl(cdalldata(1,index))) / 100.0), subcpns, cdcpncode, statetaxhandback, countrytaxhandback, cdalldata(5,index), cdalldata(6,index))
					end if
					rs2.movenext
				loop
				rs2.Close
			Next
		end if
		sSQL="SELECT DISTINCT cpnID,cpnDiscount,cpnType,cpnNumber,cpnName,cpnSitewide,cpnThreshold,cpnThresholdMax,cpnQuantity,cpnQuantityMax,cpnThresholdRepeat,cpnQuantityRepeat FROM coupons WHERE cpnNumAvail>0 AND cpnStartDate<=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND cpnEndDate>=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND (cpnIsCoupon=0"
		if cdcpncode<>"" then sSQL=sSQL & " OR (cpnIsCoupon=1 AND cpnNumber='"&cdcpncode&"')"
		sSQL=sSQL & ") AND cpnThreshold<="&(cdtotprice)&" AND cpnQuantity<="&(cdtotquant)&" AND (cpnSitewide=1 OR cpnSitewide=3) AND (cpnType=1 OR cpnType=2)" & _
			" AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			if rs("cpnSitewide")=3 then
				sSQL="SELECT cpaType,cpaAssignment FROM cpnassign WHERE (cpaType=1 OR cpaType=3) AND cpacpnID=" & rs("cpnID")
				rs2.Open sSQL,cnn,0,1
				secids="" : attributeids=""
				do while NOT rs2.EOF
					if rs2("cpaType")=1 then secids=secids & rs2("cpaAssignment")&" " else attributeids=attributeids & rs2("cpaAssignment")&" "
					rs2.movenext
				loop
				rs2.Close
				if secids<>"" OR attributeids<>"" then
					sectionidsql=" AND ("
					if secids<>"" then sectionidsql=sectionidsql&"products.pSection IN (" & getsectionids(replace(trim(secids)," ",","), FALSE) & ")"
					if attributeids<>"" then sectionidsql=sectionidsql&IIfVs(secids<>""," OR ")&"multisearchcriteria.mSCscID IN (" & replace(trim(attributeids)," ",",") & ")"
					sectionidsql=sectionidsql&")"
				else
					sectionidsql="notassigned"
				end if
			else
				sectionidsql=""
			end if
			totprice=0 : totquant=0
			if sectionidsql<>"notassigned" then
				sSQL="SELECT DISTINCT cartID,cartProdPrice,cartQuantity FROM (products INNER JOIN cart ON cart.cartProdID=products.pID) LEFT JOIN multisearchcriteria ON cart.cartProdID=multisearchcriteria.mSCpID WHERE " & getsessionsql() & sectionidsql & " AND cartProdID<>'"&giftcertificateid&"' AND cartProdID<>'"&donationid&"' AND cartProdID<>'"&giftwrappingid&"' AND cartCompleted=0 AND pExemptions<64"
				rs2.open sSQL,cnn,0,1
				do while NOT rs2.EOF
					totprice=totprice+(rs2("cartProdPrice")*rs2("cartQuantity"))
					totquant=totquant+rs2("cartQuantity")
					sSQL="SELECT coPriceDiff FROM cartoptions WHERE coCartID=" & rs2("cartID")
					rs3.Open sSQL,cnn,0,1
					do while NOT rs3.EOF
						totprice=totprice+(rs3("coPriceDiff")*rs2("cartQuantity"))
						rs3.movenext
					loop
					rs3.close
					rs2.movenext
				loop
				rs2.close
			end if
			if totquant>0 AND rs("cpnThreshold")<=totprice AND (rs("cpnThresholdMax")>totprice OR rs("cpnThresholdMax")=0) AND rs("cpnQuantity")<=totquant AND (rs("cpnQuantityMax")>totquant OR rs("cpnQuantityMax")=0) then
				if rs("cpnType")=1 then ' Flat Rate Discount
					thedisc=cdbl(rs("cpnDiscount")) * timesapply(totquant,totprice,rs("cpnQuantity"),rs("cpnThreshold"),rs("cpnQuantityRepeat"),rs("cpnThresholdRepeat"))
					if totprice < thedisc then thedisc=totprice
				elseif rs("cpnType")=2 then ' Percentage Discount
					thedisc=((cdbl(rs("cpnDiscount")) * cdbl(totprice)) / 100.0)
				end if
				call addadiscount(rs, TRUE, thedisc, subcpns, cdcpncode, statetaxhandback, countrytaxhandback, 3, 0)
				if perproducttaxrate AND cdtotprice>0 then
					if IsArray(cdalldata) then
						for index=0 to UBOUND(cdalldata,2)
							applicdisc=0
							if rs("cpnType")=1 AND cdalldata(2,index)>0 then ' Flat Rate Discount
								applicdisc=thedisc / (cdtotquant / cdalldata(2,index))
							elseif rs("cpnType")=2 AND cdalldata(1,index)>0 then ' Percentage Discount
								applicdisc=thedisc / (cdtotprice / cdalldata(1,index))
							end if
							if (cdalldata(5,index) AND 2)<>2 then countryTax=countryTax - ((applicdisc * cdalldata(6,index)) / 100.0)
						next
					end if
				end if
			end if
			rs.movenext
		loop
		rs.close
		Session.LCID=saveLCID
	end if
	if statetaxfree < 0 then statetaxfree=0
	if countrytaxfree < 0 then countrytaxfree=0
	totaldiscounts=vsround(totaldiscounts, 2)
end sub
sub calculateshippingdiscounts(subcpns)
	Session.LCID=1033
	freeshipamnt=0
	if allfreeshipexempt then freeshipmethodexists=FALSE
	SESSION("tofreeshipquant")=empty
	SESSION("tofreeshipamount")=empty
	if NOT nodiscounts then
		sSQL="SELECT cpnID,cpnName,cpnNumber,cpnDiscount,cpnThreshold,cpnCntry,cpnHandling,cpnInsurance FROM coupons WHERE cpnType=0 AND cpnSitewide=1 AND cpnNumAvail>0 AND cpnThreshold<="&(totalgoods-(shipdiscountexempt+IIfVr(shippingafterproductdiscounts,totaldiscounts,0)))&" AND (cpnThresholdMax>"&(totalgoods-(shipdiscountexempt+IIfVr(shippingafterproductdiscounts,totaldiscounts,0)))&" OR cpnThresholdMax=0) AND cpnQuantity<="&(totalquantity-numshipdiscountexempt)&" AND (cpnQuantityMax>"&(totalquantity-numshipdiscountexempt)&" OR cpnQuantityMax=0) AND cpnStartDate<=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND cpnEndDate>=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND (cpnIsCoupon=0 OR (cpnIsCoupon=1 AND cpnNumber='"&rgcpncode&"'))" & _
			" AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			if freeshipavailtodestination OR cint(rs("cpnCntry"))=0 then
				if rgcpncode<>"" AND LCase(trim(rs("cpnNumber")))=LCase(rgcpncode) then gotcpncode=TRUE : appliedcouponname=rs("cpnName")
				if freeshipmethodexists then
					if fromshipselector then
						if intShipping(4,selectedshiptype)=1 then
							freeshipamnt=intShipping(2,selectedshiptype) - intShipping(7,selectedshiptype)
							if InStr(cpnmessage,"<br>" & rs("cpnName") & "<br>")=0 then cpnmessage=cpnmessage & rs("cpnName") & "<br>"
						end if
					else
						if intShipping(3,selectedshiptype) then freeshipamnt=intShipping(2,selectedshiptype) - intShipping(7,selectedshiptype)
						if InStr(cpnmessage,"<br>" & rs("cpnName") & "<br>")=0 then cpnmessage=cpnmessage & rs("cpnName") & "<br>"
					end if
					if cint(rs("cpnHandling"))<>0 then freeshippingincludeshandling=TRUE : handling=0 : handlingchargepercent=0
					if cint(rs("cpnInsurance"))<>0 then freeshippingincludesservices=TRUE
					if subcpns then
						Set theres=cnn.execute("SELECT cpnID FROM coupons WHERE cpnNumAvail>0 AND cpnNumAvail<30000000 AND cpnID=" & rs("cpnID") & " AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))")
						if NOT theres.EOF then SESSION("couponapply")=SESSION("couponapply") & "," & rs("cpnID")
						ect_query("UPDATE coupons SET cpnNumAvail=cpnNumAvail-1 WHERE cpnNumAvail>0 AND cpnNumAvail<30000000 AND cpnID=" & rs("cpnID"))
					end if
					freeshippingapplied=TRUE
				end if
			end if
			rs.movenext
		loop
		rs.close
	end if
	if somethingToShip AND NOT fromshipselector then
		gotshipping=FALSE
		if shipType>=1 then
			if shipType=2 OR shipType=5 then sortshippingarray()
			for index=0 to UBOUND(intShipping,2)
				if intShipping(3,index)=TRUE then
					if NOT gotshipping OR (intShipping(4,index) AND freeshippingapplied) then
						shipping=intShipping(2,index)
						shipMethod=intShipping(0,index)
						selectedshiptype=index
						gotshipping=TRUE
					end if
					if intShipping(4,index) AND freeshippingapplied then freeshipamnt=intShipping(2,index) - (intShipping(7,index) + getfreeshipinsurance(indexmso))
				end if
			next
		end if
		if NOT freeshippingapplied AND freeshipmethodexists then
			sSQL="SELECT MIN(cpnQuantity) AS minquant,MIN(cpnThreshold) as minthreshold FROM coupons WHERE cpnType=0 AND cpnSitewide=1 AND cpnNumAvail>0 AND (cpnThresholdMax>"&(totalgoods-(shipdiscountexempt+IIfVr(shippingafterproductdiscounts,totaldiscounts,0)))&" OR cpnThresholdMax=0) AND (cpnQuantityMax>"&(totalquantity-numshipdiscountexempt)&" OR cpnQuantityMax=0) AND cpnStartDate<=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND cpnEndDate>=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND (cpnIsCoupon=0 OR (cpnIsCoupon=1 AND cpnNumber='"&rgcpncode&"'))" & _
				" AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"
			rs.open sSQL,cnn,0,1
			if NOT isnull(rs("minquant")) then
				if rs("minquant")-(totalquantity-numshipdiscountexempt)>0 then SESSION("tofreeshipquant")=rs("minquant")-(totalquantity-numshipdiscountexempt)
				if rs("minthreshold")-(totalgoods-(shipdiscountexempt+IIfVr(shippingafterproductdiscounts,totaldiscounts,0)))>0 then SESSION("tofreeshipamount")=rs("minthreshold")-(totalgoods-(shipdiscountexempt+IIfVr(shippingafterproductdiscounts,totaldiscounts,0)))
			end if
			rs.close
		end if
	end if
	if freeshipamnt>shipping then freeshipamnt=shipping
	Session.LCID=saveLCID
end sub
function getshiptype()
	if adminIntShipping<>0 AND shipcountry<>origCountry AND NOT ((shipCountryCode="US" OR shipCountryCode="CA") AND usandcausedomesticservice=TRUE) then
		if cartisincluded=TRUE then
			shipType=adminIntShipping
		elseif getpost("altrates")="" then
			shipType=adminIntShipping
		end if
	end if
	getshiptype=shipType
end function
sub getadminshippingparams()
	sSQL="SELECT adminPacking,AusPostAPI,adminCanPostUser,adminCanPostLogin,adminCanPostPass,adminUSPSUser,smartPostHub,adminUPSUser,adminUPSpw,adminUPSAccess,adminUPSAccount,adminUPSNegotiated,FedexAccountNo,FedexMeter,FedexUserKey,FedexUserPwd,DHLSiteID,DHLSitePW,DHLAccountNo FROM adminshipping WHERE adminShipID=1"
	rs.open sSQL,cnn,0,1
	AusPostAPI=trim(rs("AusPostAPI")&"")
	adminCanPostLogin=trim(rs("adminCanPostLogin"))
	adminCanPostPass=trim(rs("adminCanPostPass"))
	adminCanPostUser=trim(rs("adminCanPostUser"))
	packtogether=int(rs("adminPacking"))=1
	uspsUser=rs("adminUSPSUser")
	smartPostHub=rs("smartPostHub")
	upsUser=upsdecode(rs("adminUPSUser"), "")
	upsPw=upsdecode(rs("adminUPSpw"), "")
	upsAccess=rs("adminUPSAccess")
	upsAccount=rs("adminUPSAccount")
	upsnegdrates=(cint(rs("adminUPSNegotiated"))<>0)
	fedexaccount=rs("FedexAccountNo")
	fedexmeter=rs("FedexMeter")
	fedexuserkey=rs("FedexUserKey")
	fedexuserpwd=rs("FedexUserPwd")
	DHLSiteID=rs("DHLSiteID")
	DHLSitePW=rs("DHLSitePW")
	DHLAccountNo=rs("DHLAccountNo")
	rs.close
end sub
function initshippingmethods()
	initshippingmethods=TRUE
	if shipcountry<>origCountry AND NOT (shipType=3 AND shipCountryCode="PR") then international="Intl" : willpickuptext="" : willpickup_=FALSE
	if willpickup_ then
		shipType=0
		adminAltRates=0
		addshippinginsurance=0
		if willpickupcost<>"" then shipping=willpickupcost else shipping=0
		shipMethod=willpickuptext
		if willpickupnohandling then handlingchargepercent=0 : handling=0
	end if
	if adminAltRates>0 then
		rs.open "SELECT COUNT(*) AS tcnt FROM alternaterates WHERE usealtmethod"&international&"<>0",cnn,0,1
		if rs("tcnt")<2 then adminAltRates=0
		rs.close
	end if
	if altrate<>"" AND adminAltRates>0 then
		rs.open "SELECT altrateid FROM alternaterates WHERE usealtmethod"&international&"<>0 AND altrateid="&altrate,cnn,0,1
		if NOT rs.EOF then shipType=altrate
		rs.close
	end if
	if initialpackweight<>"" AND shipType<>5 then thepweight=initialpackweight
	for index3=0 to UBOUND(intShipping,2)
		intShipping(0,index3)="" ' Name
		intShipping(1,index3)="" ' Delivery
		intShipping(2,index3)=0 ' Cost
		intShipping(3,index3)=FALSE ' Used
		intShipping(4,index3)=0 ' FSA
		intShipping(5,index3)="" ' Service ID (USPS)
		intShipping(6,index3)=shipType ' shipType
		intShipping(7,index3)=0 ' Cost for Free Ship Exempt
		intShipping(8,index3)=0 ' Cost for Insurance + Services
	next
	if fromshipselector then
		if is_numeric(getpost("orderid")) AND is_numeric(getpost("shipselectoridx")) then
			numshipoptions=0
			sSQL="SELECT soMethodName,soCost,soFreeShipExempt,soFreeShip,soShipType,soDeliveryTime FROM shipoptions WHERE soOrderID=" & getpost("orderid") & " ORDER BY soIndex"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				intShipping(0,numshipoptions)=rs("soMethodName")
				intShipping(1,numshipoptions)=rs("soDeliveryTime")
				intShipping(2,numshipoptions)=rs("soCost")
				intShipping(3,numshipoptions)=TRUE
				intShipping(4,numshipoptions)=rs("soFreeShip")
				freeshipmethodexists=(freeshipmethodexists OR intShipping(4,numshipoptions))
				intShipping(6,numshipoptions)=rs("soShipType")
				intShipping(7,numshipoptions)=rs("soFreeShipExempt")
				numshipoptions=numshipoptions+1
				rs.movenext
			loop
			rs.close
			selectedshiptype=int(getpost("shipselectoridx"))
			shipping=intShipping(2,selectedshiptype)
			shipMethod=intShipping(0,selectedshiptype)
			shipType=intShipping(6,selectedshiptype)
			currShipType=intShipping(6,0)
			multipleoptions=TRUE
			numshipoptions=numshipoptions-1
		end if
	elseif shipType=1 then '  Flat rate shipping
		intShipping(0,0)=IIfVr(combineshippinghandling, xxShipHa, xxShippg)
		intShipping(3,0)=TRUE
		intShipping(4,0)=1
	elseif shipType=2 OR shipType=5 then ' Weight / Price based shipping
		allzones=""
		zoneid=0
		if splitUSZones AND shiphomecountry AND is_numeric(shipCountryID) then
			if paypalexpress AND shipCountryID=201 then
				shipstate=replace(replace(replace(replace(replace(replace(shipstate,"Argyll and Bute","Argyll"),"Greater Manchester","Manchester"),"East Riding of Yorkshire","Yorkshire"),"South Ayrshire","Ayrshire"),"North Ayrshire","Ayrshire"),"East Ayrshire","Ayrshire")
				shipstate=replace(replace(replace(replace(replace(replace(shipstate,"Armagh","County Armagh"),"Antrim","County Antrim"),"Glasgow (City of)","Glasgow"),"Clackmannan","Clackmannanshire"),"Aberdeen City","Aberdeenshire"),"Edinburgh City","Edinburgh")
				shipstate=replace(replace(replace(replace(replace(replace(shipstate,"North East Lincolnshire","Lincolnshire"),"Humberside","North Humberside"),"Tyrone","County Tyrone"),"Londonderry","County Londonderry"),"Fermanagh","County Fermanagh"),"Down","County Down")
				shipstate=replace(replace(replace(replace(replace(replace(shipstate,"South Lanarkshire","Lanarkshire"),"Perthshire and Kinross","Perthshire"),"North Lanarkshire","Lanarkshire"),"East Renfrewshire","Renfrewshire"),"East Dunbartonshire","Dunbartonshire"),"Dumfries and Galloway","Dumfriesshire")
				shipstate=replace(replace(replace(replace(replace(replace(shipstate,"Caerphilly","Glamorgan"),"Merthyr Tydfil","Glamorgan"),"Isle of Anglesey","Anglesey"),"Blaenau Gwent","Gwent"),"West Dunbartonshire","Dunbartonshire"),"Stirling","Stirlingshire")
				shipstate=replace(replace(replace(replace(replace(replace(shipstate,"Wrexham","Clwyd"),"The Vale of Glamorgan","Glamorgan"),"Torfaen","Monmouthshire"),"Swansea","Glamorgan"),"Newport","Gwent"),"Conwy","Clwyd")
				shipstate=replace(replace(replace(replace(replace(shipstate,"Falkirk","Stirlingshire"),"Inverclyde","Renfrewshire"),"Western Isles","Inverness-shire"),"Bridgend","Mid Glamorgan"),"Neath Port Talbot","West Glamorgan")
			end if
			sSQL="states INNER JOIN postalzones ON postalzones.pzID=states.stateZone WHERE stateCountryID=" & shipCountryID & " AND (stateName='"&escape_string(shipstate)&"' OR stateAbbrev='"&escape_string(shipstate)&"')"
		else
			sSQL="countries INNER JOIN postalzones ON postalzones.pzID=countries.countryZone WHERE countryName='" & escape_string(shipcountry) & "'"
		end if
		rs.open "SELECT pzID,pzMultiShipping,pzFSA,pzMethodName1,pzMethodName2,pzMethodName3,pzMethodName4,pzMethodName5 FROM "&sSQL,cnn,0,1
		if NOT (rs.EOF OR rs.BOF) then
			zoneid=rs("pzID")
			numshipoptions=rs("pzMultiShipping")
			for index3=0 to numshipoptions
				intShipping(0,index3)=rs("pzMethodName"&(index3+1))
				intShipping(3,index3)=TRUE
				intShipping(4,index3)=IIfVr((rs("pzFSA") AND (2 ^ index3))<>0, 1, 0)
			next
		else
			success=FALSE
			if splitUSZones AND shiphomecountry AND shipstate="" then errormsg=xxPlsSta else errormsg="Country / state shipping zone is unassigned."
			returntocustomerdetails=TRUE
			initshippingmethods=FALSE
		end if
		rs.close
		sSQL="SELECT zcWeight,zcRate,zcRate2,zcRate3,zcRate4,zcRate5,zcRatePC,zcRatePC2,zcRatePC3,zcRatePC4,zcRatePC5 FROM zonecharges WHERE zcZone="&zoneid&" ORDER BY zcWeight"
		rs.open sSQL,cnn,0,1
		if NOT (rs.EOF OR rs.BOF) then allzones=rs.getrows
		rs.close
	elseif shipType=3 OR shipType=4 OR shipType>=6 then ' USPS / UPS / Canada Post / Fedex / SmartPost / DHL
		if shipType=3 then
			sSQL=" FROM uspsmethods WHERE uspsID<100 AND uspsLocal="&IIfVr(international="","1","0")
		elseif shipType=4 then
			sSQL=" FROM uspsmethods WHERE uspsID>100 AND uspsID<200"
		elseif shipType=6 then
			sSQL=" FROM uspsmethods WHERE uspsID>200 AND uspsID<300"
		elseif shipType=7 then
			sSQL=",uspsLocal FROM uspsmethods WHERE uspsID>300 AND uspsID<400"&IIfVr(international="" AND commercialloc_, " AND uspsMethod<>'GROUNDHOMEDELIVERY'", "")
		elseif shipType>=8 then
			sSQL=",uspsLocal FROM uspsmethods WHERE uspsID>"&(shipType-4)&"00 AND uspsID<"&(shipType-3)&"00" & IIfVs(shipType=10, " AND uspsLocal="&IIfVr(international="","1","0"))
		end if
		rs.open "SELECT uspsMethod,uspsFSA,uspsShowAs"&sSQL&" AND uspsUseMethod=1",cnn,0,1
		if NOT rs.EOF then
			uspsmethods=rs.GetRows()
		else
			success=FALSE
			errormsg="Admin Error: " & xxNoMeth
		end if
		rs.close
	end if
	if (shipType=4 OR shipType=7 OR shipType=8) AND shipCountryCode="US" AND shipStateAbbrev="PR" then shipCountryCode="PR"
	if shipType=3 AND shipCountryCode="PR" then shipCountryCode="US" : shipStateAbbrev="PR"
	if (shipCountryCode="PR" OR (shipCountryCode="US" AND shipStateAbbrev="PR")) AND len(destZip)=3 then destZip="00"&destZip
	if shipType=3 then
		sXML="<"&international&"Rate"&IIfVr(international="","V4","V2")&"Request USERID="""&uspsUser&"""><Revision>2</Revision>"
	elseif shipType=4 then
		if shipCountryCode="US" AND shipStateAbbrev="VI" then shipCountryCode="VI"
		sXML="<" & "?xml version=""1.0""?><AccessRequest xml:lang=""en-US"">" & addtag("AccessLicenseNumber",upsAccess) & addtag("UserId",upsUser) & addtag("Password",upsPw)&"</AccessRequest><" & "?xml version=""1.0""?>" & _
			"<RatingServiceSelectionRequest xml:lang=""en-US""><Request><TransactionReference><CustomerContext>Rating and Service</CustomerContext><XpciVersion>1.0001</XpciVersion></TransactionReference>" & _
			"<RequestAction>Rate</RequestAction><RequestOption>shop</RequestOption></Request>"
		if upspickuptype<>"" then sXML=sXML & "<PickupType><Code>"&upspickuptype&"</Code></PickupType>"
		sXML=sXML & "<Shipment><Shipper>" & IIfVr(upsnegdrates=TRUE, addtag("ShipperNumber",upsAccount), "") & "<Address>" & IIfVr(upsnegdrates=TRUE, addtag("StateProvinceCode",defaultshipstate),"") & addtag("PostalCode",origZip) & addtag("CountryCode",origCountryCode)&"</Address></Shipper>" & _
			"<ShipTo><Address>" & addtag("AddressLine1",IIfVr(ordShipAddress<>"",ordShipAddress,ordAddress)) & addtag("AddressLine2",IIfVr(ordShipAddress2<>"",ordShipAddress2,ordAddress2)) & addtag("City",IIfVr(ordShipCity<>"",ordShipCity,ordCity))&IIfVr(shipCountryCode="US" OR shipCountryCode="CA",addtag("StateProvinceCode",shipStateAbbrev),"") & addtag("PostalCode",destZip) & addtag("CountryCode",shipCountryCode) & IIfVr(NOT commercialloc_, "<ResidentialAddressIndicator/>", "") & "</Address></ShipTo>"
	elseif shipType=6 then
		packtogether=TRUE : splitpackat="" : nosplitlargepacks=TRUE ' Canada Post cannot handle more than one package
		sXML="<soapenv:Envelope xmlns:rate=""http://www.canadapost.ca/ws/soap/ship/rate/v2"" xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
			"<soapenv:Header><wsse:Security soapenv:mustUnderstand=""1"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"" xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd""><wsse:UsernameToken><wsse:Username>" & adminCanPostLogin & "</wsse:Username><wsse:Password>" & adminCanPostPass & "</wsse:Password></wsse:UsernameToken></wsse:Security></soapenv:Header>" & _
			"<soapenv:Body><rate:get-rates-request><platform-id>0008107483</platform-id>" & IIfVs(storelang="fr","<locale>FR</locale>") & "<mailing-scenario><customer-number>" & adminCanPostUser & "</customer-number><origin-postal-code>" & ucase(replace(origZip," ","")) & "</origin-postal-code><destination>"
		if shipCountryCode="CA" then
			sXML=sXML & "<domestic><postal-code>" & ucase(replace(destZip," ","")) & "</postal-code></domestic>"
		elseif shipCountryCode="US" then
			sXML=sXML & "<united-states><zip-code>" & destZip & "</zip-code></united-states>"
		else
			sXML=sXML & "<international><country-code>" & shipCountryCode & "</country-code></international>"
		end if
		sXML=sXML & "</destination>"
	elseif shipType=7 OR shipType=8 then ' FedEx
		if packaging<>"" then fedexpackaging="FEDEX_" & UCase(packaging) else fedexpackaging="YOUR_PACKAGING"
		if fedexpickuptype="" then fedexpickuptype="REGULAR_PICKUP"
		fedextimestamp=""
		if saturdaydelivery_=TRUE then
			fedextimestamp="2010-03-12"
		elseif noweekendshipment=TRUE then
			if weekday(date())=1 then fedextimestamp=date()+1
			if weekday(date())=7 then fedextimestamp=date()+2
			if fedextimestamp<>"" then fedextimestamp=year(fedextimestamp)&"-"&twodp(month(fedextimestamp))&"-"&twodp(day(fedextimestamp))
		end if
		if fedextimestamp<>"" then fedextimestamp="<v9:ShipTimestamp>" & fedextimestamp & "T10:00:00-04:00</v9:ShipTimestamp>"
		sXML="<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:v9=""http://fedex.com/ws/rate/v9"">" & _
		"<soapenv:Header/><soapenv:Body><v9:RateRequest><v9:WebAuthenticationDetail><v9:CspCredential><v9:Key>mKOUqSP4CS0vxaku</v9:Key><v9:Password>IAA5db3Pmhg3lyWW6naMh4Ss2</v9:Password></v9:CspCredential>" & _
		"<v9:UserCredential><v9:Key>" & fedexuserkey & "</v9:Key><v9:Password>" & fedexuserpwd & "</v9:Password></v9:UserCredential></v9:WebAuthenticationDetail>" & _
		"<v9:ClientDetail><v9:AccountNumber>" & fedexaccount & "</v9:AccountNumber><v9:MeterNumber>" & fedexmeter & "</v9:MeterNumber><v9:ClientProductId>IBTP</v9:ClientProductId><v9:ClientProductVersion>3272</v9:ClientProductVersion></v9:ClientDetail>" & _
		"<v9:TransactionDetail><v9:CustomerTransactionId>"&IIfVr(fedexctid<>"",fedexctid,"Rate Request")&"</v9:CustomerTransactionId></v9:TransactionDetail>" & _
		"<v9:Version><v9:ServiceId>crs</v9:ServiceId><v9:Major>9</v9:Major><v9:Intermediate>0</v9:Intermediate><v9:Minor>0</v9:Minor></v9:Version>" & _
		"<v9:ReturnTransitAndCommit>true</v9:ReturnTransitAndCommit>" & _
		"<v9:RequestedShipment>" & fedextimestamp & "<v9:DropoffType>"& fedexpickuptype&"</v9:DropoffType>" & IIfVr(shipType=8,"<v9:ServiceType>SMART_POST</v9:ServiceType>","<v9:PackagingType>"&fedexpackaging&"</v9:PackagingType>") & _
		"<v9:Shipper><v9:Address><v9:PostalCode>"&origZip&"</v9:PostalCode><v9:CountryCode>" & origCountryCode & "</v9:CountryCode>" & _
		"</v9:Address></v9:Shipper><v9:Recipient><v9:Address>"
		if ordShipAddress<>"" then sXML=sXML & "<v9:StreetLines>" & vrxmlencode(ordShipAddress) & "</v9:StreetLines>"
		if ordShipCity<>"" then sXML=sXML & "<v9:City>" & ordShipCity & "</v9:City>"
		if shipCountryCode="US" OR shipCountryCode="CA" then sXML=sXML & "<v9:StateOrProvinceCode>" & shipStateAbbrev & "</v9:StateOrProvinceCode>"
		sXML=sXML & "<v9:PostalCode>" & destZip & "</v9:PostalCode>" & _
		"<v9:CountryCode>" & shipCountryCode & "</v9:CountryCode><v9:Residential>"&IIfVr(commercialloc_,"false","true")&"</v9:Residential></v9:Address></v9:Recipient>"
		if shipType=8 then
			if smartpostindicia="" then smartpostindicia="PARCEL_SELECT"
			sXML=sXML & "<v9:SmartPostDetail><v9:Indicia>" & smartpostindicia & "</v9:Indicia>" & IIfVr(smartpostancendorsement<>"","<v9:AncillaryEndorsement>" & smartpostancendorsement & "</v9:AncillaryEndorsement>","") & "<v9:HubId>" & smartPostHub & "</v9:HubId></v9:SmartPostDetail>"
		else
			sXML=sXML & "<v9:SpecialServicesRequested>"
			if saturdaydelivery_=TRUE then sXML=sXML & "<v9:SpecialServiceTypes>SATURDAY_DELIVERY</v9:SpecialServiceTypes>"
			if saturdaypickup=TRUE then sXML=sXML & "<v9:SpecialServiceTypes>SATURDAY_PICKUP</v9:SpecialServiceTypes>"
			if insidedelivery_=TRUE then sXML=sXML & "<v9:SpecialServiceTypes>INSIDE_DELIVERY</v9:SpecialServiceTypes>"
			if insidepickup=TRUE then sXML=sXML & "<v9:SpecialServiceTypes>INSIDE_PICKUP</v9:SpecialServiceTypes>"
			if emailnotification=TRUE then sXML=sXML & "<v9:SpecialServiceTypes>EMAIL_NOTIFICATION</v9:SpecialServiceTypes>"
			if homedelivery_<>"" then sXML=sXML & "<v9:SpecialServiceTypes>HOME_DELIVERY_PREMIUM</v9:SpecialServiceTypes><v9:HomeDeliveryPremiumDetail><v9:HomeDeliveryPremiumType>" & homedelivery_ & "</v9:HomeDeliveryPremiumType></v9:HomeDeliveryPremiumDetail>"
			if ordPayProvider<>"" then
				'if int(ordPayProvider)=codpaymentprovider then sXML=sXML & "<v9:SpecialServiceTypes>COD</v9:SpecialServiceTypes><v9:CodDetail><v9:CodCollectionAmount><v9:Currency>CAD</v9:Currency><v9:Amount>XXXFEDEXGRANDTOTXXX</v9:Amount></v9:CodCollectionAmount><v9:CollectionType>ANY</v9:CollectionType></v9:CodDetail>"
			end if
			if holdatlocation=TRUE then sXML=sXML & "<v9:SpecialServiceTypes>HOLD_AT_LOCATION</v9:SpecialServiceTypes><v9:HoldAtLocationDetail><v9:PhoneNumber>9052125251</v9:PhoneNumber><v9:LocationContactAndAddress><v9:Address><v9:StreetLines>HAL Address Line 1</v9:StreetLines><v9:City>St-Laurent</v9:City><v9:StateOrProvinceCode>QC</v9:StateOrProvinceCode><v9:PostalCode>H4T2A3</v9:PostalCode><v9:CountryCode>CA</v9:CountryCode></v9:Address></v9:LocationContactAndAddress></v9:HoldAtLocationDetail>"
			sXML=sXML & "</v9:SpecialServicesRequested>"
			sXML=sXML & "<v9:CustomsClearanceDetail>" & IIfVr(customsaccountnumber<>"","<v9:DutiesPayment><v9:PaymentType>SENDER</v9:PaymentType></v9:DutiesPayment>","") & "<v9:CustomsValue><v9:Currency>"&countryCurrency&"</v9:Currency><v9:Amount>XXXFEDEXGRANDTOTXXX</v9:Amount></v9:CustomsValue></v9:CustomsClearanceDetail>"
		end if
		sXML=sXML & "<v9:RateRequestTypes>ACCOUNT</v9:RateRequestTypes><v9:PackageDetail>INDIVIDUAL_PACKAGES</v9:PackageDetail>"
	elseif shipType=9 then ' DHL
		themon=DatePart("m",date()+1)
		theday=DatePart("d",date()+1)
		sXML="<?xml version=""1.0"" encoding=""utf-8"" ?><q1:DCTRequest xmlns:q1=""http://www.dhl.com""><GetQuote>" & _
		"<Request><ServiceHeader><SiteID>" & DHLSiteID & "</SiteID><Password>" & DHLSitePW & "</Password></ServiceHeader></Request>" & _
		"<From><CountryCode>" & origCountryCode & "</CountryCode><Postalcode>" & origZip & "</Postalcode></From>" & _
		"<BkgDetails><PaymentCountryCode>" & origCountryCode & "</PaymentCountryCode>" & _
		"<Date>" & DatePart("yyyy",date()+1) & "-" & IIfVs(len(themon)<2,"0") & themon & "-" & IIfVs(len(theday)<2,"0") & theday & "</Date><ReadyTime>PT9H</ReadyTime>" & _
		"<DimensionUnit>"&IIfVr((adminUnits AND 12)=4 OR ((adminUnits AND 12)=0 AND (adminUnits AND 1)=1),"IN","CM")&"</DimensionUnit><WeightUnit>"&IIfVr((adminUnits AND 1)=1,"LB","KG")&"</WeightUnit><Pieces>"
	elseif shipType=10 then ' Australia Post
		packtogether=TRUE : splitpackat="" : nosplitlargepacks=TRUE ' Australia Post cannot handle more than one package
	end if
end function
packdims=Array(0,0,0,0,0,0,0,0) ' len : wid : hei : vol : maxlen : maxwid : maxhei : items
sub zeropackdims()
	for zpd=0 to 7
		packdims(zpd)=0
	next
	thepweight=0
end sub
sub reorderpackagedimensions()
	if packdims(2)>packdims(1) then apdtemp=packdims(1) : packdims(1)=packdims(2) : packdims(2)=apdtemp
	if packdims(1)>packdims(0) then apdtemp=packdims(0) : packdims(0)=packdims(1) : packdims(1)=apdtemp
	if packdims(2)>packdims(1) then apdtemp=packdims(1) : packdims(1)=packdims(2) : packdims(2)=apdtemp
end sub
sub reorderproddims(byref pdims)
	pdims(0)=cdbl(pdims(0)) : pdims(1)=cdbl(pdims(1)) : pdims(2)=cdbl(pdims(2))
	if pdims(2)>pdims(1) then apdtemp=pdims(1) : pdims(1)=pdims(2) : pdims(2)=apdtemp
	if pdims(1)>pdims(0) then apdtemp=pdims(0) : pdims(0)=pdims(1) : pdims(1)=apdtemp
	if pdims(2)>pdims(1) then apdtemp=pdims(1) : pdims(1)=pdims(2) : pdims(2)=apdtemp
end sub
sub addpackagedimensions(dimens, apdquant)
	Session.LCID=1033
	if (adminUnits AND 12)<>0 then
		origdimens=packdims
		proddims=split(trim(dimens&""), "x")
		if UBOUND(proddims)>=2 then
			if proddims(0)<>"" AND proddims(1)<>"" AND proddims(2)<>"" then
				reorderproddims proddims
				if proddims(0)>packdims(4) then packdims(4)=proddims(0)
				if proddims(1)>packdims(5) then packdims(5)=proddims(1)
				if proddims(2)>packdims(6) then packdims(6)=proddims(2)
				proddims(2)=proddims(2) * apdquant
				thelength=proddims(0)
				reorderproddims proddims
				do while apdquant>4 AND proddims(0)>proddims(2) * 2 AND proddims(0)>thelength
					proddims(0)=proddims(0) / 2 : proddims(2)=proddims(2) * 2 : apdquant=apdquant / 2
					reorderproddims proddims
				loop
				thelength=proddims(0) : thewidth=proddims(1) : theheight=proddims(2)
				objvol=thelength * thewidth * theheight
				if thelength>packdims(0) then packdims(0)=thelength
				if thewidth>packdims(1) then packdims(1)=thewidth
				if theheight>packdims(2) then packdims(2)=theheight
				if objvol + packdims(3)>packdims(0) * packdims(1) * packdims(2) then packdims(2)=packdims(2) + IIfVr(origdimens(2)>0 AND origdimens(2) < theheight, origdimens(2),theheight)
				if objvol + packdims(3)>packdims(0) * packdims(1) * packdims(2) then packdims(1)=packdims(1) + IIfVr(origdimens(1)>0 AND origdimens(1) < thewidth, origdimens(1),thewidth)
				if objvol + packdims(3)>packdims(0) * packdims(1) * packdims(2) then packdims(0)=packdims(0) + IIfVr(origdimens(0)>0 AND origdimens(0) < thelength, origdimens(0),thelength)
				packdims(3)=packdims(3) + objvol
				reorderpackagedimensions()
			end if
		end if
	end if
	packdims(7)=packdims(7) + apdquant
	' print "Bin is len: " & packdims(0)&" wid:"& packdims(1)&" hei:"& packdims(2)&" vol:" & packdims(3) & " maxlen: " & packdims(4)&" maxwid:"& packdims(5)&" maxhei:"& packdims(6)&" items:" & packdims(7) & "<br>"
	Session.LCID=saveLCID
end sub
function splitlargepacks()
	splitlargepacks=1
	if packdims(7)<=1 AND nosplitlargepacks then exit function
	slpnumpacks=1
	slpweight=thepweight
	if shipType=6 then
		if (adminUnits AND 12)=4 then maxlenplusgirth=118 : maxlength=78 else maxlenplusgirth=300 : maxlength=200
		if (adminUnits AND 3)=1 then maxweight=66 else maxweight=30
		if adminCanPostLogin<>"" then nosplitlargepacks=TRUE
	elseif shipType=4 OR shipType=7 OR shipType=8 OR shipType=9 then
		if (adminUnits AND 12)=4 then maxlenplusgirth=165 : maxlength=108 else maxlenplusgirth=419 : maxlength=274
		if (adminUnits AND 3)=1 then maxweight=150 else maxweight=68
	else ' USPS Default
		if (adminUnits AND 12)=8 then maxlenplusgirth=330 else maxlenplusgirth=130
		maxlength=0
		if (adminUnits AND 3)=1 then maxweight=70 else maxweight=31
	end if
	if is_numeric(splitpackat) then maxweight=cdbl(splitpackat)
	if nosplitlargepacks<>TRUE AND (adminUnits AND 12)<>0 then
		if packdims(0) + ((packdims(1) + packdims(2)) * 2)>maxlenplusgirth then ' Max Length + Girth
			divisor=1
			do while (packdims(0)/sqr(divisor)) + (((packdims(1)/sqr(divisor)) + packdims(2)) * 2)>maxlenplusgirth
				divisor=divisor+1
			loop
			if packdims(0)/sqr(divisor)>maxlength AND maxlength<>0 AND (packdims(0)/divisor) + ((packdims(1) + packdims(2)) * 2)<=maxlenplusgirth then
				packdims(0)=packdims(0)/divisor
			else
				packdims(0)=packdims(0)/sqr(divisor)
				packdims(1)=packdims(1)/sqr(divisor)
			end if
			slpnumpacks=slpnumpacks * divisor
			slpweight=slpweight / divisor
			reorderpackagedimensions()
		end if
		if packdims(0)>maxlength AND maxlength<>0 then
			packdims(0)=packdims(0) / 2
			slpnumpacks=slpnumpacks * 2
			slpweight=slpweight / 2
			reorderpackagedimensions()
		end if
	end if
	if slpweight>maxweight AND (packdims(7)<=slpnumpacks OR nosplitlargepacks<>TRUE) then
		slpindex=2
		do while TRUE
			if slpweight / slpindex<=maxweight then
				packdims(0)=packdims(0) / slpindex
				slpnumpacks=slpnumpacks * slpindex
				reorderpackagedimensions()
				exit do
			end if
			slpindex=slpindex+1
		loop
	end if
	splitlargepacks=slpnumpacks
end function
packageweight=0
packagefreeexemptweight=0
function islastpacktogether(apsrs,prodindex)
	islastpacktogether=TRUE
	for tindex=prodindex+1 to UBOUND(apsrs,2)
		if (apsrs(8,tindex) AND 32)<>32 then islastpacktogether=FALSE
	next
end function
sub addproducttoshipping(apsrs, prodindex)
	shipThisProd=TRUE
	if (apsrs(8,prodindex) AND 32)=32 then
		savepacktogether=packtogether
		savethepweight=thepweight
		savepackageweight=packageweight
		savepackdims=packdims
		zeropackdims()
		packtogether=FALSE
	end if
	if (apsrs(8,prodindex) AND 4)=4 then ' No Shipping on this product
		shipThisProd=FALSE
	else
		call addpackagedimensions(apsrs(11,prodindex), IIfVr(packtogether, apsrs(4,prodindex), 1))
	end if
	if (apsrs(8,prodindex) AND 16)<>16 then allfreeshipexempt=FALSE
	if fromshipselector then
	elseif shipType=1 then ' Flat rate shipping
		if shipThisProd then
			intShipping(2,0)=intShipping(2,0) + apsrs(6,prodindex) + (apsrs(7,prodindex) * (apsrs(4,prodindex)-1))
			if (apsrs(8,prodindex) AND 16)=16 then intShipping(7,0)=intShipping(7,0) + apsrs(6,prodindex) + (apsrs(7,prodindex) * (apsrs(4,prodindex)-1))
			somethingToShip=TRUE
		end if
	elseif shipType=2 OR shipType=5 then ' Weight / Price based shipping
		havematch=FALSE
		for index3=0 to numshipoptions
			dHighest(index3)=0
		next
		if IsArray(allzones) then
			if shipThisProd then
				somethingToShip=TRUE
				if shipType=2 then tmpweight=cdbl(apsrs(5,prodindex)) else tmpweight=cdbl(apsrs(3,prodindex))
				if packtogether then
					thepweight=thepweight + (cdbl(apsrs(4,prodindex))*tmpweight)
					thePQuantity=1
				else
					thepweight=tmpweight + IIfVr(initialpackweight<>"",initialpackweight,0)
					thePQuantity=cdbl(apsrs(4,prodindex))
				end if
				packageweight=packageweight + (cdbl(apsrs(4,prodindex))*tmpweight)
				if (apsrs(8,prodindex) AND 16)=16 then packagefreeexemptweight=packagefreeexemptweight + (cdbl(apsrs(4,prodindex))*tmpweight)
			end if
			if ((NOT packtogether AND shipThisProd) OR (packtogether AND islastpacktogether(apsrs,prodindex))) AND (thepweight>0 OR shipType=5) then ' Only calculate pack together when we have the total
				for index2=0 to UBOUND(allzones,2)
					if allzones(0,index2)>=thepweight then
						havematch=TRUE
						for index3=0 to numshipoptions
							if cint(allzones(6+index3,index2))<>0 then ' by percentage
								intShipping(2,index3)=intShipping(2,index3)+((cdbl(allzones(1+index3,index2))*thePQuantity*thepweight)/100.0)
							else
								intShipping(2,index3)=intShipping(2,index3)+(cdbl(allzones(1+index3,index2))*thePQuantity)
							end if
							if cdbl(allzones(1+index3,index2))=-99999.0 then intShipping(3,index3)=FALSE
							if shipCountryCode=countryCode AND saturdaydelivery_ AND royalmail then
								if index3=2 OR index3=3 then
									if index3=2 then intShipping(2,index3)=intShipping(2,index3)*1.2
									intShipping(2,index3)=intShipping(2,index3)+3
								else
									intShipping(3,index3)=FALSE
								end if
							end if
						next
						exit for
					end if
					dHighWeight=allzones(0,index2)
					for index3=0 to numshipoptions
						if cint(allzones(6+index3,index2))<>0 then ' by percentage
							dHighest(index3)=(allzones(1+index3,index2)*dHighWeight)/100.0
						else
							dHighest(index3)=allzones(1+index3,index2)
						end if
					next
				next
				if NOT havematch then
					for index3=0 to numshipoptions
						intShipping(2,index3)=intShipping(2,index3) + dHighest(index3)
						if dHighest(index3)=-99999.0 then intShipping(3,index3)=FALSE
					next
					if allzones(0,0) < 0 then
						dHighWeight=thepweight - dHighWeight
						do while dHighWeight>0
							for index3=0 to numshipoptions
								intShipping(2,index3)=intShipping(2,index3) + (cdbl(allzones(1+index3,0))*thePQuantity)
							next
							dHighWeight=vsround(dHighWeight + allzones(0,0),4)
						loop
					end if
				end if
			end if
		end if
	elseif shipType=3 then ' USPS Shipping
		if packtogether then
			if shipThisProd then
				somethingToShip=TRUE
				thepweight=thepweight + (cdbl(apsrs(5,prodindex)) * int(apsrs(4,prodindex)))
				packageweight=packageweight + (cdbl(apsrs(5,prodindex)) * int(apsrs(4,prodindex)))
				if (apsrs(8,prodindex) AND 16)=16 then packagefreeexemptweight=packagefreeexemptweight+(cdbl(apsrs(5,prodindex)) * int(apsrs(4,prodindex)))
			end if
			if islastpacktogether(apsrs,prodindex) AND thepweight>0 then
				numpacks=splitlargepacks()
				if international<>"" then
					sXML=sXML & addUSPSInternational(rowcounter,thepweight / numpacks,numpacks,"Package",shipcountry,totalgoods-shipfreegoods)
				else
					sXML=sXML & addUSPSDomestic(rowcounter,"Parcel",origZip,destZip,thepweight / numpacks,numpacks,"REGULAR","True",totalgoods-shipfreegoods)
				end if
				rowcounter=rowcounter + 1
				zeropackdims()
			end if
		else
			if shipThisProd then
				somethingToShip=TRUE
				thepweight=apsrs(5,prodindex) + IIfVr(initialpackweight<>"",initialpackweight,0)
				packageweight=packageweight + thepweight
				if (apsrs(8,prodindex) AND 16)=16 then packagefreeexemptweight=packagefreeexemptweight+thepweight
				numpacks=splitlargepacks()
				if international<>"" then
					sXML=sXML & addUSPSInternational(rowcounter,thepweight / numpacks,apsrs(4,prodindex)*numpacks,"Package",shipcountry,apsrs(3,prodindex))
				else
					sXML=sXML & addUSPSDomestic(rowcounter,"Parcel",origZip,destZip,thepweight / numpacks,apsrs(4,prodindex)*numpacks,"REGULAR","True",apsrs(3,prodindex))
				end if
				rowcounter=rowcounter + 1
				zeropackdims()
			end if
		end if
	elseif shipType=4 OR shipType>=6 then ' UPS Shipping OR Canada Post OR FedEX OR DHL
		Session.LCID=1033
		if shipType=4 then
			packaging="02"
			if packaging<>"" then
				if packaging="envelope" then packaging="01"
				if packaging="pak" then packaging="04"
				if packaging="box" then packaging="21"
				if packaging="tube" then packaging="03"
				if packaging="10kgbox" then packaging="25"
				if packaging="25kgbox" then packaging="24"
			elseif upspacktype<>"" then
				packaging=upspacktype
			end if
		end if
		if packtogether then
			if shipThisProd then
				somethingToShip=TRUE
				thepweight=thepweight + (cdbl(apsrs(5,prodindex)) * int(apsrs(4,prodindex)))
				packageweight=packageweight+(cdbl(apsrs(5,prodindex)) * int(apsrs(4,prodindex)))
				if (apsrs(8,prodindex) AND 16)=16 then packagefreeexemptweight=packagefreeexemptweight+(cdbl(apsrs(5,prodindex)) * int(apsrs(4,prodindex)))
			end if
			if islastpacktogether(apsrs,prodindex) AND thepweight>0 then
				numpacks=splitlargepacks()
				for index3=1 to numpacks
					if shipType=4 then
						sXML=sXML & addUPSInternational(thepweight / numpacks,adminUnits,packaging,shipCountryCode,totalgoods-shipfreegoods,packdims)
					elseif shipType=6 then
						sXML=sXML & addCanadaPostPackage(thepweight / numpacks,adminUnits,packaging,shipCountryCode,totalgoods-shipfreegoods,packdims)
					elseif shipType=9 then
						sXML=sXML & addDHLPackage(thepweight / numpacks,adminUnits,packaging,shipCountryCode,totalgoods-shipfreegoods,packdims)
					elseif shipType=7 OR shipType=8 then
						sXML=sXML & addFedexPackage(thepweight / numpacks,totalgoods-shipfreegoods,packdims)
					end if
				next
				if shipType<>10 then zeropackdims()
			end if
		else
			if shipThisProd then
				somethingToShip=TRUE
				thepweight=apsrs(5,prodindex) + IIfVr(initialpackweight<>"",initialpackweight,0)
				packageweight=packageweight+thepweight
				if (apsrs(8,prodindex) AND 16)=16 then packagefreeexemptweight=packagefreeexemptweight+thepweight
				numpacks=splitlargepacks()
				for index2=0 to int(apsrs(4,prodindex))-1
					for index3=1 to numpacks
						if shipType=4 then
							sXML=sXML & addUPSInternational(thepweight / numpacks,adminUnits,packaging,shipCountryCode,apsrs(3,prodindex),packdims)
						elseif shipType=6 then
							sXML=sXML & addCanadaPostPackage(thepweight / numpacks,adminUnits,packaging,shipCountryCode,apsrs(3,prodindex),packdims)
						elseif shipType=9 then
							sXML=sXML & addDHLPackage(thepweight / numpacks,adminUnits,packaging,shipCountryCode,apsrs(3,prodindex),packdims)
						elseif shipType=7 OR shipType=8 then
							sXML=sXML & addFedexPackage(thepweight / numpacks,apsrs(3,prodindex),packdims)
						end if
					next
				next
				zeropackdims()
			end if
		end if
		Session.LCID=saveLCID
	end if
	if (apsrs(8,prodindex) AND 32)=32 then
		thepweight=savethepweight
		packageweight=savepackageweight
		packtogether=savepacktogether
		packdims=savepackdims
	end if
end sub
function calculateshipping()
	if fromshipselector then
		' Nothing
	elseif shipType=1 then
		freeshipmethodexists=TRUE
	elseif shipType=3 AND somethingToShip then
		sXML=sXML & "</"&international&"Rate"&IIfVr(international="","V4","V2")&"Request>"
		success=USPSCalculate(sXML,international,errormsg,intShipping)
		if left(errormsg, 30)="Warning - Bound Printed Matter" then success=TRUE
		if success then
			maxsopt=0
			for index=0 to UBOUND(intShipping,2)
				if iTotItems=intShipping(3,index) then
					intShipping(3,index)=TRUE
					maxsopt=index
					for index2=0 to UBOUND(uspsmethods,2)
						if replace(lcase(intShipping(5,index)),"-"," ")=replace(lcase(uspsmethods(0,index2)),"-"," ") then
							intShipping(4,index)=uspsmethods(1,index2)
						end if
					next
				else
					intShipping(3,index)=FALSE
				end if
			next
			for ssaindex2=0 to maxsopt-1
				if intShipping(3,ssaindex2)=TRUE then
					csmatch=intShipping(0,ssaindex2)
					for ssaindex=ssaindex2+1 to maxsopt
						if csmatch=intShipping(0,ssaindex) AND NOT intShipping(4,ssaindex) then intShipping(3,ssaindex)=FALSE
					next
				end if
			next
		end if
	elseif shipType=4 AND somethingToShip then
		sXML=sXML & "<ShipmentServiceOptions>" & IIfVr(saturdaydelivery_,"<SaturdayDelivery/>","") & IIfVr(saturdaypickup=TRUE,"<SaturdayPickup/>","") & "</ShipmentServiceOptions>" & IIfVr(upsnegdrates=TRUE, "<RateInformation><NegotiatedRatesIndicator /></RateInformation>", "") & "</Shipment></RatingServiceSelectionRequest>"
		if trim(upsUser)<>"" AND trim(upsPw)<>"" then
			success=UPSCalculate(sXML,international,errormsg,intShipping)
		else
			success=FALSE
			errormsg="You must register with UPS by logging on to your online admin section and clicking the &quot;Register with UPS&quot; link before you can use the UPS OnLine&reg; Shipping Rates and Services Selection"
		end if
	elseif shipType=6 AND somethingToShip then
		sXML=sXML & "</mailing-scenario></rate:get-rates-request></soapenv:Body></soapenv:Envelope>"
		success=CanadaPostCalculate(sXML,international,errormsg,intShipping)
	elseif (shipType=7 OR shipType=8) AND somethingToShip then
		session.LCID=1033
		sXML=replace(sXML,"XXXFEDEXGRANDTOTXXX",FormatNumber(totalgoods,2,-1,0,0))
		session.LCID=savelcid
		sXML=sXML & "</v9:RequestedShipment></v9:RateRequest></soapenv:Body></soapenv:Envelope>"
		if shipType=8 AND smartPostHub="" then success=FALSE : errormsg="SmartPost Hub ID not set" else success=fedexcalculate(sXML,international, errormsg, intShipping)
	elseif shipType=9 AND somethingToShip then
		if shipCountryCode="IE" AND ordCity="" then ordCity="Dublin"
		session.LCID=1033
		sXML=sXML & "</Pieces><PaymentAccountNumber>" & DHLAccountNo & "</PaymentAccountNumber><IsDutiable>" & IIfVr(origCountryCode=shipCountryCode OR (iseuropean(origCountryCode) AND iseuropean(shipCountryCode)),"N","Y") & "</IsDutiable>" & _
			"<NetworkTypeCode>AL</NetworkTypeCode></BkgDetails><To><CountryCode>" & shipCountryCode & "</CountryCode><Postalcode>" & ucase(destZip) & "</Postalcode>" & IIfVr(zipisoptional(shipCountryID) OR shipCountryID=65,"<City>"&ordCity&"</City>","") & "</To>" & _
			"<Dutiable><DeclaredCurrency>" & countryCurrency & "</DeclaredCurrency><DeclaredValue>" & totalgoods & "</DeclaredValue></Dutiable></GetQuote></q1:DCTRequest>"
		session.LCID=savelcid
		success=dhlcalculate(sXML,international,errormsg,intShipping)
	elseif shipType=10 AND somethingToShip then
		success=auspostcalculate(packageweight,international,errormsg,intShipping)
	end if
	if success AND somethingToShip AND NOT fromshipselector AND shipType>=1 then
		totShipOptions=0
		multipleoptions=TRUE
		for index=0 to UBOUND(intShipping,2)
			if intShipping(3,index)=TRUE then
				totShipOptions=totShipOptions + 1
				if intShipping(4,index) then freeshipmethodexists=TRUE
			end if
			if shipType>=2 AND packageweight>0 then
				intShipping(7,index)=intShipping(2,index)*(packagefreeexemptweight/packageweight)
			end if
		next
		if totShipOptions=0 AND NOT willpickup_ then
			multipleoptions=FALSE
			success=FALSE
			errormsg=xxNoMeth
		end if
		if willpickup_ then multipleoptions=TRUE
	end if
	calculateshipping=success
end function
sub saveshippingoptions()
	maxindex=0
	if shipType>=1 AND is_numeric(orderid) then
		sSQL="SELECT MAX(soIndex) AS maxindex FROM shipoptions WHERE soOrderID=" & orderid
		rs.open sSQL,cnn,0,1
		if NOT isnull(rs("maxindex")) then maxindex=rs("maxindex")+1
		rs.close
		for index=0 to UBOUND(intShipping,2)
			if intShipping(3,index)=TRUE then
				session.LCID=1033
				sSQL="INSERT INTO shipoptions (soOrderID,soIndex,soMethodName,soCost,soFreeShipExempt,soFreeShip,soShipType,soDeliveryTime,soDateAdded) VALUES (" & _
					orderid & "," & maxindex & ",'" & escape_string(intShipping(0,index)) & "'," & intShipping(2,index) & "," & intShipping(7,index) & "," & _
					intShipping(4,index) & "," & shipType & ",'" & escape_string(intShipping(1,index)) & "'," & vsusdate(Date()) & ")"
				ect_query(sSQL)
				maxindex=maxindex+1
				session.LCID=saveLCID
			end if
		next
	end if
end sub
numshiprate=0 : numshiprateingroup=0
sub writeshippingoption(soindex,shipcost,freeshipexempt,freeship,shipmethod,isselected,wsodelivery)
	if freeshippingapplied AND freeship=1 then wsofreeshipamnt=(shipcost-(freeshipexempt + getfreeshipinsurance(soindex))) else wsofreeshipamnt=0
	wsohandling=vsround(orighandling, 2)
	if handlingeligableitem=FALSE then
		wsohandling=0
	elseif orighandlingpercent<>0 then
		temphandling=(((totalgoods + shipcost + wsohandling) - (totaldiscounts + wsofreeshipamnt)) * orighandlingpercent / 100.0)
		if handlingeligablegoods < totalgoods AND totalgoods>0 then temphandling=temphandling * (handlingeligablegoods / totalgoods)
		wsohandling=wsohandling + temphandling
	end if
	if taxHandling=1 then wsohandling=wsohandling + (cdbl(wsohandling)*(cdbl(stateTaxRate)+cdbl(countryTaxRate)))/100.0
	if freeship=1 AND freeshippingincludeshandling then wsohandling=0
	if shippingoptionsasradios=TRUE then
		print "<div class=""shiprateline""><label class=""ectlabel shipradio""><input type=""radio"" class=""ectradio shipradio"" value=""RATE"""&IIfVr(isselected, " checked=""checked""", "")&" onclick=""updateshiprate(this,"&numshiprate&")"" /><span class=""shipratemethod"&IIfVs(numshiprateingroup=0," shipratemethodselected")&""">"&shipmethod&" "&IIfVr(combineshippinghandling, FormatEuroCurrency((shipcost+wsohandling)-wsofreeshipamnt), FormatEuroCurrency(shipcost-wsofreeshipamnt))&"</span>"&IIfVs(wsodelivery<>"","<div class=""cart3servicecommitment"">(" & wsodelivery & ")</div>")&"</label></div>"
	else
		print "<option value=""RATE"""&IIfVr(isselected, " selected=""selected""", "")&">"&shipmethod&" "&IIfVr(wsodelivery<>"" AND NOT mobilebrowser, "(" & wsodelivery & ") ", "")&IIfVr(combineshippinghandling, FormatEuroCurrency((shipcost+wsohandling)-wsofreeshipamnt), FormatEuroCurrency(shipcost-wsofreeshipamnt))&"</option>"
	end if
	numshiprate=numshiprate + 1
	numshiprateingroup=numshiprateingroup + 1
end sub
currShipType=""
function showshippingselect()
	if NOT fromshipselector then call calculateshippingdiscounts(FALSE)
	if shipType>=1 then
		if shippingoptionsasradios<>TRUE then
			print "<select class=""ectselectinput nofixw"" size=""1"" onchange=""updateshiprate(this,(this.selectedIndex"&IIfVr(fromshipselector,"-1","")&")+"&numshiprate&")"">"
			if fromshipselector then print "<option value="""">"&xxPlsSel&"</option>"
		end if
		for index=0 to UBOUND(intShipping,2)
			if intShipping(3,index) then
				if currShipType="" then currShipType=intShipping(6,index)
				if currShipType<>intShipping(6,index) then
					currShipType=intShipping(6,index)
					numshiprateingroup=0
					if shippingoptionsasradios<>TRUE then print "</select>"
					print "</div></div><div class=""shiptableline" & IIfVs(adminAltRates=2,2) & """><div class=""shiplogo" & IIfVs(adminAltRates=2,2) & """>" & getshiplogo(currShipType) & "</div><div class=""shiptablerates" & IIfVs(adminAltRates=2,2) & """>"
					if shippingoptionsasradios<>TRUE then print "<select size=""1"" onchange=""updateshiprate(this,(this.selectedIndex-1)+"&numshiprate&")""><option value="""">"&xxPlsSel&"</option>"
				end if
				call writeshippingoption(index,vsround(intShipping(2,index), 2), vsround(intShipping(7,index), 2), intShipping(4,index), intShipping(0,index), index=selectedshiptype, intShipping(1,index))
			end if
		next
		if shippingoptionsasradios<>TRUE then print "</select>"
		if royalmail AND shipType=2 AND saturdaydelivery_ then print "<div class=""nosaturdaydelivery"">First and Second Class post are not available with Saturday Delivery.</div>"
	end if
end function
function getuspsinsurancerate(theamount)
	if theamount<=0 then
		getuspsinsurancerate=0
	elseif theamount<=50 then
		getuspsinsurancerate=1.75
	elseif theamount<=100 then
		getuspsinsurancerate=2.25
	elseif theamount<=200 then
		getuspsinsurancerate=2.75
	else
		getuspsinsurancerate=4.70 + (1.0 * int((theamount-200.01) / 100.0))
	end if
end function
function getfreeshipinsurance(soindex)
	getfreeshipinsurance=0
	if (wantinsurance_ AND addshippinginsurance=2) OR addshippinginsurance=1 then
		if shipping>0 AND NOT freeshippingincludesservices then
			if (shipType=3 OR shipType=4 OR shipType=6 OR shipType=7 OR shipType=8) AND NOT nocarrierinsurancerates then
				getfreeshipinsurance=intShipping(8,soindex)
			else
				getfreeshipinsurance=(totalgoods*shipinsurancepercent)/100.0
				if getfreeshipinsurance<shipinsurancemin then getfreeshipinsurance=shipinsurancemin
			end if
		end if
	end if
end function
sub insuranceandtaxaddedtoshipping()
	if (wantinsurance_ AND addshippinginsurance=2) OR addshippinginsurance=1 then
		if shipType=3 AND NOT nocarrierinsurancerates then
			for index3=0 to UBOUND(intShipping,2)
				intShipping(2,index3)=intShipping(2,index3) + intShipping(8,index3)
			next
			shipping=shipping+intShipping(8,selectedshiptype)
		elseif shipType=2 OR shipType=5 OR nocarrierinsurancerates then
			shipinsurancetotal=(totalgoods*shipinsurancepercent)/100.0
			if shipinsurancetotal<shipinsurancemin then shipinsurancetotal=shipinsurancemin
			for index3=0 to UBOUND(intShipping,2)
				intShipping(2,index3)=intShipping(2,index3) + shipinsurancetotal
			next
			shipping=shipping+shipinsurancetotal
		end if
	end if
	if taxShipping=1 then
		for index3=0 to UBOUND(intShipping,2)
			intShipping(2,index3)=intShipping(2,index3) + (cdbl(intShipping(2,index3))*(cdbl(stateTaxRate)+cdbl(countryTaxRate)))/100.0
		next
		shipping=shipping + (cdbl(shipping)*(cdbl(stateTaxRate)+cdbl(countryTaxRate)))/100.0
	end if
end sub
sub calculatetaxandhandling()
	if countrytaxthreshold<>0 then
		if totalgoods-totaldiscounts>countrytaxthreshold then
			countryTaxRate=0
			taxShipping=0
		end if
	end if
	if handlingeligableitem=FALSE then
		handling=0
	else
		if handlingchargepercent<>0 then
			temphandling=(((totalgoods + shipping + handling) - (totaldiscounts + freeshipamnt)) * handlingchargepercent / 100.0)
			if handlingeligablegoods < totalgoods AND totalgoods>0 then temphandling=temphandling * (handlingeligablegoods / totalgoods)
			handling=handling + temphandling
		end if
		if taxHandling=1 then handling=handling + (cdbl(handling)*(cdbl(stateTaxRate)+cdbl(countryTaxRate)))/100.0
	end if
	if origCountryID=2 AND shipCountryID=2 AND (shipStateAbbrev="NB" OR shipStateAbbrev="NF" OR shipStateAbbrev="NS" OR shipStateAbbrev="ON" OR shipStateAbbrev="PE") then usehst=TRUE else usehst=FALSE
	if totalgoods>0 then
		stateTax=((cdbl(totalgoods)-(cdbl(totaldiscounts)+cdbl(statetaxfree)))*cdbl(stateTaxRate)/100.0)
		if perproducttaxrate<>TRUE then countryTax=((cdbl(totalgoods)-(cdbl(totaldiscounts)+cdbl(countrytaxfree)))*cdbl(countryTaxRate)/100.0)
		if showtaxinclusive=3 AND homeCountryTaxRate>0 then
			countryTax=vsround((totalgoods-countrytaxfree) / ((100+homeCountryTaxRate)/homeCountryTaxRate),2)
			totalgoods=totalgoods-countryTax
			if countryTaxRate<>homeCountryTaxRate then
				if countryTaxRate<>0 then countryTax=countryTax*(countryTaxRate/homeCountryTaxRate) else countryTax=0
			end if
		end if
	end if
	if taxShipping=2 AND (shipping - freeshipamnt>0) then
		if proratashippingtax=TRUE then
			if totalgoods>0 then stateTax=stateTax + (((cdbl(totalgoods)-(cdbl(totaldiscounts)+cdbl(statetaxfree))) / totalgoods) * (cdbl(shipping)-cdbl(freeshipamnt))*(cdbl(stateTaxRate)/100.0))
		else
			stateTax=stateTax + (cdbl(shipping)-cdbl(freeshipamnt))*(cdbl(stateTaxRate)/100.0)
		end if
		shippingtax=vsround((cdbl(shipping)-cdbl(freeshipamnt))*(cdbl(countryTaxRate)/100.0),2)
		countryTax=countryTax + shippingtax
	end if
	if taxHandling=2 then
		stateTax=stateTax + cdbl(handling)*(cdbl(stateTaxRate)/100.0)
		countryTax=countryTax + cdbl(handling)*(cdbl(countryTaxRate)/100.0)
	end if
	if stateTax < 0 then stateTax=0
	if countryTax < 0 then countryTax=0
	if usehst then
		countryTax=vsround(stateTax+countryTax,2)
		stateTax=0
	else
		stateTax=vsround(stateTax,2)
		countryTax=vsround(countryTax,2)
	end if
	handling=vsround(handling,2)
	if showtaxinclusive<>0 then SESSION("xscountrytax")=countryTax
	if showtaxinclusive=3 then SESSION("xscountrytax")=shippingtax
end sub
sub do_stock_check(sublevels,byref hasbackorder,byref hasstockwarning)
	Dim sameitemstock()
	redim sameitemstock(2,10)
	redim outofstockarr(4,10)
	isameitemstock=0
	gotstock=TRUE
	hasbackorder=FALSE
	sSQL="SELECT cartID,cartQuantity FROM cart WHERE cartCompleted=0 AND " & getsessionsql() & " ORDER BY cartDateAdded"
	rs3.open sSQL,cnn,0,1
	do while NOT rs3.EOF
		cartID=rs3("cartID")
		thequant=rs3("cartQuantity")
		pID=""
		sSQL="SELECT pInStock,pID,pStockByOpts,"&WSP&"pPrice,pBackOrder,pSell FROM cart INNER JOIN products ON cart.cartProdId=products.pID WHERE cartID="&cartID
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			pID=rs("pID")
			pInStock=int(rs("pInStock"))
			pStockByOpts=cint(rs("pStockByOpts"))
			pPrice=rs("pPrice")
			pBackOrder=(cint(rs("pBackOrder"))<>0)
			pSell=rs("pSell")
		end if
		rs.close
		if pID<>"" then
			if useStockManagement then
				quantity=thequant
				thisiteminstock=TRUE
				if thequant=0 then
					gotstock=FALSE
				elseif pStockByOpts<>0 then
					if mysqlserver=TRUE then
						sSQL="SELECT coID,optStock,coOptID FROM cart INNER JOIN cartoptions ON cart.cartID=cartoptions.coCartID INNER JOIN options ON cartoptions.coOptID=options.optID INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optType IN (-4,-2,-1,1,2,4) AND cartID="&cartID
					else
						sSQL="SELECT coID,optStock,coOptID FROM cart INNER JOIN (cartoptions INNER JOIN (options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID) ON cart.cartID=cartoptions.coCartID WHERE optType IN (-4,-2,-1,1,2,4) AND cartID="&cartID
					end if
					rs.open sSQL,cnn,0,1
					do while NOT rs.EOF
						sameitems=0
						for index=0 to isameitemstock
							if sameitemstock(0,index)=rs("coOptID") AND sameitemstock(2,index)=TRUE then sameitems=sameitems+sameitemstock(1,index)
						next
						pInStock=int(rs("optStock"))
						if (pInStock-sameitems) < quantity then
							thisiteminstock=FALSE
							quantity=(pInStock-sameitems)
							if quantity < 0 then quantity=0
							if sublevels AND NOT pBackOrder then ect_query("UPDATE cart SET cartQuantity="&quantity&" WHERE (cartCompleted=0 OR cartCompleted=3) AND cartID="&cartID)
							if pBackOrder then hasbackorder=TRUE else gotstock=FALSE
							outofstockarr(0, outofstockcnt)=rs("coID")
							outofstockarr(1, outofstockcnt)=TRUE ' optID
							outofstockarr(2, outofstockcnt)=pID
							outofstockarr(3, outofstockcnt)=pPrice
							outofstockarr(4, outofstockcnt)=pBackOrder
							outofstockcnt=outofstockcnt + 1
							if outofstockcnt>=UBOUND(outofstockarr, 2) then redim preserve outofstockarr(3, UBOUND(outofstockarr, 2) + 10)
						end if
						sameitemstock(0,isameitemstock)=rs("coOptID")
						sameitemstock(1,isameitemstock)=thequant
						sameitemstock(2,isameitemstock)=TRUE
						isameitemstock=isameitemstock + 1
						if isameitemstock>=UBOUND(sameitemstock, 2) then redim preserve sameitemstock(2, UBOUND(sameitemstock, 2) + 10)
						rs.movenext
					loop
					rs.close
				else
					sameitems=0
					for index=0 to isameitemstock
						if sameitemstock(0,index)=pID AND sameitemstock(2,index)=FALSE then sameitems=sameitems+sameitemstock(1,index)
					next
					if pInStock < (thequant+sameitems) then
						thisiteminstock=FALSE
						quantity=(pInStock-sameitems)
						if quantity < 0 then quantity=0
						if sublevels AND NOT pBackOrder then ect_query("UPDATE cart SET cartQuantity="&quantity&" WHERE (cartCompleted=0 OR cartCompleted=3) AND cartID="&cartID)
						if pBackOrder then hasbackorder=TRUE else gotstock=FALSE
						outofstockarr(0, outofstockcnt)=cartID
						outofstockarr(1, outofstockcnt)=FALSE ' cartID
						outofstockarr(2, outofstockcnt)=pID
						outofstockarr(3, outofstockcnt)=pPrice
						outofstockarr(4, outofstockcnt)=pBackOrder
						outofstockcnt=outofstockcnt + 1
						if outofstockcnt>=UBOUND(outofstockarr, 2) then redim preserve outofstockarr(4, UBOUND(outofstockarr, 2) + 10)
					end if
					sameitemstock(0,isameitemstock)=pID
					sameitemstock(1,isameitemstock)=thequant
					sameitemstock(2,isameitemstock)=FALSE
					isameitemstock=isameitemstock + 1
					if isameitemstock>=UBOUND(sameitemstock, 2) then redim preserve sameitemstock(2, UBOUND(sameitemstock, 2) + 10)
				end if
				for index2=0 to numaddedprods-1
					if cstr(addedprods(5,index2))=cstr(cartID) then
						if NOT thisiteminstock then addedprods(4,index2)=IIfVr(pBackOrder,3,2)
						if NOT pBackOrder then addedprods(2,index2)=quantity
					end if
				next
			elseif pSell=0 AND pBackOrder<>0 then
				hasbackorder=TRUE
				for index2=0 to numaddedprods-1
					if cstr(addedprods(5,index2))=cstr(cartID) then addedprods(4,index2)=3
				next
			end if
		end if
		rs3.movenext
	loop
	rs3.close
	hasstockwarning=NOT gotstock
	if sublevels then
		for index=0 to outofstockcnt-1
			call checkpricebreaks(outofstockarr(2, index), outofstockarr(3, index))
		next
	end if
end sub
sub checkdeletecart(thecartid)
	sSQL="SELECT cartID,cartListID,cartClientID,listOwner,cartProdID,cartProdPrice FROM cart LEFT JOIN customerlists ON cart.cartListID=customerlists.listID WHERE (cartCompleted=0 OR cartCompleted=3) AND cartID=" & IIfVr(is_numeric(thecartid), thecartid, 0) & " AND " & getsessionsql()
	rs2.open sSQL,cnn,0,1
	if NOT rs2.EOF then
		if NOT isnull(rs2("listOwner")) then listowner=clng(rs2("listOwner")) else listowner=0
		if rs2("cartListID")>0 AND listowner<>SESSION("clientID") then
			ect_query("UPDATE cart SET cartCompleted=3,cartOrderID=0,cartClientID="&rs2("listOwner")&" WHERE cartID="&thecartid)
		else
			if rs2("cartProdID")=giftwrappingid then
				ect_query("UPDATE cart SET cartGiftWrap=0 WHERE " & getsessionsql())
			else
				print "<script>gtag(""event"",""remove_from_cart"", { currency:'" & countryCurrency & "',value:" & rs2("cartProdPrice") & ",items:[" & getcartforganalytics(thecartid) & "]});</script>" & vbLf
			end if
			ect_query("DELETE FROM cartoptions WHERE coCartID="&thecartid)
			ect_query("DELETE FROM cart WHERE cartID="&thecartid)
			ect_query("DELETE FROM giftcertificate WHERE gcCartID="&thecartid)
			call updategiftwrap()
		end if
	elseif SESSION("clientID")<>"" then
		rs2.close
		sSQL="SELECT cartID FROM cart INNER JOIN customerlists ON cart.cartListID=customerlists.listID WHERE cartID=" & IIfVr(is_numeric(thecartid), thecartid, 0) & " AND listOwner=" & SESSION("clientID")
		rs2.open sSQL,cnn,0,1
		if NOT rs2.EOF then ect_query("UPDATE cart SET cartListID=0 WHERE cartID=" & IIfVr(is_numeric(thecartid), thecartid, 0))
	end if
	rs2.close
end sub
sub writeshippingflags(costage)
	hasshipflag=FALSE
	if willpickuptext<>"" OR willpickup_ OR commercialloc OR commercialloc=2 OR saturdaydelivery OR abs(addshippinginsurance)=2 OR (allowsignaturerelease AND signatureoption<>"") OR insidedelivery OR holdatlocation OR homedelivery then hasshipflag=TRUE
	if termsandconditions AND costage=3 AND (ordPayProvider="19" OR ordPayProvider="21") then hasshipflag=TRUE
	if hasshipflag then print "<div class=""coshipflagscontainer"">"
	if willpickuptext<>"" OR willpickup_ then %>
			<div class="billformrowflags"><div class="cdshipftflag cdformtwillpickup"><input type="checkbox" class="ectcheckbox" name="willpickup" id="shipflag0" value="Y" <%
				if willpickup_ then print "checked=""checked"" "
				if costage=3 then print "onchange=""setchangeflag(this.checked,'w')"" "%>/></div>
			<div class="cdshipflag cdformwillpickup"><%=labeltxt(willpickuptext,"shipflag0") & IIfVr(willpickupcost<>""," (" & FormatEuroCurrency(willpickupcost) & ")","")%></div></div>
<%	end if
	if NOT willpickup_ then
		if commercialloc=TRUE OR commercialloc=2 then %>
			<div class="billformrowflags"><div class="cdshipftflag"><input type="checkbox" class="ectcheckbox" name="commercialloc" id="shipflag1" value="Y" <%
				if ((ordComLoc AND 1)=1) OR (ordName="" AND commercialloc=2) then print "checked=""checked"" "
				if costage=3 then print "onchange=""setchangeflag(this.checked,0)"" "%>/></div>
			<div class="cdshipflag"><%=labeltxt(xxComLoc,"shipflag1")%></div></div>
<%		end if
		if saturdaydelivery=TRUE then %>
			<div class="billformrowflags"><div class="cdshipftflag"><input type="checkbox" class="ectcheckbox" name="saturdaydelivery" id="shipflag2" value="Y" <%
				if (ordComLoc AND 4)=4 then print "checked=""checked"" "
				if costage=3 then print "onchange=""setchangeflag(this.checked,2)"" "%>/></div>
			<div class="cdshipflag"><%=labeltxt(xxSatDel,"shipflag2")%></div></div>
<%		end if
		if abs(addshippinginsurance)=2 then %>
			<div class="billformrowflags"><%
				if forceinsuranceselection AND costage<>3 then
					print "<div class=""cdshipftselect""><select class=""ectselectinput"" name=""wantinsurance"" size=""1""><option value="""">"&xxPlsSel&"</option><option value="""">"&xxNo&"</option><option value=""Y"">"&xxYes&"</option></select></div>"
				else
					print "<div class=""cdshipftflag""><input type=""checkbox"" class=""ectcheckbox"" name=""wantinsurance"" id=""shipflag3"" value=""Y"" "&IIfVs((ordComLoc AND 2)=2,"checked=""checked"" ")&IIfVs(costage=3,"onchange=""setchangeflag(this.checked,1)"" ")&"/></div>"
				end if %>
			<div class="cdshipflag"><%=labeltxt(IIfVr(forceinsuranceselection,xxChoIns,xxWantIns),"shipflag3")%></div></div>
<%		end if
		if allowsignaturerelease=TRUE AND signatureoption<>"" then %>
			<div class="billformrowflags"><div class="cdshipftflag"><input type="checkbox" class="ectcheckbox" name="signaturerelease" id="shipflag4" value="Y" <%
				if (ordComLoc AND 8)=8 then print "checked=""checked"" "
				if costage=3 then print "onchange=""setchangeflag(this.checked,3)"" "%>/></div>
			<div class="cdshipflag"><%=labeltxt(xxSigRel,"shipflag4")%></div></div>
<%		end if
		if insidedelivery=TRUE then %>
			<div class="billformrowflags"><div class="cdshipftflag"><input type="checkbox" class="ectcheckbox" name="insidedelivery" id="shipflag5" value="Y" <%
				if (ordComLoc AND 16)=16 then print "checked=""checked"" "
				if costage=3 then print "onchange=""setchangeflag(this.checked,4)"" "%>/></div>
			<div class="cdshipflag"><%=labeltxt(xxInsDel,"shipflag5")%></div></div>
<%		end if
		if holdatlocation=TRUE then %>
			<div class="billformrowflags"><div class="cdshipftflag"><input type="checkbox" class="ectcheckbox" name="holdatlocation" id="shipflag6" value="Y" /></div>
			<div class="cdshipflag"><%=labeltxt("Please click here to Hold at Location","shipflag6")%>	</div></div>
<%		end if
		if homedelivery=TRUE then %>
			<div class="billformrowflags"><div class="cdshipftflag">Delivery Options</div>
			<div class="cdshipftselect"><select name="homedelivery" class="ectselectinput" size="1">
			<option value="">Standard Delivery</option>
			<option value="EVENING">Evening Home Delivery</option>
			<option value="DATE_CERTAIN">Date Certain Home Delivery</option>
			<option value="APPOINTMENT">Appointment Home Delivery</option>
			</select></div></div>
<%		end if
	end if
	if termsandconditions AND costage=3 AND (ordPayProvider="19" OR ordPayProvider="21") then %>
		<div class="billformrowflags">
		  <div class="cdshiptterms"><input type="checkbox" name="termsandconds" class="ectcheckbox" onchange="ectremoveclass(this,'ectwarning')" id="ecttnccheckbox" value="1" onclick="document.getElementById('sftermsandconds').value=this.checked?1:''" <%=IIfVs(getpost("sftermsandconds")="1","checked=""checked"" ")%>/></div>
		  <div class="cdshipterms"><%=labeltxt(xxTermsCo,"ecttnccheckbox")%></div>
		</div>
<%
	end if
	if hasshipflag then print "</div>"
end sub
sub displaycartactions(cid,tlistid)
	if NOT ispubliclist then print "<div class=""cartdelete""><a href=""#""><img class=""cartdelete"" src=""images/delete.png"" alt="""&xxDelete&""" onclick=""return dodelete("&cid&",'"&tlistid&"')"" /></a></div>"
	if haspubliclist AND ispubliclist AND NOT itemdeleted then
		print "<div class=""movetocart"">"&imageorlink(imgmovetocart,xxMovCar,"movetocart","whichcartid='"&cid&"';return dosaveitem('x')", TRUE)&"</div>"
	elseif SESSION("clientID")<>"" AND enablewishlists AND NOT itemdeleted then
		print "<div class=""cartaddtolist""><div style=""position:relative;display:inline"">"&imageorlink(imgcartaddtolist,xxMovToL,"cartaddtolist","return cartdispsavelist("&cid&",'"&tlistid&"',this)",TRUE)&"</div></div>"
	end if
end sub
sub writeestimatormenu()
	if adminAltRates=2 then
		print "<div class=""estimatorchecktext"" id=""estimatorchecktext""></div><div class=""estimatorcheckcarrier"" id=""estimatorcheckcarrier""></div>"
	elseif adminAltRates=1 then
		sSQL="SELECT altrateid,altratename,"&getlangid("altratetext",65536)&",usealtmethod,usealtmethodintl FROM alternaterates WHERE usealtmethod"&international&"<>0 ORDER BY altrateorder"&international&",altrateid"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if shippingoptionsasradios<>TRUE then print "<select id=""altratesselect"" size=""1"" onchange=""selaltrate(this[this.selectedIndex].value)"""&IIfVs(mobilebrowser," style=""font-size:10px""")&">" else print "<div"&IIfVs(mobilebrowser," style=""font-size:10px""")&">"
			do while NOT rs.EOF
				call writealtshipline(rs(getlangid("altratetext",65536)),rs("altrateid"),xxOrCom&": ",xxShEsWi&": ",TRUE)
				rs.movenext
			loop
			if shippingoptionsasradios<>TRUE then print "</select>" else print "</div>"
		end if
		rs.close
	end if
end sub
sub displaydiscounts()
	if totaldiscounts>0 then print "<div class=""cartdiscounts_cntnr""><div class=""cartdiscountstext ectdscntt"">" & xxDsApp & "</div><div class=""cartdiscounts ectdscnt"" id=""discountspan"">" & FormatEuroCurrency(totaldiscounts) & "</div></div>" & vbCrLf
	if checkoutmode<>"savedcart" AND vsround(loyaltypointsavailable*loyaltypointvalue,2)>=IIfVr(loyaltypointminimum<>"",loyaltypointminimum,0.05) then
		if SESSION("noredeempoints")<>TRUE then
			loyaltypointdiscount=loyaltypointsavailable*loyaltypointvalue
			if loyaltypointdiscount>totalgoods+IIfVr(showtaxinclusive=3,countryTax,0)-totaldiscounts then loyaltypointdiscount=totalgoods+IIfVr(showtaxinclusive=3,countryTax,0)-totaldiscounts
		end if
		print "<div class=""cartloyaltypoint_cntnr""><div class=""cartloyaltypointmenu ectdscntt"">"
		print "<select size=""1"" onchange=""document.location='cart"&extension&"?redeempoints='+(this.selectedIndex==1?'no':'yes')""><option value="""">"&xxReLPts&"</option><option value=""""" & IIfVs(SESSION("noredeempoints")=TRUE," selected=""selected""") & ">" & xxSaLPts & " (" & loyaltypointsavailable & ")</option></select>"
		print "</div><div class=""cartloyaltypoints ectdscnt"">"
		if SESSION("noredeempoints")=TRUE then print "-" else print FormatEuroCurrency(loyaltypointdiscount)
		print "</div></div>"
	end if
end sub
function displaycartclosed()
	if listid="" then
		displaycartclosed=FALSE
	elseif getget("pli")<>"" then
		displaycartclosed=getget("pli")<>listid
	else
		displaycartclosed=cstr(displaylistid)<>cstr(cartlistnumber) AND getpost("listid")<>cstr(listid)
	end if
end function
sub displaylistheader(cartlistnumber,checkoutstep,listname,isclosed,numitems,tclass)
	print "<div class=""cartlistdiv cartlistshop"">"
	print "<div class=""ectdivhead cartlistname cartlistcart"" style=""cursor:pointer"" onclick=""document.getElementById('cartlistid"&cartlistnumber&"').style.display==''?document.getElementById('cartlistimgid"&cartlistnumber&"').src='images/arrow-down.png':document.getElementById('cartlistimgid"&cartlistnumber&"').src='images/arrow-up.png';document.getElementById('cartlistid"&cartlistnumber&"').style.display=(document.getElementById('cartlistid"&cartlistnumber&"').style.display==''?'none':'')""><div class=""checkoutstep" & IIfVs(cartlistnumber=2," checkoutstepof3") & """>"&IIfVr(cartlistnumber=2,checkoutstep,"&nbsp;")&"</div><div class=""cartname"">" & listname & IIfVs(numitems>0," (" & numitems & ")") & "</div><div class=""cartlistimg""><img src=""images/arrow-" & IIfVr(isclosed,"down","up") & ".png"" id=""cartlistimgid"&cartlistnumber&""" alt=""Show / Hide Cart"" /></div></div>" & vbCrLf
	print "</div>"
	print "<div class=""" & trim("cartlistcontents " & tclass) & """ id=""cartlistid" & cartlistnumber & """" & IIfVs(isclosed," style=""display:none""") & ">"
end sub
sub showcartlines(iscartresume)
	if iscartresume then linkcartproducts=FALSE
	if NOT amazonpaycheckout then
		if (stockwarning OR backorder) AND checkoutmode<>"savedcart" then
			print "<div class=""cartstockbackorder_cntnr"">"
			if stockwarning then
				print "<div class=""cartstockwarning""><div class=""cartoutstock ectwarning"">" & xxNoStok & "</div><div class=""cartstockacceptlevel"">"&xxStkUTo&"<a class=""ectlink"" href=""cart"&extension&""">"&xxClkHere&"</a></div>"
				if getget("mode")<>"add" AND checkoutmode<>"update" then print "<div class=""cartstockjustpurchased"">("&xxJusBuy&")</div>"
				print "</div>"
			end if
			if backorder then print "<div class=""cartbackorder ectwarning"">" & xxBakOrW & "</div>"
			print "</div>"
		end if
		minquantmessage=""
		for index=0 to UBOUND(alldata,2)
			if alldata(4,index)<=alldata(21,index) then
				minquantmessage=minquantmessage&"<div class=""ectwarning"">" & replace(replace(xxMinQuw,"%pname%",trim(alldata(2,index)&"")),"%quant%",alldata(21,index)+1) & "</div>"
				if checkoutmode<>"savedcart" then minquantityerror=TRUE
			end if
			if isnull(alldata(22,index)) AND alldata(1,index)<>donationid AND alldata(1,index)<>giftcertificateid AND alldata(1,index)<>giftwrappingid then
				minquantmessage=minquantmessage&"<div class=""ectwarning"">" & replace("The product %pname% has been deleted from the product catalog.","%pname%",alldata(2,index)) & "</div>"
				if checkoutmode<>"savedcart" then deleteditemerror=TRUE
			end if
		next
		if minquantmessage<>"" then print "<div class=""cartminquant_cntnr"">" & minquantmessage & "</div>"
		print "<div class=""cartdetails_cntnr"">"
		print "<div class=""cartdetails cartdetailsid"">" & xxCODets & "</div>"
		print "<div class=""cartdetails cartdetailsname" & IIfVs(iscartresume," cartdetailsnamecr") & """>" & xxCOName & "</div>"
		if nopriceanywhere<>TRUE then print "<div class=""cartdetails cartdetailsprice"">" & xxCOUPri & "</div>"
		print "<div class=""cartdetails cartdetailsquant"">" & xxQuanty & "</div>"
		if NOT iscartresume then print "<div class=""cartdetails cartdetailscheck"">&nbsp;</div>"
		print "<div class=""cartdetails cartdetailstotal"">" & IIfVr(nopriceanywhere,"&nbsp;",xxTotal) & "</div>"
		print "</div>"
	end if
	print "<div class=""cartlineitems"">"
	for index=0 to UBOUND(alldata,2)
		cartlineunique=cartlineunique+1
		cartID=alldata(0,index)
		cartProdID=alldata(1,index)
		cartProdName=trim(alldata(2,index)&"")
		if isnull(alldata(5,index)) then alldata(5,index)=0
		if isnull(alldata(8,index)) then
			if cartProdID=giftcertificateid OR cartProdID=donationid then alldata(8,index)=15
			if cartProdID=giftwrappingid then alldata(8,index)=12
		end if
		pSection=alldata(9,index)
		topSection=alldata(10,index)
		pTax=alldata(12,index)
		pStaticPage=alldata(13,index)
		pDisplay=alldata(14,index)
		pImage=trim(alldata(15,index)&"")
		pLargeImage=trim(alldata(16,index)&"")
		cartCompleted=alldata(17,index)
		pGiftWrap=alldata(18,index)
		cartGiftWrap=alldata(19,index)
		pStaticURL=alldata(20,index)
		if isnull(alldata(21,index)) then alldata(21,index)=0
		pMinQuant=alldata(21,index)
		pID=alldata(22,index)
		cartOrigProdID=trim(alldata(23,index)&"")
		pDescription=trim(alldata(24,index)&"")
		pLongDescription=trim(alldata(25,index)&"")
		itemdeleted=isnull(pID) AND cartProdID<>donationid AND cartProdID<>giftcertificateid
		if useimageincart then
			sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageType=0 AND imageProduct='"&escape_string(cartProdID)&"' ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then pImage=trim(rs("imageSrc")&"")
			rs.close
		end if
		changechecker=changechecker & "if(document.checkoutform.quant" & alldata(0,index) & ".value!=" & alldata(4,index) & ") dowarning=true;" & vbCrLf
		theoptions=""
		theoptionspricediff=0
		isoutofstock=FALSE
		if mysqlserver=TRUE then
			sSQL="SELECT coID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff,optAltImage,optType FROM cartoptions LEFT JOIN options ON cartoptions.coOptID=options.optID LEFT JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE coCartID="&alldata(0,index) & " ORDER BY coID"
		else
			sSQL="SELECT coID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff,optAltImage,optType FROM cartoptions LEFT JOIN (options LEFT JOIN optiongroup ON options.optGroup=optiongroup.optGrpID) ON cartoptions.coOptID=options.optID WHERE coCartID="&alldata(0,index) & " ORDER BY coID"
		end if
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			optoutofstock=FALSE
			for rowcounter=0 to outofstockcnt-1
				if outofstockarr(0, rowcounter)=rs("coID") AND outofstockarr(1, rowcounter)=TRUE AND outofstockarr(4, rowcounter)=FALSE then optoutofstock=TRUE : isoutofstock=TRUE
			next
			theoptionspricediff=theoptionspricediff + rs("coPriceDiff")
			alldata(5,index)=cdbl(alldata(5,index))+cdbl(rs("coWeightDiff"))
			if trim(rs("optAltImage")&"")<>"" AND rs("optType")<>4 AND useimageincart then
				if instr(pImage, "%s")>0 then pImage=replace(pImage, "%s", rs("optAltImage")) else pImage=trim(rs("optAltImage")&"")
			end if
			theoptions=theoptions & "<div class=""cartoptionsline""><div class=""cartoptiongroup" & IIfVs(itemdeleted," ectwarning") & """>" & rs("coOptGroup") & "</div>" & _
				"<div class=""cartoption" & IIfVs(itemdeleted," ectwarning") & """>" & replace(replace(htmlspecials(rs("coCartOption")&""),vbCr,""),vbLf,"<br>") & "</div>"
			if NOT nopriceanywhere then theoptions=theoptions & "<div class=""cartoptionprice" & IIfVs(itemdeleted," ectwarning") & """>" & IIfVr(rs("coPriceDiff")=0 OR hideoptpricediffs=TRUE,"- ", FormatEuroCurrency(rs("coPriceDiff"))) & "</div>"
			theoptions=theoptions & "<div class=""cartoptionoutstock"">" & IIfVr(optoutofstock, xxLimSto, "&nbsp;") & "</div>"
			if NOT nopriceanywhere then theoptions=theoptions & "<div class=""cartoptiontotal" & IIfVs(itemdeleted," ectwarning") & """>" & IIfVr(rs("coPriceDiff")=0 OR hideoptpricediffs=TRUE,"- ", FormatEuroCurrency(rs("coPriceDiff")*alldata(4,index))) & "</div>"
			theoptions=theoptions & "<div class=""cartoptionspacer""></div></div>" & vbCrLf
			totalgoods=totalgoods + (rs("coPriceDiff")*alldata(4,index))
			if (alldata(8,index) AND 8)<>8 then handlingeligablegoods=handlingeligablegoods + (rs("coPriceDiff")*alldata(4,index))
			rs.movenext
		loop
		rs.close
		for rowcounter=0 to outofstockcnt-1
			if outofstockarr(0, rowcounter)=alldata(0,index) AND outofstockarr(1, rowcounter)=FALSE AND outofstockarr(4, rowcounter)=FALSE then isoutofstock=TRUE
		next
		rs.open "SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageType=0 AND imageProduct='" & escape_string(cartProdID) & "' ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1"),cnn,0,1
		if NOT rs.EOF then pImage=trim(rs("imageSrc")&"")
		rs.close
		rs.open "SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageType=1 AND imageProduct='" & escape_string(cartProdID) & "' ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1"),cnn,0,1
		if NOT rs.EOF then pLargeImage=trim(rs("imageSrc")&"")
		rs.close
		cartOrigProdName=cartProdName
		if cartOrigProdID<>"" then
			rs.open "SELECT pName,pDisplay,pStaticURL,pStaticPage,pLongDescription FROM products WHERE pID='" & escape_string(cartOrigProdID) & "'",cnn,0,1
			if NOT rs.EOF then
				pDisplay=rs("pDisplay")
				pStaticURL=rs("pStaticURL")
				pStaticPage=rs("pStaticPage")
				cartOrigProdName=rs("pName")
				pLongDescription=rs("pLongDescription")
			end if
			rs.close
		end if
		if pDisplay<>0 AND linkcartproducts=TRUE AND (forcedetailslink OR pLongDescription<>"" OR pLargeImage<>"") then
			thedetailslink=getdetailsurl(IIfVr(cartOrigProdID<>"",cartOrigProdID,cartProdID),pStaticPage,cartOrigProdName,trim(pStaticURL&""),"",pathtohere)
			if detailslink<>"" then
				sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageType=1 AND imageProduct='" & escape_string(cartProdID) & "' ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then pLargeImage=trim(rs("imageSrc")&"")
				rs.close
				startlink=replace(replace(detailslink,"%largeimage%", pLargeImage),"%pid%", cartProdID)
				endlink=detailsendlink
			else
				startlink="<a class=""ectlink"" href="""&thedetailslink&""">"
				endlink="</a>"
			end if
		else
			startlink=""
			endlink=""
		end if
		if NOT amazonpaycheckout then ' Cart line
			print "<div class=""cartandoptsline""><div class=""cartline"">"
				print "<div class=""cartlineid" & IIfVs(itemdeleted," ectwarning") & """>"
				if useimageincart AND NOT (pImage="" OR pImage="prodimages/") then print startlink & "<img class=""cartimage"" src=""" & pImage & """ alt=""" & strip_tags2(cartProdName) & """ />" & endlink else print startlink & cartProdID & endlink
				print "</div><div class=""cartlinename" & IIfVs(iscartresume," cartlinenamecr") & IIfVs(itemdeleted," ectwarning") & """>"
				print startlink & cartProdName & endlink
				if pGiftWrap<>0 AND NOT iscartresume then print "<div class=""giftwrap""><a href=""cart"&extension&"?mode=gw"">" & IIfVr(cartGiftWrap<>0,xxGWrSel,xxGWrAva) & "</a></div>"
				sSQL="SELECT quantity,pName,quantity FROM productpackages INNER JOIN products on productpackages.pID=products.pID WHERE packageID='"&escape_string(cartProdID)&"'"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					print "<div class=""packagediv"">"
					do while NOT rs.EOF
						print "<div class=""packagerow"">" & rs("pName") & " (" & rs("quantity") & ")</div>"
						rs.movenext
					loop
					print "</div>"
				end if
				rs.close
				print "</div>"
				if nopriceanywhere<>TRUE then print "<div class=""cartlineprice" & IIfVs(itemdeleted," ectwarning") & """>" & IIfVr(hideoptpricediffs=TRUE,FormatEuroCurrency(alldata(3,index)+theoptionspricediff),FormatEuroCurrency(alldata(3,index))) & "</div>"
				print "<div class=""cartlinequant" & IIfVs(itemdeleted," ectwarning") & """>"
				if itemdeleted then
					print "-"
				elseif getget("pla")<>"" OR iscartresume then
					print alldata(4,index)
				else
					print "<input class=""ecttextinput cartquant"" onkeydown=""showupdatebutton("&alldata(0,index)&","&cartlineunique&",'"&cartlistnumber&"')"" type=""text"" id=""quant"&alldata(0,index)&""" value="""&alldata(4,index)&""" size=""2"" maxlength=""5"" "&IIfVs(isoutofstock OR alldata(4,index)<=pMinQuant,"style=""background-color:#FF7070"" ")&IIfVs(alldata(4,index)<=pMinQuant,"title="""&htmlspecials(replace(xxMinQua,"%quant%",pMinQuant+1))&""" ")&"/>"
				end if
				print "</div>"
				if NOT iscartresume then
					print "<div class=""cartlinecheck" & IIfVs(itemdeleted," ectwarning") & """>"
					if checkoutmode<>"savedcart" then
					elseif cartCompleted=0 OR cartCompleted=2 then
						print "<div class=""wishlistpurch wishlistpurchasing"">" & xxPurcha & "</div>"
					elseif cartCompleted=1 then
						print "<div class=""wishlistpurch wishlistpurchased"">" & xxPurchd & "</div>"
					end if
					call displaycartactions(alldata(0,index),cartlistnumber)
					print "</div>"
				end if
				print "<div class=""cartlinetotal" & IIfVs(itemdeleted," ectwarning") & """ id=""cartlinetot"&cartlineunique&""">" & IIfVr(nopriceanywhere,"-",IIfVr(hideoptpricediffs=TRUE,FormatEuroCurrency((alldata(3,index)+theoptionspricediff)*alldata(4,index)),FormatEuroCurrency(alldata(3,index)*alldata(4,index)))) & "</div>"
			print "</div>" & vbCrLf
			print theoptions & "</div>"
		end if
		runTot=(alldata(3,index)*int(alldata(4,index)))
		totalquantity=totalquantity + int(alldata(4,index))
		totalgoods=totalgoods + runTot
		if NOT iscartresume then
			alldata(3,index)=alldata(3,index) + theoptionspricediff
			if trim(SESSION("clientID"))<>"" then alldata(8,index)=(alldata(8,index) OR (SESSION("clientActions") AND 7)) : if (SESSION("clientActions") AND 32)=32 then alldata(8,index)=alldata(8,index) OR 8
			if (shipType=2 OR shipType=3 OR shipType=4 OR shipType>=6) AND cdbl(alldata(5,index))<=0.0 then alldata(8,index)=(alldata(8,index) OR 4)
			if perproducttaxrate=TRUE then
				if isnull(pTax) then pTax=countryTaxRate
				if (alldata(8,index) AND 2)<>2 then countryTax=countryTax + ((pTax * alldata(3,index) * int(alldata(4,index))) / 100.0)
			else
				if (alldata(8,index) AND 2)=2 then countrytaxfree=countrytaxfree + runTot + (theoptionspricediff*int(alldata(4,index)))
			end if
			if (alldata(8,index) AND 4)=4 then shipfreegoods=shipfreegoods + runTot else somethingToShip=TRUE
			if alldata(1,index)=giftcertificateid OR alldata(1,index)=donationid then shipdiscountexempt=shipdiscountexempt+runTot : numshipdiscountexempt=numshipdiscountexempt+alldata(4,index)
			if (alldata(8,index) AND 8)<>8 then handlingeligableitem=TRUE : handlingeligablegoods=handlingeligablegoods + runTot
			if SESSION("xsshipping")="" AND checkoutmode<>"savedcart" then call addproducttoshipping(alldata, index)
		end if
	next
	print "</div>"
end sub
sub showcartresumeheader(checkoutstep)
	totalgoods=0 : totalquantity=0
	querystr="cartCompleted=0 AND "&getsessionsql()
	sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions, pSection,topSection,pDims,pTax,pStaticPage,pDisplay,'' AS pImage,'' AS pLargeImage, cartCompleted,pGiftWrap,cartGiftWrap,pStaticURL,pMinQuant,pID,cartOrigProdID,"&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)
	if mysqlserver=TRUE then
		sSQL=sSQL&" FROM cart LEFT JOIN products ON cart.cartProdID=products.pID LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE " & querystr & " ORDER BY cartID"
	else
		sSQL=sSQL&" FROM cart LEFT JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE " & querystr & " ORDER BY cartID"
	end if
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then alldata=rs.getrows else alldata=""
	rs.close
	call displaylistheader(0,checkoutstep,xxShoCar,IIfVr(cartresumeopen,FALSE,TRUE),numitems,"cartresumecontents")
	if isarray(alldata) then showcartlines(TRUE)
	print "</div>"
	if checkoutstep=3 AND trim(ordName&ordLastName&"")<>"" then
		call displaylistheader(1,checkoutstep,xxCusDet,IIfVr(custdetailsresumeopen,FALSE,TRUE),0,"custdetailsresume")
%>
<script>
function checkcustdetails(){
    var form = document.createElement("form");
    var element1 = document.createElement("input");
    form.method = "POST";
    form.action = "cart<%=extension%>";
    element1.value='checkout';
    element1.name='mode';
    form.appendChild(element1);  
    document.body.appendChild(form);
    form.submit();
	document.body.removeChild(form);
}
</script>
<%
			print "<div class=""custdetsresumehead"">" & xxContEm & "</div>"
			print "<div class=""custdetsresumeline"">" & htmlspecials(ordEmail) & "</div>"
			print "<div class=""custdetsresumehead"">" & xxBilAdd & "</div>"
			print "<div class=""custdetsresumeline"">" & IIfVs(ordExtra1<>"",extraorderfield1 & ": " & ordExtra1&"<br>") & trim(ordName&" "&ordLastName)&", "&ordAddress&IIfVs(ordAddress2<>"",", "&ordAddress2)&", "&trim(ordCity&" "&ordState&" "&ordZip) & ", " & ordCountry & IIfVs(ordExtra2<>"","<br>" & extraorderfield2 & ": " & ordExtra2) & "</div>"
			if trim(ordShipAddress)<>"" then
				print "<div class=""custdetsresumehead"">" & xxShpAdd & "</div>"
				print "<div class=""custdetsresumeline"">" & IIfVs(ordShipExtra1<>"",extraorderfield1 & ": " & ordShipExtra1&"<br>") & trim(ordShipName&" "&ordShipLastName)&", "&ordShipAddress&IIfVs(ordShipAddress2<>"",", "&ordShipAddress2)&", "&trim(ordShipCity&" "&ordShipState&" "&ordShipZip)&", "&ordShipCountry & IIfVs(ordShipExtra2<>"","<br>" & extraorderfield2 & ": " & ordShipExtra2) & "</div>"
			end if
			print "<div><input class=""ectbutton"" type=""button"" value=""Change"" onclick=""checkcustdetails()"" /></div>"
		print "</div>"
	end if
	call displaylistheader(2,checkoutstep,IIfVr(checkoutstep=2,xxCstDtl,xxChkCmp),FALSE,0,"cartmaincontents")
end sub
function requirevalidstate(tshiptype)
	requirevalidstate=FALSE
	if splitUSZones then
		if adminAltRates=2 then
			rs2.open "SELECT altrateid FROM alternaterates WHERE usealtmethod<>0 AND altrateid IN (2,5)",cnn,0,1
			requirevalidstate=NOT rs2.EOF
			rs2.close
		else
			requirevalidstate=tshiptype=2 OR tshiptype=5
		end if
	end if
end function
amzrefid_=getrequest("amzrefid")
if amzrefid_<>"" then ' Amazon Payment
	if getpayprovdetails(21,data1,data2,data3,demomode,ppmethod) then
		ordPayProvider="21"
		data2arr=split(data2,"&",2)
		if UBOUND(data2arr)>=0 then data2=data2arr(0)
		if UBOUND(data2arr)>0 then sellerid=data2arr(1)
		amazonstr=""
		ordName="" : ordLastName="" : ordAddress="" : ordAddress2="" : ordState="" : ordCountry="" : ordPhone="" : ordEmail="" : ordEmail2=""
		countryid=0 : ordComLoc=0
		checkoutmode="go"
		amazonpayment=TRUE

		scripturl="mws-eu.amazonservices.com"
		if origCountryCode="US" then scripturl="mws.amazonservices.com"
		if origCountryCode="JP" then scripturl="mws.amazonservices.jp"
		endpointpath="/OffAmazonPayments" & IIfVs(demomode,"_Sandbox") & "/2013-01-01"
		endpoint="https://" & scripturl & endpointpath

		call amazonparam2("AWSAccessKeyId",data2)
		call amazonparam2("Action","GetOrderReferenceDetails")
		call amazonparam2("AmazonOrderReferenceId",amzrefid_)
		call amazonparam2("SellerId",sellerid)
		call amazonparam2("SignatureMethod","HmacSHA256")
		call amazonparam2("SignatureVersion",2)
		call amazonparam2("Timestamp",getutcdate(0))
		call amazonparam2("Version","2013-01-01")
		call amazonparam2("Signature",b64_hmac_sha256(data3,calculateStringToSignV2()))

		if callxmlfunction(endpoint,amazonstr,res,"","WinHTTP.WinHTTPRequest.5.1",errormsg,FALSE) then
			set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
			xmlDoc.validateOnParse=FALSE
			xmlDoc.loadXML(res)
			Set nodeList=xmlDoc.getElementsByTagName("GetOrderReferenceDetailsResponse")
			Set nl=nodeList.Item(0)
			for ix1=0 to nl.childNodes.length - 1
				Set tx1=nl.childNodes.Item(ix1)
				if tx1.nodeName="Error" then
					for ix2=0 To tx1.childNodes.length - 1
						set tx2=tx1.childNodes.Item(ix2)
						if tx2.nodeName="Message" then carterror="Amazon Error: " & tx2.firstChild.nodeValue
					next
				elseif tx1.nodeName="GetOrderReferenceDetailsResult" then
					for ix2=0 To tx1.childNodes.length - 1
						set tx2=tx1.childNodes.Item(ix2)
						if tx2.nodeName="OrderReferenceDetails" then
							for ix3=0 To tx2.childNodes.length - 1
								set tx3=tx2.childNodes.Item(ix3)
								if tx3.nodeName="Constraints" then
									for ix4=0 To tx3.childNodes.length - 1
										set tx4=tx3.childNodes.Item(ix4)
										if tx4.nodeName="Constraint" then
											isamountconstraint=FALSE
											for ix5=0 To tx4.childNodes.length - 1
												set tx5=tx4.childNodes.Item(ix5)
												if tx5.nodeName="ConstraintID" then
													if tx5.firstChild.nodeValue="AmountNotSet" then isamountconstraint=TRUE
												elseif tx5.nodeName="Description" then
													thisconstraint=tx5.firstChild.nodeValue
												end if
											next
											if NOT isamountconstraint then
												amazonpayment=checkoutmode=""
												carterror=thisconstraint
											end if
										end if
									next
								elseif tx3.nodeName="Destination" then
									for ix4=0 To tx3.childNodes.length - 1
										set tx4=tx3.childNodes.Item(ix4)
										if tx4.nodeName="PhysicalDestination" then
											for ix5=0 To tx4.childNodes.length - 1
												set tx5=tx4.childNodes.Item(ix5)
												if tx5.nodeName="StateOrRegion" then
													ordState=tx5.firstChild.nodeValue
												elseif tx5.nodeName="City" then
													ordCity=tx5.firstChild.nodeValue
												elseif tx5.nodeName="CountryCode" then
													tmpcntry=replace(tx5.firstChild.nodeValue,"'","")
													sSQL="SELECT countryName,countryID FROM countries WHERE countryEnabled=1 AND "
													if tmpcntry="GB" then
														sSQL=sSQL & "countryID=201"
													elseif tmpcntry="FR" then
														sSQL=sSQL & "countryID=65"
													elseif tmpcntry="PT" then
														sSQL=sSQL & "countryID=153"
													elseif tmpcntry="ES" then
														sSQL=sSQL & "countryID=175"
													else
														sSQL=sSQL & "countryCode='"&tmpcntry&"'"
													end if
													rs.open sSQL,cnn,0,1
													if NOT rs.EOF then
														ordCountry=rs("countryName")
														countryid=rs("countryID")
														homecountry=(countryid=origCountryID)
													else
														errormsg="Purchasing from your country is not supported."
														success=FALSE
													end if
													rs.close
												elseif tx5.nodeName="PostalCode" then
													ordZip=tx5.firstChild.nodeValue
												end if
											next
										end if
									next
								end if
							next
						end if
					next
				end if
			next
		else
			print "curl failed: " & res & "<br>"
		end if
	end if
elseif left(getget("token"),2)="EC" AND checkoutmode<>"paypalcancel" then ' { PayPal Express
	call getpayprovdetails(19,username,data2pwd,data2hash,demomode,ppmethod)
	if username<>"" then username=trim(split(username,"/")(0))
	if data2pwd<>"" then data2pwd=urldecode(split(data2pwd,"&")(0))
	if data2pwd<>"" then data2pwd=trim(split(data2pwd,"/")(0))
	if data2hash<>"" then data2hash=trim(split(data2hash,"/")(0))
	if instr(username,"@AB@")<>0 then data2pwd="" : data2hash="AB"
	sXML=ppsoapheader(username, data2pwd, data2hash) & _
		"<soap:Body><GetExpressCheckoutDetailsReq xmlns=""urn:ebay:api:PayPalAPI""><GetExpressCheckoutDetailsRequest><Version xmlns=""urn:ebay:apis:eBLBaseComponents"">60.0</Version>" & _
		"  <Token>" & getget("token") & "</Token>" & _
		"</GetExpressCheckoutDetailsRequest></GetExpressCheckoutDetailsReq></soap:Body></soap:Envelope>"
	if demomode then sandbox=".sandbox" else sandbox=""
	if callxmlfunction("https://api" & IIfVr(data2hash<>"", "-3t", "") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",username), "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE) then
		countryid=0
		success=FALSE
		ordPayProvider="19"
		ordEmail="" : ordEmail2=""
		ordComLoc=0
		gotaddress=FALSE
		token=getget("token")
		if abs(addshippinginsurance)=1 then ordComLoc=ordComLoc + 2
		set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.validateOnParse=FALSE
		xmlDoc.loadXML (res)
		Set nodeList=xmlDoc.getElementsByTagName("SOAP-ENV:Body")
		Set n=nodeList.Item(0)
		for j=0 to n.childNodes.length - 1
			Set e9=n.childNodes.Item(i)
			if e9.nodeName="GetExpressCheckoutDetailsResponse" then
				for k9=0 To e9.childNodes.length - 1
					Set t=e9.childNodes.Item(k9)
					if t.nodeName="Ack" then
						if t.firstChild.nodeValue="Success" OR t.firstChild.nodeValue="SuccessWithWarning" then success=TRUE
					elseif t.nodeName="GetExpressCheckoutDetailsResponseDetails" then
						set ff=t.childNodes
						for kk=0 to ff.length - 1
							set gg=ff.item(kk)
							if gg.nodeName="PayerInfo" then
								set hh=gg.childNodes
								for ll=0 to hh.length - 1
									set ii=hh.item(ll)
									if ii.nodeName="Payer" then
										if ii.hasChildNodes then ordEmail=ii.firstChild.nodeValue : ordEmail2=ordEmail
									elseif ii.nodeName="PayerID" then
										if ii.hasChildNodes then payerid=ii.firstChild.nodeValue
									elseif ii.nodeName="PayerStatus" then
										if ii.hasChildNodes then
											ordCVV="U"
											payer_status=lcase(ii.firstChild.nodeValue)
											if payer_status="verified" then ordCVV="Y"
											if payer_status="unverified" then ordCVV="N"
										end if
									elseif ii.nodeName="PayerName" then
									elseif ii.nodeName="Address" then
										set jj=ii.childNodes
										for mm=0 to jj.length - 1
											set jjj=jj.item(mm)
											if jjj.nodeName="Name" then
												if jjj.hasChildNodes then call splitfirstlastname(trim(jjj.firstChild.nodeValue),ordName,ordLastName)
											elseif jjj.nodeName="Street1" then
												if jjj.hasChildNodes then ordAddress=jjj.firstChild.nodeValue
											elseif jjj.nodeName="Street2" then
												if jjj.hasChildNodes then ordAddress2=jjj.firstChild.nodeValue
											elseif jjj.nodeName="CityName" then
												if jjj.hasChildNodes then ordCity=jjj.firstChild.nodeValue
											elseif jjj.nodeName="StateOrProvince" then
												if jjj.hasChildNodes then ordState=jjj.firstChild.nodeValue
											elseif jjj.nodeName="Country" then
												if jjj.hasChildNodes then
													tmpcntry=replace(jjj.firstChild.nodeValue,"'","")
													sSQL="SELECT countryName,countryID FROM countries WHERE countryEnabled=1 AND "
													if tmpcntry="GB" then
														sSQL=sSQL & "countryID=201"
													elseif tmpcntry="FR" then
														sSQL=sSQL & "countryID=65"
													elseif tmpcntry="PT" then
														sSQL=sSQL & "countryID=153"
													elseif tmpcntry="ES" then
														sSQL=sSQL & "countryID=175"
													else
														sSQL=sSQL & "countryCode='"&tmpcntry&"'"
													end if
													rs.open sSQL,cnn,0,1
													if NOT rs.EOF then
														ordCountry=rs("countryName")
														countryid=rs("countryID")
														homecountry=(countryid=origCountryID)
													else
														errormsg="Purchasing from your country is not supported."
														success=FALSE
													end if
													rs.close
												end if
											elseif jjj.nodeName="PostalCode" then
												if jjj.hasChildNodes then ordZip=jjj.firstChild.nodeValue
											elseif jjj.nodeName="AddressStatus" then
												if jjj.hasChildNodes then
													ordAVS="U"
													address_status=lcase(jjj.firstChild.nodeValue)
													gotaddress=(address_status<>"none")
													if address_status="confirmed" then ordAVS="Y"
													if address_status="unconfirmed" then ordAVS="N"
												end if
											end if
										next
									end if
								next
							elseif gg.nodeName="Custom" then
								customarr=split(gg.firstChild.nodeValue, ":")
								thesessionid=customarr(0)
								if UBOUND(customarr)>0 then ordAffiliate=customarr(1) else ordAffiliate=""
								if left(thesessionid,3)="cid" then
									SESSION("clientID")=replace(right(thesessionid, len(thesessionid)-3),"'","")
									thesessionid=0
									sSQL="SELECT clID,clUserName,clActions,clLoginLevel,clPercentDiscount FROM customerlogin WHERE clID="&replace(SESSION("clientID"),"'","")
									rs.open sSQL,cnn,0,1
									if NOT rs.EOF then
										SESSION("clientUser")=rs("clUsername")
										SESSION("clientActions")=rs("clActions")
										SESSION("clientLoginLevel")=rs("clLoginLevel")
										SESSION("clientPercentDiscount")=(100.0-cdbl(rs("clPercentDiscount")))/100.0
									end if
									rs.close
								else
									thesessionid=replace(right(thesessionid, len(thesessionid)-3),"'","")
								end if
							elseif gg.nodeName="ContactPhone" then
								if gg.hasChildNodes then ordPhone=gg.firstChild.nodeValue
							end if
						next
					elseif t.nodeName="Errors" then
						set ff=t.childNodes
						for kk=0 to ff.length - 1
							set gg=ff.item(kk)
							if gg.nodeName="ShortMessage" then
								if gg.hasChildNodes then errormsg=gg.firstChild.nodeValue & "<br>" & errormsg
							elseif gg.nodeName="LongMessage" then
								if gg.hasChildNodes then errormsg= errormsg & gg.firstChild.nodeValue
							elseif gg.nodeName="ErrorCode" then
								if gg.hasChildNodes then errcode=gg.firstChild.nodeValue
							end if
						next
					end if
				next
			end if
		next
		if NOT gotaddress then
			response.redirect storeurl & "cart"&extension
			cartisincluded=TRUE
		elseif success then
			paypalexpress=TRUE
			if homecountry then
				if (countryid=1 OR countryid=2) AND usestateabbrev<>TRUE then
					sSQL="SELECT stateName FROM states WHERE (stateCountryID=1 OR stateCountryID=2) AND stateAbbrev='" & escape_string(ordState) & "'"
					rs.open sSQL,cnn,0,1
					if NOT rs.EOF then ordState=rs("stateName")
					rs.close
				end if
				if requirevalidstate(shipType) then
					sSQL="SELECT stateName FROM states WHERE stateCountryID='" & escape_string(countryid) & "' AND (stateName='" & escape_string(ordState) & "' OR stateAbbrev='" & escape_string(ordState) & "')"
					rs.open sSQL,cnn,0,1
					if rs.EOF then ordState=""
					rs.close
				end if
			end if
		else
			carterror="PayPal Express Error: "&errormsg
			checkoutmode="paypalcancel"
		end if
	else
		carterror="PayPal Express Error: "&errormsg
		checkoutmode="paypalcancel"
	end if
elseif checkoutmode="paypalexpress1" OR checkoutmode="billmelater" then ' }{ PayPal Express
	session.LCID=1033
	success=FALSE
	call getpayprovdetails(19,username,data2pwd,data2hash,demomode,ppmethod)
	if username<>"" then username=trim(split(username,"/")(0))
	if data2pwd<>"" then data2pwd=urldecode(split(data2pwd,"&")(0))
	if data2pwd<>"" then data2pwd=trim(split(data2pwd,"/")(0))
	if data2hash<>"" then data2hash=trim(split(data2hash,"/")(0))
	if instr(username,"@AB@")<>0 then data2pwd="" : data2hash="AB"
	if demomode then sandbox=".sandbox" else sandbox=""
	theestimate=vsround(cdbl(getpost("estimate")), 2)
	sXML=ppsoapheader(username, data2pwd, data2hash) & _
		"<soap:Body><SetExpressCheckoutReq xmlns=""urn:ebay:api:PayPalAPI""><SetExpressCheckoutRequest><Version xmlns=""urn:ebay:apis:eBLBaseComponents"">72.0</Version>" & _
		"<SetExpressCheckoutRequestDetails xmlns=""urn:ebay:apis:eBLBaseComponents"">" & _
		IIfVs(checkoutmode="billmelater","<FundingSourceDetails><UserSelectedFundingSource>BML</UserSelectedFundingSource></FundingSourceDetails>") & _
		"<OrderTotal currencyID=""" & countryCurrency & """>" & theestimate & "</OrderTotal>" & _
		"<ReturnURL>" & storeurlssl & "cart" & extension & "</ReturnURL><CancelURL>" & storeurlssl & "cart" & extension & "?action=paypalcancel</CancelURL>" & _
		"<Custom>" & IIfVr(SESSION("clientID")<>"", "cid"&SESSION("clientID"), "sid"&thesessionid) & ":" & strip_tags2(getpost("PARTNER")) & "</Custom>" & _
		"<PaymentAction>" & IIfVr(ppmethod=1, "Authorization", "Sale") & "</PaymentAction>"
	itemtotal=0
	sXML=sXML & "<PaymentDetails>"
	sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pShipping,pShipping2,pExemptions,pTax,pDescription FROM cart LEFT JOIN products ON cart.cartProdID=products.pId WHERE cartCompleted=0 AND " & getsessionsql()
	rs.open sSQL,cnn,0,1
	do while NOT rs.EOF
		itemtotal=itemtotal + (rs("cartProdPrice")*rs("cartQuantity"))
		description="" : addcomma=""
		optiontotal=0
		sSQL="SELECT coOptGroup,coCartOption,coPriceDiff FROM cartoptions WHERE coCartID=" & rs("cartID")
		rs2.open sSQL,cnn,0,1
		do while NOT rs2.EOF
			optiontotal=optiontotal + rs2("coPriceDiff")
			description=description & addcomma & vrxmlencode(strip_tags2(rs2("coOptGroup"))) & " : " & vrxmlencode(strip_tags2(rs2("coCartOption")))
			addcomma=", "
			rs2.movenext
		loop
		rs2.close
		itemtotal=itemtotal + (optiontotal*rs("cartQuantity"))
		sXML=sXML & "<PaymentDetailsItem><Number>" & vrxmlencode(strip_tags2(rs("cartProdID"))) & "</Number><Name>" & vrxmlencode(strip_tags2(rs("cartProdName"))) & "</Name><Description>" & left(description,122) & IIfVr(len(description)>122,"...","") & "</Description><Amount currencyID=""" & countryCurrency & """>" & (rs("cartProdPrice")+optiontotal) & "</Amount><Quantity>" & rs("cartQuantity") & "</Quantity></PaymentDetailsItem>"
		rs.movenext
	loop
	rs.close
	if itemtotal<>theestimate then
		sXML=sXML & "<PaymentDetailsItem><Name>" & vrxmlencode(xxPPEst1) & "</Name><Description>" & vrxmlencode(xxPPEst2) & "</Description><Amount currencyID=""" & countryCurrency & """>" & vsround(theestimate-itemtotal, 2) & "</Amount><Quantity>1</Quantity></PaymentDetailsItem>"
	end if
	sXML=sXML & "</PaymentDetails>"
	if paypallc<>"" then sXML=sXML & addtag("LocaleCode",paypallc)
	sXML=sXML & "  </SetExpressCheckoutRequestDetails>" & _
		"</SetExpressCheckoutRequest></SetExpressCheckoutReq></soap:Body></soap:Envelope>"
	if username="" then
		response.redirect "https://www.paypal.com/us/webapps/mpp/referral/paypal-payments-pro?partner_id=39HT54MCDMV8E"
		print "<p align=""center"">" & xxAutFo & "</p>"
		print "<p align=""center"">" & xxForAut & " <a class=""ectlink"" href=""https://www.paypal.com/us/webapps/mpp/referral/paypal-payments-pro?partner_id=39HT54MCDMV8E"">" & xxClkHere & "</a></p>"
	elseif callxmlfunction("https://api" & IIfVr(data2hash<>"", "-3t", "") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",username), "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE) then
		set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
		xmlDoc.validateOnParse=FALSE
		xmlDoc.loadXML (res)
		Set nodeList=xmlDoc.getElementsByTagName("SOAP-ENV:Body")
		Set n=nodeList.Item(0)
		for j=0 to n.childNodes.length - 1
			Set e9=n.childNodes.Item(i)
			if e9.nodeName="SetExpressCheckoutResponse" then
				for k9=0 To e9.childNodes.length - 1
					Set t=e9.childNodes.Item(k9)
					if t.nodeName="Ack" then
						if t.firstChild.nodeValue="Success" OR t.firstChild.nodeValue="SuccessWithWarning" then success=TRUE
					elseif t.nodeName="Token" then
						if t.hasChildNodes then token=t.firstChild.nodeValue
					elseif t.nodeName="Errors" then
						set ff=t.childNodes
						for kk=0 to ff.length - 1
							set gg=ff.item(kk)
							if gg.nodeName="ShortMessage" then
								if gg.hasChildNodes then errormsg=gg.firstChild.nodeValue & "<br>" & errormsg
							elseif gg.nodeName="LongMessage" then
								if gg.hasChildNodes then errormsg= errormsg & gg.firstChild.nodeValue
							elseif gg.nodeName="ErrorCode" then
								if gg.hasChildNodes then errcode=gg.firstChild.nodeValue
							end if
						next
					end if
				next
			end if
		next
		if success then
			' response.redirect "https://www" & sandbox & ".paypal.com/webscr?cmd=_express-checkout&token=" & token
			response.redirect "https://www" & sandbox & ".paypal.com/checkoutnow?useraction=commit&token=" & token
			print "<p align=""center"">" & xxAutFo & "</p>"
			print "<p align=""center"">" & xxForAut & " <a class=""ectlink"" href=""https://www" & sandbox & ".paypal.com/webscr?cmd=_express-checkout&token=" & token & """>" & xxClkHere & "</a></p>"
		else
			print "PayPal Express (3) error: " & errormsg
		end if
	else
		print "PayPal Express (4) error: " & errormsg
	end if
	session.LCID=saveLCID
elseif checkoutmode="update" OR checkoutmode="savecart" OR checkoutmode="movetocart" then ' }{
	if estimateshipping=TRUE then SESSION("xsshipping")=empty : SESSION("tofreeshipamount")=empty : SESSION("tofreeshipquant")=empty
	if NOT IsEmpty(SESSION("discounts")) then SESSION("discounts")=empty
	if NOT IsEmpty(SESSION("xscountrytax")) then SESSION("xscountrytax")=empty
	sSQL="SELECT ordID FROM orders WHERE ordStatus>1 AND ordAuthNumber='' AND " & getordersessionsql()
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then ordID=rs("ordID") else ordID=""
	rs.close
	if ordID<>"" then
		release_stock(ordID)
		ect_query("UPDATE cart SET cartSessionID='"&replace(getsessionid(),"'","")&"',cartClientID="&IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0)&" WHERE cartCompleted=0 AND cartOrderID=" & ordID)
		ect_query("UPDATE orders SET ordAuthStatus='MODWARNOPEN',ordShipType='MODWARNOPEN' WHERE ordID=" & ordID)
	end if
	listid=""
	if checkoutmode="savecart" AND is_numeric(getpost("listid")) AND SESSION("clientID")<>"" then
		rs.open "SELECT listID FROM customerlists WHERE listID="&getpost("listid")&" AND listOwner="&SESSION("clientID"),cnn,0,1
		if NOT rs.EOF then listid=rs("listID")
		rs.close
	end if
	for each objItem in request.form
		thequant=getpost(objItem)
		if NOT is_numeric(thequant) then thequant=0 else thequant=abs(int(thequant))
		if left(objItem,5)="quant" OR left(objItem,5)="delet" then
			thecartid=right(objItem, len(objItem)-5)
			if NOT is_numeric(thecartid) then thecartid=0
			pPrice=0
			pID=""
			sSQL="SELECT cartProdID,"&WSP&"pPrice FROM cart INNER JOIN products ON cart.cartProdId=products.pID WHERE cartID="&thecartid
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				pID=rs("cartProdID")
				pPrice=rs("pPrice")
			end if
			rs.close
			if checkoutmode="movetocart" then
				if Left(objItem,5)="delet" then
					sSQL="UPDATE cart SET cartCompleted=0,cartListID=0,cartDateAdded="&vsusdatetime(DateAdd("h",dateadjust,Now()))&" WHERE cartCompleted=3 AND cartID="&thecartid&" AND " & getsessionsql()
					if is_numeric(getget("pli")) AND getget("pla")<>"" then
						sSQL="SELECT listID FROM customerlists WHERE listID="&getget("pli")&" AND listAccess='"&escape_string(getget("pla"))&"'"
						rs.open sSQL,cnn,0,1
						if NOT rs.EOF then sSQL="UPDATE cart SET cartCompleted=0,cartDateAdded="&vsusdatetime(DateAdd("h",dateadjust,Now()))&",cartSessionID='"&replace(getsessionid(),"'","")&"',cartClientID="&IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0)&" WHERE cartCompleted=3 AND cartID="&thecartid&" AND cartListID=" & rs("listID") else sSQL=""
						rs.close
					end if
					if sSQL<>"" then ect_query(sSQL)
				end if
			elseif checkoutmode="savecart" AND pID<>giftwrappingid then
				if Left(objItem,5)="delet" then
					ect_query("UPDATE cart SET cartOrderID=0,cartCompleted=3,cartListID="&IIfVr(listid<>"",listid,"0")&",cartDateAdded="&vsusdatetime(DateAdd("h",dateadjust,Now()))&" WHERE (cartCompleted=0 OR cartCompleted=3) AND cartID="&thecartid&" AND " & getsessionsql())
				end if
			else
				if Left(objItem,5)="quant" AND thequant<>"" then
					if NOT is_numeric(thecartid) then thecartid=0
					if thequant=0 then
						checkdeletecart(thecartid)
					else
						if thequant>99999 then thequant=99999
						if pID<>"" AND pID<>giftwrappingid then
							ect_query("UPDATE cart SET cartQuantity="&thequant&",cartDateAdded="&vsusdatetime(DateAdd("h",dateadjust,Now()))&" WHERE cartQuantity<>"&thequant&" AND (cartCompleted=0 OR cartCompleted=3) AND cartID="&thecartid&" AND " & getsessionsql())
						end if
					end if
				elseif Left(objItem,5)="delet" then
					thecartid=Right(objItem, Len(objItem)-5)
					checkdeletecart(thecartid)
				end if
			end if
			if pID<>giftcertificateid AND pID<>donationid then call checkpricebreaks(pID,pPrice)
		end if
	next
	call updategiftwrap()
end if ' }
function additemtocart(ainame,aiprice)
	cartListID=0
	cartCompleted=0
	if getpost("listid")="0" AND SESSION("clientID")<>"" then
		cartCompleted=3
	elseif is_numeric(getpost("listid")) AND SESSION("clientID")<>"" then
		sSQL="SELECT listID FROM customerlists WHERE listOwner="&SESSION("clientID")&" AND listID="&getpost("listid")
		rs.open sSQL,cnn,0.1
		if NOT rs.EOF then cartListID=rs("listID") : cartCompleted=3
		rs.close
	end if
	floodlevel=IIfVr(blockmaxcartadds>0,blockmaxcartadds,1000)
	sSQL="SELECT COUNT(*) AS cartcnt FROM cart WHERE (cartCompleted=0 OR cartCompleted=3) AND " & getsessionsql()
	rs.open sSQL,cnn,0.1
	cartfloodcontrol=cint(rs("cartcnt"))>floodlevel
	rs.close
	if cartfloodcontrol then
		if blockmaxcartadds>0 then
			cnn.execute("INSERT INTO ipblocking (dcip1) VALUES (" & ip2long(REMOTE_ADDR) & ")")
			orderid=0 : ordauthstatus=""
			sSQL="SELECT ordID,ordAuthStatus FROM orders WHERE ordStatus>1 AND ordAuthNumber='' AND " & getordersessionsql()
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then orderid=rs("ordID") : ordauthstatus=rs("ordAuthStatus")
			rs.close
			if orderid=0 then
				if sqlserver=TRUE then
					sSQL="INSERT INTO orders (ordSessionID,ordClientID,ordName,ordPayProvider,ordAuthNumber,ordDate,ordStatusDate,ordStatus,ordAuthStatus,ordIP) VALUES (" & _
						"'" & escape_string(thesessionid) & "'," & IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0) & "," & _
						"'MAXCARTITEMS',4,''," & vsusdatetime(DateAdd("h",dateadjust,Now())) & "," & vsusdatetime(DateAdd("h",dateadjust,Now())) & ",2,'MODWARNOPEN','" & escape_string(REMOTE_ADDR) & "')"
					ect_query(sSQL)
					rs.open "SELECT @@IDENTITY AS lstIns",cnn,0,1
					orderid=int(cstr(rs("lstIns")))
					rs.close
				else
					rs.open "orders",cnn,1,3,&H0002
					rs.AddNew
					rs.Fields("ordSessionID")	= thesessionid
					rs.Fields("ordClientID")	= IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0)
					rs.Fields("ordName")		= "MAXCARTITEMS"
					rs.Fields("ordPayProvider")	= 4
					rs.Fields("ordAuthNumber")	= ""
					rs.Fields("ordStatus")		= 2
					rs.Fields("ordAuthStatus")	= "MODWARNOPEN"
					rs.Fields("ordIP")			= REMOTE_ADDR
					rs.Fields("ordDate")		= DateAdd("h",dateadjust,Now())
					rs.Fields("ordStatusDate")	= DateAdd("h",dateadjust,Now())
					rs.Update
					if mysqlserver=TRUE then
						if orderid="" then
							rs.close
							rs.open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
							orderid=rs("lstIns")
						end if
					else
						orderid=rs.Fields("ordID")
					end if
					rs.close
				end if
				sSQL="UPDATE cart SET cartOrderID="&orderid&" WHERE cartCompleted=0 AND " & getsessionsql()
				ect_query(sSQL)
			elseif ordauthstatus<>"MODWARNOPEN" then
				release_stock(orderid)
				ect_query("UPDATE cart SET cartSessionID='"&replace(getsessionid(),"'","")&"',cartClientID="&IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0)&" WHERE cartCompleted=0 AND cartOrderID=" & orderid)
				ect_query("UPDATE orders SET ordAuthStatus='MODWARNOPEN',ordShipType='MODWARNOPEN' WHERE ordID=" & orderid)
			end if
		end if
		additemtocart=0
	elseif sqlserver=TRUE then
		sSQL="INSERT INTO cart (cartSessionID,cartClientID,cartProdID,cartOrigProdID,cartQuantity,cartCompleted,cartProdName,cartProdPrice,cartOrderID,cartDateAdded,cartListID) VALUES (" & _
		"'" & escape_string(thesessionid) & "'," & IIfVr(SESSION("clientID")<>"", SESSION("clientID"), 0) & ",'" & escape_string(theid) & "','" & IIfVs(theid<>origid,escape_string(origid)) & "'," & quantity & "," & cartCompleted & ",'" & escape_string(strip_tags2(ainame)) & "'," & vsround(aiprice,2) & ",0," & vsusdate(DateAdd("h",dateadjust,Now())) & "," & cartListID & ")"
		ect_query(sSQL)
		rs.open "SELECT @@IDENTITY AS lstIns",cnn,0,1
		additemtocart=int(cstr(rs("lstIns")))
		rs.close
	else
		rs.open "cart",cnn,1,3,&H0002
		rs.AddNew
		rs.Fields("cartSessionID")		= thesessionid
		if SESSION("clientID")<>"" then rs.Fields("cartClientID")=SESSION("clientID") else rs.Fields("cartClientID")=0
		rs.Fields("cartProdID")			= theid
		rs.Fields("cartOrigProdID")		= IIfVs(theid<>origid,origid)
		rs.Fields("cartQuantity")		= quantity
		rs.Fields("cartCompleted")		= cartCompleted
		rs.Fields("cartProdName")		= strip_tags2(ainame)
		rs.Fields("cartProdPrice")		= vsround(aiprice,2)
		rs.Fields("cartOrderID")		= 0
		rs.Fields("cartDateAdded")		= DateAdd("h",dateadjust,Now())
		rs.Fields("cartListID")			= cartListID
		rs.Update
		if mysqlserver=TRUE then
			rs.close
			rs.open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
			additemtocart=int(cstr(rs("lstIns")))
		else
			additemtocart=int(cstr(rs.Fields("cartID")))
		end if
		rs.close
	end if
end function
function addoption(opttoadd)
	optvalue=getpost(opttoadd)
	if NOT is_numeric(optvalue) then optvalue=""
	if (left(opttoadd,4)="optn" OR left(opttoadd,4)="optm") AND optvalue<>"" then
		if left(opttoadd,4)="optm" then
			optID=right(opttoadd, len(opttoadd)-4)
			quantity=optvalue
			if quantity<>"" AND is_numeric(optID) AND is_numeric(quantity) then
				if quantity>0 then
					totalquantity=totalquantity + quantity
					if theid=origid OR addalternateoptions then
						Session.LCID=1033
						sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&","&OWSP&"optPriceDiff,optWeightDiff,optType,optFlags FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&optID
						rs.open sSQL,cnn,0,1
						sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&escape_string(rs(getlangid("optGrpName",16)))&"','"&escape_string(rs(getlangid("optName",32)))&"',"
						if (rs("optFlags") AND 1)=0 then sSQL=sSQL & IIfVr(theid=origid,rs("optPriceDiff"),0) & "," else sSQL=sSQL & vsround((rs("optPriceDiff")*thepprice)/100.0, 2) & ","
						if (rs("optFlags") AND 2)=0 then sSQL=sSQL & rs("optWeightDiff") & ")" else sSQL=sSQL & multShipWeight(thepweight,rs("optWeightDiff")) & ")"
						rs.close
						ect_query(sSQL)
						Session.LCID=saveLCID
					end if
					call checkpricebreaks(theid, thepprice)
				end if
			end if
		elseif getpost("v"&opttoadd)="" then
			sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&","&OWSP&"optPriceDiff,optWeightDiff,optType,optFlags,optRegExp FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&optvalue
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if addalternateoptions<>TRUE AND trim(rs("optRegExp")&"")<>"" AND left(trim(rs("optRegExp")&""),1)<>"!" then
					' Do Nothing
				elseif abs(rs("optType"))<>3 AND abs(rs("optType"))<>5 then
					sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&escape_string(rs(getlangid("optGrpName",16)))&"','"&escape_string(rs(getlangid("optName",32)))&"',"
					if (rs("optFlags") AND 1)=0 then sSQL=sSQL & IIfVr(trim(rs("optRegExp")&"")<>"", 0, rs("optPriceDiff")) & "," else sSQL=sSQL & vsround((rs("optPriceDiff")*thepprice)/100.0, 2) & ","
					if (rs("optFlags") AND 2)=0 then sSQL=sSQL & rs("optWeightDiff") & ")" else sSQL=sSQL & multShipWeight(thepweight,rs("optWeightDiff")) & ")"
				else
					sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff) VALUES ("&cartID&","&rs("optID")&",'"&escape_string(rs(getlangid("optGrpName",16)))&"','',0,0)"
				end if
			end if
			rs.close
			ect_query(sSQL)
		else
			sSQL="SELECT optID,"&getlangid("optGrpName",16)&","&getlangid("optName",32)&",optTxtCharge,optMultiply,optAcceptChars,optType FROM options INNER JOIN optiongroup ON options.optGroup=optiongroup.optGrpID WHERE optID="&optvalue
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				theopttoadd=getpost("v"&opttoadd)
				optPriceDiff=IIfVr(rs("optTxtCharge")<0 AND theopttoadd<>"",abs(rs("optTxtCharge")),rs("optTxtCharge")*len(theopttoadd))
				optmultiply=0
				if rs("optMultiply")<>0 AND abs(rs("optType"))=3 then
					if is_numeric(theopttoadd) then optmultiply=cdbl(theopttoadd) else theopttoadd="#NAN"
				end if
				sSQL="INSERT INTO cartoptions (coCartID,coOptID,coOptGroup,coCartOption,coPriceDiff,coWeightDiff,coMultiply) VALUES ("&cartID&","&rs("optID")&",'"&escape_string(rs(getlangid("optGrpName",16)))&"','"&escape_string(left(theopttoadd,txtcollen))&"',"&optPriceDiff&",0," & IIfVr(abs(rs("optType"))=3,IIfVr(rs("optMultiply")<>0,1,0),0) & ")"
				ect_query(sSQL)
			end if
			rs.close
		end if
	end if
end function
sub addproduct(theid)
	sSQL="SELECT "&getlangid("pName",1)&","&WSP&"pPrice,pWeight FROM products WHERE "&IIfVr(NOT useStockManagement,"(pSell<>0 OR pBackOrder<>0) AND", "")&" pID='"&escape_string(theid)&"'"
	rs.open sSQL,cnn,0,1
	idexists=IIfVr(rs.EOF,0,1)
	if NOT rs.EOF then
		thepname=rs(getlangid("pName",1))
		thepprice=vsround(rs("pPrice"),2)
		thepweight=rs("pWeight")
	else
		thepname="Product ID Error"
		thepprice=0
		thepweight=0
	end if
	rs.close
	addedprods(0,numaddedprods)=theid
	addedprods(1,numaddedprods)=thepname
	addedprods(2,numaddedprods)=quantity
	addedprods(3,numaddedprods)=thepprice
	addedprods(4,numaddedprods)=idexists
	addedprods(5,numaddedprods)=""
	if idexists then
		cartID=additemtocart(thepname,thepprice)
		if cartID=0 then exit sub
		for index=0 to numoptions-1
			if optarray(index)="multioption" then addoption(objForm) else addoption(optarray(index))
		next
		addedprods(5,numaddedprods)=cartID
	end if
	numaddedprods=numaddedprods+1
	if numaddedprods>=UBOUND(addedprods, 2) then redim preserve addedprods(5, UBOUND(addedprods, 2) + 20)
end sub
if checkoutmode="add" then ' {
	Dim optarray(),addedprods()
	redim optarray(20)
	redim addedprods(5,20)
	origid=theid
	thesessionid=getsessionid()
	if SESSION("clientID")<>"" AND is_numeric(getpost("listid")) then listid=getpost("listid") else listid=""
	if estimateshipping=TRUE then SESSION("xsshipping")=empty : SESSION("tofreeshipamount")=empty : SESSION("tofreeshipquant")=empty
	if NOT IsEmpty(SESSION("discounts")) then SESSION("discounts")=empty
	if NOT IsEmpty(SESSION("xscountrytax")) then SESSION("xscountrytax")=empty
	sSQL="SELECT ordID FROM orders WHERE ordStatus>1 AND ordAuthNumber='' AND " & getordersessionsql()
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then ordID=rs("ordID") else ordID=""
	rs.close
	if ordID<>"" then
		release_stock(ordID)
		ect_query("UPDATE cart SET cartSessionID='"&replace(getsessionid(),"'","")&"',cartClientID="&IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0)&" WHERE cartCompleted=0 AND cartOrderID=" & ordID)
		ect_query("UPDATE orders SET ordAuthStatus='MODWARNOPEN',ordShipType='MODWARNOPEN' WHERE ordID=" & ordID)
	end if
	Session.LCID=1033
	if NOT is_numeric(getpost("quant")) then quantity=1 else quantity=abs(int(getpost("quant")))
	if quantity<1 then quantity=1
	if quantity>99999 then quantity=99999
	hasmultioption=FALSE
	origquantity=quantity
	altids=""
	numoptions=0
	numaddedprods=0
	for jj=1 to request.form.Count
		for each objElem in request.form
			if request.form(objElem) is request.form(jj) then objForm=objElem : exit for
		next
		if left(objForm,4)="optn" AND is_numeric(getpost(objForm)) then
			doaddoption=FALSE
			sSQL="SELECT optRegExp FROM options WHERE optID="&replace(getpost(objForm),"'","")
			rs2.Open sSQL,cnn,0,1
			if rs2.EOF then theexp="" else theexp=trim(rs2("optRegExp")&"")
			if theexp<>"" AND Left(theexp,1)<>"!" then
				theexp=replace(theexp, "%s", theid)
				altids=altids&":"&getpost(objForm)&":"
				if InStr(theexp, " ")>0 then ' Search and replace
					exparr=split(theexp, " ", 2)
					theid=replace(theid, exparr(0), exparr(1), 1, 1)
				else
					theid=theexp
				end if
				if addalternateoptions=TRUE then doaddoption=TRUE
			else
				doaddoption=TRUE
			end if
			if doaddoption then
				if numoptions>=UBOUND(optarray) then redim preserve optarray(UBOUND(optarray)+20)
				optarray(numoptions)=objForm
				numoptions=numoptions+1
			end if
			rs2.Close
		elseif left(objForm,4)="optm" AND is_numeric(getpost(objForm)) then
			if is_numeric(right(objForm, len(objForm)-4)) then
				if NOT hasmultioption then
					if numoptions>=UBOUND(optarray) then redim preserve optarray(UBOUND(optarray)+20)
					optarray(numoptions)="multioption"
					numoptions=numoptions+1
				end if
				hasmultioption=TRUE
			end if
		end if
	next
	if hasmultioption then
		for jj=1 to request.form.Count
			for each objElem in request.form
				if request.form(objElem) is request.form(jj) then objForm=objElem : exit for
			next
			if left(objForm,4)="optm" AND is_numeric(getpost(objForm)) then
				if is_numeric(right(objForm, len(objForm)-4)) then
					quantity=abs(int(getpost(objForm)))
					if quantity>99999 then quantity=99999
					theid=origid
					sSQL="SELECT optRegExp FROM options WHERE optID="&replace(right(objForm, len(objForm)-4),"'","")
					rs2.Open sSQL,cnn,0,1
					if rs2.EOF then theexp="" else theexp=trim(rs2("optRegExp")&"")
					if theexp<>"" AND Left(theexp,1)<>"!" then
						theexp=replace(theexp, "%s", theid)
						if InStr(theexp, " ")>0 then ' Search and replace
							exparr=split(theexp, " ", 2)
							theid=replace(theid, exparr(0), exparr(1), 1, 1)
						else
							theid=theexp
						end if
					end if
					rs2.Close
					call addproduct(theid)
				end if
			end if
		next
	else
		call addproduct(theid)
	end if
	' Check duplicates
	sSQL="SELECT cartID,cartProdID,cartQuantity FROM cart WHERE cartCompleted="&IIfVr(listid="",0,3)&" AND " & getsessionsql() & IIfVr(listid="",""," AND cartListID="&listid) & " ORDER BY cartID"
	rs.open sSQL,cnn,0,1
		if NOT rs.EOF then cartarr=rs.getRows else cartarr=""
	rs.close
	if isarray(cartarr) then
		for index=0 to UBOUND(cartarr, 2)
			hasoptions=FALSE
			sSQL="SELECT cartID,cartQuantity FROM cart WHERE cartID>" & cartarr(0,index) & " AND cartCompleted=0 AND " & getsessionsql() & IIfVr(listid="",""," AND cartListID="&listid) & " AND cartProdID='" & escape_string(cartarr(1,index)) & "'"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				thecartid=rs("cartID")
				thequant=rs("cartQuantity")
				hasoptions=TRUE
				if thecartid<>"" then ' check options
					sSQL="SELECT coOptID,coCartOption FROM cartoptions WHERE coCartID=" & cartarr(0, index)
					rs2.Open sSQL,cnn,0,1
						if NOT rs2.EOF then optarr1=rs2.getRows else optarr1=""
					rs2.Close
					sSQL="SELECT coOptID,coCartOption FROM cartoptions WHERE coCartID=" & thecartid
					rs2.Open sSQL,cnn,0,1
						if NOT rs2.EOF then optarr2=rs2.getRows else optarr2=""
					rs2.Close
					if (isarray(optarr1) AND NOT isarray(optarr2)) OR (NOT isarray(optarr1) AND isarray(optarr2)) then hasoptions=FALSE
					if isarray(optarr1) AND isarray(optarr2) then
						if UBOUND(optarr1,2)<>UBOUND(optarr2,2) then hasoptions=FALSE
						if hasoptions then
							for index2=0 to UBOUND(optarr1,2)
								hasthisoption=FALSE
								for index3=0 to UBOUND(optarr2,2)
									if optarr1(0,index2)=optarr2(0,index3) AND optarr1(1,index2)=optarr2(1,index3) then hasthisoption=TRUE
								next
								if NOT hasthisoption then hasoptions=FALSE
							next
						end if
					end if
				end if
				if hasoptions then exit do
				rs.movenext
			loop
			rs.close
			if thecartid<>"" AND hasoptions then
				ect_query("DELETE FROM cartoptions WHERE coCartID="&thecartid)
				ect_query("DELETE FROM cart WHERE cartID="&thecartid)
				ect_query("UPDATE cart SET cartQuantity=cartQuantity+"&thequant&" WHERE cartID="&cartarr(0,index))
				for index2=0 to numaddedprods-1
					if cstr(addedprods(5,index2))=cstr(thecartid) then addedprods(5,index2)=cartarr(0,index)
				next
			end if
		next
	end if
	for index=0 to numaddedprods-1
		if addedprods(4,index) then call checkpricebreaks(addedprods(0,index), addedprods(3,index)) else actionaftercart=0 : cartrefreshseconds=3
	next
	if getpost("ajaxadd")="true" then
		failedprods=0
		call do_stock_check(TRUE,backorder,stockwarning)
		scidnoexistflag="" : scbackorderflag=0 : scinstockflag=0
		for index2=0 to numaddedprods-1
			if addedprods(4,index2)=0 then scidnoexistflag=addedprods(0,index2) : failedprods=failedprods+1
			if addedprods(4,index2)=2 then scinstockflag=1
			if addedprods(4,index2)=3 then scbackorderflag=1
		next
		response.clear
		print jsurlencode(scidnoexistflag) & "&" & scinstockflag & "&" & scbackorderflag & "&" & (numaddedprods-failedprods)
		sSQL="SELECT cartID,cartProdPrice,cartQuantity,pExemptions,pTax FROM cart INNER JOIN products ON cart.cartProdId=products.pID WHERE cartCompleted=0 AND " & getsessionsql()
		rs.open sSQL,cnn,0,1
		sctotquant=0 : totalgoods=0
		do while NOT rs.EOF
			optPriceDiff=0
			pexemptions=rs("pExemptions")
			thetax=homeCountryTaxRate
			if perproducttaxrate AND NOT isnull(rs("pTax")) then thetax=rs("pTax")
			sSQL="SELECT SUM(coPriceDiff) AS sumDiff FROM cartoptions WHERE coCartID="&rs("cartID")
			rs2.Open sSQL,cnn,0,1
			if NOT isnull(rs2("sumDiff")) then optPriceDiff=rs2("sumDiff")
			rs2.Close
			subtot=((rs("cartProdPrice")+optPriceDiff)*int(rs("cartQuantity")))
			sctotquant=sctotquant+int(rs("cartQuantity"))
			totalgoods=totalgoods+subtot
			if perproducttaxrate then
				if (pexemptions AND 2)<>2 then countryTax=countryTax+(subtot*thetax/100.0)
			else
				if (pexemptions AND 2)=2 then countrytaxfree=countrytaxfree+subtot
			end if
			for index=0 to numaddedprods-1
				if addedprods(5,index)=rs("cartID") then addedprods(3,index)=rs("cartProdPrice")
			next
			rs.movenext
		loop
		rs.close
		call calculatediscounts(totalgoods,FALSE,rgcpncode)
		if totaldiscounts>totalgoods then totaldiscounts=totalgoods
		if showtaxinclusive<>0 then calculatetaxandhandling() else countryTax=0
		print "&"&sctotquant&"&"&jsurlencode(FormatEuroCurrency(IIfVr(nopriceanywhere,0,(totalgoods-totaldiscounts)+countryTax)))&"&"
		sSQL="SELECT "&IIfVs(mysqlserver<>TRUE,"TOP 1 ")&"imageSrc FROM productimages WHERE imageType=0 AND imageProduct='"&escape_string(theid)&"' ORDER BY imageNumber"&IIfVs(mysqlserver=TRUE," LIMIT 0,1")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then print jsurlencode(rs("imageSrc"))
		rs.close
		for index=0 to numaddedprods-1
			if addedprods(5,index)<>"" then
				pexemptions=0 : totoptdiff=0
				thetax=homeCountryTaxRate
				sSQL="SELECT pExemptions,pTax FROM products WHERE pID='"&escape_string(theid)&"'"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					pexemptions=rs("pExemptions")
					if perproducttaxrate AND NOT isnull(rs("pTax")) then thetax=rs("pTax")
				end if
				rs.close
				sSQL="SELECT coOptGroup,coCartOption,coPriceDiff FROM cartoptions WHERE coCartID=" & addedprods(5,index) & " ORDER BY coID"
				rs.cursorlocation=3
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then
					optresult="&"&rs.recordcount
					do while NOT rs.EOF
						totoptdiff=totoptdiff+rs("coPriceDiff")
						optresult=optresult&"&"&jsurlencode(rs("coOptGroup"))&"&"&jsurlencode(rs("coCartOption"))
						rs.movenext
					loop
				else
					optresult="&0"
				end if
				rs.close
				totitemcost=addedprods(3,index)+totoptdiff
				for index2=0 to 4
					print "&"
					if index2=3 then print jsurlencode(FormatEuroCurrency(IIfVr(nopriceanywhere,0,totitemcost))) else print jsurlencode(addedprods(index2,index))
				next
				print optresult
			end if
		next
		print "&"&jsurlencode(FormatEuroCurrency(IIfVr(nopriceanywhere,0,countryTax)))&"&"&jsurlencode(FormatEuroCurrency(IIfVr(nopriceanywhere,0,totaldiscounts)))
		if totaldiscounts=0 then SESSION("discounts")=empty : print "&0&" else SESSION("discounts")=totaldiscounts : print "&1&"
		sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity FROM cart WHERE cartCompleted=0 AND "&getsessionsql()
		rs.open sSQL,cnn,0,1
		do while NOT rs.EOF
			print jsurlencode("<div class=""minicartcnt"">"&rs("cartQuantity")&" "&rs("cartProdName")&"</div>")
			rs.movenext
		loop
		rs.close
	else
		print "<div style=""text-align:center;padding:40px"">"
		call do_stock_check(FALSE,backorder,stockwarning)
		if stockwarning then actionaftercart=4
		if cartrefreshseconds="" then cartrefreshseconds=3
		if listid<>"" then listidurl="&mode=sc" & IIfVr(listid<>"0","&lid="&listid,"") else listidurl=""
		if thefrompage<>"" AND actionaftercart=3 then
			if cartrefreshseconds=0 then response.redirect thefrompage else print "<meta http-equiv=""Refresh"" content="""&cartrefreshseconds&"; URL="&thefrompage&""">"
		elseif actionaftercart=4 OR cartrefreshseconds=0 then
			urllink="?rp="&urlencode(thefrompage)
			if listid<>"" AND SESSION("clientID")<>"" then
				urllink=urllink & listidurl
			elseif stockwarning then
				urllink=urllink & "&mode=add"
			end if
			if getpost("PARTNER")<>"" then urllink=urllink & "&PARTNER="&strip_tags2(getpost("partner"))
			response.redirect "cart" & extension & urllink
		else
			print "<meta http-equiv=""Refresh"" content="""&cartrefreshseconds&"; URL=cart" & extension & "?rp="&urlencode(thefrompage)&IIfVr(getpost("PARTNER")<>"","&PARTNER="&strip_tags2(getpost("partner")),"")&listidurl&""">"
		end if
		print "<div class=""hardcarttable"" style=""display:table;padding:10px;width:auto;margin-left:auto;margin-right:auto;border:1px solid grey"">"
		if stockwarning then print "<div style=""text-align:center"" class=""hardcartstockwarn"">" & xxInsMul & "</div>"
		for index=0 to numaddedprods-1
			print "<div class=""hardcartaddproductline"" style=""display:table-row""><div class=""hardcartaddproductquant"" style=""display:table-cell;text-align:right;padding:6px"">" & IIfVr(addedprods(4,index),addedprods(2,index),"X") & "</div><div class=""hardcartaddproduct"" style=""display:table-cell;text-align:left;padding:6px;"">" & IIfVr(addedprods(4,index),addedprods(1,index) & " " & xxAddOrd,"<span class=""ectwarning"">The product id <span style=""color:#000000"">" & htmlspecials(addedprods(0,index)) & "</span> does not exist in the product database.</span></span>") & "</div></div>"
		next
		print "</div>"
		print "<div class=""hardcartpleasewait"" style=""padding:10px;margin:50px 10px"">" & xxPlsWait & " <a class=""ectlink"" href="""
		if thefrompage<>"" AND actionaftercart=3 then print thefrompage else print "cart" & extension & "?rp="&urlencode(thefrompage)&listidurl
		print """><strong>" & xxClkHere & "</strong></a>.</div>"
		print "</div>"
	end if
elseif checkoutmode="go" OR paypalexpress OR amazonpayment then ' }{
	call getadminshippingparams()
%>
<!--#include file="uspsshipping.asp"-->
<%	function setcheckouterr(coerrmsg)
		success=FALSE : checkoutmode="checkout" : returntocustomerdetails=TRUE : errormsg=errormsg&"<div>" & coerrmsg & "</div>"
		setcheckouterr=TRUE
	end function
	if is_numeric(getpost("orderid")) AND getpost("sessionid")<>"" then
		call retrieveorderdetails(getpost("orderid"), getpost("sessionid"))
	elseif NOT (paypalexpress OR amazonpayment) then
		if enableclientlogin AND SESSION("clientID")<>"" then
			sSQL="SELECT clEmail FROM customerlogin WHERE clEmail<>'' AND clID=" & replace(SESSION("clientID"),"'","")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then ordEmail=trim(rs("clEmail")&"") else ordEmail=cleanupemail(getpost("email"))
			rs.close
		else
			ordEmail=cleanupemail(getpost("email"))
			ordEmail2=cleanupemail(getpost("email2"))
		end if
		if enableclientlogin AND is_numeric(getpost("addressid")) AND getpost("addaddress")="" AND SESSION("clientID")<>"" then
			sSQL="SELECT addName,addLastName,addAddress,addAddress2,addCity,addState,addZip,addCountry,addPhone,addExtra1,addExtra2 FROM address WHERE addCustID="&replace(SESSION("clientID"),"'","")&" AND addID="&replace(getpost("addressid"),"'","")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				ordName=rs("addName")
				ordLastName=rs("addLastName")
				ordAddress=rs("addAddress")
				ordAddress2=rs("addAddress2")
				ordCity=rs("addCity")
				ordState=rs("addState")
				ordZip=rs("addZip")
				ordCountry=rs("addCountry")
				ordPhone=rs("addPhone")
				ordExtra1=rs("addExtra1")
				ordExtra2=rs("addExtra2")
				ect_query("UPDATE address SET addIsDefault=0 WHERE addCustID="&replace(SESSION("clientID"),"'",""))
				ect_query("UPDATE address SET addIsDefault=1 WHERE addCustID="&replace(SESSION("clientID"),"'","")&" AND addID="&replace(getpost("addressid"),"'",""))
			end if
			rs.close
		else
			ordName=left(strip_tags2(getpost("ordname")),100)
			ordLastName=left(strip_tags2(getpost("lastname")),100)
			ordAddress=left(strip_tags2(getpost("address")),150)
			ordAddress2=left(strip_tags2(getpost("address2")),150)
			ordCity=left(strip_tags2(getpost("city")),75)
			ordState=getstatefromid(left(strip_tags2(getpost("state"&IIfVr(getpost("state")="","2",""))),75))
			ordZip=left(strip_tags2(getpost("zip")),50)
			ordCountry=left(getcountryfromid(getpost("country")),50)
			ordPhone=left(strip_tags2(getpost("phone")),50)
			ordExtra1=left(strip_tags2(getpost("ordextra1")),255)
			ordExtra2=left(strip_tags2(getpost("ordextra2")),255)
		end if
		if getpost("allowemail")="ON" then call addtomailinglist(ordEmail,trim(ordName&" "&ordLastName))
		if enableclientlogin AND is_numeric(getpost("saddressid")) AND getpost("saddaddress")="" AND SESSION("clientID")<>"" then
			sSQL="SELECT addName,addLastName,addAddress,addAddress2,addCity,addState,addZip,addCountry,addPhone,addExtra1,addExtra2 FROM address WHERE addCustID="&replace(SESSION("clientID"),"'","")&" AND addID="&replace(getpost("saddressid"),"'","")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				ordShipName=rs("addName")
				ordShipLastName=rs("addLastName")
				ordShipAddress=rs("addAddress")
				ordShipAddress2=rs("addAddress2")
				ordShipCity=rs("addCity")
				ordShipState=rs("addState")
				ordShipZip=rs("addZip")
				ordShipCountry=rs("addCountry")
				ordShipPhone=rs("addPhone")
				ordShipExtra1=rs("addExtra1")
				ordShipExtra2=rs("addExtra2")
			end if
			rs.close
		else
			if getpost("shipdiff")="1" OR (getpost("addressarray")="1" AND (getpost("saddaddress")="add" OR getpost("saddaddress")="edit")) then
				ordShipName=left(strip_tags2(getpost("sordname")),100)
				ordShipLastName=left(strip_tags2(getpost("slastname")),100)
				ordShipAddress=left(strip_tags2(getpost("saddress")),150)
				ordShipAddress2=left(strip_tags2(getpost("saddress2")),150)
				ordShipCity=left(strip_tags2(getpost("scity")),75)
				ordShipState=getstatefromid(left(strip_tags2(getpost("sstate"&IIfVr(getpost("sstate")="","2",""))),75))
				ordShipZip=left(strip_tags2(getpost("szip")),50)
				ordShipCountry=left(getcountryfromid(getpost("scountry")),50)
				ordShipPhone=left(strip_tags2(getpost("sphone")),50)
				ordShipExtra1=left(strip_tags2(getpost("ordsextra1")),255)
				ordShipExtra2=left(strip_tags2(getpost("ordsextra2")),255)
				if (trim(ordShipName & ordShipLastName)="" OR ordShipAddress="" OR ordShipCity="" OR ordShipZip="") AND trim(ordShipName & ordShipLastName & ordShipAddress & ordShipCity & ordShipState & ordShipZip & ordShipPhone & "")<>"" then
					errordshipaddress=setcheckouterr("If you enter a separate ship address you must enter all address details.")
				end if
			end if
		end if
		ordZip=ucase(ordZip)
		ordShipZip=ucase(ordShipZip)
		if SESSION("clientID")<>"" then
			if trim(ordName & ordLastName)<>"" AND ordAddress<>"" AND ordCity<>"" AND ordState<>"" AND ordCountry<>"" AND ordZip<>"" AND ordPhone<>"" then
				if getpost("addaddress")="add" then
					sSQL="SELECT addID FROM address WHERE addCustID="&replace(SESSION("clientID"),"'","")&" AND addName='"&escape_string(ordName)&"' AND addLastName='"&escape_string(ordLastName)&"' AND addAddress='"&escape_string(ordAddress)&"' AND addAddress2='"&escape_string(ordAddress2)&"' AND addCity='"&escape_string(ordCity)&"' AND addState='"&escape_string(ordState)&"' AND addZip='"&escape_string(ordZip)&"' AND addCountry='"&escape_string(ordCountry)&"' AND addPhone='"&escape_string(ordPhone)&"' AND addExtra1='"&escape_string(ordExtra1)&"' AND addExtra2='"&escape_string(ordExtra2)&"'"
					rs.open sSQL,cnn,0,1
					hasaddress=NOT rs.EOF
					rs.close
					sSQL="INSERT INTO address (addCustID,addIsDefault,addName,addLastName,addAddress,addAddress2,addCity,addState,addZip,addCountry,addPhone,addExtra1,addExtra2) VALUES ("&replace(SESSION("clientID"),"'","")&",0,'"&escape_string(ordName)&"','"&escape_string(ordLastName)&"','"&escape_string(ordAddress)&"','"&escape_string(ordAddress2)&"','"&escape_string(ordCity)&"','"&escape_string(ordState)&"','"&escape_string(ordZip)&"','"&escape_string(ordCountry)&"','"&escape_string(ordPhone)&"','"&escape_string(ordExtra1)&"','"&escape_string(ordExtra2)&"')"
					if NOT hasaddress then ect_query(sSQL)
				elseif getpost("addaddress")="edit" then
					sSQL="UPDATE address SET addName='"&escape_string(ordName)&"',addLastName='"&escape_string(ordLastName)&"',addAddress='"&escape_string(ordAddress)&"',addAddress2='"&escape_string(ordAddress2)&"',addCity='"&escape_string(ordCity)&"',addState='"&escape_string(ordState)&"',addZip='"&escape_string(ordZip)&"',addCountry='"&escape_string(ordCountry)&"',addPhone='"&escape_string(ordPhone)&"',addExtra1='"&escape_string(ordExtra1)&"',addExtra2='"&escape_string(ordExtra2)&"' WHERE addCustID="&replace(SESSION("clientID"),"'","")&" AND addID=" & replace(getpost("addressid"),"'","")
					ect_query(sSQL)
				end if
			end if
			if trim(ordShipName & ordShipLastName)<>"" AND ordShipAddress<>"" AND ordShipCity<>"" AND ordShipState<>"" AND ordShipCountry<>"" AND ordShipZip<>"" then
				if getpost("saddaddress")="add" then
					sSQL="SELECT addID FROM address WHERE addCustID="&replace(SESSION("clientID"),"'","")&" AND addName='"&escape_string(ordShipName)&"' AND addLastName='"&escape_string(ordShipLastName)&"' AND addAddress='"&escape_string(ordShipAddress)&"' AND addAddress2='"&escape_string(ordShipAddress2)&"' AND addCity='"&escape_string(ordShipCity)&"' AND addState='"&escape_string(ordShipState)&"' AND addZip='"&escape_string(ordShipZip)&"' AND addCountry='"&escape_string(ordShipCountry)&"' AND addPhone='"&escape_string(ordShipPhone)&"' AND addExtra1='"&escape_string(ordShipExtra1)&"' AND addExtra2='"&escape_string(ordShipExtra2)&"'"
					rs.open sSQL,cnn,0,1
					hasaddress=NOT rs.EOF
					rs.close
					sSQL="INSERT INTO address (addCustID,addIsDefault,addName,addLastName,addAddress,addAddress2,addCity,addState,addZip,addCountry,addPhone,addExtra1,addExtra2) VALUES ("&replace(SESSION("clientID"),"'","")&",0,'"&escape_string(ordShipName)&"','"&escape_string(ordShipLastName)&"','"&escape_string(ordShipAddress)&"','"&escape_string(ordShipAddress2)&"','"&escape_string(ordShipCity)&"','"&escape_string(ordShipState)&"','"&escape_string(ordShipZip)&"','"&escape_string(ordShipCountry)&"','"&escape_string(ordShipPhone)&"','"&escape_string(ordShipExtra1)&"','"&escape_string(ordShipExtra2)&"')"
					if NOT hasaddress then ect_query(sSQL)
				elseif getpost("saddaddress")="edit" then
					sSQL="UPDATE address SET addName='"&escape_string(ordShipName)&"',addLastName='"&escape_string(ordShipLastName)&"',addAddress='"&escape_string(ordShipAddress)&"',addAddress2='"&escape_string(ordShipAddress2)&"',addCity='"&escape_string(ordShipCity)&"',addState='"&escape_string(ordShipState)&"',addZip='"&escape_string(ordShipZip)&"',addCountry='"&escape_string(ordShipCountry)&"',addPhone='"&escape_string(ordShipPhone)&"',addExtra1='"&escape_string(ordShipExtra1)&"',addExtra2='"&escape_string(ordShipExtra2)&"' WHERE addCustID="&replace(SESSION("clientID"),"'","")&" AND addID=" & replace(getpost("saddressid"),"'","")
					ect_query(sSQL)
				end if
			end if
		end if
		ordAddInfo=left(strip_tags2(getpost("ordAddInfo")),4096)
		if commercialloc_ then ordComLoc=1 else ordComLoc=0
		if wantinsurance_ OR abs(addshippinginsurance)=1 then ordComLoc=ordComLoc+2
		if saturdaydelivery_ then ordComLoc=ordComLoc+4
		if signaturerelease_ then ordComLoc=ordComLoc+8
		if insidedelivery_ then ordComLoc=ordComLoc+16
		ordAffiliate=strip_tags2(left(getpost("PARTNER"),48))
		ordCheckoutExtra1=strip_tags2(left(getpost("ordcheckoutextra1"),255))
		ordCheckoutExtra2=strip_tags2(left(getpost("ordcheckoutextra2"),255))
	end if
	if ordShipAddress<>"" then
		shipcountry=ordShipCountry
		shipstate=ordShipState
		destZip=ordShipZip
	else
		shipcountry=ordCountry
		shipstate=ordState
		destZip=ordZip
		if autobillingtoshipping=TRUE then
			ordShipName=ordName
			ordShipLastName=ordLastName
			ordShipAddress=ordAddress
			ordShipAddress2=ordAddress2
			ordShipCity=ordCity
			ordShipState=ordState
			ordShipZip=ordZip
			ordShipCountry=ordCountry
			ordShipPhone=ordPhone
			ordShipExtra1=ordExtra1
			ordShipExtra2=ordExtra2
		end if
	end if
	sSQL="SELECT countryID,countryCode,countryCode3,loadStates FROM countries WHERE countryName='"&escape_string(ordCountry)&"'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		countryID=rs("countryID")
		countryCode=rs("countryCode")
		countryCode3=rs("countryCode3")
		homecountry=(rs("countryID")=origCountryID)
		if rs("loadStates")<>-1 AND ordState="" then errordstate=setcheckouterr(replace(xxMusSta,"%s",getstatetext(countryID)))
	else
		success=FALSE
	end if
	rs.close
	'******* Modified International Handling Fee by DLSS ********
	if NOT homecountry then
		perproducttaxrate=FALSE
		handling=IntlHandling
	End If
	'************************************************************
	sSQL="SELECT countryID,countryTax,countryTaxThreshold,countryCode,countryCode3,countryFreeShip,loadStates FROM countries WHERE countryName='"&escape_string(shipcountry)&"'"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		countryTaxRate=rs("countryTax")
		countrytaxthreshold=rs("countryTaxThreshold")
		shipCountryID=rs("countryID")
		shipCountryCode=rs("countryCode")
		shipCountryCode3=rs("countryCode3")
		freeshipavailtodestination=(rs("countryFreeShip")=1)
		shiphomecountry=(rs("countryID")=origCountryID) OR ((rs("countryID")=1 OR rs("countryID")=2) AND usandcasplitzones)
		if rs("loadStates")<>-1 AND shipstate="" AND ordShipAddress<>"" then errordshipstate=setcheckouterr(replace(xxMusSta,"%s",getstatetext(shipCountryID)))
	else
		if shipcountry<>ordCountry then errordshipcountry=setcheckouterr("You must select a ship country.")
		if shipcountry<>"" then errormsg=errormsg & "<div>Invalid countryName:" & htmldisplay(shipcountry) & "</div>"
	end if
	rs.close
	sSQL="SELECT shipInsurance"&IIfVr(shiphomecountry,"Dom","Int")&",insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Min,insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Percent,noCarrier"&IIfVr(shiphomecountry,"Dom","Int")&"Ins FROM admin WHERE adminID=1"
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then
		addshippinginsurance=rs("shipInsurance"&IIfVr(shiphomecountry,"Dom","Int"))
		shipinsurancemin=rs("insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Min")
		shipinsurancepercent=rs("insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Percent")
		nocarrierinsurancerates=rs("noCarrier"&IIfVr(shiphomecountry,"Dom","Int")&"Ins")<>0
	end if
	if addshippinginsurance=3 then forceinsuranceselection=TRUE : addshippinginsurance=2
	rs.close
	orderid="" : ordauthstatus=""
	if success then
		if countryID=1 OR countryID=2 then stateAbbrev=getstateabbrev(ordState)
		if shipCountryID=1 OR shipCountryID=2 then shipStateAbbrev=getstateabbrev(shipstate)
		if shipCountryID<>"" then
			sSQL="SELECT stateID,stateTax,stateAbbrev,stateFreeShip FROM states WHERE stateCountryID=" & shipCountryID & " AND (stateName='"&escape_string(shipstate)&"'"
			if shipCountryID=1 OR shipCountryID=2 then sSQL=sSQL & " OR stateAbbrev='"&escape_string(shipstate)&"')" else sSQL=sSQL & ")"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if shiphomecountry then
					if shipCountryID=origCountryID OR ((shipCountryID=1 OR shipCountryID=2) AND usandcasplitzones) then stateTaxRate=rs("stateTax") else stateTaxRate=0
					freeshipavailtodestination=(freeshipavailtodestination AND (rs("stateFreeShip")=1))
				end if
				shipstateid=rs("stateID")
			end if
			rs.close
			if shiphomecountry AND willpickup_ then
				if NOT isempty(homestatetaxrate) then
					stateTaxRate=homestatetaxrate
				else
					rs.open "SELECT MAX(stateTax) as maxtax FROM states WHERE stateCountryID=" & shipCountryID & " AND stateEnabled=1",cnn,0,1
					if NOT rs.EOF then stateTaxRate=rs("maxtax")
					rs.close
				end if
			end if
		end if
		if (shipType=4 OR shipType=7 OR shipType=8) AND shipCountryID=1 AND shipStateAbbrev="GU" then shipCountryCode="GU"
		if trim(SESSION("clientID"))<>"" then
			if (SESSION("clientActions") AND 1)=1 then stateTaxRate=0
			if (SESSION("clientActions") AND 2)=2 then countryTaxRate=0
		end if
		getpayprovhandling()
		shipType=getshiptype()
		if NOT initshippingmethods() then success=FALSE : checkoutmode="checkout"
		sSQL="SELECT ordID,ordAuthStatus FROM orders WHERE ordStatus>1 AND ordAuthNumber='' AND " & getordersessionsql()
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then orderid=rs("ordID") : ordauthstatus=rs("ordAuthStatus")
		rs.close
		if mysqlserver=TRUE then
			sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax,pMinQuant,pID FROM cart LEFT JOIN products ON cart.cartProdID=products.pId LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE cartCompleted=0 AND " & getsessionsql() & " ORDER BY cartID"
		else
			sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions,pSection,topSection,pDims,pTax,pMinQuant,pID FROM cart LEFT JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE cartCompleted=0 AND " & getsessionsql() & " ORDER BY cartID"
		end if
		rs.open sSQL,cnn,0,1
		if NOT (rs.EOF OR rs.BOF) then alldata=rs.getrows
		rs.close
	end if
	if ordCountry="" then
		errordcountry=setcheckouterr("You must select a country.")
	else
		if ordEmail="" AND NOT amazonpayment then errordemail=setcheckouterr("You must enter a valid email address.")
		if verifyemail AND SESSION("clientID")="" AND ordEmail2 <> ordEmail AND getpost("shipselectoraction")="" AND NOT amazonpayment then errordemailv=setcheckouterr("Email verify does not match.")
		if trim(ordName&ordLastName)="" AND NOT amazonpayment then errordname=setcheckouterr("You must enter your name.")
		if ordAddress="" AND NOT amazonpayment then errordaddress=setcheckouterr("You must enter your address.")
		if ordCity="" then errordcity=setcheckouterr("You must enter a city.")
		if ordZip="" AND NOT zipisoptional(shipCountryID) then errordzip=setcheckouterr("You must enter a zip / postal code.")
		if ordPhone="" AND NOT amazonpayment AND NOT paypalexpress then errordphone=setcheckouterr("You must enter a phone number.")
	end if
	if NOT is_numeric(ordPayProvider) then errordpayprovider=setcheckouterr("You must select a payment method.")
	if (orderid="" AND getpost("shipselectoraction")<>"") then
		success=FALSE
		errormsg="Invalid Data"
	end if
	if termsandconditions AND getpost("license")<>"1" AND getpost("shipselectoraction")="" AND ordPayProvider<>"19" AND ordPayProvider<>"21" then errtermsandconditions=setcheckouterr("Please proceed only if you are in acceptance of our terms and conditions.")
	if success AND isarray(alldata) then
		rowcounter=0
		for index=0 to UBOUND(alldata,2)
			if isnull(alldata(5,index)) then alldata(5,index)=0
			if (alldata(1,index)=giftcertificateid OR alldata(1,index)=donationid) AND isnull(alldata(8,index)) then alldata(8,index)=15
			if alldata(1,index)=giftwrappingid AND isnull(alldata(8,index)) then alldata(8,index)=12
			sSQL="SELECT SUM(coPriceDiff) AS coPrDff FROM cartoptions WHERE coCartID="&alldata(0,index)
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if NOT isnull(rs("coPrDff")) then alldata(3,index)=cdbl(alldata(3,index))+cdbl(rs("coPrDff"))
			end if
			rs.close
			sSQL="SELECT SUM(coWeightDiff) AS coWghtDff FROM cartoptions WHERE coCartID="&alldata(0,index)
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if NOT isnull(rs("coWghtDff")) then alldata(5,index)=cdbl(alldata(5,index))+cdbl(rs("coWghtDff"))
			end if
			rs.close
			runTot=(alldata(3,index)*int(alldata(4,index)))
			totalquantity=totalquantity+alldata(4,index)
			totalgoods=totalgoods+runTot
			thistopcat=0
			if trim(SESSION("clientID"))<>"" then alldata(8,index)=(alldata(8,index) OR (SESSION("clientActions") AND 7)) : if (SESSION("clientActions") AND 32)=32 then alldata(8,index)=alldata(8,index) OR 8
			if (shipType=2 OR shipType=3 OR shipType=4 OR shipType>=6) AND cdbl(alldata(5,index))<=0.0 then alldata(8,index)=(alldata(8,index) OR 4)
			if (alldata(8,index) AND 1)=1 then statetaxfree=statetaxfree+runTot
			if (alldata(8,index) AND 8)<>8 then handlingeligableitem=TRUE : handlingeligablegoods=handlingeligablegoods+runTot
			if perproducttaxrate=TRUE then
				if isnull(alldata(12,index)) then alldata(12,index)=countryTaxRate
				if (alldata(8,index) AND 2)<>2 then countryTax=countryTax + ((alldata(12,index) * runTot) / 100.0)
			else
				if (alldata(8,index) AND 2)=2 then countrytaxfree=countrytaxfree + runTot
			end if
			if (alldata(8,index) AND 4)=4 then shipfreegoods=shipfreegoods + runTot
			if alldata(1,index)=giftcertificateid OR alldata(1,index)=donationid OR (alldata(8,index) AND 64)=64 then shipdiscountexempt=shipdiscountexempt+runTot : numshipdiscountexempt=numshipdiscountexempt+alldata(4,index)
			call addproducttoshipping(alldata, index)
			if alldata(4,index)<=alldata(13,index) then setcheckouterr(replace(replace(xxMinQuw,"%pname%",trim(alldata(2,index)&"")),"%quant%",alldata(13,index)+1))
			if isnull(alldata(14,index)) AND alldata(1,index)<>giftcertificateid AND alldata(1,index)<>donationid AND alldata(1,index)<>giftwrappingid then setcheckouterr(xxHasDel)
		next
		loyaltypointsused=0
		if success then
			call calculatediscounts(vsround(totalgoods,currDecimals),TRUE,rgcpncode)
			calculateshipping()
			if NOT fromshipselector then insuranceandtaxaddedtoshipping()
			call calculateshippingdiscounts(TRUE)
			if SESSION("clientID")<>"" AND SESSION("clientActions")<>0 then cpnmessage=cpnmessage & xxLIDis & htmlspecials(strip_tags2(replace(SESSION("clientUser"),"""",""))) & "<br>"
			cpnmessage=Right(cpnmessage,Len(cpnmessage)-4)
			calculatetaxandhandling()
			totalgoods=vsround(totalgoods,currDecimals)
			shipping=vsround(shipping,currDecimals)
			stateTax=vsround(stateTax,currDecimals)
			countryTax=vsround(countryTax,currDecimals)
			handling=vsround(handling,currDecimals)
			freeshipamnt=vsround(freeshipamnt,currDecimals)
			if loyaltypoints<>"" AND SESSION("clientID")<>"" AND SESSION("noredeempoints")<>TRUE then
				if NOT ((loyaltypointsnowholesale AND (SESSION("clientActions") AND 8)=8) OR (loyaltypointsnopercentdiscount AND (SESSION("clientActions") AND 16)=16)) then
					if orderid<>"" then
						pointsRedeemed=0
						rs.open "SELECT pointsRedeemed FROM orders WHERE ordID="&orderid,cnn,0,1
						if NOT rs.EOF then pointsRedeemed=rs("pointsRedeemed")
						rs.close
						if pointsRedeemed>0 then
							ect_query("UPDATE customerlogin SET loyaltyPoints=loyaltyPoints+" & pointsRedeemed & " WHERE clID=" & SESSION("clientID"))
							ect_query("UPDATE orders SET loyaltyPoints=0 WHERE ordID="&orderid)
						end if
					end if
					sSQL="SELECT loyaltyPoints FROM customerlogin WHERE clID=" & SESSION("clientID")
					rs.open sSQL,cnn,0,1
					if NOT rs.EOF then loyaltypointsused=rs("loyaltyPoints")
					rs.close
					if vsround(loyaltypointsused*loyaltypointvalue,2)>=IIfVr(loyaltypointminimum<>"",loyaltypointminimum,0.05) then
						loyaltypointdiscount=loyaltypointsused*loyaltypointvalue
						if loyaltypointdiscount>totalgoods+IIfVr(showtaxinclusive=3,countryTax,0)-totaldiscounts then loyaltypointdiscount=totalgoods+IIfVr(showtaxinclusive=3,countryTax,0)-totaldiscounts : loyaltypointsused=int(loyaltypointdiscount/loyaltypointvalue)
						totaldiscounts=totaldiscounts+vsround(loyaltypointdiscount,2)
						ect_query("UPDATE customerlogin SET loyaltyPoints=loyaltyPoints-" & loyaltypointsused & " WHERE clID=" & SESSION("clientID"))
						cpnmessage=cpnmessage & xxLoyPod & ": " & FormatEuroCurrency(loyaltypointdiscount) & "<br>"
					else
						loyaltypointsused=0
					end if
				end if
			end if
			if totaldiscounts>totalgoods+IIfVr(showtaxinclusive=3,countryTax,0) then totaldiscounts=totalgoods+IIfVr(showtaxinclusive=3,countryTax,0)
			if addshippingtodiscounts=TRUE then totaldiscounts=totaldiscounts + freeshipamnt : freeshipamnt=0
			totaldiscounts=vsround(totaldiscounts, 2)
			grandtotal=vsround((totalgoods + shipping + stateTax + countryTax + handling) - (totaldiscounts + freeshipamnt), 2)
			if grandtotal < 0 then grandtotal=0
			call do_stock_check(FALSE,backorder,stockwarning)
			if getpost("shipselectoraction")<>"" then stockwarning=FALSE : backorder=FALSE
			if stockwarning then
				checkoutmode=""
				response.redirect storeurl & "cart" & extension
			end if
		end if
		if (success OR getpost("shipselectoraction")="") AND NOT stockwarning then
			referer=SESSION("httpreferer")
			storeurlpos=instr(lcase(referer), parse_url(lcase(storeurl),1))
			storeurlsslpos=instr(lcase(referer), parse_url(lcase(storeurlssl),1))
			if (storeurlpos>0 AND storeurlpos<10) OR (storeurlssl<>"" AND storeurlsslpos>0 AND storeurlsslpos<10) then referer=""
			referarr=split(referer, "?", 2)
			if ordShipName="" AND ordShipLastName="" AND ordShipAddress="" AND ordShipAddress2="" AND ordShipCity="" then ordShipCountry=""
			if UBOUND(referarr)>=0 then ordReferer=left(referarr(0), 255) else ordReferer=""
			if UBOUND(referarr)>=1 then ordQuerystr=left(referarr(1), 255) else ordQuerystr=""
			ordIP=REMOTE_ADDR
			if sqlserver=TRUE then
				session.LCID=1033
				if orderid="" then
					isneworder=TRUE
					sSQL="INSERT INTO orders (ordSessionID,ordClientID,ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordShipPhone,ordPayProvider,ordAuthNumber,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordShipType,ordShipCarrier,ordTotal,ordDate,ordStatus,ordAuthStatus,pointsRedeemed,ordStatusDate,ordComLoc,ordIP,ordAffiliate,ordExtra1,ordExtra2,ordShipExtra1,ordShipExtra2,ordCheckoutExtra1,ordCheckoutExtra2,ordAVS,ordCVV,ordLang,ordReferer,ordQuerystr,ordDiscount,ordDiscountText,ordUserAgent,ordAddInfo) VALUES (" & _
						"'" & escape_string(thesessionid) & "'," & IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0) & "," & _
						"'" & escape_string(ordName) & "','" & escape_string(ordLastName) & "','" & escape_string(ordAddress) & "','" & escape_string(ordAddress2) & "'," & _
						"'" & escape_string(ordCity) & "','" & escape_string(ordState) & "','" & escape_string(ordZip) & "','" & escape_string(ordCountry) & "'," & _
						"'" & escape_string(ordEmail) & "','" & escape_string(ordPhone) & "'," & _
						"'" & escape_string(ordShipName) & "','" & escape_string(ordShipLastName) & "','" & escape_string(ordShipAddress) & "','" & escape_string(ordShipAddress2) & "'," & _
						"'" & escape_string(ordShipCity) & "','" & escape_string(ordShipState) & "','" & escape_string(ordShipZip) & "','" & escape_string(ordShipCountry) & "','" & escape_string(ordShipPhone) & "'," & _
						ordPayProvider & ",''," & _
						(shipping-freeshipamnt) & "," & _
						IIfVr(usehst, "0,0," & (stateTax+countryTax) & ",", stateTax & "," & countryTax & ",0,") & _
						handling & ",'" & escape_string(shipMethod) & "'," & shipType & "," & totalgoods & "," & _
						vsusdatetime(DateAdd("h",dateadjust,Now())) & ",2,''," & loyaltypointsused & "," & vsusdatetime(DateAdd("h",dateadjust,Now())) & "," & _
						ordComLoc & ",'" & escape_string(ordIP) & "','" & escape_string(ordAffiliate) & "'," & _
						"'" & escape_string(ordExtra1) & "','" & escape_string(ordExtra2) & "','" & escape_string(ordShipExtra1) & "','" & escape_string(ordShipExtra2) & "','" & escape_string(ordCheckoutExtra1) & "','" & escape_string(ordCheckoutExtra2) & "'," & _
						"'" & escape_string(ordAVS) & "','" & escape_string(ordCVV) & "'," & _
						IIfVr(languageid="",1,languageid)-1 & "," & _
						"'" & escape_string(ordReferer) & "','" & escape_string(ordQuerystr) & "'," & _
						totaldiscounts & ",'" & escape_string(left(cpnmessage,255)) & "'," & _
						"'" & IIfVs(captureuseragent,escape_string(left(request.servervariables("HTTP_USER_AGENT"),255))) & "'," & _
						"'" & escape_string(ordAddInfo) & "')"
					ect_query(sSQL)
					rs.open "SELECT @@IDENTITY AS lstIns",cnn,0,1
					orderid=int(cstr(rs("lstIns")))
					rs.close
				else
					isneworder=FALSE
					sSQL="UPDATE orders SET "
					if getpost("shipselectoraction")="" then
						sSQL=sSQL&"ordSessionID='" & escape_string(thesessionid) & "'," & _
							"ordClientID=" & IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0) & "," & _
							"ordName='" & escape_string(ordName) & "',ordLastName='" & escape_string(ordLastName) & "',ordAddress='" & escape_string(ordAddress) & "',ordAddress2='" & escape_string(ordAddress2) & "'," & _
							"ordCity='" & escape_string(ordCity) & "',ordState='" & escape_string(ordState) & "',ordZip='" & escape_string(ordZip) & "',ordCountry='" & escape_string(ordCountry) & "'," & _
							"ordEmail='" & escape_string(ordEmail) & "',ordPhone='" & escape_string(ordPhone) & "'," & _
							"ordShipName='" & escape_string(ordShipName) & "',ordShipLastName='" & escape_string(ordShipLastName) & "',ordShipAddress='" & escape_string(ordShipAddress) & "',ordShipAddress2='" & escape_string(ordShipAddress2) & "'," & _
							"ordShipCity='" & escape_string(ordShipCity) & "',ordShipState='" & escape_string(ordShipState) & "',ordShipZip='" & escape_string(ordShipZip) & "',ordShipCountry='" & escape_string(ordShipCountry) & "',ordShipPhone='" & escape_string(ordShipPhone) & "'," & _
							"ordPayProvider=" & ordPayProvider & ",ordAuthNumber=''," & _
							"ordComLoc=" & ordComLoc & ",ordIP='" & escape_string(ordIP) & "',ordAffiliate='" & escape_string(ordAffiliate) & "'," & _
							"ordExtra1='" & escape_string(ordExtra1) & "',ordExtra2='" & escape_string(ordExtra2) & "'," & _
							"ordShipExtra1='" & escape_string(ordShipExtra1) & "',ordShipExtra2='" & escape_string(ordShipExtra2) & "',ordCheckoutExtra1='" & escape_string(ordCheckoutExtra1) & "',ordCheckoutExtra2='" & escape_string(ordCheckoutExtra2) & "'," & _
							"ordAVS='" & escape_string(ordAVS) & "',ordCVV='" & escape_string(ordCVV) & "'," & _
							"ordLang=" & IIfVr(languageid="",1,languageid)-1 & "," & _
							"ordDiscount=" & totaldiscounts & "," & _
							"ordAddInfo='" & escape_string(ordAddInfo) & "',"
					end if
					sSQL=sSQL&"ordDate=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & ",ordStatusDate=" & vsusdatetime(DateAdd("h",dateadjust,Now())) & "," & _
						"ordShipping=" & (shipping-freeshipamnt) & "," & _
						"ordDiscountText='" & escape_string(left(cpnmessage,255)) & "'," & _
						"ordTotal=" & totalgoods & ",ordStateTax=" & IIfVr(usehst, "0,ordCountryTax=0,ordHSTTax=" & (stateTax+countryTax), stateTax & ",ordCountryTax=" & countryTax & ",ordHSTTax=0") & ",ordHandling=" & handling & "," & _
						"ordShipType='" & escape_string(shipMethod) & "',ordShipCarrier=" & shipType & ",ordAuthStatus='',pointsRedeemed=" & loyaltypointsused & _
						" WHERE ordID=" & orderid
					ect_query(sSQL)
				end if
				session.LCID=saveLCID
			else
				if orderid="" then
					rs.open "orders",cnn,1,3,&H0002
					rs.AddNew
					isneworder=TRUE
				else
					if mysqlserver then rs.CursorLocation=3
					rs.open "SELECT * FROM orders WHERE ordID="&orderid,cnn,1,3,&H0001
					isneworder=FALSE
				end if
				if getpost("shipselectoraction")="" then
					rs.Fields("ordSessionID")	= thesessionid
					rs.Fields("ordClientID")	= IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0)
					rs.Fields("ordName")		= ordName
					rs.Fields("ordLastName")	= ordLastName
					rs.Fields("ordAddress")		= ordAddress
					rs.Fields("ordAddress2")	= ordAddress2
					rs.Fields("ordCity")		= ordCity
					rs.Fields("ordState")		= ordState
					rs.Fields("ordZip")			= ordZip
					rs.Fields("ordCountry")		= ordCountry
					rs.Fields("ordEmail")		= ordEmail
					rs.Fields("ordPhone")		= ordPhone
					rs.Fields("ordShipName")	= ordShipName
					rs.Fields("ordShipLastName")= ordShipLastName
					rs.Fields("ordShipAddress")	= ordShipAddress
					rs.Fields("ordShipAddress2")= ordShipAddress2
					rs.Fields("ordShipCity")	= ordShipCity
					rs.Fields("ordShipState")	= ordShipState
					rs.Fields("ordShipZip")		= ordShipZip
					rs.Fields("ordShipCountry")	= ordShipCountry
					rs.Fields("ordShipPhone")	= ordShipPhone
					rs.Fields("ordPayProvider")	= ordPayProvider
					rs.Fields("ordAuthNumber")	= "" ' Not yet authorized
					rs.Fields("ordStatus")		= 2
					rs.Fields("ordIP")			= REMOTE_ADDR
					rs.Fields("ordComLoc")		= ordComLoc
					rs.Fields("ordAffiliate")	= ordAffiliate
					rs.Fields("ordAddInfo")		= ordAddInfo
					rs.Fields("ordDiscount")	= totaldiscounts
					rs.Fields("ordExtra1")		= ordExtra1
					rs.Fields("ordExtra2")		= ordExtra2
					rs.Fields("ordShipExtra1")	= ordShipExtra1
					rs.Fields("ordShipExtra2")	= ordShipExtra2
					rs.Fields("ordCheckoutExtra1")	= ordCheckoutExtra1
					rs.Fields("ordCheckoutExtra2")	= ordCheckoutExtra2
					rs.Fields("ordAVS")			= ordAVS
					rs.Fields("ordCVV")			= ordCVV
					rs.Fields("ordLang")		= IIfVr(languageid="",1,languageid)-1
					rs.Fields("ordReferer")		= ordReferer
					rs.Fields("ordQuerystr")	= ordQuerystr
				end if
				rs.Fields("ordDate")		= DateAdd("h",dateadjust,Now())
				rs.Fields("ordStatusDate")	= DateAdd("h",dateadjust,Now())
				rs.Fields("ordShipping")	= shipping - freeshipamnt
				rs.Fields("ordDiscountText")= left(cpnmessage,255)
				rs.Fields("ordTotal")		= totalgoods
				rs.Fields("ordHSTTax")		= IIfVr(usehst,stateTax+countryTax,0)
				rs.Fields("ordStateTax")	= IIfVr(usehst,0,stateTax)
				rs.Fields("ordCountryTax")	= IIfVr(usehst,0,countryTax)
				rs.Fields("ordHandling")	= handling
				rs.Fields("ordShipType")	= shipMethod
				rs.Fields("ordShipCarrier")	= shipType
				rs.Fields("ordAuthStatus")	= ""
				rs.Fields("pointsRedeemed")	= loyaltypointsused
				if orderid="" then rs.Fields("ordUserAgent")=IIfVs(captureuseragent,escape_string(left(request.servervariables("HTTP_USER_AGENT"),255)))
				rs.Update
				if mysqlserver=TRUE then
					if orderid="" then
						rs.close
						rs.open "SELECT LAST_INSERT_ID() AS lstIns",cnn,0,1
						orderid=rs("lstIns")
					end if
				else
					orderid=rs.Fields("ordID")
				end if
				rs.close
			end if
			sSQL="UPDATE cart SET cartOrderID="&orderid&" WHERE cartCompleted=0 AND " & getsessionsql()
			ect_query(sSQL)
			if isneworder OR ordauthstatus="MODWARNOPEN" then stock_subtract(orderid)
			sSQL="SELECT gcaGCID,gcaAmount FROM giftcertsapplied WHERE gcaOrdID="&orderid
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				ect_query("UPDATE giftcertificate SET gcRemaining=gcRemaining+"&vsround(rs("gcaAmount"), 2)&" WHERE gcID='"&rs("gcaGCID")&"'")
				rs.movenext
			loop
			rs.close
			ect_query("DELETE FROM giftcertsapplied WHERE gcaOrdID="&orderid)
			if SESSION("giftcerts")<>"" AND grandtotal>0 then
				sSQL="SELECT gcID,gcRemaining FROM giftcertificate WHERE gcRemaining>0 AND gcAuthorized<>0 AND gcID IN ('" & replace(escape_string(SESSION("giftcerts"))," ","','") & "')"
				rs.open sSQL,cnn,0,1
				do while NOT rs.EOF
					if giftcertsamount>=grandtotal then exit do
					thiscertamount=vrmin(grandtotal-giftcertsamount, rs("gcRemaining"))
					ect_query("INSERT INTO giftcertsapplied (gcaGCID,gcaOrdID,gcaAmount) VALUES ('"&rs("gcID")&"',"&orderid&","&thiscertamount&")")
					ect_query("UPDATE giftcertificate SET gcRemaining=gcRemaining-"&vsround(thiscertamount, 2)&",gcDateUsed=" & vsusdate(DateAdd("h",dateadjust,Now()))&" WHERE gcID='"&rs("gcID")&"'")
					giftcertsamount=giftcertsamount + thiscertamount
					rs.movenext
				loop
				rs.close
				totaldiscounts=totaldiscounts + giftcertsamount
				grandtotal=grandtotal - giftcertsamount
				cpnmessage=cpnmessage & xxAppGC & " " & FormatEuroCurrency(giftcertsamount) & IIfVr(cpnmessage<>"", "<br>", "")
				sSQL="UPDATE orders SET ordDiscount="& totaldiscounts & ",ordDiscountText='" & escape_string(cpnmessage) & "' WHERE ordID=" & orderid
				ect_query(sSQL)
			end if
			descstr=""
			addcomma=""
			sSQL="SELECT cartID,cartProdID,cartQuantity,cartProdName FROM cart WHERE cartOrderID="&orderid&" AND cartCompleted=0"
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				if rs("cartProdID")=giftcertificateid then
					sSQL="SELECT gcID FROM giftcertificate WHERE gcCartID="&rs("cartID")
					rs2.open sSQL,cnn,0,1
					if rs2.EOF then
						errormsg="You have a gift certificate added to the cart but unfortunately the certificate was not created correctly. Please remove the certificate from the cart and add it again."
						success=FALSE
					end if
					rs2.close
					ect_query("UPDATE giftcertificate SET gcOrderID="&orderid&" WHERE gcCartID="&rs("cartID"))
				end if
				descstr=descstr&addcomma&rs("cartQuantity")&" "&strip_tags2(rs("cartProdName"))
				addcomma=", "
				rs.movenext
			loop
			rs.close
			descstr=Replace(descstr,"""","")
			if NOT fromshipselector then
				ect_query("DELETE FROM shipoptions WHERE soOrderID="&orderid&" OR soDateAdded<" & vsusdate(Date()-1))
				saveshippingoptions()
			end if
			if SESSION("clientID")="" then
				call setacookie("ectordid",orderid,0)
				call setacookie("ectsessid",thesessionid,0)
				call setacookie("ecthash",sha256(orderid&thesessionid&adminSecret),0)
			end if
		end if
	else
		success=FALSE
	end if
	if stockwarning OR returntocustomerdetails then
		success=FALSE
	elseif success AND ordPayProvider<>"" then
		blockuser=checkuserblock(ordPayProvider)
		if blockuser then
			orderid=0
			thesessionid=0
			xxMstClk=multipurchaseblockmessage
		else
			call getpayprovdetx(ordPayProvider,data1,data2,data3,data4,data5,data6,ppflag1,ppflag2,ppflag3,ppbits,demomode,ppmethod)
		end if
		if wpconfirmpage="" then wpconfirmpage="wpconfirm.asp"
		if nopriceanywhere then grandtotal=0
		if success=FALSE then
			print "<form method=""post"" action=""cart" & extension & """>"
		elseif grandtotal=0 OR ordPayProvider="4" OR ordPayProvider="17" then ' Email
			if grandtotal=0 AND ordPayProvider<>"17" then ordPayProvider="4"
			print "<form method=""post"" name=""checkoutform"" action=""thanks" & extension & """" & IIfVs(recaptchaenabled(256)," onsubmit=""return docheckform(this)""") & ">"
			print whv(IIfVs(grandtotal>0 AND ordPayProvider="17","second")&"emailorder", orderid)
			print whv("thesessionid", thesessionid)
			payprovextraparams=IIfVr(grandtotal>0 AND ordPayProvider="17",payprovextraparams17,payprovextraparams4)
		elseif ordPayProvider="1" then ' PayPal
			if instr(data1,"/")>0 then
				data1arr=split(data1,"/")
				if grandtotal<12 then data1=trim(data1arr(1)) else data1=trim(data1arr(0))
			end if
			if paypalhostedsolution then
				print "<form method=""post"" action=""https://securepayments." & IIfVs(demomode,"sandbox.") & "paypal.com/cgi-bin/acquiringweb"">" & vbCrLf
				print whv("cmd","_hosted-payment")
			else
				print "<form method=""post"" action=""https://www." & IIfVs(demomode,"sandbox.") & "paypal.com/cgi-bin/webscr"">" & vbCrLf
				print whv("cmd","_ext-enter") & whv("redirect_cmd","_xclick") & whv("rm","2")
			end if
			print whv("business",data1) & whv("return",storeurlssl&"thanks"&extension)
			print whv("notify_url",storeurlssl&"vsadmin/ppconfirm.asp") & whv("item_name",left(IIfVr(cartdescription<>"",cartdescription,descstr),127)) & whv("custom",orderid) & whv("invoice",orderid) & whv("no_note","1")
			if paypallc<>"" then print whv("lc",paypallc)
			Session.LCID=1033
			if paypalhostedsolution then
				print whv("subtotal",FormatNumber(grandtotal,getDPs(countryCurrency),-1,0,0))
			elseif splitpaypalshipping then
				print whv("shipping",FormatNumber(vsround((shipping + handling) - freeshipamnt,2),getDPs(countryCurrency),-1,0,0))
				print whv("amount",FormatNumber(vsround((totalgoods + stateTax + countryTax) - totaldiscounts,2),getDPs(countryCurrency),-1,0,0))
			else
				print whv("amount",FormatNumber(grandtotal,getDPs(countryCurrency),-1,0,0))
			end if
			session.LCID=saveLCID
			print whv("currency_code",countryCurrency) & whv("bn","ecommercetemplates_Cart_WPS_US")
			if usefirstlastname then
				print whv("first_name",ordName) & whv("last_name",ordLastName)
				if paypalhostedsolution then print whv("billing_first_name",ordName) & whv("billing_last_name",ordLastName)
			elseif instr(trim(ordName)," ")>0 then
				namearr=split(trim(ordName)," ",2)
				print whv("first_name",namearr(0)) & whv("last_name",namearr(1))
				if paypalhostedsolution then print whv("billing_first_name",namearr(0)) & whv("billing_last_name",namearr(1))
			else
				print whv("last_name",ordName)
				if paypalhostedsolution then print whv("billing_last_name",ordName)
			end if
			if (trim(ordShipName)<>"" OR trim(ordShipLastName)<>"" OR trim(ordShipAddress)<>"") AND paypalhostedsolution then
				print whv("address1",ordShipAddress) & whv("address2",ordShipAddress2) & whv("city",ordShipCity) & whv("state",IIfVr(shipCountryID=1 AND shipStateAbbrev<>"",shipStateAbbrev,ordShipState)) & whv("country",shipCountryCode) & whv("zip",ordShipZip)
			else
				print whv("address1",ordAddress) & whv("address2",ordAddress2) & whv("city",ordCity) & whv("state",IIfVr(countryID=1 AND stateAbbrev<>"",stateAbbrev,ordState)) & whv("country",countryCode) & whv("zip",ordZip)
			end if
			print whv("email",ordEmail)
			if paypalhostedsolution then print whv("billing_address1",ordAddress) & whv("billing_address2",ordAddress2) & whv("billing_city",ordCity) & whv("billing_state",IIfVr(countryID=1 AND stateAbbrev<>"",stateAbbrev,ordState)) & whv("billing_country",countryCode) & whv("buyer_email",ordEmail) & whv("billing_zip",ordZip)
			print whv("cancel_return",storeurlssl&"cart"&extension)
			if countryCode<>"US" AND countryCode<>"CA" then print whv("night_phone_b",ordPhone)
			if ppmethod=1 then print whv("paymentaction","authorization")
			payprovextraparams=payprovextraparams1
		elseif ordPayProvider="2" then ' 2Checkout
			function twocoready(tcoinput)
				twocoready=left(trim(htmlspecials(replace(replace(replace(strip_tags2(tcoinput),vbNewLine,"\n"),"<",""),">",""))), 200)
			end function
			courl="https://" & IIfVr(demomode,"sandbox","www") & ".2checkout.com/checkout/purchase"
			print "<form method=""post"" action="""&courl&""">"
			print whv("merchant_order_id", orderid) & whv("sid", data1) & whv("mode", "2CO")
			print whv("card_holder_name", trim(ordName&" "&ordLastName)) & whv("street_address", ordAddress) & IIfVs(trim(ordAddress2)<>"",whv("ship_street_address2", ordAddress2))
			print whv("city", ordCity) & whv("state", ordState) & whv("zip", ordZip) & whv("country", countryCode) & whv("email", ordEmail) & whv("phone", ordPhone)
			index=0
			sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,"&IIfVr(digidownloads=TRUE,"pDownload,","")&"pDescription FROM cart LEFT JOIN products on cart.cartProdID=products.pID WHERE cartCompleted=0 AND " & getsessionsql()
			rs.open sSQL,cnn,0,1
			do while NOT rs.EOF
				thedesc=twocoready(rs("pDescription"))
				print whv("li_" & index & "_product_id", twocoready(rs("cartProdID")))
				print whv("li_" & index & "_name", twocoready(rs("cartProdName")))
				if thedesc<>"" then print whv("li_" & index & "_description", thedesc)
				print whv("li_" & index & "_price", FormatNumber(rs("cartProdPrice"),2,-1,0,0))
				print whv("li_" & index & "_quantity",rs("cartQuantity"))
				if digidownloads=TRUE then print whv("li_" & index & "_tangible",IIfVr(trim(rs("pDownload")&"")<>"","N","Y")) else print whv("li_" & index & "_tangible","Y")
					sSQL="SELECT coOptGroup,coCartOption,coPriceDiff FROM cartoptions WHERE coCartID=" & rs("cartID")
					rs2.open sSQL,cnn,0,1
					index2=0
					do while NOT rs2.EOF
						print whv("li_" & index & "_option_" & index2 & "_name", twocoready(rs2("coOptGroup")))
						print whv("li_" & index & "_option_" & index2 & "_value", twocoready(rs2("coCartOption")))
						print whv("li_" & index & "_option_" & index2 & "_surcharge", FormatNumber(rs2("coPriceDiff"),2,-1,0,0))
						rs2.movenext
					loop
					rs2.close
				index=index+1
				rs.movenext
			loop
			rs.close
			print whv("li_" & index & "_type", "shipping")
			print whv("li_" & index & "_price", FormatNumber(shipping + handling,2,-1,0,0))
			index=index+1
			if showtaxinclusive<>3 AND stateTax+countryTax>0 then
				print whv("li_" & index & "_type", "tax")
				print whv("li_" & index & "_price", FormatNumber(stateTax+countryTax,2,-1,0,0))
				index=index+1
			end if
			if totaldiscounts + freeshipamnt>0 then
				print whv("li_" & index & "_type", "coupon")
				print whv("li_" & index & "_price", FormatNumber(totaldiscounts + freeshipamnt,2,-1,0,0))
				print whv("li_" & index & "_name", "Discounts")
				index=index+1
			end if
			if trim(ordShipName)<>"" OR trim(ordShipLastName)<>"" OR trim(ordShipAddress)<>"" then
				print whv("ship_name", trim(ordShipName&" "&ordShipLastName)) & whv("ship_street_address", ordShipAddress) & IIfVs(trim(ordShipAddress2)<>"",whv("ship_street_address2", ordShipAddress2)) & whv("ship_city", ordShipCity) & whv("ship_state", ordShipState) & whv("ship_zip", ordShipZip) & whv("ship_country", ordShipCountry)
			else
				print whv("ship_name", trim(ordName&" "&ordLastName)) & whv("ship_street_address", ordAddress) & IIfVs(trim(ordAddress2)<>"",whv("ship_street_address2", ordAddress2)) & whv("ship_city", ordCity) & whv("ship_state", ordState) & whv("ship_zip", ordZip) & whv("ship_country", ordCountry)
			end if
			if demomode then print whv("demo", "Y")
			print whv("purchase_step", "payment-method") & whv("currency_code", countryCurrency) & whv("x_receipt_link_url", storeurl & "thanks" & extension)
			payprovextraparams=payprovextraparams2
		elseif ordPayProvider="3" then ' Authorize.net SIM
			if authnetemulateurl="" then authnetemulateurl="https://secure.authorize.net/gateway/transact.dll"
			if secretword<>"" then
				data1=upsdecode(data1, secretword)
				data2=upsdecode(data2, secretword)
			end if
			if ppflag1 then
				firstname="" : lastname=trim(ordName)
				if usefirstlastname then
					firstname=ordName : lastname=ordLastName
				elseif instr(trim(ordName)," ")>0 then
					namearr=split(trim(ordName)," ",2)
					firstname=namearr(0) : lastname=namearr(1)
				end if
				fingerprint=UCASE(calcHMACSha512(data3,data1 & "^" & orderid & "^" & adminSecret & "^" & data2 & "^","TEXT","TEXT"))
				sjson="{""getHostedPaymentPageRequest"": {" & _
		"""merchantAuthentication"":{""name"":" & json_encode(data1) & ",""transactionKey"":" & json_encode(data2) & "}," & _
		"""transactionRequest"":{" & _
			"""transactionType"":""auth" & IIfVr(ppmethod=1,"Only","Capture") & "Transaction""," & _
			"""amount"":""" & FormatNumber(grandtotal,2,-1,0,0) & """," & _
			IIfVs(NOT demomode,"""solution"":{""id"":""AAA172582""},") & _
			"""order"":{""invoiceNumber"":" & json_encode(orderid) & ",""description"":" & json_encode(left(descstr,255)) & "}," & _
			"""tax"":{""amount"":" & FormatNumber(stateTax+countryTax,2,-1,0,0) & "}," & _
			"""shipping"":{""amount"":" & FormatNumber(shipping + IIfVr(combineshippinghandling,handling,0) - freeshipamnt,2,-1,0,0) & "}," & _
			"""customer"":{""email"":" & json_encode(ordEmail) & "}," & _
			"""billTo"":{" & _
				"""firstName"":" & json_encode(firstname) & ",""lastName"":" & json_encode(lastname) & "," & _
				"""address"":" & json_encode(ordAddress) & ",""city"":" & json_encode(ordCity) & ",""state"":" & json_encode(IIfVr(countryID=1 AND stateAbbrev<>"",stateAbbrev,ordState)) & ",""zip"":" & json_encode(ordZip) & ",""country"":" & json_encode(countryCode3) & ",""phoneNumber"":" & json_encode(ordPhone) & _
			"}"
			if trim(ordShipName)<>"" OR trim(ordShipLastName)<>"" OR trim(ordShipAddress)<>"" then
				firstname="" : lastname=trim(ordShipName)
				if usefirstlastname then
					firstname=ordShipName : lastname=ordShipLastName
				elseif instr(trim(ordShipName)," ")>0 then
					namearr=split(trim(ordShipName)," ",2)
					firstname=namearr(0) : lastname=namearr(1)
				end if
				sjson=sjson&",""shipTo"":{" & _
					"""firstName"":" & json_encode(firstname) & ",""lastName"":" & json_encode(lastname) & "," & _
					"""address"":" & json_encode(ordShipAddress) & ",""city"":" & json_encode(ordShipCity) & ",""state"":" & json_encode(IIfVr(shipCountryID=1 AND shipStateAbbrev<>"",shipStateAbbrev,ordShipState)) & ",""zip"":" & json_encode(ordShipZip) & ",""country"":" & json_encode(shipCountryCode3) & _
				"}"
			end if
		sjson=sjson&"},""hostedPaymentSettings"":{" & _
			"""setting"": [{""settingName"":""hostedPaymentReturnOptions"",""settingValue"":""{\""showReceipt\"":false,\""url\"":\""" & replace(storeurlssl,"/","\/") & "thanks.asp?method=anethosted%26ordid=" & orderid & "%26fp=" & fingerprint & "\"", \""urlText\"": \""" & xxContin & "\"", \""cancelUrl\"": \""" & replace(storeurl,"/","\/") & "cart" & extension & "\"", \""cancelUrlText\"": \""" & xxCancel & "\""}""" & _
			"}, {""settingName"":""hostedPaymentButtonOptions"",""settingValue"":""{\""text\"": \""Pay\""}""" & _
			"}, {""settingName"":""hostedPaymentPaymentOptions"",""settingValue"":""{\""cardCodeRequired\"":true,\""showCreditCard\"":true,\""showBankAccount\"":false}""" & _
			"}, {""settingName"":""hostedPaymentSecurityOptions"",""settingValue"":""{\""captcha\"": false}""" & _
			"}, {""settingName"":""hostedPaymentBillingAddressOptions"",""settingValue"":""{\""show\"": true, \""required\"": true}""" & _
			"}]}}}"
				success=callxmlfunction("https://api" & IIfVs(demomode,"test") & ".authorize.net/xml/v1/request.api",sjson,jres,"","Msxml2.ServerXMLHTTP",vsRESPMSG,FALSE)
				if success then
					if get_json_val(jres,"resultCode","")="Ok" then
						token=get_json_val(jres,"token","")
					else
						aneterr=get_json_val(jres,"text","")
						if aneterr<>"" then
							errormsg=aneterr
						elseif trim(jres)<>"" then
							errormsg=trim(jres)
						end if
						success=FALSE
					end if
				end if
				print "<form method=""post"" action=""https://" & IIfVr(demomode,"test","accept") & ".authorize.net/payment/payment"">"
				print "<input type=""hidden"" id=""popupToken"" name=""token"" value=""" & token & """ />"
			else
				print "<form method=""post"" action="""&authnetemulateurl&""">"
				print whv("x_Version", "3.0") & whv("x_Login", data1) & whv("x_Show_Form", "PAYMENT_FORM")
				if ppmethod=1 then print whv("x_type", "AUTH_ONLY")
				if usefirstlastname then
					print whv("x_first_name", ordName) & whv("x_last_name", ordLastName)
				elseif InStr(trim(ordName)," ")>0 then
					namearr=Split(trim(ordName)," ",2)
					print whv("x_first_name", namearr(0)) & whv("x_last_name", namearr(1))
				else
					print whv("x_last_name", trim(ordName))
				end if
				Randomize
				sequence=int(1000 * Rnd)
				if authnetadjust<>"" then tstamp=GetSecondsSince1970() + authnetadjust else tstamp=GetSecondsSince1970()
				if len(data3)>100 then
					fingerprint=UCASE(calcHMACSha512(data3,data1 & "^" & sequence & "^" & tstamp & "^" & FormatNumber(grandtotal,2,-1,0,0) & "^","TEXT","HEX"))
				else
					fingerprint=HMAC(data2,data1 & "^" & sequence & "^" & tstamp & "^" & FormatNumber(grandtotal,2,-1,0,0) & "^")
				end if
				print whv("x_fp_sequence", sequence) & whv("x_fp_timestamp", tstamp) & whv("x_fp_hash", fingerprint)
				print whv("x_address", ordAddress& IIfVr(trim(ordAddress2)<>"",", " & ordAddress2, "")) & whv("x_city", ordCity) & whv("x_country", ordCountry)
				print whv("x_phone", ordPhone) & whv("x_state", ordState) & whv("x_zip", ordZip)
				print whv("x_invoice_num", orderid) & whv("x_email", ordEmail) & whv("x_description", left(descstr,255))
				if SESSION("clientID")<>"" then print whv("x_cust_id", SESSION("clientID"))
				if trim(ordShipName)<>"" OR trim(ordShipLastName)<>"" OR trim(ordShipAddress)<>"" then
					if usefirstlastname then
						print whv("x_ship_to_first_name", ordShipName) & whv("x_ship_to_last_name", ordShipLastName)
					elseif InStr(trim(ordShipName)," ")>0 then
						namearr=Split(trim(ordShipName)," ",2)
						print whv("x_ship_to_first_name", namearr(0)) & whv("x_ship_to_last_name", namearr(1))
					else
						print whv("x_ship_to_last_name", trim(ordShipName))
					end if
					print whv("x_ship_to_address", ordShipAddress& IIfVr(trim(ordShipAddress2)<>"",", " & ordShipAddress2, "")) & whv("x_ship_to_city", ordShipCity) & whv("x_ship_to_country", ordShipCountry) & whv("x_ship_to_state", ordShipState) & whv("x_ship_to_zip", ordShipZip)
				end if
				print whv("x_Amount",FormatNumber(grandtotal,2,-1,0,0)) & whv("x_Relay_Response","TRUE") & whv("x_Relay_URL",storeurl&"vsadmin/"&wpconfirmpage) & whv("x_solution_id","AAA172582")
				if demomode then print whv("x_Test_Request", "TRUE")
				payprovextraparams=payprovextraparams3
			end if
		elseif ordPayProvider="5" then ' WorldPay
			print "<form method=""post"" action=""https://secure" & IIfVr(demomode, "-test", "") & ".worldpay.com/wcc/purchase"">"
			print whv("instId", data1) & whv("cartId", orderid)
			Session.LCID=1033
			print whv("amount", FormatNumber(grandtotal,2,-1,0,0))
			Session.LCID=saveLCID
			print whv("currency", countryCurrency)
			print whv("desc", Left(descstr,255))
			print whv("name", trim(ordName&" "&ordLastName)) & whv("address", ordAddress & IIfVr(trim(ordAddress2)<>"",", " & ordAddress2, "") & vbCrLf & ordCity & vbCrLf & ordState) & whv("postcode", ordZip) & whv("country", countryCode) & whv("tel", ordPhone) & whv("email", ordEmail)
			print whv("authMode", IIfVr(ppmethod=1, "E", "A")) & whv("testMode", IIfVr(demomode, "100", "0"))
			data2arr=split(data2,"&",2)
			if UBOUND(data2arr)>=0 then data2=data2arr(0)
			if data2<>"" then
				sigfields="amount:currency:cartId:testMode"
				print whv("signatureFields", sigfields)
				Session.LCID=1033
				print whv("signature", calcmd5(data2&";"&sigfields&";"&FormatNumber(grandtotal,2,-1,0,0)&";"&countryCurrency&";"&orderid&";"&IIfVr(demomode,"100","0")))
				Session.LCID=saveLCID
			end if
			payprovextraparams=payprovextraparams5
		elseif ordPayProvider="6" then ' NOCHEX
			print "<form method=""post"" action=""https://secure.nochex.com/"">"
			print whv("merchant_id", data1)
			print whv("success_url", storeurl & "thanks" & extension & "?ncretval="&orderid&"&ncsessid="&thesessionid) & whv("callback_url", storeurl&"vsadmin/ncconfirm.asp")
			print whv("description", left(descstr,255))
			print whv("order_id", orderid) & whv("amount", FormatNumber(grandtotal,2,-1,0,0))
			print whv("billing_fullname", trim(ordName&" "&ordLastName)) & whv("billing_address", ordAddress& IIfVr(trim(ordAddress2)<>"",", " & ordAddress2, "")) & whv("billing_postcode", ordZip) & whv("email_address", ordEmail) & whv("customer_phone_number", ordPhone)
			if trim(ordShipName)<>"" OR trim(ordShipAddress)<>"" then
				print whv("delivery_fullname", trim(ordShipName&" "&ordShipLastName)) & whv("delivery_address", ordShipAddress& IIfVr(trim(ordShipAddress2)<>"",", " & ordShipAddress2, "")) & whv("delivery_postcode", ordShipZip)
			end if
			if demomode then print whv("test_transaction", "100")
			payprovextraparams=payprovextraparams6
		elseif ordPayProvider="7" then ' Payflow Pro
			print "<form method=""post"" action=""cart" & extension & """ onsubmit=""return isvalidcard(this)"">"
			print whv("mode", "authorize") & whv("method", "7") & whv("ordernumber", orderid)
			payprovextraparams=payprovextraparams7
		elseif ordPayProvider="8" then ' Payflow Link
			if instr(data1,"&")>0 then
				print "<form method=""post"" action=""cart" & extension & """ onsubmit=""return isvalidcard(this)"">"
				print whv("mode", "authorize") & whv("method", "8") & whv("ordernumber", orderid)
			else
				paymentlink="https://payflowlink.paypal.com"
				print "<form method=""post"" action="""&paymentlink&""">"
				print whv("LOGIN", data1) & whv("PARTNER", data2) & whv("CUSTID", orderid)
				print whv("AMOUNT", FormatNumber(grandtotal,2,-1,0,0))
				print whv("TYPE", IIfVr(ppmethod=1,"A","S"))
				print whv("DESCRIPTION", Left(descstr,255))
				print whv("NAME", trim(ordName&" "&ordLastName)) & whv("ADDRESS", ordAddress& IIfVr(trim(ordAddress2)<>"",", " & ordAddress2, "")) & whv("CITY", ordCity) & whv("STATE", ordState) & whv("ZIP", ordZip) & whv("COUNTRY", IIfVr(countryCode="US", "USA", ordCountry))
				print whv("EMAIL", ordEmail) & whv("PHONE", ordPhone)
				print whv("METHOD", "CC") & whv("ORDERFORM", "TRUE") & whv("SHOWCONFIRM", "FALSE") & whv("BUTTONSOURCE", "EcommerceTemplatesUS_Cart_PPA")
				if trim(ordShipName)<>"" OR trim(ordShipAddress)<>"" then
					print whv("NAMETOSHIP", trim(ordShipName&" "&ordShipLastName)) & whv("ADDRESSTOSHIP", ordShipAddress& IIfVr(trim(ordShipAddress2)<>"",", " & ordShipAddress2, "")) & whv("CITYTOSHIP", ordShipCity) & whv("STATETOSHIP", ordShipState) & whv("ZIPTOSHIP", ordShipZip) & whv("COUNTRYTOSHIP", IIfVr(shipCountryCode="US", "USA", ordShipCountry))
				end if
			end if
			payprovextraparams=payprovextraparams8
		elseif ordPayProvider="9" then ' PayPoint.net
			print "<form method=""post"" action=""https://www.secpay.com/java-bin/ValCard"">"
			print whv("merchant", data1) & whv("trans_id", orderid)
			print whv("amount", FormatNumber(grandtotal,2,-1,0,0))
			print whv("callback", storeurl&"vsadmin/"&wpconfirmpage) & whv("currency", countryCurrency) & whv("cb_post", "true")
			print whv("bill_name", trim(ordName&" "&ordLastName)) & whv("bill_addr_1", ordAddress) & whv("bill_addr_2", ordAddress2) & whv("bill_city", ordCity) & whv("bill_state", ordState) & whv("bill_post_code", ordZip) & whv("bill_country", ordCountry)
			print whv("bill_email", ordEmail) & whv("bill_tel", ordPhone)
			if trim(ordShipName)<>"" OR trim(ordShipLastName)<>"" OR trim(ordShipAddress)<>"" then
				print whv("ship_name", trim(ordShipName&" "&ordShipLastName)) & whv("ship_addr_1", ordShipAddress) & whv("ship_addr_2", ordShipAddress2) & whv("ship_city", ordShipCity) & whv("ship_state", ordShipState) & whv("ship_post_code", ordShipZip) & whv("ship_country", ordShipCountry)
			end if
			data2arr=split(data2,"&",2)
			if UBOUND(data2arr)>=0 then data2md5=data2arr(0)
			if UBOUND(data2arr)>0 then data2tpl=data2arr(1)
			if trim(data2md5)<>"" then
				Session.LCID=1033
				print whv("digest", calcmd5(orderid & FormatNumber(grandtotal,2,-1,0,0) & data2md5)) & whv("md_flds", "trans_id:amount:callback")
				Session.LCID=saveLCID
			end if
			print whv("mpi_description", left(descstr,125))
			if trim(data2tpl)<>"" then print whv("template", urldecode(data2tpl))
			if ppmethod=1 then print whv("deferred", "reuse:5:5")
			print whv("req_cv2", "true")
			if data3="1" then print whv("ssl_cb", "true")
			if demomode then print whv("options", "test_status=true,dups=false")
			payprovextraparams=payprovextraparams9
		elseif ordPayProvider="10" then ' Capture Card
			print "DISABLED!!<br>"
		elseif ordPayProvider="11" OR ordPayProvider="12" then ' PSiGate
			print "<form method=""post"" action=""https://"&IIfVr(demomode,"staging","checkout")&".psigate.com/HTMLPost/HTMLMessenger""" & IIfVr(ordPayProvider="12", " onsubmit=""return isvalidcard(this)""", "") & ">"
			print whv("MerchantID", data1) & whv("Oid", orderid)
			Session.LCID=1033
			print whv("FullTotal", FormatNumber(grandtotal,2,-1,0,0))
			Session.LCID=saveLCID
			print whv("ThanksURL", storeurl&"thanks"&extension) & whv("NoThanksURL", storeurl&"thanks"&extension) & whv("CustomerRefNo", left(calcmd5(orderid&":"&secretword),24)) & whv("ChargeType", IIfVr(ppmethod=1, "1", "0"))
			if ordPayProvider="11" then print whv("Bname", trim(ordName&" "&ordLastName))
			print whv("Baddr1", ordAddress) & whv("Baddr2", ordAddress2) & whv("Bcity", ordCity)
			print whv("IP", REMOTE_ADDR)
			print whv("Bstate", IIfVr(countryID=1 AND stateAbbrev<>"", stateAbbrev, ordState)) & whv("Bzip", ordZip) & whv("Bcountry", countryCode)
			print whv("Email", ordEmail) & whv("Phone", ordPhone)
			if trim(ordShipName)<>"" OR trim(ordShipLastName)<>"" OR trim(ordShipAddress)<>"" then
			print whv("Sname", trim(ordShipName&" "&ordShipLastName)) & whv("Saddr1", ordShipAddress) & whv("Saddr2", ordShipAddress2) & whv("Scity", ordShipCity) & whv("Sstate", ordShipState) & whv("Szip", ordShipZip) & whv("Scountry", ordShipCountry)
			end if
			payprovextraparams=payprovextraparams11
			if ordPayProvider="12" then payprovextraparams=payprovextraparams12
		elseif ordPayProvider="13" then ' Authorize.net AIM
			print "<form method=""post"" action=""cart" & extension & """ onsubmit=""return isvalidcard(this)"">"
			print whv("mode", "authorize") & whv("method", "13") & whv("ordernumber", orderid) & whv("description", Left(descstr,254))
			payprovextraparams=payprovextraparams13
		elseif ordPayProvider="14" then ' Custom Pay Provider
%>
<!--#include file="customppsend.asp"-->
<%			payprovextraparams=payprovextraparams14
		elseif ordPayProvider="15" then ' Netbanx
			randomize
			sequence=int(1000000 * Rnd) + 1000000
			print "<form method=""post"" action=""https://pay.netbanx.com/"&data1&""">"
			print whv("nbx_merchant_reference", orderid&"."&sequence) & whv("nbx_payment_amount", int(grandtotal*100)) & whv("nbx_currency_code", countryCurrency) & whv("nbx_cardholder_name", trim(ordName&" "&ordLastName)) & whv("nbx_email", ordEmail) & whv("nbx_postcode", ordZip)
			print whv("nbx_return_url", storeurl&"categories"&extension)
			print whv("nbx_success_url", storeurl&"vsadmin/ncconfirm.asp")
			if trim(data2)<>"" then print whv("nbx_checksum", hex_sha1(int(grandtotal*100)&countryCurrency&orderid&"."&sequence&data2))
			payprovextraparams=payprovextraparams15
		elseif ordPayProvider="16" then ' Linkpoint
			lpsubtotal=vsround(totalgoods - totaldiscounts, 2)
			lpshipping=vsround((shipping + handling) - freeshipamnt, 2)
			lptax=vsround(stateTax + countryTax, 2)
			randomize
			sequence="."&int(1000000 * rnd) + 1000000
			if data3<>"" then
				payurl="https://connect."&IIfVr(demomode,"merchanttest.","")&"firstdataglobalgateway.com/IPGConnect/gateway/processing"
			else
				payurl="https://www."&IIfVr(demomode,"staging.","")&"linkpointcentral.com/lpc/servlet/lppay"
			end if
			print "<form action=""" & payurl & """ method=""post"""&IIfVr(data2="1"," onsubmit=""return isvalidcard(this)""","")&">"
			print whv("storename", data1) & whv("mode", "payonly") & whv("ponumber", orderid) & whv("oid", orderid&sequence) & whv("responseURL", storeurl&"thanks"&extension)
			print whv("subtotal", FormatNumber(lpsubtotal,2,-1,0,0)) & whv("chargetotal", FormatNumber(lpsubtotal+lpshipping+lptax,2,-1,0,0)) & whv("shipping", FormatNumber(lpshipping,2,-1,0,0)) & whv("tax", FormatNumber(lptax,2,-1,0,0))
			if data2<>"1" then print whv("bname", trim(ordName&" "&ordLastName))
			print whv("baddr1", ordAddress) & whv("baddr2", ordAddress2) & whv("bcity", ordCity)
			if countryID=1 AND stateAbbrev<>"" then print whv("bstate", stateAbbrev) else print whv("bstate2", ordState)
			print whv("bzip", ordZip) & whv("bcountry", countryCode) & whv("email", ordEmail) & whv("phone", ordPhone) & whv("txntype", IIfVr(ppmethod=1, "preauth", "sale"))
			if trim(ordShipName)<>"" OR trim(ordShipLastName)<>"" OR trim(ordShipAddress)<>"" then
				print whv("sname", trim(ordShipName&" "&ordShipLastName)) & whv("saddr1", ordShipAddress) & whv("saddr2", ordShipAddress2) & whv("scity", ordShipCity) & whv("sstate", ordShipState) & whv("szip", ordShipZip) & whv("scountry", shipCountryCode)
			end if
			if data3<>"" then
				formattedDate=replace(replace(replace(getutcdate(0),"-",":"),"T","-"),"Z","")
				str=data1 & formattedDate & FormatNumber(lpsubtotal+lpshipping+lptax,2,-1,0,0) & data3
				hex_str=""
				for ilp=1 to len(str)
					hex_str=hex_str+lcase(cstr(hex(asc(mid(str,ilp,1)))))
				next
				print whv("txndatetime" ,formattedDate)
				ect_query("UPDATE orders SET ordPrivateStatus='" & escape_string(formattedDate) & "' WHERE ordID=" & orderid)
				print whv("timezone", "UTC")
				print whv("hash", SHA256(hex_str))
			end if
			payprovextraparams=payprovextraparams16
		elseif ordPayProvider="18" then ' PayPal Payment Pro
			print "<form method=""post"" action=""cart" & extension & """ onsubmit=""return isvalidcard(this)"">"
			print whv("mode", "authorize") & whv("method", "18") & whv("ordernumber", orderid) & whv("description", left(descstr,254))
			payprovextraparams=payprovextraparams18
		elseif ordPayProvider="19" then ' PayPal Express Payment
			print "<form method=""post"" action=""thanks" & extension & """ onsubmit=""return docheckform(this)"">"
			print whv("token", token) & whv("method", "paypalexpress") & whv("ordernumber", orderid) & whv("payerid", payerid) & whv("email", ordEmail)
			payprovextraparams=payprovextraparams19
		elseif ordPayProvider="21" then ' Amazon Pay
			print "<form method=""post"" action=""thanks" & extension & """ onsubmit=""return docheckform(this)"">"
			print whv("pprov", "21") & whv("ordernumber", orderid) & whv("amzrefid", amzrefid_)
			sSQL="UPDATE orders SET ordTransID='" & escape_string(amzrefid_) & "' WHERE ordID=" & escape_string(orderid) & " AND ordSessionID='"&escape_string(thesessionid)&"'"
			if amazonpayment then cnn.execute(sSQL)
			payprovextraparams=payprovextraparams21
		elseif ordPayProvider="22" then ' PayPal Advanced
			print "<form method=""post"" action=""cart" & extension & """>"
			print whv("mode", "authorize") & whv("method", "22") & whv("ordernumber", orderid) & whv("sessionid", thesessionid)
			payprovextraparams=payprovextraparams22
		elseif ordPayProvider="23" then ' Stripe.com
			if ppflag1=0 then
				print "<form method=""post"" action=""thanks" & extension & """>"
				print whv("pprov", "23") & whv("ordernumber", orderid)
				payprovextraparams=payprovextraparams23
			else
				sXML=""
				if (ppbits AND 1)=1 OR ppbits=0 then sXML=sXML&"payment_method_types[]=card"
				if (ppbits AND 2)=2 then sXML=sXML&IIfVs(sXML<>"","&")&"payment_method_types[]=ideal"
				if (ppbits AND 4)=4 then sXML=sXML&IIfVs(sXML<>"","&")&"payment_method_types[]=bancontact"
				if (ppbits AND 8)=8 then sXML=sXML&IIfVs(sXML<>"","&")&"payment_method_types[]=giropay"
				if (ppbits AND 16)=16 then sXML=sXML&IIfVs(sXML<>"","&")&"payment_method_types[]=p24"
				if (ppbits AND 32)=32 then sXML=sXML&IIfVs(sXML<>"","&")&"payment_method_types[]=eps"
				if (ppbits AND 64)=64 then sXML=sXML&IIfVs(sXML<>"","&")&"payment_method_types[]=fpx"
				if (ppbits AND 128)=128 then sXML=sXML&IIfVs(sXML<>"","&")&"payment_method_types[]=bacs_debit"
				if (ppbits AND 256)=256 then sXML=sXML&IIfVs(sXML<>"","&")&"payment_method_types[]=klarna"
				sXML=sXML&"&line_items[0][price_data][product_data][name]=Cart Items Total" & _
					"&line_items[0][price_data][unit_amount]=" & vsround(grandtotal*100,0) & _
					"&line_items[0][price_data][currency]=" & lcase(countryCurrency) & _
					"&line_items[0][quantity]=1"
					
				sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
				rs.open sSQL,cnn,0,1
				itemno=1
				tempXML=""
				do while NOT rs.EOF
					sSQL="SELECT coOptGroup,coCartOption FROM cartoptions WHERE coCartID=" & rs("cartID")
					rs2.open sSQL,cnn,0,1
					optionnames=""
					do while NOT rs2.EOF
						optionnames=optionnames & " " & rs2("coCartOption") & ","
						rs2.movenext
					loop
					rs2.close
					if optionnames<>"" then optionnames=left(trim(optionnames),len(trim(optionnames))-1)
					tempXML=tempXML&"&line_items[" & itemno & "][price_data][product_data][name]=" & urlencode(left(rs("cartProdName") & IIfVs(optionnames<>""," ("&optionnames&")"), 255)) & _
						"&line_items[" & itemno & "][price_data][unit_amount]=0" & _
						"&line_items[" & itemno & "][price_data][currency]=" & lcase(countryCurrency) & _
						"&line_items[" & itemno & "][quantity]=" & rs("cartQuantity")
					sSQL="SELECT imageSrc FROM productimages WHERE imageProduct='" & escape_string(rs("cartProdID")) & "' AND imageSrc NOT LIKE '% %' ORDER BY imageType,imageNumber"
					rs2.open sSQL,cnn,0,1
					optionnames=""
					if NOT rs2.EOF then
						tempXML=tempXML&"&line_items[" & itemno & "][price_data][product_data][images][]=" & urlencode(IIfVs(instr(rs2("imageSrc"),"//")=0, storeurl) & rs2("imageSrc"))
					end if
					rs2.close
					itemno=itemno+1
					if itemno>32 then tempXML="" : exit do
					rs.movenext
				loop
				rs.close
				sXML=sXML&tempXML&"&payment_intent_data[capture_method]=" & IIfVr(ppmethod=1,"manual","automatic") & _
					"&payment_intent_data[metadata[order_id]]=" & orderid & _
					"&billing_address_collection=auto" & _
					"&customer_email=" & urlencode(ordEmail) & _
					"&mode=payment" & _
					"&success_url=" & storeurlssl & urlencode("thanks" & extension & "?method=stripe&sid={CHECKOUT_SESSION_ID}&soid=" & orderid) & _
					"&cancel_url=" & storeurlssl & "cart" & extension
				xmlfnheaders=array(array("User-Agent","Stripe/v1 RubyBindings/1.12.0"),array("Authorization","Bearer "&data1),array("Content-Type","application/x-www-form-urlencoded"))
				iscanadapost=TRUE
				if callxmlfunction("https://api.stripe.com/v1/checkout/sessions",sXML,xmlres,"","Msxml2.ServerXMLHTTP", errtext, FALSE) then
					if instr(xmlres,"""error"":") > 0 then
						success=FALSE
						errormsg=get_json_val(xmlres,"message","")
					else
						stripeid=get_json_val(xmlres,"id","")
						paymentintent=get_json_val(xmlres,"payment_intent","")
						SESSION("stripeid")=stripeid
						sSQL="UPDATE orders SET ordTransID='" & escape_string(paymentintent) & "' WHERE ordID=" & orderid
						ect_query(sSQL)
					end if
				else
					success=FALSE
				end if
			end if
		elseif ordPayProvider="24" then ' SagePay
			function spchkalphanum(ectinstr)
				Set toregexp=new RegExp
				toregexp.pattern="[\W_]"
				toregexp.ignorecase=TRUE
				toregexp.global=TRUE
				spchkalphanum=toregexp.replace(ectinstr, "")
				Set toregexp=nothing
			end function
			if trim(ordShipAddress&"")="" then
				ordShipName=ordName
				ordShipLastName=ordLastName
				ordShipAddress=ordAddress
				ordShipAddress2=ordAddress2
				ordShipCity=ordCity
				ordShipState=ordState
				ordShipZip=ordZip
				ordShipCountry=ordCountry
				ordShipPhone=ordPhone
				ordShipExtra1=ordExtra1
				ordShipExtra2=ordExtra2
				shipStateAbbrev=stateAbbrev
			end if
			Randomize
			upperbound="9999999"
			lowerbound="1000000"
			sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
			rs.open sSQL,cnn,0,1
			index=0
			thecart=""
			do while NOT rs.EOF
				sSQL="SELECT SUM(coPriceDiff) as totpricediff FROM cartoptions WHERE coCartID=" & rs("cartID")
				rs2.open sSQL,cnn,0,1
				if NOT isnull(rs2("totpricediff")) then totpricediff=rs2("totpricediff") else totpricediff=0
				rs2.close
				thecart=thecart & ":[" & replace(replace(rs("cartProdID"),":",""),"&","") & "]" & strip_tags2(replace(replace(rs("cartProdName"),":",""),"&","")) & ":" & rs("cartQuantity") & ":" & formatnumber(rs("cartProdPrice")+totpricediff,2,-1,0,0) & ":0.00:" & formatnumber(rs("cartProdPrice")+totpricediff,2,-1,0,0) & ":" & formatnumber((rs("cartProdPrice")+totpricediff)*rs("cartQuantity"),2,-1,0,0)
				index=index+1
				rs.MoveNext
			loop
			rs.close
			if stateTax + countryTax > 0 then
				thecart=thecart & ":Taxes:---:---:---:---:" & formatnumber(stateTax + countryTax,2,-1,0,0)
				index=index + 1
			end if
			if totaldiscounts > 0 then
				thecart=thecart & ":Discounts:---:---:---:---:" & formatnumber(0 - totaldiscounts,2,-1,0,0)
				index=index + 1
			end if
			if (shipping + handling) - freeshipamnt > 0 then
				thecart=thecart & ":Delivery:---:---:---:---:" & formatnumber((shipping + handling) - freeshipamnt,2,-1,0,0)
				index=index + 1
			end if
			thecart=replace(thecart,"&pound;","£")
			thecart=index&thecart
			session.LCID=1033
			spzipcode=trim(ordZip&"")
			if spzipcode="" then spzipcode="NA"
			spshipzipcode=trim(ordShipZip&"")
			if spshipzipcode="" then spshipzipcode="NA"
			stuff="VendorTxCode=" & orderid & "-" & Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
			stuff=stuff & "&Amount=" & FormatNumber(grandtotal,2,-1,0,0) & "&Currency=" & countryCurrency & "&Description=" & left(descstr,99)
			stuff=stuff & "&Basket=" & thecart & "&SuccessURL=" & storeurl & "thanks" & extension & "&FailureURL=" & storeurl & "thanks" & extension
			stuff=stuff & IIfVs(NOT nosagepayemail,"&CustomerEMail=" & ordEmail) & "&CustomerName=" & trim(ordName&" "&ordLastName)
			stuff=stuff & "&ContactNumber=" & spchkalphanum(ordPhone) & IIfVs(NOT nosagepayemail,"&VendorEMail=" & emailAddr)
			if usefirstlastname then
				stuff=stuff & "&BillingFirstnames=" & ordName & "&BillingSurname=" & ordLastName
			elseif InStr(trim(ordName)," ") > 0 then
				namearr=Split(trim(ordName)," ",2)
				stuff=stuff & "&BillingFirstnames=" & namearr(0) & "&BillingSurname=" & namearr(1)
			else
				stuff=stuff & "&BillingFirstnames=&BillingSurname=" & ordName
			end if
			stuff=stuff & "&BillingAddress1=" & ordAddress
			if trim(ordAddress2)<>"" then stuff=stuff & "&BillingAddress2=" & ordAddress2
			stuff=stuff & "&BillingCity=" & ordCity
			if countryID=1 then
				spstatefield=ucase(IIfVr(stateAbbrev<>"",stateAbbrev,ordState))
				if len(spstatefield)<>2 then
					success=FALSE
					errormsg=twoletterstateerr & "<br><div style=""text-align:center""><input type=""button"" class=""ectbutton"" onclick=""history.go(-1)"" value="""&xxGoBack&""" /></div>"
				end if
				stuff=stuff & "&BillingState=" & spstatefield
			end if
			stuff=stuff & "&BillingCountry=" & countryCode & "&BillingPostCode=" & spzipcode & "&BillingPhone=" & spchkalphanum(ordPhone)
			if usefirstlastname AND ordShipLastName<>"" then
				stuff=stuff & "&DeliveryFirstnames=" & ordShipName & "&DeliverySurname=" & ordShipLastName
			elseif InStr(trim(ordShipName)," ") > 0 then
				namearr=Split(trim(ordShipName)," ",2)
				stuff=stuff & "&DeliveryFirstnames=" & namearr(0) & "&DeliverySurname=" & namearr(1)
			else
				stuff=stuff & "&DeliveryFirstnames=" & ordShipName & "&DeliverySurname=" & ordShipName
			end if
			stuff=stuff & "&DeliveryAddress1=" & ordShipAddress
			if trim(ordShipAddress2)<>"" then stuff=stuff & "&DeliveryAddress2=" & ordShipAddress2
			stuff=stuff & "&DeliveryCity=" & ordShipCity
			if shipCountryID=1 then
				spstatefield=ucase(IIfVr(shipStateAbbrev<>"",shipStateAbbrev,ordShipState))
				if len(spstatefield)<>2 then
					success=FALSE
					errormsg=twoletterstateerr & "<br><div style=""text-align:center""><input type=""button"" class=""ectbutton"" onclick=""history.go(-1)"" value="""&xxGoBack&""" /></div>"
				end if
				stuff=stuff & "&DeliveryState=" & spstatefield
			end if
			stuff=stuff & "&DeliveryCountry=" & shipCountryCode & "&DeliveryPostCode=" & spshipzipcode & "&DeliveryPhone=" & spchkalphanum(ordShipPhone)
			stuff=stuff & "&ReferrerID=7B0AD331-0388-44EA-BE3A-D05D3FB9FE28"
			crypt="@" & AESEncrypt(stuff,data2)
			print "<form method=""post"" action=""https://" & IIfVr(demomode,"test","live") & ".sagepay.com/gateway/service/vspform-register.vsp"">"
			print whv("VPSProtocol","3.00") & whv("TxType",IIfVr(ppmethod=1,"DEFERRED","PAYMENT")) & whv("Vendor",data1) & whv("Crypt",crypt)
			session.LCID=saveLCID
			payprovextraparams=payprovextraparams24
		elseif ordPayProvider="28" then ' SquareUp
			print "<form method=""post"" id=""ectcheckoutform"" action=""thanks" & extension & """>"
			print whv("mode", "authorize") & whv("method",28) & whv("ordernumber", orderid) & whv("sessionid", thesessionid)
			call writehiddenidvar("payment_method_nonce","")
			call writehiddenidvar("txnid","")
		elseif ordPayProvider="29" then ' NMI
			xmlfnheaders=array(array("Content-Type","text/xml"))
			sXML="<?xml version=""1.0"" encoding=""UTF-8""?><sale>" & vrxmltag("api-key",data1)
			sXML=sXML&vrxmltag("redirect-url",storeurlssl & "thanks" & extension & "?mode=authorize&method=29&ordernumber=" & orderid)
			sXML=sXML&vrxmltag("amount",grandtotal) & vrxmltag("ip-address",REMOTE_ADDR) & vrxmltag("currency",countryCurrency)
			sXML=sXML&vrxmltag("order-id",orderid)
			sXML=sXML&vrxmltag("order-description",left(descstr, 99))
			sXML=sXML&vrxmltag("tax-amount",formatnumber(stateTax+countryTax,2,-1,0,0))
			sXML=sXML&vrxmltag("shipping-amount",formatnumber((shipping + handling) - freeshipamnt,2,-1,0,0))
			sXML=sXML&"<billing>"
				call splitname(trim(ordName & " " & ordLastName), firstname, lastname)
				sXML=sXML&vrxmltag("first-name",firstname) & vrxmltag("last-name",lastname)
				sXML=sXML&vrxmltag("address1",ordAddress) & vrxmltag("address2",ordAddress2) & vrxmltag("city",ordCity) & vrxmltag("state",ordState) & vrxmltag("postal",ordZip) & vrxmltag("country",countryCode)
				sXML=sXML&vrxmltag("email",ordEmail) & vrxmltag("phone",ordPhone)
			sXML=sXML&"</billing></sale>"
			if callxmlfunction("https://secure.networkmerchants.com/api/v2/three-step",sXML,xmlres,"","Msxml2.ServerXMLHTTP",errormsg,FALSE) then
				set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
				xmlDoc.validateOnParse=FALSE
				xmlDoc.loadXML(xmlres)
				set nmiresultobj=xmlDoc.getElementsByTagName("result")
				resultcode=nmiresultobj.item(0).firstChild.nodeValue
				if resultcode="1" then
					set nmiformobj=xmlDoc.getElementsByTagName("form-url")
					print "<form method=""post"" action=""" & nmiformobj.item(0).firstChild.nodeValue & """ onsubmit=""return isvalidcard(this)"">"
				else
					set nmiformobj=xmlDoc.getElementsByTagName("result-text")
					success=FALSE
					errormsg=nmiformobj.item(0).firstChild.nodeValue
				end if
			else
				success=FALSE
			end if
		elseif ordPayProvider="30" then // eWay
			xmlfnheaders=array(array("Content-Type","application/json"),array("Authorization", "Basic " & vrbase64_encrypt(data1 & ":" & data2)))
			call splitname(trim(ordName&" "&ordLastName), firstname, lastname)
			sjson="{""Customer"":{" & _ 
					"""FirstName"":" & json_encode(firstname) & ",""LastName"":" & json_encode(lastname) & "," & _
					"""Street1"":" & json_encode(ordAddress) & ",""Street2"":" & json_encode(ordAddress2) & "," & _
					"""City"":" & json_encode(ordCity) & ",""State"":" & json_encode(ordState) & ",""PostalCode"":" & json_encode(ordZip) & ",""Country"":" & json_encode(lcase(countryCode)) & "," & _
					"""Phone"":" & json_encode(ordPhone) & ",""Email"":" & json_encode(ordEmail) & "," & _
				"},""Payment"":{" & _
					"""TotalAmount"":" & (grandtotal*100) & ",""InvoiceReference"":""" & orderid & """,""CurrencyCode"":""" & countryCurrency & """" & _
				"}," & _
				"""RedirectUrl"":""" & storeurlssl & "thanks" & extension & """,""Method"":""ProcessPayment"",""TransactionType"":""Purchase""}"
			if callxmlfunction("https://api." & IIfVr(demomode,"sandbox.","") & "ewaypayments.com/AccessCodes",sjson,jres,"","Msxml2.ServerXMLHTTP",errormsg,TRUE) then
				accesscode=get_json_val(jres,"AccessCode","")
				formactionurl=get_json_val(jres,"FormActionURL","")
				formerrors=get_json_val(jres,"Errors","")
%>				<form method="POST" action="<%=formactionurl%>" id="payment_form" onsubmit="return isvalidcard(this)">
					<input type="hidden" name="EWAY_ACCESSCODE" value="<%=accesscode%>" />
					<input type="hidden" name="EWAY_PAYMENTTYPE" value="Credit Card" />
<%				if formerrors<>"" then
					success=FALSE
					errormsg=formerrors
				end if
			else
				success=FALSE
			end if
		elseif ordPayProvider="31" then ' Pay360
%>
<div id="ppcoverdiv" class="ectopaque" style="display:none"><img src="images/preloader.gif" alt="" style="margin-top:250px"></div>
<script>
var p3ajaxobj;
function ppajaxcallback(){
	if(p3ajaxobj.readyState==4){
		var restxt=p3ajaxobj.responseText;
		if(restxt.substr(0,1)=='1')
			document.location=restxt.substr(2);
		else
			alert(restxt.substr(2));
	}
}
function payprovscript(){
	document.getElementById('ppcoverdiv').style.display='';
	p3ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	p3ajaxobj.onreadystatechange=ppajaxcallback;
	p3ajaxobj.open("GET", "vsadmin/ajaxservice.asp?action=pay360&orderid=<%=orderid%>", true);
	p3ajaxobj.send(null);
}
</script>
<%		elseif ordPayProvider="32" then ' Global Payments
			td=now()
			timestamp=datepart("yyyy",td) & check2d(datepart("m",td)) & check2d(datepart("d",td)) & check2d(datepart("h",td)) & check2d(datepart("n",td)) & check2d(datepart("s",td))
			grandtotalrnd=vsround(grandtotal*100,0)
			randomize
			gporderid=orderid&"-"&(int(10000 * rnd) + 10000)
			hashstring=timestamp&"."&data1&"."&gporderid&"."&grandtotalrnd&"."&countryCurrency
			hashstring=hex_sha1(hex_sha1(hashstring)&"."&data2)
%>
<script>
function payprovscript(){
	document.getElementById('gpiform').submit();
	document.getElementById('gpicoverdiv').style.display='';
}
</script>
<div id="gpicoverdiv" class="ectopaque" style="display:none">
	<iframe id="gpiframe" name="gpiframe" style="margin-top:100px;width:600px;height:580px;border:none"></iframe>
</div>
			<form id="gpiform" action="https://pay.<%=IIfVs(demomode,"sandbox.")%>realexpayments.com/pay" method="POST" target="gpiframe">
			<input type="hidden" name="TIMESTAMP" value="<%=timestamp%>">
			<input type="hidden" name="MERCHANT_ID" value="<%=data1%>">
			<input type="hidden" name="ORDER_ID" value="<%=gporderid%>">
			<input type="hidden" name="AMOUNT" value="<%=grandtotalrnd%>">
			<input type="hidden" name="CURRENCY" value="<%=countryCurrency%>">
			<input type="hidden" name="AUTO_SETTLE_FLAG" value="<%=ppmethod%>">
			<input type="hidden" name="HPP_VERSION" value="2">
			<input type="hidden" name="HPP_CHANNEL" value="ECOM">
			<input type="hidden" name="HPP_CUSTOMER_EMAIL" value="<%=htmlspecials(ordEmail)%>">
			<input type="hidden" name="HPP_BILLING_STREET1" value="<%=htmlspecials(ordAddress)%>">
			<input type="hidden" name="HPP_BILLING_STREET2" value="<%=htmlspecials(ordAddress2)%>">
			<input type="hidden" name="HPP_BILLING_CITY" value="<%=htmlspecials(ordCity)%>">
<%		if countryCode="US" OR countryCode="CA" then print whv("HPP_BILLING_STATE",getstateabbrev(ordState)) %>
			<input type="hidden" name="HPP_BILLING_POSTALCODE" value="<%=htmlspecials(ordZip)%>">
<%			if data3<>"" then print whv("ACCOUNT",data3)
			sSQL="SELECT countryNumCurrency FROM countries WHERE countryCode='" & escape_string(countryCode) & "'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then print whv("HPP_BILLING_COUNTRY",rs("countryNumCurrency"))
			rs.close
			if shipCountryCode<>"" AND ordShipName<>"" AND ordShipAddress<>"" then %>
			<input type="hidden" name="HPP_SHIPPING_STREET1" value="<%=htmlspecials(ordShipAddress)%>">
			<input type="hidden" name="HPP_SHIPPING_STREET2" value="<%=htmlspecials(ordShipAddress2)%>">
			<input type="hidden" name="HPP_SHIPPING_CITY" value="<%=htmlspecials(ordShipCity)%>">
<%		if countryCode="US" OR countryCode="CA" then print whv("HPP_SHIPPING_STATE",getstateabbrev(ordShipState)) %>
			<input type="hidden" name="HPP_SHIPPING_POSTALCODE" value="<%=htmlspecials(ordShipZip)%>">
<%				sSQL="SELECT countryNumCurrency FROM countries WHERE countryCode='" & escape_string(shipCountryCode) & "'"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then print whv("HPP_SHIPPING_COUNTRY",rs("countryNumCurrency"))
				rs.close
			end if %>
			<input type="hidden" name="HPP_ADDRESS_MATCH_INDICATOR" value="FALSE">
			<input type="hidden" name="HPP_CHALLENGE_REQUEST_INDICATOR" value="NO_PREFERENCE">
			<input type="hidden" name="MERCHANT_RESPONSE_URL" value="<%=storeurlssl & "vsadmin/ajaxservice.asp?action=globalpayments"%>">
			<input type="hidden" name="SHA1HASH" value="<%=hashstring%>">
<%		end if
		print payprovextraparams
	end if
	if NOT returntocustomerdetails AND NOT (nopriceanywhere AND success) then ' {
		if xxCoStp3<>"" then print "<div class=""checkoutsteps"">" & xxCoStp3 & "</div>"%>
			<div class="cart3details">
<%		showcartresumeheader(3)
		if (rgcpncode<>"" OR ordPayProvider="19") AND (NOT gotcpncode OR cpnerror<>"") AND nogiftcertificate<>TRUE then %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3"><% if rgcpncode<>"" AND ordPayProvider="19" AND NOT gotcpncode then print "<span style=""color:#FF0000"">" & xxCpnNoF & "</span>" else print labeltxt(xxGifCer,"cpncode") & ":"%></div>
				<div class="cobll cobll3"><%
			if ordPayProvider="19" AND NOT gotcpncode then
				print "<input type=""text"" name=""cpncode"" id=""cpncode"" size=""20"" value=""" & htmlspecials(rgcpncode) & """ autocomplete=""off"" /> <input type=""button"" class=""ectbutton"" value=""" & xxAppCpn & """ onclick=""document.location='cart" & extension & "?token=" & getget("token") & "&cpncode='+document.getElementById('cpncode').value"" />"
			else
				print cpnerror
				if rgcpncode<>"" AND NOT gotcpncode then print "<div class=""expiredcoupon"">" & replace(replace(xxNoGfCr,"history.go(-%s)","checkcustdetails()"),"%s",rgcpncode,1,1) & "</div>"
			end if %></div>
			  </div>
<%		end if
		if ordPayProvider="19" AND noadditionalinfo<>TRUE then
			if getpost("shipformaddinfo")<>"" then ordAddInfo=getpost("shipformaddinfo") %>
			  <div class="cart3row">
				<div class="cobhl cobhl3"><%=xxAddInf%></div>
				<div class="cobll cobll3"><textarea name="ordAddInfo" id="ordAddInfo" class="addinfo"<% if addinfplaceholder<>"" then print " placeholder=""" & addinfplaceholder & """"%>><%=htmlspecials(ordAddInfo)%></textarea></div>
			  </div>
<%		end if
		if backorder then %>
			  <div class="cart3row">
				<div class="cobll cart2column ectwarning"><%=xxBakOrW%></div>
			  </div>
<%		end if
		if (warncheckspamfolder=TRUE OR getpost("warncheckspamfolder")="true") AND noconfirmationemail<>TRUE then %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3"><%=xxThkSub%></div>
				<div class="cobll cobll3 ectwarning"><%=xxSpmWrn%></div>
			  </div>
<%		end if
		if cpnmessage<>"" then %>
			  <div class="cart3row cart3discountsrow">
			    <div class="cobhl cobhl3"><%=xxAppDs%></div>
				<div class="cobll cobll3"><%=cpnmessage%></div>
			  </div>
<%		end if %>
			  <div class="cart3row cart3totgoodsrow">
			    <div class="cobhl cobhl3 cart3totgoodst"><%=xxTotGds%></div>
				<div class="cobll cobll3 cart3totgoods"><%=FormatEuroCurrency(totalgoods)%>
<script>/* <![CDATA[ */
function updateshiprate(obj,theselector){
	if(obj.value!=''){
		document.getElementById("shipselectoridx").value=theselector;
		document.getElementById("shipselectoraction").value="selector";
		if(document.getElementById('ordAddInfo')) document.getElementById("shipformaddinfo").value=document.getElementById('ordAddInfo').value;
		document.forms.shipform.submit();
	}
}
function selaltrate(id){
	document.getElementById('altrates').value=id;
	document.getElementById('shipselectoraction').value='altrates';
	document.forms.shipform.submit();
}
function setchangeflag(tisset,tname){
	if(tname=='w')
		document.getElementById('willpickup').value=tisset?'Y':'';
	else
		document.getElementById('changeaction').value=tname+(tisset?'y':'n');
	document.getElementById('shipselectoraction').value='altrates';
	document.forms.shipform.submit();
}
<%	if closeorderimmediately then
		SESSION("sessionid")=thesessionid %>
function docloseorder(){
	ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
	ajaxobj.open("GET", "vsadmin/ajaxservice.asp?action=clord", false);
	ajaxobj.send(null);
}
<%	end if
	if adminAltRates=2 then
		sSQL="SELECT altrateid FROM alternaterates WHERE usealtmethod"&international&"<>0 AND altrateid<>"&shipType&" ORDER BY altrateorder"&international&",altrateid"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			print "var extraship=["
			addcomma=""
			do while NOT rs.EOF
				print addcomma & rs("altrateid")
				addcomma=","
				rs.movenext
			loop
			print "];" & vbCrLf %>
function acajaxcallback(){
	if(ajaxobj.readyState==4){
		var restxt=ajaxobj.responseText;
		var gssr=restxt.split('SHIPSELPARAM=');
		document.getElementById('shipoptionstable').innerHTML+='<div class="shiptableline2"><div class="shiplogo2">'+decodeURIComponent(gssr[1])+'</div><div class="shiptablerates2">'+gssr[0]+'</div></div>';
		if(decodeURIComponent(gssr[2])!='ERROR')
			document.getElementById('numshiprate').value=gssr[4];
		getalternatecarriers();
	}
}
function getalternatecarriers(){
	if(extraship.length>0){
		var shiptype=extraship.shift();
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=acajaxcallback;
		ajaxobj.open("GET", "vsadmin/shipservice.asp?shiptype="+shiptype+"&numshiprate="+document.getElementById('numshiprate').value+"&sessionid=<%=urlencode(thesessionid)%>&destzip=<%=urlencode(destZip)%>&sc=<%=urlencode(shipCountryID)%>&scc=<%=urlencode(shipCountryCode)%>&sta=<%=urlencode(shipstateid)%>&orderid=<%=orderid%>", true);
		ajaxobj.send(null);
	}
}
<%		end if
		rs.close
	end if %>
function docheckform(){
<%		if termsandconditions AND (ordPayProvider="19" OR ordPayProvider="21") then %>
if(document.getElementById('ecttnccheckbox').checked==false){
	alert("<%=jscheck(xxPlsProc)%>");
	document.getElementById('ecttnccheckbox').focus();
	return(false);
}
<%		elseif recaptchaenabled(256) AND (grandtotal=0 OR ordPayProvider="4" OR ordPayProvider="17") then
			print "if(!cardentrycaptchaok){ alert(""" & jscheck(xxRecapt) & """);return(false); }"
		end if %>
return(true);
}
<%	if termsandconditions AND (ordPayProvider="19" OR ordPayProvider="21") then call gettermsjsfunction() %>
/* ]]> */</script>
				</div>
			  </div>
<%		if shipType=0 then combineshippinghandling=FALSE
		if shipType<>0 OR willpickup_ then
			if currShipType="" then currShipType=shipType
%>			  <div class="<%=IIfVs(adminAltRates=0,"cart3row ")%>cart3shipselrow">
<%			if adminAltRates=0 then
				print "<div class=""cobhl cobhl3 cart3shippingt""><div class=""shiplogo"">" & getshiplogo(currShipType) & "</div></div>"
%>				<div class="cobll cobll3 cart3shipping"><%
				print "<div class=""shipoptionstable" & IIfVs(NOT success," ectwarning") & """ id=""shipoptionstable"">"
				if NOT success then
					print errormsg
				else
					if shipType<>0 OR (shipping-freeshipamnt)<>0 OR willpickup_ then
						if NOT multipleoptions then print FormatEuroCurrency((cdbl(shipping)+IIfVr(combineshippinghandling=TRUE, handling, 0))-freeshipamnt) & IIfVr(shipMethod<>""," - " & shipMethod,"") else showshippingselect()
					end if
				end if
				print "</div></div>"
			elseif adminAltRates=1 then
				print "<div class=""cobhl cobhl3 cart3shippingt""><div class=""shipaltrates"">"
				sSQL="SELECT altrateid,altratename,"&getlangid("altratetext",65536)&",usealtmethod,usealtmethodintl FROM alternaterates WHERE usealtmethod"&international&"<>0 ORDER BY altrateorder"&international&",altrateid"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then %>
					<div class="cart3shiphead"><%="Compare Carriers"%></div>
					<div class="cart3alratelines"><%
					do while NOT rs.EOF
						call writealtshipline(rs(getlangid("altratetext",65536)),rs("altrateid"),"","",FALSE)
						rs.movenext
					loop
%>					</div>
<%					end if
				rs.close
				print "</div></div>"
%>				<div class="cobll cobll3 cart3shipping">
					<div class="flexvertalign"><div class="shiplogo"><%=getshiplogo(currShipType)%></div><div class="cart3shiphead">Select Shipping Method</div></div><%
				print "<div class=""shipoptionstable" & IIfVs(NOT success," ectwarning") & """ id=""shipoptionstable"">"
				if NOT success then
					print errormsg
				else
					if shipType<>0 OR (shipping-freeshipamnt)<>0 OR willpickup_ then
						if NOT multipleoptions then print FormatEuroCurrency((cdbl(shipping)+IIfVr(combineshippinghandling=TRUE, handling, 0))-freeshipamnt) & IIfVr(shipMethod<>""," - " & shipMethod,"") else showshippingselect()
					end if
				end if
				print "</div></div>"
			else
				print "<div class=""shipoptionstable2"" id=""shipoptionstable""><div class=""shiptableline2""><div class=""shiplogo2"">" & getshiplogo(currShipType) & "</div><div class=""shiptablerates2" & IIfVs(NOT success," ectwarning") & """>"
				if NOT success then
					print errormsg
				else
					if shipType<>0 OR (shipping-freeshipamnt)<>0 OR willpickup_ then
						if NOT multipleoptions then print FormatEuroCurrency((cdbl(shipping)+IIfVr(combineshippinghandling=TRUE, handling, 0))-freeshipamnt) & IIfVr(shipMethod<>""," - " & shipMethod,"") else showshippingselect()
					end if
				end if
				print "</div></div></div>"
			end if %>
			  </div>
<%			call writeshippingflags(3) %>
			  <div class="cart3row">
				<div class="cobhl cobhl3 cart3shippingt"><%=IIfVr(combineshippinghandling, xxShipHa, xxShippg)%></div>
				<div class="cobll cobll3 cart3shipping"><%
			if NOT success then
				print "<span class=""ectwarning"">" & errormsg & "</span>"
			else
				print FormatEuroCurrency((cdbl(shipping)+IIfVr(combineshippinghandling=TRUE, handling, 0))-freeshipamnt) & IIfVr(shipMethod<>""," - " & shipMethod,"")
			end if %></div>
			  </div>
<%		end if
		if success AND handling<>0 AND combineshippinghandling<>TRUE then %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3 cart3handlingt"><%=xxHndlg%></div>
				<div class="cobll cobll3 cart3handling"><%=FormatEuroCurrency(handling)%></div>
			  </div>
<%		end if
		if totaldiscounts<>0 AND ((totalgoods+cdbl(shipping)+handling)-(totaldiscounts+freeshipamnt))>=0 AND showtaxinclusive<>3 then %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3 ectdscntt"><%=xxTotDs%></div>
				<div class="cobll cobll3 ectdscnt"><%=FormatEuroCurrency(totaldiscounts)%></div>
			  </div>
			  <div class="cart3row">
			    <div class="cobhl cobhl3"><%=xxSubTot%></div>
				<div class="cobll cobll3 cart3sustot"><%=FormatEuroCurrency((totalgoods+cdbl(shipping)+handling)-(totaldiscounts+freeshipamnt))%></div>
			  </div>
<%		end if
		if usehst then %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3"><%=xxHST%></div>
				<div class="cobll cobll3"><%=FormatEuroCurrency(stateTax+countryTax)%></div>
			  </div>
<%		else
			if stateTax<>0.0 then %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3"><%=xxStaTax%></div>
				<div class="cobll cobll3"><%=FormatEuroCurrency(stateTax)%></div>
			  </div>
<%			end if
			if countryTax<>0.0 OR alwaysdisplaycountrytax then %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3"><%=xxCntTax%></div>
				<div class="cobll cobll3"><%=FormatEuroCurrency(countryTax)%></div>
			  </div>
<%			end if
		end if
		if totaldiscounts<>0 AND (((totalgoods+cdbl(shipping)+handling)-(totaldiscounts+freeshipamnt))<0 OR showtaxinclusive=3) then %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3 ectdscntt"><%=xxTotDs%></div>
				<div class="cobll cobll3 ectdscnt"><%=FormatEuroCurrency(totaldiscounts)%></div>
			  </div>
<%		end if %>
			  <div class="cart3row">
			    <div class="cobhl cobhl3 cart3gndtot<%=IIfVs(NOT success," ectwarning")%>"><%=IIfVr(success,xxGndTot,"Error")%></div>
				<div class="cobll cobll3 cart3gndtott<%=IIfVs(NOT success," ectwarning")%>"><%
					if success=FALSE then
						print IIfVr(trim(errormsg&"")<>"",errormsg,"-")
					else
						print FormatEuroCurrency(grandtotal)
					end if %></div>
			  </div>
<%		if NOT (ordPayProvider="7" OR ordPayProvider="13" OR ordPayProvider="18") then cardinalprocessor=""
		checkcarddetails=FALSE
		if success AND grandtotal>0 AND (ordPayProvider="7" OR ordPayProvider="10" OR ordPayProvider="12" OR ordPayProvider="13" OR (ordPayProvider="14" AND customppacceptcc) OR (ordPayProvider="16" AND (data2&"")="1") OR ordPayProvider="18" OR ordPayProvider="29" OR ordPayProvider="30") then ' { Payflow Pro OR PSiGate SSL OR Auth.NET AIM OR PayPal Pro OR NMI OR eWay
			checkcarddetails=TRUE
			if ordPayProvider="7" then
				vsdetails=Split(data1, "&")
			end if
			if ordPayProvider<>"10" then
				data1="XXXXXXX0XXXXXXXXXXXXXXXXX"
				if origCountryCode="GB" OR origCountryCode="IE" then data1="XXXXXXXXXXXXXXXXXXXXXXXXX"
			end if
			isPSiGate=(ordPayProvider="12")
			isLinkpoint=(ordPayProvider="16")
			isnmi=ordPayProvider="29"
			iseway=ordPayProvider="30"
			if isPSiGate then
				sscardname="Bname"
				sscardnum="CardNumber"
				ssexmon="CardExpMonth"
				ssexyear="CardExpYear"
				sscvv2="CardIDNumber"
			elseif isLinkpoint then
				sscardname="bname"
				sscardnum="cardnumber"
				ssexmon="expmonth"
				ssexyear="expyear"
				sscvv2="cvm"
			elseif isnmi then
				sscardname="nminame"
				sscardnum="billing-cc-number"
				ssexmon="nmimon"
				ssexyear="nmiyear"
				sscvv2="cvv"
				call writehiddenidvar("billing-cc-exp","")
			elseif iseway then
				sscardname="EWAY_CARDNAME"
				sscardnum="EWAY_CARDNUMBER"
				ssexmon="EWAY_CARDEXPIRYMONTH"
				ssexyear="EWAY_CARDEXPIRYYEAR"
				sscvv2="EWAY_CARDCVN"
			else
				sscardname="cardname"
				sscardnum="ACCT"
				ssexmon="EXMON"
				ssexyear="EXYEAR"
				sscvv2="CVV2"
			end if
			acceptecheck=((ppbits AND 1)=1 AND ordPayProvider="13") OR (customppacceptecheck AND ordPayProvider="14")
%>
<input type="hidden" name="sessionid" value="<%=thesessionid%>" />
<script>/* <![CDATA[ */
function setnmidate(){
	document.getElementById('billing-cc-exp').value=document.getElementById('nmimon').value+document.getElementById('nmiyear').value;
}
var isswitchcard=false;
function clearcc(){
	document.getElementById("<%=sscardnum%>").value="";
	document.getElementById("<%=sscvv2%>").value="";
	document.getElementById("<%=ssexmon%>").selectedIndex=0;
	document.getElementById("<%=ssexyear%>").selectedIndex=0;
}
function donecc(){
	return true;
}
if(window.addEventListener){
	window.addEventListener("load", clearcc, false);
	window.addEventListener("unload",donecc,false);
}else if(window.attachEvent){
	window.attachEvent("onload", clearcc);
}
function isCreditCard(st){
  if(st.length>19)return(false);
  sum=0; mul=1; l=st.length;
  for(i=0; i < l; i++){
	digit=st.substring(l-i-1,l-i);
	tproduct=parseInt(digit ,10)*mul;
	if(tproduct>=10)
		sum+=(tproduct % 10) + 1;
	else
		sum+=tproduct;
	if(mul==1)mul++;else mul--;
  }
  return((sum % 10)==0);
}
function isVisa(cc){
  if(((cc.length==16) || (cc.length==13)) && (cc.substr(0,1)==4))
	return isCreditCard(cc);
  return false;
}
function isMasterCard(cc){
  firstdig=cc.substr(0,1);
  seconddig=cc.substr(1,1);
  first4digs=parseInt(cc.substr(0,4));
  if((cc.length==16) && ((firstdig==5 && seconddig>=1 && seconddig<=5) || (first4digs>=2221 && first4digs<=2729)))
	return isCreditCard(cc);
  return false;
}
function isAmericanExpress(cc){
  firstdig=cc.substr(0,1);
  seconddig=cc.substr(1,1);
  if(cc.length==15 && firstdig==3 && (seconddig==4 || seconddig==7))
	return isCreditCard(cc);
  return false;
}
function isDinersClub(cc){
  firstdig=cc.substr(0,1);
  seconddig=cc.substr(1,1);
  if(cc.length==14 && firstdig==3 && (seconddig==0 || seconddig==6 || seconddig==8))
	return isCreditCard(cc);
  return false;
}
function isDiscover(cc){
  first4digs=cc.substr(0,4);
  if(cc.length==16 && (first4digs=="6011" || cc.substr(0,3)=="622" || cc.substr(0,2)=="64" || cc.substr(0,2)=="65"))
	return isCreditCard(cc);
  return false;
}
function isAusBankcard(cc){
  first4digs=cc.substr(0,4);
  if(cc.length==16 && (first4digs=="5610"||first4digs=="5602"))
	return isCreditCard(cc);
  return false;
}
function isEnRoute(cc){
  first4digs=cc.substr(0,4);
  if(cc.length==15 && (first4digs=="2014" || first4digs=="2149"))
	return isCreditCard(cc);
  return false;
}
function isJCB(cc){
  first4digs=cc.substr(0,4);
  if(cc.length==16 && (first4digs=="3088" || first4digs=="3096" || first4digs=="3112" || first4digs=="3158" || first4digs=="3337" || first4digs=="3528" || first4digs=="3589"))
	return isCreditCard(cc);
  return false;
}
function isSwitch(cc){
  first4digs=cc.substr(0,4);
  if((cc.length>=16 && cc.length<=19) && (first4digs=="4903" || first4digs=="4911" || first4digs=="4936" || first4digs=="5018" || first4digs=="5020" || first4digs=="5038" || first4digs=="5641" || first4digs=="6304" || first4digs=="6333" || first4digs=="6334" || first4digs=="6759" || first4digs=="6761" || first4digs=="6763" || first4digs=="6767")){
	isswitchcard=true;
	return(isCreditCard(cc));
  }
  return false;
}
function isLaser(cc){
  first4digs=cc.substr(0,4);
  if((cc.length>=16 && cc.length<=19) && (first4digs=="6304" || first4digs=="6706" || first4digs=="6771" || first4digs=="6709"))
	return(isCreditCard(cc));
  return false;
}
function isvalidcard(theForm){
  cc=theForm.elements['<%=sscardnum%>'].value;
  newcode="";
  isswitchcard=false;
  var l=cc.length;
  for(i=0;i<l;i++){
	digit=cc.substring(i,i+1);
	digit=parseInt(digit ,10);
	if(!isNaN(digit)) newcode += digit;
  }
  cc=newcode;
  if(theForm.<%=sscardname%>.value==""){
	alert("<%=jscheck(xxPlsEntr & " """ & xxCCName & """") %>");
	theForm.<%=sscardname%>.focus();
	return false;
  }
<% if acceptecheck then %>
if(cc!="" && theForm.accountnum.value!=""){
alert("Please enter either Credit Card OR ECheck details");
return(false);
}else if(theForm.accountnum.value!=""){
  if(theForm.accountname.value==""){
	alert("<%=jscheck(xxPlsEntr)%> \"Account Name\".");
	theForm.accountname.focus();
	return false;
  }
  if(theForm.bankname.value==""){
	alert("<%=jscheck(xxPlsEntr)%> \"Bank Name\".");
	theForm.bankname.focus();
	return false;
  }
  if(theForm.routenumber.value==""){
	alert("<%=jscheck(xxPlsEntr)%> \"Routing Number\".");
	theForm.routenumber.focus();
	return false;
  }
  if(theForm.accounttype.selectedIndex==0){
	alert("Please select your account type: (Checking / Savings).");
	theForm.accounttype.focus();
	return false;
  }
<%		if wellsfargo=true then %>
  if(theForm.orgtype.selectedIndex==0){
	alert("Please select your account type: (Personal / Business).");
	theForm.orgtype.focus();
	return false;
  }
  if(theForm.taxid.value=="" && theForm.licensenumber.value==""){
	alert("Please enter either a Tax ID number or Drivers License Details.");
	theForm.taxid.focus();
	return false;
  }
  if(theForm.taxid.value==""){
	if(theForm.licensestate.selectedIndex==0){
		alert("Please select your Drivers License State.");
		theForm.licensestate.focus();
		return false;
	}
	if(theForm.dldobmon.selectedIndex==0){
		alert("Please select your Drivers License D.O.B. Month.");
		theForm.dldobmon.focus();
		return false;
	}
	if(theForm.dldobday.selectedIndex==0){
		alert("Please select your Drivers License D.O.B. Day.");
		theForm.dldobday.focus();
		return false;
	}
	if(theForm.dldobyear.selectedIndex==0){
		alert("Please select your Drivers License D.O.B. year.");
		theForm.dldobyear.focus();
		return false;
	}
  }
<%		end if %>
}else{
<% end if %>
  if(true <% 
		if Mid(data1,8,1)="X" then print "&& !isSwitch(cc) "
		if Mid(data1,1,1)="X" then print "&& !isVisa(cc) "
		if Mid(data1,2,1)="X" then print "&& !isMasterCard(cc) "
		if Mid(data1,3,1)="X" then print "&& !isAmericanExpress(cc) "
		if Mid(data1,4,1)="X" then print "&& !isDinersClub(cc) "
		if Mid(data1,5,1)="X" then print "&& !isDiscover(cc) "
		if Mid(data1,6,1)="X" then print "&& !isEnRoute(cc) "
		if Mid(data1,7,1)="X" then print "&& !isJCB(cc) "
		if Mid(data1,9,1)="X" then print "&& !isAusBankcard(cc)"
		if Mid(data1,10,1)="X" then print "&& !isLaser(cc)" %>){
	<% if acceptecheck then xxValCC="Please enter a valid credit card number or bank account details if paying by ECheck." %>
	alert("<%=jscheck(xxValCC)%>");
	theForm.elements['<%=sscardnum%>'].focus();
	return false;
  }
  if(theForm.<%=ssexmon%>.selectedIndex==0){
	alert("<%=jscheck(xxCCMon)%>");
	theForm.<%=ssexmon%>.focus();
	return false;
  }
  if(theForm.<%=ssexyear%>.selectedIndex==0){
	alert("<%=jscheck(xxCCYear)%>");
	theForm.<%=ssexyear%>.focus();
	return false;
  }
<%	if Mid(data1,8,1)="X" then %>
	theForm.IssNum.value=theForm.IssNum.value.replace(/[^0-9]/g, '');
  if(theForm.IssNum.value=="" && isswitchcard){
	alert("Please enter an issue number / start date for Maestro/Solo cards.");
	theForm.IssNum.focus();
	return false;
  }
<%	end if %>
  if(theForm.<%=sscvv2%>.value==""){
	alert("<%=jscheck(xxPlsEntr & " """ & xx34code & """")%>");
	theForm.<%=sscvv2%>.focus();
	return false;
  }
<%	if acceptecheck then print "}"
	if recaptchaenabled(1) then print "if(!cardentrycaptchaok){ alert(""" & jscheck(xxRecapt) & """);return(false); }" %>
	theForm.elements['<%=sscardnum%>'].value=cc;
	return true;
}
<%	if cardinalprocessor<>"" AND cardinalmerchant<>"" AND cardinalpwd<>"" then %>
vbvtext='<html><head><title>Verified by Visa</title><style type="text/css">body {font-family: verdana,sans-serif;font-size:10pt;}</style></head><body><p><h3><%=replace(xxVBV1,"'","\'")%></h3></p><p><%=replace(xxVBV2,"'","\'")%><img src="images/vbv_logo.gif" border="0" style="float:<%=tright%>;margin:4px;" /></p><p><%=replace(xxVBV3,"'","\'")%></p><p><%=replace(xxVBV4,"'","\'")%></p><p><%=replace(xxVBV5,"'","\'")%></p><p align="center"><input type="button" class="ectbutton" value="<%=replace(xxClsWin,"'","\'")%>" onclick="window.close()"></p></body></'+'html>';
<%	end if %>
/* ]]> */</script>
<%			if request.servervariables("HTTPS")<>"on" AND (Request.ServerVariables("SERVER_PORT_SECURE")<>"1") AND nochecksslserver<>TRUE then %>
			  <div class="cart4row">
			    <div class="cobhl cart2column ectwarning">This site may not be secure. Do not enter real Credit Card numbers.</div>
			  </div>
<%			end if %>
			  <div><div class="cobhl cart4header cartheader"><%=xxCCDets%></div></div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt(xxCCName,sscardname)%></div>
				<div class="cobll cobll4"><input type="text" class="cdform3fixw" name="<%=sscardname%>" name="<%=sscardname%>" size="21" value="<%=trim(ordName&" "&ordLastName)%>" AUTOCOMPLETE="off" /></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt(xxCrdNum,sscardnum)%></div>
				<div class="cobll cobll4"><input type="text" class="cdform3fixw" name="<%=sscardnum%>" id="<%=sscardnum%>" size="21" AUTOCOMPLETE="off" /></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=xxExpEnd%></div>
				<div class="cobll cobll4">
				  <select class="ectselectinput nofixw" name="<%=ssexmon%>" id="<%=ssexmon%>" size="1"<%=IIfVs(isnmi," onchange=""setnmidate()""")%>>
					<option value=""><%=xxMonth%></option>
					<%	for index=1 to 12
							if index < 10 then themonth="0" & index else themonth=index
							print "<option value='"&themonth&"'>"&themonth&"</option>"&vbCrLf
						next %>
				  </select> / <select class="ectselectinput nofixw" name="<%=ssexyear%>" id="<%=ssexyear%>" size="1"<%=IIfVs(isnmi," onchange=""setnmidate()""")%>>
					<option value=""><%=xxYear%></option>
					<%	thisyear=DatePart("yyyy", Date())
						for index=thisyear to thisyear+10
							print "<option value='"&IIfVr(isPSiGate OR isnmi,right(index,2),index)&"'>"&index&"</option>"&vbCrLf
						next %></select>
				</div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt(xx34code,sscvv2)%></div>
				<div class="cobll cobll4"><input class="ecttextinput nofixw" type="text" name="<%=sscvv2%>" id="<%=sscvv2%>" size="4" AUTOCOMPLETE="off" /></div>
			  </div>
<%			if Mid(data1,8,1)="X" then %>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Issue Number / Start Date (mmyy)",IssNum)%></div>
				<div class="cobll cobll4"><input class="ecttextinput nofixw" type="text" name="IssNum" id="IssNum" size="4" AUTOCOMPLETE="off" /> (Maestro/Solo Only)</div>
			  </div>
<%			end if
			if acceptecheck then ' Auth.net
%>			  <div>
			    <div class="cobhl cart2column"><div class="cartecheck cart4header">ECheck Details</div><div class="echeckeither ectwarning">Please enter either Credit Card OR ECheck details</div></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Account Name","accountname")%></div>
				<div class="cobll cobll4"><input type="text" name="accountname" id="accountname" size="21" AUTOCOMPLETE="off" value="<%=trim(ordName&" "&ordLastName)%>" /></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Account Number","accountnum")%></div>
				<div class="cobll cobll4"><input type="text" name="accountnum" id="accountnum" size="21" AUTOCOMPLETE="off" /></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Bank Name","bankname")%></div>
				<div class="cobll cobll4"><input type="text" name="bankname" id="bankname" size="21" AUTOCOMPLETE="off" /></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Routing Number","routenumber")%></div>
				<div class="cobll cobll4"><input type="text" name="routenumber" id="routenumber" size="10" AUTOCOMPLETE="off" /></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Account Type","accounttype")%></div>
				<div class="cobll cobll4"><select name="accounttype" id="accounttype" size="1"><option value=""><%=xxPlsSel%></option><option value="checking">Checking</option><option value="savings">Savings</option><option value="businessChecking">Business Checking</option></select></div>
			  </div>
<%				if wellsfargo=TRUE then %>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Personal or Business Acct.","orgtype")%></div>
				<div class="cobll cobll4"><select name="orgtype" id="orgtype" size="1"><option value=""><%=xxPlsSel%></option><option value="I">Personal</option><option value="B">Business</option></select></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Tax ID","taxid")%></div>
				<div class="cobll cobll4"><input type="text" name="taxid" id="taxid" size="21" AUTOCOMPLETE="off" /></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4 cart2column carttaxidnot">If you have provided a Tax ID then the following information is not necessary</div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Drivers License Number","licensenumber")%></div>
				<div class="cobll cobll4"><input type="text" name="licensenumber" id="licensenumber" size="21" AUTOCOMPLETE="off" /></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Drivers License State","licensestate")%></div>
				<div class="cobll cobll4"><select size="1" name="licensestate" id="licensestate"><option value=""><%=xxPlsSel%></option><%
					sSQL="SELECT stateName,stateAbbrev FROM states WHERE stateCountryID=1 ORDER BY stateName"
					rs.open sSQL,cnn,0,1
					do while not rs.EOF
						print "<option value=""" & htmlspecials(rs("stateAbbrev")) & """>"&rs("stateName")&"</option>"
						rs.movenext
					loop
					rs.close %></select></div>
			  </div>
			  <div class="cart4row">
			    <div class="cobhl cobhl4"><%=labeltxt("Date Of Birth On License","dldobmon")%></div>
				<div class="cobll cobll4"><select name="dldobmon" id="dldobmon" size="1"><option value=""><%=xxMonth%></option>
<%					for index=1 to 12 %><option value="<%=index%>"><%=MonthName(index)%></option><% next %>
				</select> <select name="dldobday" size="1"><option value="">Day</option>
<%					for index=1 to 31 %><option value="<%=index%>"><%=index%></option><% next %>
				</select> <select name="dldobyear" size="1"><option value=""><%=xxYear%></option>
<%					for index=Year(date())-100 to Year(date()) %><option value="<%=index%>"><%=index%></option><% next %>
				</select></div>
			  </div>
<%				end if
			end if
		elseif success AND grandtotal>0 AND ordPayProvider="28" then ' SquareUp
%>
<script src="https://js.squareup<%=IIfVs(demomode,"sandbox")%>.com/v2/paymentform"></script>
<div id="squarecover" class="ectopaque" style="display:none"><img src="images/preloader.gif" alt="" style="margin-top:250px"></div>
<div class="sq-payment-form">
	<div id="sq-walletbox">
		<button type="button" id="sq-google-pay" class="button-google-pay" style="background-color:#000;background-origin:content-box;background-position:center;background-repeat:no-repeat;background-size:contain;background-image:url(data:image/svg+xml,%3Csvg%20width%3D%22103%22%20height%3D%2217%22%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%3E%3Cg%20fill%3D%22none%22%20fill-rule%3D%22evenodd%22%3E%3Cpath%20d%3D%22M.148%202.976h3.766c.532%200%201.024.117%201.477.35.453.233.814.555%201.085.966.27.41.406.863.406%201.358%200%20.495-.124.924-.371%201.288s-.572.64-.973.826v.084c.504.177.912.471%201.225.882.313.41.469.891.469%201.442a2.6%202.6%200%200%201-.427%201.47c-.285.43-.667.763-1.148%201.001A3.5%203.5%200%200%201%204.082%2013H.148V2.976zm3.696%204.2c.448%200%20.81-.14%201.085-.42.275-.28.413-.602.413-.966s-.133-.684-.399-.959c-.266-.275-.614-.413-1.043-.413H1.716v2.758h2.128zm.238%204.368c.476%200%20.856-.15%201.141-.448.285-.299.427-.644.427-1.036%200-.401-.147-.749-.441-1.043-.294-.294-.688-.441-1.183-.441h-2.31v2.968h2.366zm5.379.903c-.453-.518-.679-1.239-.679-2.163V5.86h1.54v4.214c0%20.579.138%201.013.413%201.302.275.29.637.434%201.085.434.364%200%20.686-.096.966-.287.28-.191.495-.446.644-.763a2.37%202.37%200%200%200%20.224-1.022V5.86h1.54V13h-1.456v-.924h-.084c-.196.336-.5.611-.91.826-.41.215-.845.322-1.302.322-.868%200-1.528-.259-1.981-.777zm9.859.161L16.352%205.86h1.722l2.016%204.858h.056l1.96-4.858H23.8l-4.41%2010.164h-1.624l1.554-3.416zm8.266-6.748h1.666l1.442%205.11h.056l1.61-5.11h1.582l1.596%205.11h.056l1.442-5.11h1.638L36.392%2013h-1.624L33.13%207.876h-.042L31.464%2013h-1.596l-2.282-7.14zm12.379-1.337a1%201%200%200%201-.301-.735%201%201%200%200%201%20.301-.735%201%201%200%200%201%20.735-.301%201%201%200%200%201%20.735.301%201%201%200%200%201%20.301.735%201%201%200%200%201-.301.735%201%201%200%200%201-.735.301%201%201%200%200%201-.735-.301zM39.93%205.86h1.54V13h-1.54V5.86zm5.568%207.098a1.967%201.967%200%200%201-.686-.406c-.401-.401-.602-.947-.602-1.638V7.218h-1.246V5.86h1.246V3.844h1.54V5.86h1.736v1.358H45.75v3.36c0%20.383.075.653.224.812.14.187.383.28.728.28.159%200%20.299-.021.42-.063.121-.042.252-.11.392-.203v1.498c-.308.14-.681.21-1.12.21-.317%200-.616-.051-.896-.154zm3.678-9.982h1.54v2.73l-.07%201.092h.07c.205-.336.511-.614.917-.833.406-.22.842-.329%201.309-.329.868%200%201.53.254%201.988.763.457.509.686%201.202.686%202.079V13h-1.54V8.688c0-.541-.142-.947-.427-1.218-.285-.27-.656-.406-1.113-.406-.345%200-.656.098-.931.294a2.042%202.042%200%200%200-.651.777%202.297%202.297%200%200%200-.238%201.029V13h-1.54V2.976zm32.35-.341v4.083h2.518c.6%200%201.096-.202%201.488-.605.403-.402.605-.882.605-1.437%200-.544-.202-1.018-.605-1.422-.392-.413-.888-.62-1.488-.62h-2.518zm0%205.52v4.736h-1.504V1.198h3.99c1.013%200%201.873.337%202.582%201.012.72.675%201.08%201.497%201.08%202.466%200%20.991-.36%201.819-1.08%202.482-.697.665-1.559.996-2.583.996h-2.485v.001zm7.668%202.287c0%20.392.166.718.499.98.332.26.722.391%201.168.391.633%200%201.196-.234%201.692-.701.497-.469.744-1.019.744-1.65-.469-.37-1.123-.555-1.962-.555-.61%200-1.12.148-1.528.442-.409.294-.613.657-.613%201.093m1.946-5.815c1.112%200%201.989.297%202.633.89.642.594.964%201.408.964%202.442v4.932h-1.439v-1.11h-.065c-.622.914-1.45%201.372-2.486%201.372-.882%200-1.621-.262-2.215-.784-.594-.523-.891-1.176-.891-1.96%200-.828.313-1.486.94-1.976s1.463-.735%202.51-.735c.892%200%201.629.163%202.206.49v-.344c0-.522-.207-.966-.621-1.33a2.132%202.132%200%200%200-1.455-.547c-.84%200-1.504.353-1.995%201.062l-1.324-.834c.73-1.045%201.81-1.568%203.238-1.568m11.853.262l-5.02%2011.53H96.42l1.864-4.034-3.302-7.496h1.635l2.387%205.749h.032l2.322-5.75z%22%20fill%3D%22%23FFF%22%2F%3E%3Cpath%20d%3D%22M75.448%207.134c0-.473-.04-.93-.116-1.366h-6.344v2.588h3.634a3.11%203.11%200%200%201-1.344%202.042v1.68h2.169c1.27-1.17%202.001-2.9%202.001-4.944%22%20fill%3D%22%234285F4%22%2F%3E%3Cpath%20d%3D%22M68.988%2013.7c1.816%200%203.344-.595%204.459-1.621l-2.169-1.681c-.603.406-1.38.643-2.29.643-1.754%200-3.244-1.182-3.776-2.774h-2.234v1.731a6.728%206.728%200%200%200%206.01%203.703%22%20fill%3D%22%2334A853%22%2F%3E%3Cpath%20d%3D%22M65.212%208.267a4.034%204.034%200%200%201%200-2.572V3.964h-2.234a6.678%206.678%200%200%200-.717%203.017c0%201.085.26%202.11.717%203.017l2.234-1.731z%22%20fill%3D%22%23FABB05%22%2F%3E%3Cpath%20d%3D%22M68.988%202.921c.992%200%201.88.34%202.58%201.008v.001l1.92-1.918c-1.165-1.084-2.685-1.75-4.5-1.75a6.728%206.728%200%200%200-6.01%203.702l2.234%201.731c.532-1.592%202.022-2.774%203.776-2.774%22%20fill%3D%22%23E94235%22%2F%3E%3C%2Fg%3E%3C%2Fsvg%3E);"></button>
		<button type="button" id="sq-apple-pay" class="sq-apple-pay"></button>
		<button type="button" id="sq-masterpass" class="sq-masterpass"></button>
		<div class="sq-wallet-divider"><span class="sq-wallet-divider__text">Or</span></div>
	</div>
	<div id="sq-ccbox">
		<div class="sq-field">
			<label class="ectlabel sq-label">Card Number</label>
			<div id="sq-card-number"></div>
		</div>
		<div class="sq-field-wrapper">
			<div class="sq-field sq-field--in-wrapper">
				<label class="ectlabel sq-label">CVV</label>
				<div id="sq-cvv"></div>
			</div>
			<div class="sq-field sq-field--in-wrapper">
				<label class="ectlabel sq-label">Expiration</label>
				<div id="sq-expiration-date"></div>
			</div>
			<div class="sq-field sq-field--in-wrapper">
				<label class="ectlabel sq-label">Postal</label>
				<div id="sq-postal-code"></div>
			</div>
		</div>
		<div class="sq-field">
			<button type="button" id="sq-creditcard" class="ectbutton sq-button" onclick="onGetCardNonce(event)">Pay <%=FormatEuroCurrency(grandtotal)%> Now</button>
		</div>
	</div>
</div>
<script>
<%	if data3<>"" then
		call splitname(trim(ordName&" "&ordLastName), firstname, lastname)
%>
const verificationDetails={
	intent:'CHARGE',
	amount:'<%=FormatNumber(grandtotal,2,-1,0,0)%>',
	currencyCode:'<%=countryCurrency%>',
	billingContact:{
		givenName:"<%=jscheck(firstname)%>",
		familyName:"<%=jscheck(lastname)%>",
		email:"<%=jscheck(ordEmail)%>",
		country:"<%=jscheck(countryCode)%>",
		postalCode:"<%=jscheck(ordZip)%>",
		region:"<%=jscheck(ordState)%>",
		city:"<%=jscheck(ordCity)%>",
		addressLines:["<%=jscheck(ordAddress)%>"<% if ordAddress2<>"" then print ",""" & jscheck(ordAddress2) & """"%>]
	}
};
<%	end if %>
const paymentForm = new SqPaymentForm({
	applicationId:"<%=data1%>",
	locationId:"<%=data3%>",
	inputClass:'sq-input',
	autoBuild:false,
	inputStyles:[{fontSize:'16px',lineHeight:'24px',padding:'16px',placeholderColor:'#a0a0a0',backgroundColor:'transparent',}],
	googlePay:{elementId:'sq-google-pay'},applePay:{elementId:'sq-apple-pay'},masterpass:{elementId:'sq-masterpass'},
	cardNumber:{elementId:'sq-card-number',placeholder:'Card Number'},
	cvv:{elementId:'sq-cvv',placeholder:'CVV'},
	expirationDate:{elementId:'sq-expiration-date',placeholder:'MM/YY'},
	postalCode:{elementId:'sq-postal-code',placeholder:'Postal'},
	callbacks:{
		methodsSupported:function(methods){
			document.getElementById('sq-walletbox').style.display = methods.masterpass || methods.applePay || methods.googlePay ? 'block' :'none';
			if (methods.googlePay === true) document.getElementById('sq-google-pay').style.display = 'inline-block';
			if (methods.applePay === true) document.getElementById('sq-apple-pay').style.display = 'inline-block';
			if (methods.masterpass === true) document.getElementById('sq-masterpass').style.display = 'inline-block';
		},
		createPaymentRequest:function(){
			var paymentRequestJson = {
				requestShippingAddress:false,
				requestBillingInfo:false,
				currencyCode:"<%=countryCurrency%>",
				total:{
					label:"Cart Checkout",
					amount:"<%=FormatNumber(grandtotal,2,-1,0,0)%>",
					pending:false
				}
			};
			return paymentRequestJson;
		},
		cardNonceResponseReceived:function (errors, nonce, cardData) {
			if(errors){
				handlesquareerror();
				alert(errors[0].message);
				return;
			}
			document.getElementById('squarecover').style.display='';
<%	if data3<>"" then %>
			paymentForm.verifyBuyer(nonce, verificationDetails, function(err, verificationResult){
				if (err == null) {
<%	end if %>
					document.getElementById('payment_method_nonce').value=nonce;
					fetch('vsadmin/ajaxservice.asp?action=squareup',{
						method:'POST',
						headers:{'Content-Type':'application/x-www-form-urlencoded'},
						body:'ordernumber=<%=orderid%>&nonce='+nonce<% if data3<>"" then print "+'&verification='+verificationResult.token"%>
					}).catch(err => {
						handlesquareerror();
						alert('Network error:' + err);
					}).then(response => {
						if (!response.ok) return response.text().then(errorInfo => Promise.reject(errorInfo));
						return response.text();
					}).then(data => {
						if(data.substr(0,7)=='SUCCESS'){
							document.getElementById('payment_method_nonce').value=nonce;
							document.getElementById('txnid').value=data.substr(8);
							document.getElementById('ectcheckoutform').submit();
						}else{
							handlesquareerror();
							alert(data.substr(8));
						}
					}).catch(err => {
						handlesquareerror();
						alert('Payment failed to complete:' + err);
					});
<%	if data3<>"" then %>
				}else{
					handlesquareerror();
					alert(err.message);
				}
			});
<%	end if %>
		}
	
	}
});
function handlesquareerror(){
	document.getElementById('sq-creditcard').disabled=false;
	document.getElementById('squarecover').style.display='none';
}
function onGetCardNonce(event){
	document.getElementById('sq-creditcard').disabled=true;
	document.getElementById('squarecover').style.display='';
	event.preventDefault();
	paymentForm.requestCardNonce();
}
paymentForm.build();
</script>
<%		end if ' }
		if success AND ordPayProvider<>"28" then
			if cardinalprocessor<>"" AND cardinalmerchant<>"" AND cardinalpwd<>"" then%>
				<div class="cart4row"><div class="cobhl cart2column"><%=xxCentl%></div></div>
<%			end if
			if (checkcarddetails AND recaptchaenabled(1)) OR ((grandtotal=0 OR ordPayProvider="4" OR ordPayProvider="17") AND recaptchaenabled(256)) then
				call displayrecaptchajs("cardentrycaptcha",TRUE,FALSE)
				print "<div style=""text-align:center""><div id=""cardentrycaptcha"" class=""g-recaptcha reCAPTCHAcheckout"" style=""margin-bottom:10px;display:inline-block""></div></div>"
			end if %>
				<div class="cobll cart2column checkoutbutton3"><%
			if orderid<>0 then
				if ordPayProvider="27" then
					ppfunding="" : ppfundingd="" : ppcostyle="size:'responsive'"
					buttonstyle=split(data3,"|")
					if UBOUND(buttonstyle)>=0 then buttonstyle0=buttonstyle(0) else buttonstyle0=""
					if UBOUND(buttonstyle)>=1 then buttonstyle1=buttonstyle(1) else buttonstyle1=""
					if UBOUND(buttonstyle)>=2 then buttonstyle2=buttonstyle(2) else buttonstyle2=""
					if UBOUND(buttonstyle)>=3 then buttonstyle3=buttonstyle(3) else buttonstyle3=""
					if UBOUND(buttonstyle)>=4 then buttonstyle4=buttonstyle(4) else buttonstyle4=""
					if UBOUND(buttonstyle)>=5 then buttonstyle5=buttonstyle(5) else buttonstyle5=""
					if UBOUND(buttonstyle)>=6 then buttonstyle6=buttonstyle(6) else buttonstyle6=""
					if UBOUND(buttonstyle)>=7 then buttonstyle7=buttonstyle(7) else buttonstyle7=""

					if buttonstyle0<>"" then ppcostyle=ppcostyle&",shape:""" & buttonstyle0 & """"
					if buttonstyle1<>"" then ppcostyle=ppcostyle&",size:""" & buttonstyle1 & """"
					if buttonstyle2<>"" then ppcostyle=ppcostyle&",color:""" & buttonstyle2 & """"
					if buttonstyle3<>"" then ppcostyle=ppcostyle&",layout:""" & buttonstyle3 & """"
					ppcostyle="style:{" & ppcostyle & "}," & vbLf
					if buttonstyle4="hide" then ppfundingd="paypal.FUNDING.CREDIT" else ppfunding="paypal.FUNDING.CREDIT"
					if buttonstyle5="hide" then ppfundingd=ppfundingd&IIfVs(ppfundingd<>"",",") & "paypal.FUNDING.CARD" else ppfunding=ppfunding&IIfVs(ppfunding<>"",",") & "paypal.FUNDING.CARD"
					if buttonstyle6="hide" then ppfundingd=ppfundingd&IIfVs(ppfundingd<>"",",") & "paypal.FUNDING.ELV" else ppfunding=ppfunding&IIfVs(ppfunding<>"",",") & "paypal.FUNDING.ELV"
					if buttonstyle7<>"display" then ppfundingd=ppfundingd&IIfVs(ppfundingd<>"",",") & "paypal.FUNDING.VENMO" else ppfunding=ppfunding&IIfVs(ppfunding<>"",",") & "paypal.FUNDING.VENMO"
					ppfunding="funding:{allowed:[" & ppfunding & "],disallowed:[" & ppfundingd & "]}," & vbLf
%>
<script src="https://www.paypalobjects.com/api/checkout.js"></script>
<div id="paypal-button"></div>
<script>
	paypal.Button.render({
		<%=ppcostyle&ppfunding%>meta:{partner_attribution_id:'ecommercetemplates_Cart_EC_US'},
		env:'<%=IIfVr(demomode,"sandbox","production")%>',
		payment:function(data, actions){ // Set up the payment:
			return actions.request.post('<%=storeurlssl%>vsadmin/ajaxservice.asp?action=createppsale&ordid=<%=orderid%>').then(function(res){
				return res.id;
			});
		},
		onAuthorize:function(data, actions){ // Execute the payment:
			return actions.request.post('<%=storeurlssl%>vsadmin/ajaxservice.asp?action=executeppsale&ordid=<%=orderid%>',{
				paymentID:data.paymentID,
				payerID:data.payerID
			}).then(function(res) {
				document.location='thanks.asp?action=termpp&paymentID='+data.paymentID+'&ordid=<%=orderid%>';
			});
		},
		onError:function(err){
			alert("PayPal Checkout Error: "+err);
		}
	}, '#paypal-button');
</script>
<%				elseif ordPayProvider="23" AND grandtotal>0 then
					if ppflag1=0 then
						print "<script src=""https://checkout.stripe.com/checkout.js"" class=""stripe-button"" data-key="""&data2&""" data-amount="""&round(grandtotal*100)&""" data-currency="""&countryCurrency&""" data-email="""&ordEmail&""" " & IIfVs(data3<>"","data-name="""&data3&""" ") & "data-description=""" & left(htmlspecials(descstr),255) & """ data-zip-code=""true"" data-billing-address=""true""></script>"
					else
						print "<script src=""https://js.stripe.com/v3""></script>"
						print "<button class=""ectbutton widecheckout3"" id=""checkout-button"" onclick=""stripe.redirectToCheckout({sessionId:'" & stripeid & "'}).then(function(result){alert(result.error.message)})"">" & xxPayCar & "</button>"
						print "<script>var stripe=Stripe('" & data2 & "');</script>"
					end if
				elseif ordPayProvider="31" OR ordPayProvider="32" then
					print imageorbutton(imgcheckoutbutton3,IIfVr(xxCOTxt3<>"",xxCOTxt3,xxCOTxt),"widecheckout3","payprovscript()",TRUE)
				else
					print imageorsubmit(imgcheckoutbutton3,IIfVr(xxCOTxt3<>"",xxCOTxt3,xxCOTxt)&IIfVs(closeorderimmediately,""" onclick=""docloseorder()"),"widecheckout3")
				end if
			end if %>
				</div>
<%		end if %>
			</div>
		</div>
<%		if shipType=4 then %>
			<div class="carriertm" style="text-align:center;font-size:10px;padding:6px"><%=xxUPStm%></div>
<%		elseif shipType=7 OR shipType=8 then %>
			<div class="carriertm" style="text-align:center;font-size:10px;padding:6px"><%=fedexcopyright%></div>
<%		end if
		print "</form>" & vbCrLf & "<form method=""post"" name=""shipform"" id=""shipform"" action=""cart" & extension & """>"
		print whv("mode","go") & whv("sessionid",thesessionid) & whv("orderid",orderid) & whv("cpncode",rgcpncode) & whv("token",token) & whv("payerid",payerid)
		call writehiddenidvar("altrates",shipType)
		call writehiddenidvar("shipselectoridx","")
		call writehiddenidvar("shipselectoraction","")
		call writehiddenidvar("numshiprate",numshiprate)
		call writehiddenidvar("changeaction","")
		call writehiddenidvar("willpickup","")
		call writehiddenidvar("shipformaddinfo","")
		call writehiddenidvar("sftermsandconds",getpost("sftermsandconds"))
		if amazonpayment then call writehiddenidvar("amzrefid",amzrefid_)
		if warncheckspamfolder then print whv("warncheckspamfolder","true")
		print "</form>"
		SESSION("shipselectoridx")=IIfVr(is_numeric(getpost("shipselectoridx")),getpost("shipselectoridx"),"")
		SESSION("shipselectoraction")=IIfVr(getpost("shipselectoraction")="selector" OR getpost("shipselectoraction")="altrates",getpost("shipselectoraction"),"")
		if NOT fromshipselector AND adminAltRates=2 then print "<script>getalternatecarriers();</script>"
	end if ' }
	if nopriceanywhere AND success then
		if recaptchaenabled(256) then
			print "<script>function docheckform(){if(!cardentrycaptchaok){ alert(""" & jscheck(xxRecapt) & """);return(false); }return(true);}</script>"
			call displayrecaptchajs("cardentrycaptcha",TRUE,FALSE)
			print "<div class=""cart3header"">" & xxRecapt & "</div>"
			print "<div class=""cart2column""><div id=""cardentrycaptcha"" class=""g-recaptcha reCAPTCHAcheckout""></div></div>"
			print "<div class=""cart2column"">"&imageorsubmit(imgcheckoutbutton3,IIfVr(xxCOTxt3<>"",xxCOTxt3,xxCOTxt),"checkoutbutton checkoutbutton3")&"</div>"
		end if
		print "</form>"
		if NOT recaptchaenabled(256) then print "<script>document.checkoutform.submit();</script>"
	end if
	if googletagid<>"" then
		print "<script>gtag(""event"",""add_payment_info"",{ currency:'" & countryCurrency & "',value:" & totalgoods & ",items:[" & getcartforganalytics("") & "]});</script>" & vbLf
	end if
elseif getget("amazon")="logout" then
	SESSION("AmazonLogin")=""
	SESSION("AmazonLoginTimeout")=""
end if ' }
if checkoutmode="checkout" then ' }{
	Dim ordName,ordLastName,ordAddress,ordAddress2,ordCity,ordState,ordZip,ordCountry,ordEmail,ordPhone,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipZip,ordShipCountry,ordAddInfo
	Dim havestate,allcountries
	sSQL="SELECT ordID FROM orders WHERE ordStatus>1 AND ordAuthNumber='' AND " & getordersessionsql()
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then ordID=rs("ordID") else ordID=""
	rs.close
	if ordID<>"" then
		release_stock(ordID)
		ect_query("UPDATE cart SET cartSessionID='"&replace(thesessionid,"'","")&"',cartClientID="&IIfVr(SESSION("clientID")<>"",SESSION("clientID"),0)&" WHERE cartCompleted=0 AND cartOrderID=" & ordID)
		ect_query("UPDATE orders SET ordAuthStatus='MODWARNOPEN',ordShipType='MODWARNOPEN' WHERE ordID=" & ordID)
	end if
	allcountries=""
	if getpost("checktmplogin")="x" then
		SESSION("clientID")=empty : SESSION("clientUser")=empty : SESSION("clientActions")=empty : SESSION("clientLoginLevel")=empty : SESSION("clientPercentDiscount")=empty
	elseif is_numeric(getpost("checktmplogin")) AND getpost("checktmplogin")<>"" then
		sSQL="SELECT tmploginname FROM tmplogin WHERE tmploginid='" & escape_string(getpost("sessionid")) & "' AND tmploginchk=" & replace(getpost("checktmplogin"),"'","")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			SESSION("clientID")=rs("tmploginname")
			rs.close
			sSQL="SELECT clUserName,clActions,clLoginLevel,clPercentDiscount,clEmail,clPW FROM customerlogin WHERE clID="&replace(SESSION("clientID"),"'","")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				SESSION("clientUser")=rs("clUserName")
				SESSION("clientActions")=rs("clActions")
				SESSION("clientLoginLevel")=rs("clLoginLevel")
				SESSION("clientPercentDiscount")=(100.0-cdbl(rs("clPercentDiscount")))/100.0
				get_wholesaleprice_sql()
				if rs("clEmail")<>request.cookies("WRITECLL") OR rs("clPW")<>request.cookies("WRITECLP") then
					call setacookie("WRITECLL",rs("clEmail"),0)
					call setacookie("WRITECLP",rs("clPW"),0)
				end if
			end if
		end if
		rs.close
	end if
	if SESSION("clientID")="" AND request.cookies("ectsessid")<>"" AND request.cookies("ecthash")<>"" AND is_numeric(request.cookies("ectordid")) AND NOT returntocustomerdetails then
		if request.cookies("ecthash")=sha256(request.cookies("ectordid")&request.cookies("ectsessid")&adminSecret) then call retrieveorderdetails(request.cookies("ectordid"),request.cookies("ectsessid"))
	end if
	if success then
		if ordZip="" then ordZip=SESSION("zip")
		if ordState="" then ordState=SESSION("state")
		if ordCountry="" then ordCountry=SESSION("country")
	end if
	sSQL="SELECT stateID FROM states INNER JOIN countries ON states.stateCountryID=countries.countryID WHERE countryEnabled<>0 AND stateEnabled<>0 AND loadStates=2 ORDER BY stateCountryID,stateName"
	rs.open sSQL,cnn,0,1
	hasstates=(NOT rs.EOF)
	rs.close
	sSQL="SELECT countryName,countryOrder,"&getlangid("countryName",8)&" AS cnameshow,countryID,loadStates FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC,"&getlangid("countryName",8)
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then allcountries=rs.getrows
	rs.close
	addresses=""
	if enableclientlogin AND SESSION("clientID")<>"" then
		sSQL="SELECT addID,addIsDefault,addName,addAddress,addAddress2,addState,addCity,addZip,addPhone,addCountry,addExtra1,addExtra2,addLastName FROM address INNER JOIN countries ON address.addCountry=countries.countryName WHERE countries.countryEnabled<>0 AND addCustID=" & replace(SESSION("clientID"),"'","") & " ORDER BY addName,addLastName,addAddress"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then addresses=rs.getrows
		rs.close
	end if %>
		    <form method="post" name="mainform" action="cart<%=extension%>" onsubmit="return checkform(this)">
<%	if IsArray(addresses) then %>
<script>/* <![CDATA[ */
var addrs=new Array();
addrs[0]={'name':'','lastname':'','address':'','address2':'','city':'','state':'','zip':'','phone':'','country':'','extra1':'','extra2':''};
function checkeditbutton(isshipping){
	adidobj=document.getElementById(isshipping + 'addressid');
	theaddy=adidobj[adidobj.selectedIndex].value;
	if(theaddy=='') document.getElementById(isshipping + 'editbutton').disabled=true; else document.getElementById(isshipping + 'editbutton').disabled=false;
}
function editaddress(isshipping,isaddnew){
	eval(isshipping+'checkaddress=true;');
	adidobj=document.getElementById(isshipping + 'addressid');
	theaddy=adidobj[adidobj.selectedIndex].value;
	if(isaddnew)theaddy=0;
	document.getElementById(isshipping + 'ordname').value=addrs[theaddy]['name'];
<%		if usefirstlastname then print "document.getElementById(isshipping + 'lastname').value=addrs[theaddy]['lastname'];" %>
	document.getElementById(isshipping + 'address').value=addrs[theaddy]['address'];
<%		if useaddressline2=TRUE then print "document.getElementById(isshipping + 'address2').value=addrs[theaddy]['address2'];" %>
	document.getElementById(isshipping + 'city').value=addrs[theaddy]['city'];
	document.getElementById(isshipping + 'zip').value=addrs[theaddy]['zip'];
	document.getElementById(isshipping + 'phone').value=addrs[theaddy]['phone'];
<%		if trim(extraorderfield1)<>"" then print "setdefs(document.getElementById('ord'+isshipping+'extra1'),addrs[theaddy]['extra1']);"
		if trim(extraorderfield2)<>"" then print "setdefs(document.getElementById('ord'+isshipping+'extra2'),addrs[theaddy]['extra2']);" %>
	thecntry=document.getElementById(isshipping + 'country')
	foundcntry=9999;
	for(var ind=0; ind < thecntry.length; ind++){
		if(thecntry[ind].value==addrs[theaddy]['countryid']){
			thecntry.selectedIndex=ind;
			foundcntry=ind;
		}
	}
	if(foundcntry==9999)thecntry.selectedIndex=0;
	dynamiccountries(document.getElementById(isshipping+'country'),isshipping);
	foundstate=0;
	checkoutspan(isshipping);
<%		if hasstates then %>
	thestate=document.getElementById(isshipping + 'state');
	if(countryhasstates[addrs[theaddy]['countryid']]){
		for(var ind=0; ind < thestate.length; ind++){
			if(thestate[ind].value==addrs[theaddy]['stateid'])
				foundstate=ind;
		}
	}else
		document.getElementById(isshipping+'state2').value=addrs[theaddy]['state'];
	thestate.selectedIndex=foundstate;
<%		end if %>
	showshipform(1,thecntry);
}
<%		for ii=0 to UBOUND(addresses,2)
			countryid=getidfromcountry(addresses(9,ii))
			stateid=getidfromstate(addresses(5,ii),countryid)
			print "addrs[" & addresses(0,ii) & "]={'name':'" & jschk(addresses(2,ii)) & "','lastname':'" & jschk(addresses(12,ii)) & "','address':'" & jschk(addresses(3,ii)) & "','address2':'" & jschk(addresses(4,ii)) & "','state':'" & jschk(addresses(5,ii)) & "','stateid':'" & stateid & "','city':'" & jschk(addresses(6,ii)) & "','zip':'" & jschk(addresses(7,ii)) & "','phone':'" & jschk(addresses(8,ii)) & "','country':'" & jschk(addresses(9,ii)) & "','countryid':" & countryid & ",'extra1':'" & jschk(addresses(10,ii)) & "','extra2':'" & jschk(addresses(11,ii)) & "'};" & vbLf
		next %>
/* ]]> */</script>
<%	end if
	print whv("mode","go")
	print whv("sessionid",strip_tags2(trim(thesessionid)))
	print whv("PARTNER",strip_tags2(getpost("PARTNER")))
	print whv("altrates",strip_tags2(getpost("altrates")))
	colspan2=""
	colspan3=""
%>
	<input type="hidden" name="addaddress" id="addaddress" value="<%=IIfVr(IsArray(addresses),"","add")%>" />
	<input type="hidden" name="saddaddress" id="saddaddress" value="<%=IIfVr(IsArray(addresses),"","add")%>" />
<%	if xxCoStp2<>"" then print "<div class=""checkoutsteps"">" & xxCoStp2 & "</div>"%>
		<div class="cart2details">
<%	showcartresumeheader(2)
	if returntocustomerdetails AND errormsg<>"" then print "<div class=""cobhl cart2column ectwarning"">" & errormsg & "</div>" & vbCrLf
	if isarray(addresses) then
		print whv("addressarray",1) %>
				<div class="cartcheckoutsavedaddr">
				  <div class="cartheader cdformtitle"><%=xxBilAdd%></div>
				  <div class="cdformtitlell">
<%		sub writeaddressspans(isshp) %>
		<span id="<%=isshp%>addressspan1" style="display:block"><select class="ectselectinput cdform2fixw" name="<%=isshp%>addressid" id="<%=isshp%>addressid" size="1" onchange="checkeditbutton('<%=isshp%>')"><%
		if isshp="s" then print "<option value="""">" & xxSamAs & "</option>"
		for index=0 to UBOUND(addresses,2)
			print "<option value=""" & addresses(0, index) & """" & IIfVr(addresses(1, index)=IIfVr(isshp="s",2,1), " selected=""selected""", "") & ">" & htmlspecials(addresses(2, index)&""&IIfVs(usefirstlastname," "&addresses(12, index))) & ", " & htmlspecials(addresses(3, index)&"") & IIfVr(trim(addresses(4, index)&"")<>"", ", " & htmlspecials(addresses(4, index)&""), "") & ", " & htmldisplay(addresses(5, index)&"") & "</option>"
		next %></select> <div class="editaddressbuttons"><input type="button" class="ectbutton" value="<%=xxEdit%>" id="<%=isshp%>editbutton" onclick="editaddress('<%=isshp%>',false);document.getElementById('<%=isshp%>addressspan1').style.display='none';document.getElementById('<%=isshp%>addressspan2').style.display='block';document.getElementById('<%=isshp%>addaddress').value='edit';"> <input type="button" class="ectbutton" value="<%=xxNew%>" onclick="editaddress('<%=isshp%>',true);document.getElementById('<%=isshp%>addressspan1').style.display='none';document.getElementById('<%=isshp%>addressspan2').style.display='block';document.getElementById('<%=isshp%>addaddress').value='add';" /></div>
		</span><div id="<%=isshp%>addressspan2" style="display:none">
			<%	if trim(extraorderfield1)<>"" then %>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtextra1"><%=IIfVr(extraorderfield1required=true,redstar,"") & labeltxt(extraorderfield1,"ord"&isshp&"extra1")%>:</div><div class="cobll cobll2 cdformextra1"><% if extraorderfield1html<>"" then print replace(replace(extraorderfield1html,"ectfield","ord"&isshp&"extra1"),"ordextra1","ord"&isshp&"extra1") else print "<input type=""text"" name=""ord"&isshp&"extra1"" id=""ord"&isshp&"extra1"" class=""ecttextinput cdform1fixw"" placeholder="""&stripnspecials(extraorderfield1)&""" autocomplete=""false"" />"%></div></div>
			<%	end if %>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtname<%=IIfVs(isshp="" AND errordname," ectwarning")%>"><%=redstar & labeltxt(xxName,isshp&"ordname")%>:</div><div class="cobll cobll2 cdformname"><%
			if usefirstlastname then
				print "<input type=""text"" name="""&isshp&"ordname"" id="""&isshp&"ordname"" class=""ecttextinput cdformsmfixw"" placeholder="""&stripnspecials(xxFirNam)&""" autocomplete=""given-name"" /> <input type=""text"" name="""&isshp&"lastname"" id="""&isshp&"lastname"" class=""ecttextinput cdformsmfixw"" placeholder="""&stripnspecials(xxLasNam)&""" autocomplete=""family-name"" />"
			else
				print "<input type=""text"" name="""&isshp&"ordname"" id="""&isshp&"ordname"" class=""ecttextinput cdform1fixw"" placeholder="""&stripnspecials(xxName)&""" />"
			end if
			%></div></div>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtaddress<%=IIfVs(isshp="" AND errordaddress," ectwarning")%>"><%=redstar & labeltxt(xxAddress,isshp&"address")%>:</div><div class="cobll cobll2 cdformaddress"><input type="text" name="<%=isshp%>address" id="<%=isshp%>address" class="ecttextinput cdform1fixw" placeholder="<%=stripnspecials(xxAddress)%>" /></div></div>
			<%	if useaddressline2=TRUE then %>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtaddress2"><%=labeltxt(xxAddress2,isshp&"address2")%>:</div><div class="cobll cobll2 cdformaddress2"><input type="text" name="<%=isshp%>address2" id="<%=isshp%>address2" class="ecttextinput cdform1fixw" placeholder="<%=stripnspecials(xxAddress2)%>" /></div></div>
			<%	end if %>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtcity<%=IIfVs(isshp="" AND errordcity," ectwarning")%>"><%=redstar & labeltxt(xxCity,isshp&"city")%>:</div><div class="cobll cobll2 cdformcity"><input type="text" name="<%=isshp%>city" id="<%=isshp%>city" class="ecttextinput cdform1fixw" placeholder="<%=stripnspecials(xxCity)%>" /></div></div>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtstate<%=IIfVs((isshp="s" AND errordshipstate) OR (isshp<>"s" AND errordstate)," ectwarning")%>"><%=replace(redstar,"<span","<span id="""&isshp&"statestar""")%><label class="ectlabel" for="<%=isshp%>state" id="<%=isshp%>statetxt"><%=xxState%></label>:</div><div class="cobll cobll2 cdformstate"><% if hasstates then %><select name="<%=isshp%>state" id="<%=isshp%>state" size="1" onchange="remwarning(this);dosavestate('')" class="ectselectinput cdform1fixw"><% havestate=show_states(-1) %></select><% end if %><input type="text" name="<%=isshp%>state2" id="<%=isshp%>state2" class="ecttextinput cdform1fixw" /></div></div>
			<div class="nohidebillrow"><div class="cobhl cobhl2 cdformtcountry<%=IIfVs(isshp="" AND errordcountry," ectwarning")%>"><%=redstar & labeltxt(xxCountry,isshp&"country")%>:</div><div class="cobll cobll2 cdformcountry"><select name="<%=isshp%>country" id="<%=isshp%>country" size="1" onchange="remwarning(this);checkoutspan('<%=isshp%>');showshipform(1,this)" class="ectselectinput cdform1fixw"><% call show_countries(-1,TRUE) %></select></div></div>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtzip<%=IIfVs(isshp="" AND errordzip," ectwarning")%>"><%=replace(redstar,"<span","<span id="""&isshp&"zipstar""") & "<label class=""ectlabel"" for="""&isshp&"zip"" id="""&isshp&"ziptxt"">" & xxZip & "</label>"%>:</div><div class="cobll cobll2 cdformzip"><input type="text" name="<%=isshp%>zip" id="<%=isshp%>zip" class="ecttextinput cdform1fixw" placeholder="<%=stripnspecials(xxZip)%>" autocapitalize="characters" /></div></div>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtphone<%=IIfVs(isshp="" AND errordphone," ectwarning")%>"><% if isshp="" then print redstar %><%=labeltxt(xxPhone,isshp&"phone")%>:</div><div class="cobll cobll2 cdformphone"><input type="tel" name="<%=isshp%>phone" id="<%=isshp%>phone" class="ecttextinput cdform1fixw" placeholder="<%=stripnspecials(xxPhone)%>" /></div></div>
			<%	if trim(extraorderfield2)<>"" then %>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtextra2"><%=IIfVr(extraorderfield2required=true,redstar,"") & labeltxt(extraorderfield2,"ord"&isshp&"extra2")%>:</div><div class="cobll cobll2 cdformextra2"><% if extraorderfield2html<>"" then print replace(replace(extraorderfield2html,"ectfield","ord"&isshp&"extra2"),"ordextra2","ord"&isshp&"extra2") else print "<input type=""text"" name=""ord"&isshp&"extra2"" id=""ord"&isshp&"extra2"" class=""ecttextinput cdform1fixw"" placeholder="""&stripnspecials(extraorderfield2)&""" autocomplete=""false"" />"%></div></div>
			<%	end if %>
			<div class="billformrow"><div class="cobhl cobhl2 cdformtextra2">&nbsp;</div><div class="cobll cobll2 cdformextra2"><input type="button" class="ectbutton cdform2fixw" value="<%=xxCancel%>" onclick="document.getElementById('<%=isshp%>addressspan2').style.display='none';document.getElementById('<%=isshp%>addressspan1').style.display='block';document.getElementById('<%=isshp%>addaddress').value='';<%=isshp%>checkaddress=false;"></div></div>
		</div>
<%		end sub
		call writeaddressspans("") %>
				  </div>
				</div>
<%		if noshipaddress<>TRUE then %>
				<div class="cartcheckoutsavedaddr cartshipsavedaddr">
				  <div class="cartheader cdformtitle"><%=xxShpAdd%></div>
				  <div class="cdformtitlell"><% call writeaddressspans("s") %></div>
				</div>
<%		end if
		call writeshippingflags(2)
	else
		if SESSION("clientID")<>"" then
			rs.open "SELECT clUserName,clEmail FROM customerlogin WHERE clID=" & SESSION("clientID")
			if NOT rs.EOF then
				ordName=trim(rs("clUserName")&"")
				if usefirstlastname then
					if instr(ordName," ")>0 then
						namearr=Split(ordName," ",2)
						ordName=namearr(0)
						ordLastName=namearr(1)
					else
						ordName=""
					end if
				end if
				ordEmail=rs("clEmail")
			end if
			rs.close
		end if
		function displayzip(isship) %>
				<div class="<%=IIfVr(isship="s","ship","bill")%>formrow">
				  <div class="cobhl cobhl2 cdformtzip<%=IIfVs((isship="" AND errordzip) OR (isship="s" AND errordshipaddress AND ordShipZip="")," ectwarning")%>"><%=replace(redstar,"<span","<span id="""&isship&"zipstar""") %><label class="ectlabel" for="<%=isship&"zip"%>" id="<%=isship%>ziptxt"><%=xxZip%></label></div>
				  <div class="cobll cobll2 cdformzip"><input type="text" name="<%=isship%>zip" class="cdform2fixw ecttextinput cdformzip" id="<%=isship%>zip" value="<%=htmlspecials(IIfVr(isship="s",ordShipZip,ordZip))%>" placeholder="<%=stripnspecials(xxZip)%>" autocapitalize="characters" /></div>
				</div>
<%		end function
		if trim(extraorderfield1)<>"" then %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtextra1"><% if extraorderfield1required=TRUE then print redstar
									print labeltxt(extraorderfield1,"ordextra1")%></div>
				  <div class="cobll cobll2 cdformextra1"><% if extraorderfield1html<>"" then print replace(extraorderfield1html,"ectfield","ordextra1") else print "<input type=""text"" name=""ordextra1"" class=""cdform2fixw ecttextinput cdformextra1"" id=""ordextra1"" placeholder="""&stripnspecials(extraorderfield1)&""" value="""&htmlspecials(ordExtra1)&""" autocomplete=""false"" />"%></div>
				</div>
<%		end if %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtname<%=IIfVs(errordname," ectwarning")%>"><%=redstar & labeltxt(xxName,"ordname")%></div>
				  <div class="cobll cobll2 cdformname" style="white-space:nowrap"><%
		if usefirstlastname then
			print "<input type=""text"" name=""ordname"" class=""cdformsmfixw ecttextinput cdformname"" id=""ordname"" value="""&htmlspecials(ordName)&""" placeholder="""&stripnspecials(xxFirNam)&""" autocomplete=""given-name"" /> <input type=""text"" name=""lastname"" id=""lastname"" class=""cdformsmfixw ecttextinput cdformlastname"" value="""&htmlspecials(ordLastName)&""" placeholder="""&stripnspecials(xxLasNam)&""" autocomplete=""family-name"" />"
		else
			print "<input type=""text"" name=""ordname"" class=""cdform2fixw ecttextinput cdformname"" id=""ordname"" value="""&htmlspecials(ordName)&""" placeholder=""" & stripnspecials(xxName) & """ />"
		end if %></div>
				</div>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtemail<%=IIfVs(errordemail," ectwarning")%>"><%=redstar & labeltxt(xxEmail,"email")%></div>
				  <div class="cobll cobll2 cdformemail"><input type="email" name="email" class="cdform2fixw ecttextinput cdformemail" id="email" placeholder="<%=stripnspecials(xxEmail)%>" value="<%=htmlspecials(ordEmail)%>" /></div>
				</div>
<%		if verifyemail then %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtemail<%=IIfVs(errordemailv," ectwarning")%>"><%=redstar & labeltxt(xxEmVerf,"email2")%></div>
				  <div class="cobll cobll2 cdformemail"><input type="email" name="email2" class="cdform2fixw ecttextinput cdformemail" id="email2" placeholder="<%=stripnspecials(xxEmVerf)%>" value="<%=htmlspecials(ordEmail2)%>" /></div>
				</div>
<%		end if %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtaddress<%=IIfVs(errordaddress," ectwarning")%>"><%=redstar & labeltxt(xxAddress,"address")%></div>
				  <div class="cobll cobll2 cdformaddress"><input type="text" name="address" class="cdform2fixw ecttextinput cdformaddress" id="address" placeholder="<%=stripnspecials(xxAddress)%>" value="<%=htmlspecials(ordAddress)%>" /></div>
				</div>
<%		if useaddressline2=TRUE then %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtaddress2"><%=labeltxt(xxAddress2,"address2")%></div>
				  <div class="cobll cobll2 cdformaddress2"><input type="text" name="address2" class="cdform2fixw ecttextinput cdformaddress2" id="address2" placeholder="<%=stripnspecials(xxAddress2)%>" value="<%=htmlspecials(ordAddress2)%>" /></div>
				</div>
<%		end if
		if zipposition=4 then call displayzip("") %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtcity<%=IIfVs(errordcity," ectwarning")%>"><%=redstar & labeltxt(xxCity,"city")%></div>
				  <div class="cobll cobll2 cdformcity"><input type="text" name="city" class="cdform2fixw ecttextinput cdformcity" id="city" placeholder="<%=stripnspecials(xxCity)%>" value="<%=htmlspecials(ordCity)%>" /></div>
				</div>
<%		if zipposition=3 then call displayzip("") %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtstate<%=IIfVs(errordstate," ectwarning")%>"><%=replace(redstar,"<span","<span id=""statestar""")%><label class="ectlabel" for="state" id="statetxt"><%=xxState%></label></div>
				  <div class="cobll cobll2 cdformstate"><% if hasstates then %><select name="state" class="cdform2fixw ectselectinput cdformstate" id="state" size="1" onchange="remwarning(this);dosavestate('')"><% havestate=show_states(ordState) %></select><% end if %><input type="text" name="state2" class="cdform2fixw ecttextinput cdformstate" id="state2" style="display:none" placeholder="<%=stripnspecials(xxState)%>" value="<% if not havestate then print htmlspecials(ordState)%>" /></div>
				</div>
<%		if zipposition=2 then call displayzip("") %>
				<div class="nohidebillrow">
				  <div class="cobhl cobhl2 cdformtcountry<%=IIfVs(errordcountry," ectwarning")%>"><%=redstar & labeltxt(xxCountry,"country")%></div>
				  <div class="cobll cobll2 cdformcountry"><select name="country" class="cdform2fixw ectselectinput cdformcountry" id="country" size="1" onchange="remwarning(this);checkoutspan('');showshipform(1,this)" ><% call show_countries(ordCountry,TRUE) %></select></div>
				</div>
<%		if zipposition=1 then call displayzip("") %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtphone<%=IIfVs(errordphone," ectwarning")%>"><%=redstar & labeltxt(xxPhone,"phone")%></div>
				  <div class="cobll cobll2 cdformphone"><input type="tel" name="phone" class="cdform2fixw ecttextinput cdformphone" id="phone" placeholder="<%=stripnspecials(xxPhone)%>" value="<%=htmlspecials(ordPhone)%>" /></div>
				</div>
<%		if trim(extraorderfield2)<>"" then %>
				<div class="billformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtextra2"><% if extraorderfield2required=TRUE then print redstar
									print labeltxt(extraorderfield2,"ordextra2")%></div>
				  <div class="cobll cobll2 cdformextra2"><% if extraorderfield2html<>"" then print replace(extraorderfield2html,"ectfield","ordextra2") else print "<input type=""text"" name=""ordextra2"" class=""cdform2fixw ecttextinput cdformextra2"" id=""ordextra2"" placeholder="""&stripnspecials(extraorderfield2)&""" value="""&ordExtra2&""" autocomplete=""false"" />"%></div>
				</div>
<%		end if
		call writeshippingflags(2)
		if noshipaddress<>TRUE then %>
				<div>
				  <div class="cdformshipdiff">
					<input type="hidden" name="shipdiff" id="shipdiff" value="<%=IIfVs(getpost("shipdiff")="1" OR (trim(ordShipName&ordShipLastName)<>"" AND trim(ordShipAddress)<>""),"1")%>" />
					<input type="button" class="ectbutton widecheckout2 shipdiff" value="<%=xxShpDff%>" onclick="document.getElementById('shipdiff').value=document.getElementById('shipdiff').value=='1'?'':'1';checkoutspan('s');showshipform(2,document.getElementById('scountry'))" />
				  </div>
				</div>
		<%	if trim(extraorderfield1)<>"" then %>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtextra1"><% if extraorderfield1required=TRUE then print redstar
									print labeltxt(extraorderfield1,"ordsextra1")%></div>
				  <div class="cobll cobll2 cdformextra1"><% if extraorderfield1html<>"" then print replace(replace(extraorderfield1html,"ordextra1","ordsextra1"),"ectfield","ordsextra1") else print "<input type=""text"" name=""ordsextra1"" class=""cdform2fixw ecttextinput cdformextra1"" id=""ordsextra1"" placeholder="""&stripnspecials(extraorderfield1)&""" value="""&htmlspecials(ordShipExtra1)&""" autocomplete=""false"" />"%></div>
				</div>
<%			end if %>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtname<%=IIfVs(errordshipaddress AND ordShipName=""," ectwarning")%>"><%=redstar & labeltxt(xxName,"sordname")%></div>
				  <div class="cobll cobll2 cdformname"><%
			if usefirstlastname then
				print "<input type=""text"" name=""sordname"" class=""cdformsmfixw ecttextinput cdformname"" id=""sordname"" value="""&htmlspecials(ordShipName)&""" placeholder="""&stripnspecials(xxFirNam)&""" /> <input type=""text"" name=""slastname"" class=""cdformsmfixw ecttextinput cdformlastname"" id=""slastname"" value="""&htmlspecials(ordShipLastName)&""" placeholder="""&stripnspecials(xxLasNam)&""" />"
			else
				print "<input type=""text"" name=""sordname"" class=""cdform2fixw ecttextinput cdformname"" id=""sordname"" value="""&htmlspecials(ordShipName)&""" placeholder=""" & stripnspecials(xxName) & """ />"
			end if %></div>
				</div>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtaddress<%=IIfVs(errordshipaddress AND ordShipAddress=""," ectwarning")%>"><%=redstar & labeltxt(xxAddress,"saddress")%></div>
				  <div class="cobll cobll2 cdformaddress"><input type="text" name="saddress" class="cdform2fixw ecttextinput cdformaddress" id="saddress" placeholder="<%=stripnspecials(xxAddress)%>" value="<%=htmlspecials(ordShipAddress)%>" /></div>
				</div>
<%			if useaddressline2=TRUE then %>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtaddress2"><%=labeltxt(xxAddress2,"saddress2")%></div>
				  <div class="cobll cobll2 cdformaddress2"><input type="text" name="saddress2" class="cdform2fixw ecttextinput cdformaddress2" id="saddress2" placeholder="<%=stripnspecials(xxAddress2)%>" value="<%=htmlspecials(ordShipAddress2)%>" /></div>
				</div>
<%			end if
			if zipposition=4 then call displayzip("s") %>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtcity<%=IIfVs(errordshipaddress AND ordShipCity=""," ectwarning")%>"><%=redstar & labeltxt(xxCity,"scity")%></div>
				  <div class="cobll cobll2 cdformcity"><input type="text" name="scity" class="cdform2fixw ecttextinput cdformcity" id="scity" placeholder="<%=stripnspecials(xxCity)%>" value="<%=htmlspecials(ordShipCity)%>" /></div>
				</div>
<%			if zipposition=2 then call displayzip("s") %>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtcountry<%=IIfVs(errordshipcountry," ectwarning")%>"><%=redstar & labeltxt(xxCountry,"scountry")%></div>
				  <div class="cobll cobll2 cdformcountry"><select name="scountry" class="cdform2fixw ectselectinput cdformcountry" id="scountry" size="1" onchange="remwarning(this);checkoutspan('s')"><option value=""><%=xxPlsSel%>...</option><% call show_countries(ordShipCountry,FALSE) %></select></div>
				</div>
<%			if zipposition=3 then call displayzip("s") %>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtstate<%=IIfVs(errordshipstate," ectwarning")%>"><%=replace(redstar,"<span","<span id=""sstatestar""")%><label class="ectlabel" for="sstate" id="sstatetxt"><%=xxState%></label></div>
				  <div class="cobll cobll2 cdformstate"><% if hasstates then %><select name="sstate" class="cdform2fixw ectselectinput cdformstate" id="sstate" size="1" onchange="remwarning(this);dosavestate('s')"><% havestate=show_states(ordShipState) %></select><% end if %><input type="text" name="sstate2" class="cdform2fixw ecttextinput cdformstate" id="sstate2" style="display:none" placeholder="<%=stripnspecials(xxState)%>" value="<% if not havestate then print htmlspecials(ordShipState)%>" /></div>
				</div>
<%			if zipposition=1 then call displayzip("s") %>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtphone"><%=labeltxt(xxPhone,"sphone")%></div>
				  <div class="cobll cobll2 cdformphone"><input type="tel" name="sphone" class="cdform2fixw ecttextinput cdformphone" id="sphone" placeholder="<%=stripnspecials(xxPhone)%>" value="<%=htmlspecials(ordShipPhone)%>" /></div>
				</div>
<%			if trim(extraorderfield2)<>"" then %>
				<div class="shipformrow" style="display:none">
				  <div class="cobhl cobhl2 cdformtextra2"><% if extraorderfield2required=TRUE then print redstar
									print labeltxt(extraorderfield2,"ordsextra2")%></div>
				  <div class="cobll cobll2 cdformextra2"><% if extraorderfield2html<>"" then print replace(replace(extraorderfield2html,"ordextra2","ordsextra2"),"ectfield","ordsextra2") else print "<input type=""text"" name=""ordsextra2"" class=""cdform2fixw ecttextinput cdformextra2"" id=""ordsextra2"" placeholder="""&stripnspecials(extraorderfield2)&""" value="""&ordShipExtra2&""" autocomplete=""false"" />"%></div>
				</div>
<%			end if
		end if ' noshipaddress
	end if ' isarray(addresses)
%>				<div class="cobhl cartheader cart2subheader cart2column"><%=xxMisc%></div>
<%	if noadditionalinfo<>TRUE then %>
				<div class="checkoutadditionals checkoutadds2col">
				  <div class="cobhl cobhl2 cdformtaddinfo"><%=xxAddInf%></div>
				  <div class="cobll cobll2 cdformaddinfo"><textarea name="ordAddInfo" class="addinfo"<% if addinfplaceholder<>"" then print " placeholder=""" & addinfplaceholder & """"%>><%=htmlspecials(ordAddInfo)%></textarea></div>
				</div>
<%	end if
	if trim(extracheckoutfield1)<>"" then
		checkoutfield1=IIfVr(extracheckoutfield1required=TRUE, redstar, "") & labeltxt(extracheckoutfield1,"ordcheckoutextra1")
		checkoutfield2=IIfVr(extracheckoutfield1html<>"", replace(extracheckoutfield1html,"ectfield","ordcheckoutextra1"), "<input type=""text"" name=""ordcheckoutextra1"" class=""cdform2fixw ecttextinput cdformextraco1"" id=""ordcheckoutextra1"" size=""20"" value="""&htmlspecials(ordCheckoutExtra1)&""" placeholder=""" & htmlspecials(strip_tags2(extracheckoutfield1)) & """autocomplete=""false"" />")
%>				<div class="checkoutadditionals checkoutadds2col">
				  <div class="cobhl cobhl2 cdformtextraco1"><% if extracheckoutfield1reverse then print checkoutfield2 else print checkoutfield1 %></div>
				  <div class="cobll cobll2 cdformextraco1"><% if extracheckoutfield1reverse then print checkoutfield1 else print checkoutfield2 %></div>
				</div>
<%	end if
	if trim(extracheckoutfield2)<>"" then
		checkoutfield1=IIfVr(extracheckoutfield2required=TRUE, redstar, "") & labeltxt(extracheckoutfield2,"ordcheckoutextra2")
		checkoutfield2=IIfVr(extracheckoutfield2html<>"", replace(extracheckoutfield2html,"ectfield","ordcheckoutextra2"), "<input type=""text"" name=""ordcheckoutextra2"" class=""cdform2fixw ecttextinput cdformextraco2"" id=""ordcheckoutextra2"" size=""20"" value="""&htmlspecials(ordCheckoutExtra2)&""" placeholder=""" & htmlspecials(strip_tags2(extracheckoutfield2)) & """ autocomplete=""false"" />")
%>				<div class="checkoutadditionals checkoutadds2col">
				  <div class="cobhl cobhl2 cdformtextraco2"><% if extracheckoutfield2reverse then print checkoutfield2 else print checkoutfield1 %></div>
				  <div class="cobll cobll2 cdformextraco2"><% if extracheckoutfield2reverse then print checkoutfield1 else print checkoutfield2 %></div>
				</div>
<%	end if
	if SESSION("clientID")="" then
		if enableclientlogin AND allowclientregistration AND NOT nocreateaccountoncheckout then %>
				<div class="checkoutadditionals checkoutadds2col">
				  <div id="co2newacctbutton" class="cobhl cobhl2 cdformtnewaccount"><%
			if displaysoftlogindone="" then displaysoftlogindone=""
			if enableclientlogin then call displaysoftlogin()  
			print imageorbutton(imgcreateaccount,xxCreAcc,"createaccount","co2displaynewaccount()",TRUE)
%>				  </div>
				  <div id="co2newaccttxt" class="cobll cobll2 cdformnewaccount"><%=xxClkCrA%><br><span class="ectsmallnote"><%=xxCrASvD%></span>
				  </div>
				</div>
<%		end if
	end if
	if nomailinglist<>TRUE then %>
				<div class="billformrowflags">
				  <div class="cdaddtflag cdformtmailing<%=IIfVs(mailinglistdropdown,"d")%>">
<%		if mailinglistdropdown then %>
					<select name="allowemail" id="allowemail" size="1">
						<option value=""><%=xxPlsSel%></option>
						<option value="ON"><%=xxYes%></option>
						<option value=""><%=xxNo%></option>
					</select>
<%		elseif mailinglistradios then %>
					<div><%=xxYes%>&nbsp;<input type="radio" class="ectradio" id="allowemailradioy" name="allowemail" value="ON" /></div>
					<div><%=xxNo%>&nbsp;<input type="radio" class="ectradio" id="allowemailradion" name="allowemail" value="" /></div>
<%		else %>
					<input type="checkbox" name="allowemail" class="ectcheckbox cdformcb cdformmailing" value="ON" <% if allowemaildefaulton=TRUE then print "checked=""checked"""%> />
<%		end if %>
				  </div>
				  <div class="cdaddflag cdformmailing<%=IIfVs(mailinglistdropdown,"d")%>"><%=xxAlPrEm%><br><span class="ectsmallnote"><%=xxNevDiv%></span></div>
				</div>
<%	end if
	if termsandconditions then %>
				<div class="billformrowflags">
				  <div class="cdaddtflag cdformtterms"><input id="ecttnccheckbox" type="checkbox" name="license" class="ectcheckbox cdformcb cdformterms" onchange="ectremoveclass(this,'ectwarning')" value="1" /></div>
				  <div class="cdaddflag cdformterms<%=IIfVs(errtermsandconditions," ectwarning")%>"><%=xxTermsCo%></div>
				</div>
<%	end if
	if nogiftcertificate<>TRUE then %>
				<div class="checkoutadditionals checkoutadds2col"><div class="cobhl cobhl2 cdformtcoupon"><%=labeltxt(xxGifNum,"cpncode")%></div><div class="cobll cobll2 cdformcoupon">
			<div><input type="text" name="cpncode" class="cdform2fixw ecttextinput cdformcoupon" id="cpncode" placeholder="<%=stripnspecials(xxGifNum)%>" size="<%=IIfVr(mobilebrowser,14,20)%>" autocomplete="off" /> <%=imageorbutton(imgapplycoupon,xxApply,"applycoupon applycoupon2","applycert()",TRUE)%></div>
		<div id="cpncodespan"><%
		if SESSION("giftcerts")<>"" OR SESSION("cpncode")<>"" then
			print "<div style=""display:table"">"
			gcarr=split(trim(SESSION("giftcerts")), " ")
			for index=0 to UBOUND(gcarr)
				print "<div style=""display:table-row""><div style=""display:table-cell"">" & xxAppGC & "</div><div style=""display:table-cell"">" & gcarr(index) & "</div><div style=""display:table-cell"">(<a href=""#"" onclick=""return removecert('"&gcarr(index)&"')"">"&xxRemove&"</a>)</div></div>"
			next
			cpnarr=split(trim(SESSION("cpncode")), " ")
			for index=0 to UBOUND(cpnarr)
				print "<div style=""display:table-row""><div style=""display:table-cell"">" & xxApdCpn & "</div><div style=""display:table-cell"">" & cpnarr(index) & "</div><div style=""display:table-cell"">(<a href=""#"" onclick=""return removecert('"&cpnarr(index)&"')"">"&xxRemove&"</a>)</div></div>"
			next
			print "</div>"
		end if %>
		</div>
		</div></div>
<%	end if
	ppsuccess=TRUE
	if returntocustomerdetails then print whv("token", token) & whv("payerid", payerid) & whv("checktmplogin", getpost("checktmplogin"))
	if IsEmpty(noemailgiftcertorders) then noemailgiftcertorders="4"
	sSQL="SELECT cartID FROM cart WHERE cartCompleted=0 AND (cartProdID='"&giftcertificateid&"' OR cartProdID='"&donationid&"') AND " & getsessionsql()
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF AND noemailgiftcertorders<>"" then exclemail=noemailgiftcertorders&"," else exclemail=""
	rs.close
	sSQL="SELECT payProvID,"&getlangid("PayProvShow",128)&" FROM payprovider WHERE payProvEnabled=1 AND payProvLevel<="&minloglevel&IIfVr(ordPayProvider="19" AND getget("token")<>""," AND payProvID=19 "," AND payProvID NOT IN ("&exclemail&"19,20,21" & IIfVs(paypalhostedsolution,",18") & ")") & " ORDER BY payProvOrder"
	rs.open sSQL,cnn,0,1
	alldata=""
	if NOT rs.EOF then alldata=rs.getrows else ppsuccess=FALSE
	rs.close
	if NOT IsArray(alldata) then %>
				<div class="cobhl cart2column ectwarning"><%=IIfVr(errormsg<>"",errormsg,xxNoPay)%></div>
<%	elseif UBOUND(alldata,2)=0 then
		print whv("payprovider", alldata(0,0))
		nodefaultpayprovider=FALSE
		payproviderradios=""
	else
		if payproviderradios=1 then
			showtitle=TRUE
			for rowcounter=0 to UBOUND(alldata,2)
				print "<div class=""checkoutadditionals copayradio1""><div class=""cobhl cdformpayradio1"">"
				print IIfVs(showtitle, "<div class=""cdformtpayradio1"">" & xxPlsChz & "</div>")
				print "<input type=""radio"" name=""payprovider"" id=""pprovid" & alldata(0,rowcounter) & """ class=""ectradio cdformradio cdformpayment"" value=""" & alldata(0,rowcounter) & """"
				if ordPayProvider=alldata(0,rowcounter) OR (ordPayProvider="" AND nodefaultpayprovider<>TRUE) then print " checked=""checked""" : ordPayProvider="-1"
				print " /></div><div class=""cobll cobll2 cdformpayment""><label class=""ectlabel cdformradio"" for=""pprovid" & alldata(0,rowcounter) & """>" & alldata(1,rowcounter) & IIfVs(alldata(0,rowcounter)=1,"<img src=""images/paypalacceptmark.gif"" style=""vertical-align:middle;padding-left:15px"" alt=""PayPal Payments"" />") & "</label></div>"
				print "</div>" & vbLf
				showtitle=FALSE
			next
		else %>
				<div class="checkoutadditionals checkoutadds2col"><div class="cobhl cobhl2 cdformtpayment"><%=xxPlsChz%></div>
				  <div class="cobll cobll2 cdformpayment"><%
			if payproviderradios=2 then
				print "<div class=""payprovider"">"
				for rowcounter=0 to UBOUND(alldata,2)
					print "<label class=""ectlabel cdformradio""><input type=""radio"" name=""payprovider"" class=""ectradio cdformradio cdformpayment"" value="""&alldata(0,rowcounter)&""""
					if ordPayProvider=alldata(0,rowcounter) OR (ordPayProvider="" AND nodefaultpayprovider<>TRUE) then print " checked=""checked""" : ordPayProvider="-1"
					print " />"&alldata(1,rowcounter)&IIfVs(alldata(0,rowcounter)=1,"<img src=""images/paypalacceptmark.gif"" alt=""PayPal Payments"" style=""vertical-align:bottom;margin-left:3px"" />") & "</label>"
				next
				print "</div>"
			else
				print "<select name=""payprovider"" class=""cdform2fixw ectselectinput cdformpayment"" size=""1"" onchange=""remwarning(this)"">"
				if nodefaultpayprovider=TRUE then print "<option value="""">"&xxPlsSel&"</option>"
				for rowcounter=0 to UBOUND(alldata,2)
					print "<option value='"&alldata(0,rowcounter)&"'" & IIfVs(ordPayProvider=alldata(0,rowcounter), " selected=""selected""") & ">"&alldata(1,rowcounter)&"</option>"&vbCrLf
				next
				print "</select>"
			end if %></div>
				</div>
<%		end if
	end if
	if ppsuccess then %>
				<div>
			      <div class="cobll cart2column checkoutbutton2"><%=imageorsubmit(imgcheckoutbutton2,IIfVr(xxCOTxt2<>"",xxCoTxt2,xxCOTxt),"widecheckout2")%></div>
				</div><%
	end if %> </div>
			</div>
		</form>
<script>
/* <![CDATA[ */
var allinputs=document.forms.mainform.getElementsByTagName('input');
for(var i=0; i<allinputs.length; i++){
	fieldtype=(allinputs[i].type?allinputs[i].type:'radio');
	if(fieldtype=='text'||fieldtype=='textarea'||fieldtype=='email'||fieldtype=='tel') allinputs[i].onkeyup=function(){if(this.value!='')remwarning(this);};
}
var globcurobj=[];
function doshowshipform(itm){
	var elem=document.getElementsByTagName('div');
	for(var i=0; i<elem.length; i++){
		var classes=elem[i].className;
		var issel=((itm==2?document.getElementById('shipdiff').value=='1':globcurobj[itm].selectedIndex!=0||globcurobj[itm].options.length<=1)?'':'none');
		if(classes.indexOf((itm==1?'bill':'ship')+'formrow')!=-1){
			if(elem[i].style.display!=issel){
				elem[i].style.display=issel;
				setTimeout("doshowshipform("+itm+");", 30);
				return;
			}
		}
	}
}
function showshipform(itm,curobj){
	globcurobj[itm]=curobj;
	doshowshipform(itm);
}
var checkedfullname=false,isshipcheckouterr=false;
var checkaddress=true,scheckaddress=true;
var checkouterrtxt='';
function chkextra(isship,ob,fldtxt){
	var hasselected=false,fieldtype='';
	if(ob){
		fieldtype=(ob.type?ob.type:'radio');
		if(fieldtype=='text'||fieldtype=='textarea'||fieldtype=='password'){
			hasselected=ob.value!='';
		}else if(fieldtype=='select-one'){
			hasselected=ob.selectedIndex!=0;
		}else if(fieldtype=='radio'){
			for(var ii=0;ii<ob.length;ii++)if(ob[ii].checked)hasselected=true;
		}else if(fieldtype=='checkbox')
			hasselected=ob.checked;
		if(!hasselected){
			if(checkouterrtxt==''){
				if(ob.focus)ob.focus();else ob[0].focus();
				if(isship){
					chkconfship(true,ob,"<%=jscheck(xxPlsEntr&" """&xxShpDet & " / ")%>"+fldtxt+"\".\n\n<%=jscheck(xxNoShip)%>");
				}else
					chkfocus(true,ob,"<%=jscheck(xxPlsEntr)%> \""+fldtxt+"\".");
				return(false);
			}
			ectaddclass(ob,'ectwarning');
		}else
			ectremoveclass(ob,'ectwarning');
	}else{
		alert("Invalid: " + fldtxt);
		return(false);
	}
	return(true);
}
function setdefs(ob,deftxt){
	var fieldtype='';
	if(ob)fieldtype=(ob.type?ob.type:'radio');<% if debugmode then print "else alert('Extra order field id not found');"%>
	if(fieldtype=='text'||fieldtype=='textarea'||fieldtype=='password'){
		ob.value=deftxt;
	}else if(fieldtype=='select-one'){
		for(var ii=0;ii<ob.length;ii++)if(ob[ii].value==deftxt)ob[ii].selected=true;
	}else if(fieldtype=='radio'){
		for(var ii=0;ii<ob.length;ii++)if(ob[ii].value==deftxt)ob[ii].checked=true;
	}else if(fieldtype=='checkbox'){
		if(ob.value==deftxt)ob.checked=true;
	}
}
function zipoptional(cntobj){
var cntid=cntobj[cntobj.selectedIndex].value;
if(cntid==0<%	for each objitem in zipoptional
					print "||cntid==" & objitem
				next %>)return true; else return false;
}
function stateoptional(cntobj){
var cntid=cntobj[cntobj.selectedIndex].value;
if(false<%
rs.open "SELECT countryID FROM countries WHERE countryEnabled<>0 AND loadStates<0",cnn,0,1
do while NOT rs.EOF
	print "||cntid==" & rs("countryID")
	rs.movenext
loop
rs.close
%>)return true; else return false;
}
<%	if NOT IsArray(addresses) then
		if trim(extraorderfield1)<>"" AND trim(ordExtra1)<>"" then print "setdefs(document.forms.mainform.ordextra1,'" & jsspecials(ordExtra1) & "');" & vbCrLf
		if trim(extraorderfield2)<>"" AND trim(ordExtra2)<>"" then print "setdefs(document.forms.mainform.ordextra2,'" & jsspecials(ordExtra2) & "');" & vbCrLf
		if noshipaddress<>TRUE then
			if trim(extraorderfield1)<>"" AND trim(ordShipExtra1)<>"" then print "setdefs(document.forms.mainform.ordsextra1,'" & jsspecials(ordShipExtra1) & "');" & vbCrLf
			if trim(extraorderfield2)<>"" AND trim(ordShipExtra2)<>"" then print "setdefs(document.forms.mainform.ordsextra2,'" & jsspecials(ordShipExtra1) & "');" & vbCrLf
		end if
	end if
	if trim(extracheckoutfield1)<>"" AND trim(ordCheckoutExtra1)<>"" then print "setdefs(document.forms.mainform.ordcheckoutextra1,'" & jsspecials(ordCheckoutExtra1) & "');" & vbCrLf
	if trim(extracheckoutfield2)<>"" AND trim(ordCheckoutExtra2)<>"" then print "setdefs(document.forms.mainform.ordcheckoutextra2,'" & jsspecials(ordCheckoutExtra2) & "');" & vbCrLf %>
function chkfocus(ttest,tobj,ttxt){
	if(ttest){
		if(checkouterrtxt==''){
			tobj.focus();
			checkouterrtxt=ttxt;
		}
		ectaddclass(tobj,'ectwarning');
	}else
		ectremoveclass(tobj,'ectwarning');
	return false;
}
function chkconfship(ttest,tobj,ttxt){
	if(ttest) isshipcheckouterr=true;
	return chkfocus(ttest,tobj,ttxt);
}
function remwarning(tobj){
	ectremoveclass(tobj,'ectwarning');
}
function checkform(frm){
	var cntelem=document.getElementById('country');
	var scntelem=document.getElementById('scountry');
	checkouterrtxt='';
if(checkaddress){
if(frm.country[frm.country.selectedIndex].value=='') return(chkfocus(true,frm.country,"<%=jscheck(xxPlsSlct & " " & xxCountry)%>"));
<%	if trim(extraorderfield1)<>"" AND extraorderfield1required then print "chkextra(false,frm.ordextra1,"""&jscheck(strip_tags2(extraorderfield1))&""");" & vbLf %>
chkfocus(frm.ordname.value=="",frm.ordname,"<%=jscheck(xxPlsEntr&" """&IIfVr(usefirstlastname, xxFirNam, xxName))%>\".");
<%	if usefirstlastname then %>
chkfocus(frm.lastname.value=="",frm.lastname,"<%=jscheck(xxPlsEntr&" """&xxLasNam)%>\".");
<%	end if
	if NOT IsArray(addresses) then %>
var regex=/[^@]+@[^@]+\.[a-z]{2,}$/i;
chkfocus(!regex.test(frm.email.value),frm.email,"<%=jscheck(xxValEm)%>");
<%		if verifyemail then %>
chkfocus(!regex.test(frm.email2.value),frm.email2,"<%=jscheck(xxEmVerf&"\n\n"&xxValEm)%>");
chkfocus(frm.email.value!=frm.email2.value,frm.email2,"<%=jscheck(xxEmNoMa)%>");
<%		end if
	end if %>
chkfocus(frm.address.value=="",frm.address,"<%=jscheck(xxPlsEntr&" """&xxAddress)%>\".");
chkfocus(frm.city.value=="",frm.city,"<%=jscheck(xxPlsEntr&" """&xxCity)%>\".");
	if(stateoptional(cntelem)){
	}else if(stateselectordisabled[0]==false){
<%	if hasstates then %>
		chkfocus(frm.state.selectedIndex==0,frm.state,"<%=jscheck(xxPlsSlct & " ")%>"+document.getElementById('statetxt').innerHTML);
<%	end if %>
	}else
		chkfocus(frm.state2.value=="",frm.state2,"<%=jscheck(xxPlsEntr)%> \""+document.getElementById('statetxt').innerHTML+"\".");
chkfocus(frm.zip.value==""&&!zipoptional(cntelem),frm.zip,"<%=jscheck(xxPlsEntr)%> \""+getziptext(cntelem[cntelem.selectedIndex].value)+"\".");
chkfocus(frm.phone.value=="",frm.phone,"<%=jscheck(xxPlsEntr&" """&xxPhone)%>\".");
<%	if trim(extraorderfield2)<>"" AND extraorderfield2required then print "chkextra(false,frm.ordextra2,"""&jscheck(strip_tags2(extraorderfield2))&""");" & vbLf %>
}
<% if abs(addshippinginsurance)=2 AND forceinsuranceselection then %>
chkfocus(frm.wantinsurance.selectedIndex==0,frm.wantinsurance,"<%=jscheck(strip_tags2(replace(xxChoIns,"<br>","\n")))%>");
<% end if
   if noshipaddress<>true then %>
var xxnoship="\n\n<%=jscheck(xxNoShip)%>";
if(scheckaddress<% if NOT isarray(addresses) then print "&&document.getElementById('shipdiff').value=='1'"%>){
	chkconfship(frm.scountry[frm.scountry.selectedIndex].value=='',frm.scountry,"<%=jscheck(xxPlsSlct&" """&xxShpDet&" / "&xxCountry)%>\"."+xxnoship);
<%	if trim(extraorderfield1)<>"" AND extraorderfield1required then print "if(!chkextra(true,frm.ordsextra1,"""&jscheck(strip_tags2(extraorderfield1))&""");" & vbLf %>
	chkconfship(frm.sordname.value=="",frm.sordname,"<%=jscheck(xxPlsEntr&" """&xxShpDet&" / "&xxName)%>\"."+xxnoship);
<%	if usefirstlastname then %>
	chkconfship(frm.slastname.value=="",frm.slastname,"<%=jscheck(xxPlsEntr&" """&xxLasNam)%>\"."+xxnoship);
<%	end if %>
	chkconfship(frm.saddress.value=="",frm.saddress,"<%=jscheck(xxPlsEntr&" """&xxShpDet&" / "&xxAddress)%>\"."+xxnoship);
	chkconfship(frm.scity.value=="",frm.scity,"<%=jscheck(xxPlsEntr&" """&xxShpDet&" / "&xxCity)%>\"."+xxnoship);
	if(stateoptional(scntelem)){
	}else if(stateselectordisabled[1]==false){
<%	if hasstates then %>
		chkconfship(frm.sstate.selectedIndex==0,frm.sstate,"<%=jscheck(xxPlsSlct&" """&xxShpDet&" / ")%>"+document.getElementById('sstatetxt').innerHTML+"\"."+xxnoship);
<%	end if %>
	}elseelse
		chkconfship(frm.sstate2.value=="",frm.sstate2,"<%=jscheck(xxPlsEntr&" """&xxShpDet&" / ")%>"+document.getElementById('sstatetxt').innerHTML+"\"."+xxnoship);
	chkconfship(frm.szip.value==""&&!zipoptional(scntelem),frm.szip,"<%=jscheck(xxPlsEntr&" """&xxShpDet&" / ")%>"+getziptext(scntelem[scntelem.selectedIndex].value)+"\"."+xxnoship);
<%	if trim(extraorderfield2)<>"" AND extraorderfield2required then print "chkextra(true,frm.ordsextra2,"""&jscheck(strip_tags2(extraorderfield2))&""");" & vbLf %>
}
<% end if
	if trim(extracheckoutfield1)<>"" AND extracheckoutfield1required then print "chkextra(false,frm.ordcheckoutextra1,"""&jscheck(strip_tags2(extracheckoutfield1))&""");" & vbLf
	if trim(extracheckoutfield2)<>"" AND extracheckoutfield2required then print "chkextra(false,frm.ordcheckoutextra2,"""&jscheck(strip_tags2(extracheckoutfield2))&""");" & vbLf
	if mailinglistdropdown then %>
chkfocus(frm.allowemail.selectedIndex==0,frm.allowemail,"<%=jscheck(xxPlSePE)%>");
<%	elseif mailinglistradios then %>
chkfocus(!(document.getElementById('allowemailradioy').checked||document.getElementById('allowemailradion').checked),document.getElementById('allowemailradioy'),"<%=jscheck(xxPlSePE)%>");
<%	end if
	if termsandconditions then %>
chkfocus(frm.license.checked==false,frm.license,"<%=jscheck(xxPlsProc)%>");
<%	end if
	if payproviderradios<>"" then %>
hasselected=false;
for(var ii=0;ii<frm.payprovider.length;ii++)if(frm.payprovider[ii].checked)hasselected=true;
chkfocus(!hasselected,frm.payprovider[0],"<%=jscheck(xxPlsEntr&" """&xxPlsChz)%>\".");
<%	elseif nodefaultpayprovider then %>
chkfocus(frm.payprovider.selectedIndex==0,frm.payprovider,"<%=jscheck(xxPlsEntr&" """&xxPlsChz)%>\".");
<%	end if %>
if(checkouterrtxt!=''){
	if(isshipcheckouterr){
		if(!confirm(checkouterrtxt)){
			document.getElementById('shipdiff').value='';
			showshipform(2,document.getElementById('scountry'));
			return true;
		}
	}else
		alert(checkouterrtxt);
	return false;
}
<%	if NOT usefirstlastname then %>
var regex=/ /;
if(checkaddress&&!checkedfullname&&!regex.test(frm.ordname.value)){
	alert("<%=jscheck(xxFulNam&" """&xxName)%>\".");
	frm.ordname.focus();
	checkedfullname=true;
	return(false);
}
<%	end if %>
return true;
}
<%	if termsandconditions then call gettermsjsfunction() %>
var savestate=0;
var ssavestate=0;
function applycertcallback(){
	if(ajaxobj.readyState==4){
		document.getElementById("cpncodespan").innerHTML=ajaxobj.responseText;
	}
}
function applycert(){
	cpncode=document.getElementById("cpncode").value;
	if(cpncode!=""){
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=applycertcallback;
		document.getElementById("cpncodespan").innerHTML="<%=xxAplyng%>...";
		ajaxobj.open("GET", "vsadmin/ajaxservice.asp?action=applycert&cpncode="+cpncode, true);
		ajaxobj.send(null);
	}
}
function removecert(cpncode){
	if(cpncode!=""){
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=applycertcallback;
		document.getElementById("cpncodespan").innerHTML="<%=xxDeltng%>...";
		ajaxobj.open("GET", "vsadmin/ajaxservice.asp?action=applycert&act=delete&cpncode="+cpncode, true);
		ajaxobj.send(null);
		document.getElementById("cpncode").value="";
	}
	return false;
}
function dosavestate(shp){
	thestate=eval('document.forms.mainform.'+shp+'state');
	eval(shp+'savestate=thestate.selectedIndex');
}
function checkoutspan(shp){
	document.getElementById(shp+'zipstar').style.display=(zipoptional(document.getElementById(shp+'country'))?'none':'');
	document.getElementById(shp+'statestar').style.display=(stateoptional(document.getElementById(shp+'country'))?'none':'');<%
	if hasstates then print "thestate=document.getElementById(shp+'state');"&vbCrLf
	print "dynamiccountries(document.getElementById(shp+'country'),shp);" & vbCrLf
	print "if(stateselectordisabled[shp=='s'?1:0]==false&&!stateoptional(document.getElementById(shp+'country'))){" & vbCrLf
	print "if(document.getElementById(shp+'state2'))document.getElementById(shp+'state2').style.display='none';document.getElementById('statetxt').htmlFor=shp+'state';"&vbCrLf
	if hasstates then
		print "thestate.disabled=false;"&vbCrLf
		print "eval('thestate.selectedIndex='+shp+'savestate');"&vbCrLf
		print "document.getElementById(shp+'state').style.display='';"&vbCrLf
	end if %>
}else{<%
	print "if(document.getElementById(shp+'state2')){document.getElementById(shp+'state2').style.display='';document.getElementById('statetxt').htmlFor=shp+'state2';}"&vbCrLf
	if hasstates then %>
		document.getElementById(shp+'state').style.display='none';
		if(thestate.disabled==false){
		thestate.disabled=true;
		eval(shp+'savestate=thestate.selectedIndex');
		thestate.selectedIndex=0;}
<%	end if %>
}}
<%	createdynamicstates("SELECT stateID,stateAbbrev,stateName,stateName2,stateName3,stateCountryID,countryName FROM states INNER JOIN countries ON states.stateCountryID=countries.countryID WHERE countryEnabled<>0 AND stateEnabled<>0 AND loadStates=2 ORDER BY stateCountryID," & getlangid("stateName",1048576))
	if IsArray(addresses) then print "checkaddress=false;scheckaddress=false;" & vbCrLf
	if IsArray(addresses) AND noshipaddress<>TRUE then print "checkeditbutton('s');"
	print "checkoutspan('');" & vbCrLf
	if noshipaddress<>TRUE then print "checkoutspan('s');" & vbCrLf
	print "setinitialstate('');setinitialstate('s');" & vbCrLf
	if NOT IsArray(addresses) then
		print "showshipform(1,document.getElementById('country'));" & vbCrLf
		if noshipaddress<>TRUE then print "showshipform(2,document.getElementById('scountry'));" & vbCrLf
	end if
	if SESSION("clientID")="" AND enableclientlogin AND allowclientregistration then %>
function co2newacctcallback(){
	if(liajaxobj.readyState==4){
		postdata="email=" + encodeURIComponent(document.getElementById('naemail').value) + "&pw=" + encodeURIComponent(document.getElementById('pass').value);
		document.getElementById('newacctpl').style.display='none';
		document.getElementById('newacctdiv').style.display='';
		if(liajaxobj.responseText.substr(0,7)=='SUCCESS'){
			document.getElementById('acopaquediv').style.display='none';
			document.getElementById('co2newacctbutton').innerHTML='';
			document.getElementById('co2newaccttxt').innerHTML='<%=jscheck(xxAccSuc)%>';
		}else{
			document.getElementById('accounterrordiv').innerHTML=liajaxobj.responseText.substr(6);
			var className=document.getElementById('accounterrordiv').className;
			if(className.indexOf('ectwarning')==-1)document.getElementById('accounterrordiv').className+=' ectwarning cartnewaccloginerror';
<%	if recaptchaenabled(8) then print "nacaptchaok=false;grecaptcha.reset(nacaptchawidgetid);" %>
		}
	}
}
function ecttrim(x) {
	return x.replace(/^\s+|\s+$/gm,'');
}
function co2displaynewaccount(){
	displaynewaccount();
	document.getElementById('naname').value=ecttrim(document.getElementById('name').value+(document.getElementById('lastname')?' '+document.getElementById('lastname').value:''));
	document.getElementById('naemail').value=document.getElementById('email').value;
}
<%	end if
	if googletagid<>"" then
		sSQL="SELECT SUM(cartProdPrice*cartQuantity) AS thePrice FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then totalgoods=rs("thePrice") else totalgoods=0
		rs.close
		sSQL="SELECT SUM(coPriceDiff*cartQuantity) AS thePrice FROM cart INNER JOIN cartoptions ON cart.cartID=cartoptions.coCartID WHERE cartCompleted=0 AND " & getsessionsql()
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT isnull(rs("thePrice")) then totalgoods=totalgoods + rs("thePrice")
		end if
		rs.close
		print "gtag(""event"",""add_shipping_info"",{ currency:'" & countryCurrency & "',value:" & totalgoods & ",items:[" & getcartforganalytics("") & "]});" & vbLf
	end if
%>/* ]]> */</script><%
elseif checkoutmode="authorize" then ' }{
	iframe="" : ordauthstatus=""
	blockuser=checkuserblock("")
	ordID=replace(getpost("ordernumber"),"'","")
	centinelenrolled="N"
	recaptchasuccess=TRUE
	if is_numeric(ordID) AND is_numeric(getpost("method")) then
		if getpayprovdetx(getpost("method"),data1,data2,data3,data4,data5,data6,ppflag1,ppflag2,ppflag3,ppbits,demomode,ppmethod) then
			sSQL="SELECT ordID,ordAuthStatus FROM orders WHERE ordID=" & ordID & " AND " & getordersessionsql()
			rs.open sSQL,cnn,0,1
			if rs.EOF then ordID=0 else ordauthstatus=rs("ordAuthStatus")
			rs.close
			centinelerror=SESSION("ErrorDesc")
			if getpost("method")="14" AND custompp3ds<>TRUE then cardinalprocessor=""
		else
			ordID=0
		end if
	else
		ordID=0
	end if
	if recaptchaenabled(1) AND NOT (getpost("method")="8" OR getpost("method")="22") then
		recaptchasuccess=checkrecaptcha(errormsg)
		if NOT recaptchasuccess then ordID=0 : vsRESPMSG="reCAPTCHA failure. If you inadvertently refreshed this page, please contact customer support.<br>" & errormsg
	end if
	if ordID<>0 AND ordauthstatus<>"MODWARNOPEN" AND cardinalprocessor<>"" AND cardinalmerchant<>"" AND cardinalpwd<>"" AND SESSION("centinelok")="" then
		cardnum=replace(getpost("ACCT"), " ", "")
		exmon=getpost("EXMON")
		exyear=getpost("EXYEAR")
		cardname=getpost("cardname")
		cvv2=getpost("CVV2")
		issuenum=getpost("IssNum")
		sSQL="SELECT ordID,ordName,ordLastName,ordCity,ordState,ordCountry,ordPhone,ordHandling,ordZip,ordEmail,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordTotal,ordDiscount,ordAddress,ordAddress2,ordIP,ordAuthNumber,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipCountry,ordShipZip FROM orders WHERE ordID=" & ordID
		rs.open sSQL,cnn,0,1
		sXML="<CardinalMPI>" & _
			addtag("Version","1.7") & addtag("MsgType","cmpi_lookup") & addtag("ProcessorId",cardinalprocessor) & addtag("MerchantId",cardinalmerchant) & addtag("TransactionPwd",cardinalpwd) & addtag("TransactionType","C") & _
			addtag("Amount",int(((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling")+0.001)-rs("ordDiscount"))*100)) & _
			addtag("CurrencyCode",countryNumCurrency) & addtag("OrderNumber",ordID) & addtag("OrderDescription","Order id " & ordID) & addtag("EMail",rs("ordEmail")) & _
			addtag("UserAgent",Request.ServerVariables("HTTP_USER_AGENT")) & addtag("BrowserHeader",Request.ServerVariables("HTTP_ACCEPT")) & addtag("IPAddress",REMOTE_ADDR) & _
			addtag("CardNumber",cardnum) & addtag("CardExpMonth",exmon) & addtag("CardExpYear",IIfVr(len(exyear)=2,"20","")&exyear) & _
			"</CardinalMPI>"
		rs.close
		theurl="https://"&IIfVr(getpost("method")="7" OR getpost("method")="18","paypal","centinel400")&".cardinalcommerce.com/maps/txns.asp"
		if cardinaltestmode then theurl="https://centineltest.cardinalcommerce.com/maps/txns.asp"
		if cardinalurl<>"" then theurl=cardinalurl
		if callxmlfunction(theurl, "cmpi_msg=" & urlencode(sXML), res, "", "WinHTTP.WinHTTPRequest.5.1", vsRESPMSG, 12) then
			set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
			xmlDoc.validateOnParse=FALSE
			xmlDoc.loadXML(res)
			set oNodeList=xmlDoc.documentElement.childNodes
			for i=0 To oNodeList.length - 1
				Set Item=oNodeList.item(i)
				if Item.nodeName="ACSUrl" then acsurl=Item.Text
				if Item.nodeName="Payload" then SESSION("cardinal_pareq")=Item.Text
				if Item.nodeName="Enrolled" then centinelenrolled=Item.Text : SESSION("centinel_enrolled")=centinelenrolled
				if Item.nodeName="OrderId" then SESSION("cardinal_orderid")=Item.Text
				if Item.nodeName="TransactionId" then SESSION("cardinal_transaction")=Item.Text
				if Item.nodeName="EciFlag" then SESSION("EciFlag")=Item.Text
				if Item.nodeName="ErrorDesc" then centinelerror=Item.Text
				if Item.nodeName="ErrorNo" then if Item.Text="1360" then centinelerror="" : exit for
			next
			set xmlDoc=nothing
			if centinelenrolled="Y" then
				SESSION("cardinal_method")=getpost("method")
				SESSION("cardinal_ordernum")=ordID
				SESSION("cardinal_sessionid")=thesessionid
				SESSION("cardinal_cardnum")=cardnum
				SESSION("cardinal_exmon")=getpost("EXMON")
				SESSION("cardinal_exyear")=getpost("EXYEAR")
				SESSION("cardinal_cardname")=getpost("cardname")
				SESSION("cardinal_cvv2")=getpost("CVV2")
				SESSION("cardinal_issnum")=getpost("IssNum")
				print "<div style=""font-weight:bold;padding:5px;margin:5px;text-align:center;"">" & xxComOrd & "<br><br>" & xxNoBack & "<br><br><iframe id=""centinelwin"" src=""vsadmin/ajaxservice.asp?action=centinel&url="&urlencode(acsurl)&""" width=""440"" height=""400"">Browser error.</iframe><br>&nbsp;</div>"
			end if
		end if
	elseif ordID<>0 AND SESSION("centinelok")="Y" then
		cardnum=SESSION("cardinal_cardnum")
		exmon=SESSION("cardinal_exmon")
		exyear=SESSION("cardinal_exyear")
		cardname=SESSION("cardinal_cardname")
		cvv2=SESSION("cardinal_cvv2")
		issuenum=SESSION("cardinal_issnum")
		SESSION("cardinal_cardnum")=empty
		SESSION("cardinal_exmon")=empty
		SESSION("cardinal_exyear")=empty
		SESSION("cardinal_cardname")=empty
		SESSION("cardinal_cvv2")=empty
		SESSION("cardinal_issnum")=empty
	elseif ordID<>0 then
		cardnum=replace(getpost("ACCT"), " ", "")
		exmon=getpost("EXMON")
		exyear=getpost("EXYEAR")
		cardname=getpost("cardname")
		cvv2=getpost("CVV2")
		issuenum=getpost("IssNum")
	end if
	if ordID=0 OR ordauthstatus="MODWARNOPEN" then
		if recaptchasuccess then
			if ordauthstatus="MODWARNOPEN" then vsRESPMSG=xxMoWnRC else vsRESPMSG="Error"
		end if
	elseif centinelenrolled="Y" then
		' Do Nothing
	elseif SESSION("centinelok")="N" OR centinelerror<>"" then
		vsRESPMSG=IIfVr(centinelerror<>"",centinelerror&"<br>","")&xx3DSFai
	elseif getpost("method")="7" OR getpost("method")="8" OR getpost("method")="22" then ' Payflow Pro / Payflow Link / PayPal Advanced
		if getpost("method")="7" then authorizeextraparams=authorizeextraparams7
		if getpost("method")="8" then authorizeextraparams=authorizeextraparams8
		if getpost("method")="22" then authorizeextraparams=authorizeextraparams22
		vsdetails=Split(data1, "&")
		if UBOUND(vsdetails)>0 then
			vs1=vsdetails(0)
			vs2=vsdetails(1)
			vs3=vsdetails(2)
			vs4=vsdetails(3)
		end if
		sSQL="SELECT ordName,ordLastName,ordZip,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordHandling,ordTotal,ordDiscount,ordAddress,ordAddress2,ordCity,ordState,ordCountry,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipCountry,ordShipZip,ordAuthNumber,ordEmail FROM orders WHERE ordID="&ordID
		rs.open sSQL,cnn,0,1
		vsAUTHCODE=(rs("ordAuthNumber")&"")
		sSQL="SELECT countryID,countryCode FROM countries WHERE countryName='" & escape_string(rs("ordCountry")) & "'"
		rs2.Open sSQL,cnn,0,1
			countryid=rs2("countryID")
			countryCode=rs2("countryCode")
			homecountry=(countryid=origCountryID)
		rs2.Close
		sSQL="SELECT countryCode FROM countries WHERE countryName='" & IIfVr(trim(rs("ordShipAddress"))<>"", escape_string(rs("ordShipCountry")), escape_string(rs("ordCountry"))) & "'"
		rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then shipCountryCode=rs2("countryCode")
		rs2.Close
		if trim(rs("ordShipAddress"))<>"" then isshp="Ship" else isshp=""
		ordState=rs("ordState")
		ordShipState=rs("ord" & isshp & "State")
		if (countryid=1 OR countryid=2) AND homecountry AND usestateabbrev<>TRUE then
			ordState=getstateabbrev(ordState)
			ordShipState=getstateabbrev(ordShipState)
		end if
		call splitname(IIfVr(getpost("method")="8" OR getpost("method")="22",trim(rs("ordName")&" "&rs("ordLastName")),cardname), firstname, lastname)
		call splitname(trim(rs("ord"&isshp&"Name")&" "&rs("ord"&isshp&"LastName")), shipfirstname, shiplastname)
		sXML="PARTNER="&vs3&"&VENDOR="&vs2&"&TRXTYPE="&IIfVr(ppmethod=1,"A","S")&"&TENDER=C&ZIP["&Len(rs("ordZip"))&"]="&rs("ordZip")&"&STREET["&len(rs("ordAddress"))&"]="&rs("ordAddress")& IIfVr(rs("ordAddress2")<>"", "&STREET2["&len(rs("ordAddress2"))&"]="&rs("ordAddress2"), "") & "&CITY["&len(rs("ordCity"))&"]="&rs("ordCity")&"&STATE["&len(ordState)&"]="&ordState&"&BILLTOCOUNTRY["&len(countryCode)&"]="&countryCode&"&FIRSTNAME["&len(firstname)&"]="&firstname&"&LASTNAME["&len(lastname)&"]="&lastname&"&EMAIL="&rs("ordEmail")
		sXML=sXML & "&SHIPTOZIP["&len(rs("ord"&isshp&"Zip"))&"]="&rs("ord"&isshp&"Zip")&"&SHIPTOSTREET["&len(rs("ord"&isshp&"Address"))&"]="&rs("ord"&isshp&"Address")& IIfVr(rs("ord"&isshp&"Address2")<>"", "&SHIPTOSTREET2["&len(rs("ord"&isshp&"Address2"))&"]="&rs("ord"&isshp&"Address2"), "") & "&SHIPTOCITY["&len(rs("ord"&isshp&"City"))&"]="&rs("ord"&isshp&"City")&"&SHIPTOSTATE["&len(ordShipState)&"]="&ordShipState&"&SHIPTOCOUNTRYCODE["&len(shipCountryCode)&"]="&shipCountryCode&"&SHIPTOCOUNTRY["&len(shipCountryCode)&"]="&shipCountryCode&"&SHIPTOFIRSTNAME["&len(shipfirstname)&"]="&shipfirstname&"&SHIPTOLASTNAME["&len(shiplastname)&"]="&shiplastname
		if issuenum<>"" then
			if len(issuenum)=2 then sXML=sXML & "&CARDISSUE=" & issuenum else sXML=sXML & "&CARDSTART=" & issuenum
		end if
		sXML=sXML & "&CUSTIP=" & ipv6to4(REMOTE_ADDR) & "&PWD=" & vs4 & "&USER=" & vs1 & "&CURRENCY=" & countryCurrency & "&AMT=" & FormatNumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0) & "&BUTTONSOURCE=EcommerceTemplatesUS_Cart_PPA" & authorizeextraparams
		rs.close
		if getpost("method")="8" OR getpost("method")="22" then
			securetokenid="ECTP" & calcmd5(ordID & timer() & adminSecret & vs4 & lastname & firstname)
			sXML=sXML & "&INVNUM=" & ordID & "&RETURNURL="&storeurlssl&"thanks"&extension&"&ERRORURL="&storeurlssl&"thanks"&extension&"&NOTIFYURL="&storeurlssl&"vsadmin/ppconfirm.asp&CREATESECURETOKEN=Y&TEMPLATE=MINLAYOUT&DISABLERECEIPT=TRUE&SECURETOKENID=" & securetokenid
			success=callxmlfunction("https://" & IIfVr(demomode, "pilot-", "") & "payflowpro.paypal.com", sXML, curString, "", "WinHTTP.WinHTTPRequest.5.1", errormsg, FALSE)
			resparr=split(curString,"&")
			for each objItem in resparr
				itemarr=split(objItem,"=")
				if itemarr(0)="SECURETOKEN" then SECURETOKEN=itemarr(1)
				if itemarr(0)="RESPMSG" then RESPMSG=itemarr(1)
			next
			if RESPMSG="Approved" then
				iframe="<iframe style=""border:none"" src=""https://payflowlink.paypal.com/?SECURETOKEN=" & SECURETOKEN & "&SECURETOKENID=" & securetokenid & IIfVs(demomode,"&MODE=test") & IIfVs(mobilebrowser,"&TEMPLATE=MOBILE") & """ width=""510"" height=""610""></iframe>"
				print iframe
			else
				vsRESPMSG=RESPMSG
			end if
		else
			sXML=sXML & "&COMMENT1="&ordID & "&ACCT=" & cardnum & "&CVV2="&cvv2&"&EXPDATE=" & exmon & right(exyear,2)
			if cardinalprocessor<>"" AND cardinalmerchant<>"" AND cardinalpwd<>"" then
				sXML=sXML & "&AUTHSTATUS3DS="&SESSION("PAResStatus") & "&MPIVENDOR3DS=" & SESSION("centinel_enrolled") & "&CAVV=" & SESSION("Cavv") & "&ECI=" & SESSION("EciFlag") & "&XID=" & SESSION("Xid")
			end if
			if vsAUTHCODE="" then
				success=TRUE
				if blockuser then
					success=FALSE
					vsRESPMSG=multipurchaseblockmessage
				else
					randomize
					xmlfnheaders=array(array("X-VPS-REQUEST-ID",ordID&"."&(int(1000000 * Rnd) + 1000000)))
					success=callxmlfunction("https://" & IIfVr(demomode, "pilot-", "") & "payflowpro.paypal.com", sXML, curString, "", "WinHTTP.WinHTTPRequest.5.1", cferr, FALSE)
					if success then
						do while Len(curString)<>0
							if InStr(curString,"&") then varString=Left(curString, InStr(curString , "&") -1) else varString=curString
							name=Left(varString, InStr(varString, "=" ) -1)
							value=Right(varString, Len(varString) - (Len(name)+1))
							if name="RESULT" then
								vsRESULT=value
							elseif name="PNREF" OR name="PPREF" then
								vsTRANSID=value
							elseif name="RESPMSG" then
								vsRESPMSG=value
							elseif name="AUTHCODE" then
								vsAUTHCODE=value
							elseif name="AVSADDR" OR name="AVSCODE" then
								vsAVSADDR=value
							elseif name="AVSZIP" then
								vsAVSZIP=value
							elseif name="IAVS" then
								vsIAVS=value
							elseif name="CVV2MATCH" then
								vsCVV2=value
							elseif name="ACK" then
								if value="Success" then vsRESULT="0" : vsRESPMSG=xxTranAp else vsRESULT=""
							elseif name="L_LONGMESSAGE0" then
								vsRESPMSG=urldecode(value&" ")
							elseif name="L_ERRORCODE0" then
								vsERRCODE=value
							elseif name="DUPLICATE" then
								if value="1" then vsRESPMSG="DUPLICATE" : success=FALSE : vsRESULT=""
							end if
							if Len(curString)<>Len(varString) then curString=Right(curString, Len(curString) - (Len(varString)+1)) else curString=""
						loop
					else
						vsRESPMSG=cferr
					end if
				end if
				if success then
					if vsRESULT="0" OR vsRESULT="126" then
						if vsRESULT="126" then underreview="Fraud Review:<br>" : vsRESPMSG="Approved" else underreview=""
						ect_query("UPDATE cart SET cartDateAdded=" & vsusdate(DateAdd("h",dateadjust,Now()))&",cartCompleted=1 WHERE cartCompleted<>1 AND cartOrderID="&ordID)
						ect_query("UPDATE orders SET ordStatus=3,ordAuthStatus='',ordAVS='"&replace(vsAVSADDR&vsAVSZIP,"'","")&"',ordCVV='"&replace(vsCVV2,"'","")&"',ordAuthNumber='"&replace(underreview&vsAUTHCODE,"'","")&"',ordTransID='"&vsTRANSID&"',ordDate="&vsusdatetime(DateAdd("h",dateadjust,Now()))&",ordPrivateStatus='' WHERE ordAuthNumber='' AND ordID="&ordID)
						call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
						vsRESULT="0"
					elseif vsERRCODE<>"" then
						vsERRCODE=int(vsRESULT)
						ect_query("UPDATE orders SET ordPrivateStatus='"&escape_string(strip_tags2("(" & vsERRCODE & ") " & vsRESPMSG))&"' WHERE ordAuthNumber='' AND ordID="&ordID)
						if vsERRCODE=12 OR vsERRCODE=24 OR vsERRCODE=114 then vsRESPMSG=xxCCErro
					end if
				end if
				set client=nothing
			else
				vsRESULT="0"
				vsRESPMSG="Approved"
			end if
		end if
	elseif getpost("method")="13" then ' Auth.net AIM
		acceptecheck=(ppbits AND 1)=1
		if secretword<>"" then
			data1=upsdecode(data1, secretword)
			data2=upsdecode(data2, secretword)
		end if
		sSQL="SELECT ordID,ordStatus,ordCity,ordState,ordCountry,ordPhone,ordHandling,ordZip,ordEmail,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordTotal,ordDiscount,ordAddress,ordAddress2,ordIP,ordAuthNumber,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipCountry,ordShipZip FROM orders WHERE ordID="&ordID
		rs.open sSQL,cnn,0,1
		vsAUTHCODE=trim(rs("ordAuthNumber")&"")
		ordstatus=rs("ordStatus")
		gotvalidresult=FALSE
		vsRESULT=-1
		saveerrcode=""
		if vsAUTHCODE="" OR ordstatus<3 then
			if authnetemulateurl<>"" then
				sXML="x_version=3.1&x_delim_data=TRUE&x_relay_response=FALSE&x_delim_char=|&x_duplicate_window=15&x_solution_id=AAA172582" & _
					"&x_login="&data1&"&x_tran_key="&data2&IIfVs(SESSION("clientID")<>"","&x_cust_id="&SESSION("clientID"))&"&x_invoice_num="&rs("ordID") & _
					"&x_amount="&FormatNumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0) & _
					"&x_currency_code="&countryCurrency&"&x_description=" & left(urlencode(replace(getpost("description"),"&quot;","""")),254)
				if getpost("accountnum")<>"" then
					sXML=sXML & "&x_method=ECHECK&x_echeck_type=WEB&x_recurring_billing=NO" & _
						"&x_bank_acct_name=" & urlencode(getpost("accountname")) & "&x_bank_acct_num=" & urlencode(getpost("accountnum")) & _
						"&x_bank_name=" & urlencode(getpost("bankname")) & "&x_bank_aba_code=" & urlencode(getpost("routenumber")) & _
						"&x_bank_acct_type=" & urlencode(getpost("accounttype")) & "&x_type=AUTH_CAPTURE"
					if wellsfargo=TRUE then
						sXML=sXML & "&x_customer_organization_type=" & getpost("orgtype")
						if getpost("taxid")<>"" then
							sXML=sXML & "&x_customer_tax_id=" & urlencode(getpost("taxid"))
						else
							sXML=sXML & "&x_drivers_license_num=" & urlencode(getpost("licensenumber")) & "&x_drivers_license_state=" & urlencode(getpost("licensestate")) & "&x_drivers_license_dob=" & urlencode(getpost("dldobyear") & "/" & getpost("dldobmon") & "/" & getpost("dldobday"))
						end if
					end if
				else
					sXML=sXML & "&x_method=CC&x_card_num=" & urlencode(cardnum) & "&x_exp_date=" & exmon & right(exyear,2)
					if cvv2<>"" then sXML=sXML & "&x_card_code=" & urlencode(cvv2)
					if ppmethod=1 then sXML=sXML & "&x_type=AUTH_ONLY" else sXML=sXML & "&x_type=AUTH_CAPTURE"
				end if
				if cardinalprocessor<>"" AND cardinalmerchant<>"" AND cardinalpwd<>"" then	
					sXML=sXML & "&x_cardholder_authentication_value=" & urldecode(SESSION("Cavv")) & "&x_authentication_indicator=" & SESSION("EciFlag")
				end if
				if cardname<>"" then
					if InStr(cardname," ")>0 then
						namearr=Split(cardname," ",2)
						sXML=sXML & "&x_first_name=" & urlencode(namearr(0)) & "&x_last_name=" & urlencode(namearr(1))
					else
						sXML=sXML & "&x_last_name=" & urlencode(cardname)
					end if
				end if
				sXML=sXML & "&x_address="&urlencode(rs("ordAddress"))
				if trim(rs("ordAddress2")&"")<>"" then sXML=sXML & urlencode(", "&rs("ordAddress2"))
				sXML=sXML & "&x_city="&urlencode(rs("ordCity")) & "&x_state="&urlencode(rs("ordState")) & "&x_zip="&urlencode(rs("ordZip")) & "&x_country="&urlencode(rs("ordCountry")) & "&x_phone="&urlencode(rs("ordPhone")) & "&x_email="&urlencode(rs("ordEmail"))
				if trim(rs("ordShipName")&"")<>"" OR trim(rs("ordShipLastName")&"")<>"" OR rs("ordShipAddress")<>"" then
					if usefirstlastname then
						sXML=sXML & "&x_ship_to_first_name=" & urlencode(rs("ordShipName")) & "&x_ship_to_last_name=" & urlencode(rs("ordShipLastName"))
					elseif InStr(trim(rs("ordShipName")&"")," ")>0 then
						namearr=Split(trim(rs("ordShipName")&"")," ",2)
						sXML=sXML & "&x_ship_to_first_name=" & urlencode(namearr(0)) & "&x_ship_to_last_name=" & urlencode(namearr(1))
					else
						sXML=sXML & "&x_ship_to_last_name=" & urlencode(rs("ordShipName"))
					end if
					sXML=sXML & "&x_ship_to_address="&urlencode(rs("ordShipAddress"))
					if trim(rs("ordShipAddress2")&"")<>"" then sXML=sXML & urlencode(", "&rs("ordShipAddress2"))
					sXML=sXML & "&x_ship_to_city="&urlencode(rs("ordShipCity")) & "&x_ship_to_state="&urlencode(rs("ordShipState")) & "&x_ship_to_zip="&urlencode(rs("ordShipZip")) & "&x_ship_to_country="&urlencode(rs("ordShipCountry"))
				end if
				if trim(rs("ordIP"))<>"" then sXML=sXML & "&x_customer_ip="&urlencode(rs("ordIP"))
				if demomode then sXML=sXML & "&x_test_request=TRUE"
				success=TRUE
				if blockuser then
					success=FALSE
					vsRESPMSG=multipurchaseblockmessage
				else
					if authnetemulateurl="" then authnetemulateurl="https://secure.authorize.net/gateway/transact.dll"
					if callxmlfunction(authnetemulateurl, sXML & authorizeextraparams13, res, "", "Msxml2.ServerXMLHTTP", vsRESPMSG, FALSE) then
						varString=split(res, "|")
						if UBOUND(varString)<38 then
							vsRESPMSG="Invalid response: " & res
						else
							vsRESULT=varString(0)
							vsERRCODE=int(varString(2))
							vsRESPMSG=varString(3)
							vsAUTHCODE=varString(4)
							vsAVSADDR=varString(5)
							vsTRANSID=varString(6)
							vsCVV2=varString(38)
							gotvalidresult=TRUE
						end if
					end if
				end if
			else
				call splitname(cardname, firstname, lastname)
				sSQL="SELECT countryID,countryCode3,loadStates FROM countries WHERE countryName='" & escape_string(rs("ordCountry")) & "'"
				rs2.open sSQL,cnn,0,1
				if NOT rs2.EOF then
					countryid=rs2("countryID")
					countryCode3=rs2("countryCode3")
					homecountry=(rs2("countryID")=origCountryID)
				end if
				rs2.close
				sjson="{""createTransactionRequest"":{" & _
					"""merchantAuthentication"":{""name"":" & json_encode(data1) & ",""transactionKey"":" & json_encode(data2) & "}," & _
					"""transactionRequest"":{" & _
						"""transactionType"":""auth" & IIfVr(ppmethod=1,"Only","Capture") & "Transaction""," & _
						"""amount"":""" & FormatNumber((rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount"),2,-1,0,0) & ""","
				if getpost("accountnum")<>"" AND acceptecheck then
					sjson=sjson&"""payment"":{""bankAccount"":{""accountType"":" & json_encode(getpost("accounttype")) & ",""routingNumber"":" & json_encode(getpost("routenumber")) & ",""accountNumber"":" & json_encode(getpost("accountnum")) & ",""nameOnAccount"":" & json_encode(getpost("accountname")) & ",""echeckType"":" & json_encode("WEB") & ",""bankName"":" & json_encode(getpost("bankname")) & "}},"
				else
					sjson=sjson&"""payment"":{""creditCard"":{""cardNumber"":" & json_encode(cardnum) & ",""expirationDate"":" & json_encode(exyear & "-" & exmon) & ",""cardCode"":" & json_encode(cvv2) & "}},"
				end if
				sjson=sjson&IIfVs(NOT demomode,"""solution"":{""id"":""AAA172582""},") & _
					"""order"":{""invoiceNumber"":" & json_encode(ordID) & ",""description"":" & json_encode(left(getpost("description"),254)) & "}," & _
					"""tax"":{""amount"":" & json_encode(rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")) & "}," & _
					"""shipping"":{""amount"":" & json_encode(rs("ordShipping") + IIfVr(combineshippinghandling,rs("ordHandling"),0)) & "}," & _
					"""customer"":{""email"":" & json_encode(rs("ordEmail")) & "}," & _
					"""billTo"":{" & _
						"""firstName"":" & json_encode(firstname) & ",""lastName"":" & json_encode(lastname) & "," & _
						"""address"":" & json_encode(rs("ordAddress")) & ",""city"":" & json_encode(rs("ordCity")) & ",""state"":" & json_encode(IIfVr(countryid=1 AND homecountry AND NOT usestateabbrev,getstateabbrev(rs("ordState")),rs("ordState"))) & ",""zip"":" & json_encode(rs("ordZip")) & ",""country"":" & json_encode(countryCode3) & ",""phoneNumber"":" & json_encode(rs("ordPhone")) & _
					"},"
				if trim(rs("ordShipName"))<>"" OR trim(rs("ordShipLastName"))<>"" OR trim(rs("ordShipAddress"))<>"" then
					sSQL="SELECT countryID,countryCode3,loadStates FROM countries WHERE countryName='" & escape_string(rs("ordShipCountry")) & "'"
					rs2.open sSQL,cnn,0,1
					if NOT rs2.EOF then
						shipcountryid=rs2("countryID")
						shipCountryCode3=rs2("countryCode3")
						shiphomecountry=(rs2("countryID")=origCountryID)
					end if
					rs2.close
					firstname=rs("ordShipName") : lastname=rs("ordShipLastName")
					if NOT usefirstlastname then call splitname(rs("ordShipName"),firstname,lastname)
					sjson=sjson&"""shipTo"":{" & _
						"""firstName"":" & json_encode(firstname) & ",""lastName"":" & json_encode(lastname) & "," & _
						"""address"":" & json_encode(rs("ordShipAddress")) & ",""city"":" & json_encode(rs("ordShipCity")) & ",""state"":" & json_encode(IIfVr(shipcountryid=1 AND shiphomecountry AND NOT usestateabbrev,getstateabbrev(rs("ordShipState")),rs("ordShipState"))) & ",""zip"":" & json_encode(rs("ordShipZip")) & ",""country"":" & json_encode(shipCountryCode3) & _
					"},"
				end if
				sjson=sjson&"""customerIP"":" & json_encode(ipv6to4(rs("ordIP"))) & ",""transactionSettings"":{""setting"":[{""settingName"":""duplicateWindow"",""settingValue"":""60""}]}" & _
					"}}}"
				success=callxmlfunction("https://api" & IIfVs(demomode,"test") & ".authorize.net/xml/v1/request.api",sjson,jres,"","Msxml2.ServerXMLHTTP",vsRESPMSG,FALSE)
				if success then
					vsRESULT=get_json_val(jres,"responseCode","")
					if NOT is_numeric(vsRESULT) then vsRESULT=-1 else vsRESULT=int(vsRESULT)
					vsERRCODE=0
					if instr(jres,"""errors""") then
						vsERRCODE=get_json_val(jres,"errorCode","")
						vsRESPMSG=get_json_val(jres,"errorText","")
						vsRESULT=-1
					elseif vsRESULT=4 then
						vsERRCODE=get_json_val(jres,"code","messages")
						vsRESPMSG=get_json_val(jres,"text","messages")
					elseif vsRESULT=-1 then
						vsERRCODE=get_json_val(jres,"code","message")
						vsRESPMSG=get_json_val(jres,"text","message")
					else
						vsRESPMSG=get_json_val(jres,"description","")
					end if
					if is_numeric(vsERRCODE) then vsERRCODE=int(vsERRCODE) else saveerrcode=vsERRCODE : vsERRCODE=0
					vsAUTHCODE=get_json_val(jres,"authCode","")
					vsAVSADDR=get_json_val(jres,"avsResultCode","")
					vsTRANSID=get_json_val(jres,"transId","")
					vsCVV2=get_json_val(jres,"cvvResultCode","")
					gotvalidresult=TRUE
				end if
			end if
		else
			vsRESULT="0"
			vsRESPMSG=xxTranAp
			if InStr(vsAUTHCODE,"-")>0 then vsAUTHCODE=Right(vsAUTHCODE,Len(vsAUTHCODE)-InStr(vsAUTHCODE,"-"))
		end if
		rs.close
		if gotvalidresult then
			if int(vsRESULT)=1 OR vsERRCODE=253 then
				if vsERRCODE=253 then pendingreason="Pending: FDS" else pendingreason=""
				if getpost("accountnum")<>"" then vsAUTHCODE="eCheck"
				vsRESULT="0" ' Keep in sync with Payflow Pro
				ect_query("UPDATE cart SET cartDateAdded=" & vsusdate(DateAdd("h",dateadjust,Now()))&",cartCompleted=1 WHERE cartCompleted<>1 AND cartOrderID="&ordID)
				ect_query("UPDATE orders SET ordStatus=3,ordAuthStatus='"&pendingreason&"',ordAVS='"&vsAVSADDR&"',ordCVV='"&vsCVV2&"',ordAuthNumber='"&vsAUTHCODE&"',ordTransID='"&vsTRANSID&"',ordDate="&vsusdatetime(DateAdd("h",dateadjust,Now()))&",ordPrivateStatus='' WHERE ordAuthNumber='' AND ordID="&ordID)
				call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
			elseif vsERRCODE=252 then
				ect_query("UPDATE orders SET ordPrivateStatus='"&escape_string(strip_tags2("(" & vsERRCODE & ") " & vsRESPMSG))&"' WHERE ordAuthNumber='' AND ordID="&ordID)
				ect_query("UPDATE cart SET cartDateAdded=" & vsusdate(DateAdd("h",dateadjust,Now()))&",cartCompleted=2 WHERE cartCompleted=0 AND cartOrderID="&ordID)
				ect_query("UPDATE orders SET ordAuthNumber='FDS Review',ordDate="&vsusdatetime(DateAdd("h",dateadjust,Now()))&" WHERE ordAuthNumber='' AND ordID="&ordID)
				vsRESPMSG=xxAuNetR
			elseif vsERRCODE=27 OR vsERRCODE=127 then
				ect_query("UPDATE orders SET ordPrivateStatus='"&escape_string(strip_tags2("(" & vsERRCODE & ") " & vsRESPMSG))&"' WHERE ordAuthNumber='' AND ordID="&ordID)
				isavsmismatch=TRUE
			elseif vsERRCODE=6 OR vsERRCODE=7 OR vsERRCODE=8 OR vsERRCODE=78 then
				ect_query("UPDATE orders SET ordPrivateStatus='"&escape_string(strip_tags2("(" & vsERRCODE & ") " & vsRESPMSG))&"' WHERE ordAuthNumber='' AND ordID="&ordID)
				vsRESPMSG=xxCCErro
			end if
		end if
		if saveerrcode<>"" then vsERRCODE=saveerrcode
	elseif getpost("method")="14" then ' Custom Payment Processor
		call retrieveorderdetails(ordID, thesessionid)
		sSQL="SELECT ordID,ordHandling,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordTotal,ordDiscount,ordIP,ordAuthNumber FROM orders WHERE ordID=" & ordID
		rs.open sSQL,cnn,0,1
		ordShipping=rs("ordShipping")
		ordStateTax=rs("ordStateTax")
		ordCountryTax=rs("ordCountryTax")
		ordHSTTax=rs("ordHSTTax")
		ordTotal=rs("ordTotal")
		ordHandling=rs("ordHandling")
		ordDiscount=rs("ordDiscount")
		ordIP=rs("ordIP")
		ordAuthNumber=trim(rs("ordAuthNumber")&"")
		vsAUTHCODE=ordAuthNumber
		rs.close
		grandtotal=(ordShipping+ordStateTax+ordCountryTax+ordHSTTax+ordTotal+ordHandling)-ordDiscount
		if vsAUTHCODE="" then
%>
<!--#include file="customppreturn.asp"-->
<%		else
			vsRESULT="0"
			vsRESPMSG=xxTranAp
		end if
	elseif getpost("method")="18" then ' PayPal Direct
		on error resume next
		Server.ScriptTimeout=120
		on error goto 0
		sSQL="SELECT ordID,ordName,ordLastName,ordCity,ordState,ordCountry,ordPhone,ordHandling,ordZip,ordEmail,ordShipping,ordStateTax,ordCountryTax,ordHSTTax,ordTotal,ordDiscount,ordAddress,ordAddress2,ordIP,ordAuthNumber,ordShipName,ordShipLastName,ordShipAddress,ordShipAddress2,ordShipCity,ordShipState,ordShipCountry,ordShipZip FROM orders WHERE ordID=" & ordID
		rs.open sSQL,cnn,0,1
		ordState=rs("ordState")
		ordShipState=rs("ordShipState")
		grandtotal=(rs("ordShipping")+rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax")+rs("ordTotal")+rs("ordHandling"))-rs("ordDiscount")
		if data2<>"" then data2=urldecode(split(data2,"&")(0))
		if instr(data1,"/")>0 then
			data1arr=split(data1,"/")
			if grandtotal<12 then data1=trim(data1arr(1)) else data1=trim(data1arr(0))
			data1arr=split(data2,"/")
			if grandtotal<12 AND ubound(data1arr)>0 then data2=trim(data1arr(1)) else data2=trim(data1arr(0))
			data1arr=split(data3,"/")
			if grandtotal<12 AND ubound(data1arr)>0 then data3=trim(data1arr(1)) else data3=trim(data1arr(0))
		end if
		sSQL="SELECT countryCode FROM countries WHERE countryName='" & escape_string(rs("ordCountry")) & "'"
		rs2.Open sSQL,cnn,0,1
			countryCode=rs2("countryCode")
		rs2.Close
		sSQL="SELECT countryCode FROM countries WHERE countryName='" & escape_string(rs("ordShipCountry")) & "'"
		rs2.Open sSQL,cnn,0,1
			if NOT rs2.EOF then shipCountryCode=rs2("countryCode")
		rs2.Close
		if countryCode="US" OR countryCode="CA" then
			sSQL="SELECT stateAbbrev FROM states WHERE (stateCountryID=1 OR stateCountryID=2) AND stateName='" & escape_string(ordState) & "'"
			rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF then ordState=rs2("stateAbbrev")
			rs2.Close
		end if
		if shipCountryCode="US" OR shipCountryCode="CA" then
			sSQL="SELECT stateAbbrev FROM states WHERE (stateCountryID=1 OR stateCountryID=2) AND stateName='" & escape_string(ordShipState) & "'"
			rs2.Open sSQL,cnn,0,1
				if NOT rs2.EOF then ordShipState=rs2("stateAbbrev")
			rs2.Close
		end if
		vsAUTHCODE=trim(rs("ordAuthNumber")&"")
		call splitname(cardname, firstname, lastname)
		cardtype=getcctypefromnum(cardnum)
		data2hash=data3
		session.lcid=1033
		if trim(rs("ordShipAddress"))<>"" then doship="Ship" else doship=""
		sXML=ppsoapheader(data1, data2, data2hash) & _
			"<soap:Body><DoDirectPaymentReq xmlns=""urn:ebay:api:PayPalAPI"">" & _
			"<DoDirectPaymentRequest><Version xmlns=""urn:ebay:apis:eBLBaseComponents"">60.0</Version>" & _
			"  <DoDirectPaymentRequestDetails xmlns=""urn:ebay:apis:eBLBaseComponents"">" & _
			addtag("PaymentAction",IIfVr(ppmethod=1, "Authorization", "Sale")) & _
			"    <PaymentDetails>" & _
			"      <OrderTotal currencyID=""" & countryCurrency & """>" & FormatNumber(grandtotal,getDPs(countryCurrency),-1,0,0) & "</OrderTotal>"
			if rs("ordTotal")>=rs("ordDiscount") then
				sXML=sXML&"<ItemTotal currencyID=""" & countryCurrency & """>" & FormatNumber(rs("ordTotal")-rs("ordDiscount"),getDPs(countryCurrency),-1,0,0) & "</ItemTotal>" & _
					"<ShippingTotal currencyID=""" & countryCurrency & """>" & FormatNumber(rs("ordShipping"),getDPs(countryCurrency),-1,0,0) & "</ShippingTotal>" & _
					"<HandlingTotal currencyID=""" & countryCurrency & """>" & FormatNumber(rs("ordHandling"),getDPs(countryCurrency),-1,0,0) & "</HandlingTotal>" & _
					"<TaxTotal currencyID=""" & countryCurrency & """>" & FormatNumber(rs("ordStateTax")+rs("ordCountryTax")+rs("ordHSTTax"),getDPs(countryCurrency),-1,0,0) & "</TaxTotal>"
			end if
			sXML=sXML&addtag("ButtonSource","ecommercetemplates_Cart_DP_US") & _
			addtag("NotifyURL",storeurl & "vsadmin/ppconfirm.asp") & _
			addtag("Custom",ordID) & _
			"      <ShipToAddress><Name>" & vrxmlencode(trim(rs("ord"&doship&"Name")&" "&rs("ord"&doship&"LastName"))) & "</Name><Street1>" & vrxmlencode(rs("ord"&doship&"Address")) & "</Street1><Street2>" & vrxmlencode(rs("ord"&doship&"Address2")) & "</Street2><CityName>" & rs("ord"&doship&"City") & "</CityName><StateOrProvince>" & IIfVr(doship<>"", ordShipState, ordState) & "</StateOrProvince><Country>" & IIfVr(doship<>"", shipCountryCode, countryCode) & "</Country><PostalCode>" & rs("ord"&doship&"Zip") & "</PostalCode></ShipToAddress>" & _
			"    </PaymentDetails><CreditCard>" & _
			addtag("CreditCardType",cardtype) & addtag("CreditCardNumber",vrxmlencode(cardnum)) & addtag("ExpMonth",exmon) & addtag("ExpYear",exyear) & _
			"      <CardOwner>" & _
			addtag("Payer",vrxmlencode(rs("ordEmail"))) & _
			"<PayerName>" & addtag("FirstName",firstname) & addtag("LastName",lastname) & "</PayerName>" & addtag("PayerCountry",countryCode) & _
			"        <Address>" & addtag("Street1",vrxmlencode(rs("ordAddress"))) & addtag("Street2",vrxmlencode(rs("ordAddress2"))) & addtag("CityName",rs("ordCity")) & addtag("StateOrProvince",ordState) & addtag("Country",countryCode) & addtag("PostalCode",rs("ordZip")) & "</Address>" & _
			"      </CardOwner>" & _
			addtag("CVV2",cvv2)&""
		if issuenum<>"" then
			if len(issuenum)=2 then sXML=sXML & addtag("IssueNumber",issuenum) else sXML=sXML & addtag("StartMonth",left(issuenum,2)) & addtag("StartYear",right(issuenum,2))
		end if
		if cardinalprocessor<>"" AND cardinalmerchant<>"" AND cardinalpwd<>"" then
			sXML=sXML & "<ThreeDSecureRequest>" & addtag("AuthStatus3ds",SESSION("PAResStatus")) & addtag("MpiVendor3ds",SESSION("centinel_enrolled")) & addtag("Cavv",SESSION("Cavv")) & addtag("Eci3ds",SESSION("EciFlag")) & addtag("Xid",SESSION("Xid")) & "</ThreeDSecureRequest>"
		end if
		sXML=sXML & "</CreditCard>" & _
			addtag("IPAddress",trim(rs("ordIP"))) & addtag("MerchantSessionId",rs("ordID")) & _
			"  </DoDirectPaymentRequestDetails>" & _
			"</DoDirectPaymentRequest></DoDirectPaymentReq></soap:Body></soap:Envelope>"
		session.lcid=saveLCID
		rs.close
		if demomode then sandbox=".sandbox" else sandbox=""
		vsRESULT="-1"
		if vsAUTHCODE="" then
			if blockuser then
				success=FALSE
				vsRESPMSG=multipurchaseblockmessage
			else
				success=callxmlfunction("https://api" & IIfVs(data2hash<>"","-3t") & sandbox & ".paypal.com/2.0/", sXML, res, IIfVr(data2hash<>"","",data1), "WinHTTP.WinHTTPRequest.5.1", vsRESPMSG, TRUE)
			end if
			if success then
				vsAUTHCODE="":vsERRCODE="":vsRESPMSG="":vsAVSADDR="":vsTRANSID="":vsCVV2=""
				set xmlDoc=Server.CreateObject("MSXML2.DOMDocument")
				xmlDoc.validateOnParse=FALSE
				xmlDoc.loadXML (res)
				Set nodeList=xmlDoc.getElementsByTagName("SOAP-ENV:Body")
				Set n=nodeList.Item(0)
				for j=0 to n.childNodes.length - 1
					Set e9=n.childNodes.Item(j)
					if e9.nodeName="DoDirectPaymentResponse" then
						for k9=0 To e9.childNodes.length - 1
							Set t=e9.childNodes.Item(k9)
							if t.nodeName="Ack" then
								if t.firstChild.nodeValue="Success" OR t.firstChild.nodeValue="SuccessWithWarning" then
									vsRESULT=1
									vsRESPMSG=xxTranAp
								end if
							elseif t.nodeName="TransactionID" then
								if t.hasChildNodes then vsAUTHCODE=t.firstChild.nodeValue
							elseif t.nodeName="AVSCode" then
								if t.hasChildNodes then vsAVSADDR=t.firstChild.nodeValue
							elseif t.nodeName="CVV2Code" then
								if t.hasChildNodes then vsCVV2=t.firstChild.nodeValue
							elseif t.nodeName="Errors" then
								shortmsg="" : themsg="" : thecode=""
								iswarning=FALSE
								set ff=t.childNodes
								for kk=0 to ff.length - 1
									set gg=ff.item(kk)
									if gg.nodeName="ShortMessage" then
										if gg.hasChildNodes then shortmsg=gg.firstChild.nodeValue
									elseif gg.nodeName="LongMessage" then
										if gg.hasChildNodes then themsg=gg.firstChild.nodeValue
									elseif gg.nodeName="ErrorCode" then
										if gg.hasChildNodes then thecode=gg.firstChild.nodeValue
									elseif gg.nodeName="SeverityCode" then
										if gg.hasChildNodes then iswarning=(gg.firstChild.nodeValue="Warning")
									end if
								next
								if NOT iswarning then
									vsRESPMSG=IIfVr(themsg<>"",themsg,shortmsg) & IIfVs(vsRESPMSG<>"","<br>" & vsRESPMSG)
									vsERRCODE=thecode
								end if
							end if
						next
					end if
				next
				if int(vsRESULT)=1 then
					vsRESULT="0" ' Keep in sync with Payflow Pro
					ect_query("UPDATE cart SET cartDateAdded=" & vsusdate(DateAdd("h",dateadjust,Now()))&",cartCompleted=1 WHERE cartCompleted<>1 AND cartOrderID="&ordID)
					ect_query("UPDATE orders SET ordStatus=3,ordAuthStatus='',ordAVS='"&vsAVSADDR&"',ordCVV='"&vsCVV2&"',ordAuthNumber='"&vsAUTHCODE&"',ordTransID='"&vsTRANSID&"',ordDate="&vsusdatetime(DateAdd("h",dateadjust,Now()))&",ordPrivateStatus='' WHERE ordAuthNumber='' AND ordID="&ordID)
					call do_order_success(ordID,emailAddr,sendEmail,FALSE,TRUE,TRUE,TRUE)
				elseif vsERRCODE<>"" then
					vsERRCODE=int(vsERRCODE)
					ect_query("UPDATE orders SET ordPrivateStatus='"&escape_string(strip_tags2("(" & vsERRCODE & ") " & vsRESPMSG))&"' WHERE ordAuthNumber='' AND ordID="&ordID)
					if vsERRCODE=10502 OR vsERRCODE=10504 OR vsERRCODE=10508 OR vsERRCODE=10521 OR vsERRCODE=10527 OR vsERRCODE=10534 OR vsERRCODE=10535 OR vsERRCODE=10541 OR vsERRCODE=12000 OR vsERRCODE=12001 OR vsERRCODE=15004 then vsRESPMSG=xxCCErro
				end if
			end if
		else
			vsRESULT="0"
			vsRESPMSG=xxTranAp
			if InStr(vsAUTHCODE,"-")>0 then vsAUTHCODE=Right(vsAUTHCODE,Len(vsAUTHCODE)-InStr(vsAUTHCODE,"-"))
		end if
	elseif getpost("method")="10" then ' Capture Card
		print "DISABLED!!<br>"
	else
		print "Error"
		response.end
	end if
	SESSION("centinelok")=""
	if centinelenrolled<>"Y" AND ((getpost("method")<>"8" AND getpost("method")<>"22") OR iframe="") then
		call logevent(REMOTE_ADDR,"TRANSACTION",vsRESULT="0","cart"&extension,"ORDERS")
		if vsRESPMSG=xxAuNetR then %>
		<div class="cart4details">
			<div class="cart4header cartheader"><%=xxTnxOrd%></div>
			<div class="copayresultrow">
				<div class="cobhl cobhl4"><%=xxTrnRes%></div><div class="cobll cobll4"><%=vsRESPMSG%></div>
			</div>
			<div class="copayresultrow">
				<div class="cobhl cobhl4"><%=xxOrdNum%></div><div class="cobll cobll4"><%=ordID%></div>
			</div>
			<div class="copayresultrow">
				<div class="cobhl cobhl4"><%=xxAutCod%></div><div class="cobll cobll4">FDS Review</div>
			</div>
		</div>
<%		elseif vsRESULT="0" then %>
	<form method="post" action="thanks<%=extension%>" name="checkoutform">
		<input type="hidden" name="xxpreauth" value="<%=ordID%>" />
		<input type="hidden" name="xxpreauthmethod" value="<%=IIfVr(is_numeric(getpost("method")),getpost("method"),"")%>" />
		<input type="hidden" name="thesessionid" value="<%=thesessionid%>" />
		<div class="cart4details">
			<div class="cart4header cartheader"><%=xxTnxOrd%></div>
			<div class="copayresultrow">
				<div class="cobhl cobhl4"><%=xxTrnRes%></div><div class="cobll cobll4"><%=vsRESPMSG%></div>
			</div>
			<div class="copayresultrow">
				<div class="cobhl cobhl4"><%=xxOrdNum%></div><div class="cobll cobll4"><%=ordID%></div>
			</div>
			<div class="copayresultrow">
				<div class="cobhl cobhl4"><%=xxAutCod%></div><div class="cobll cobll4"><%=vsAUTHCODE%></div>
			</div>
			<div class="cobll cart2column cart4buttons"><%=imageorsubmit(imgclickforreceipt,xxCliCon,"clickforreceipt")%></div>
		</div>
	</form>
	<script>setTimeout("document.checkoutform.submit()",5000);</script>
<%		else %>
	<form method="post" action="cart<%=extension%>" name="checkoutform">
		<input type="hidden" name="mode" value="<%=IIfVr(isavsmismatch,"checkout","go")%>" />
		<input type="hidden" name="orderid" value="<%=ordID%>" />
		<input type="hidden" name="sessionid" value="<%=thesessionid%>" />
		<input type="hidden" name="shipselectoridx" value="<%=SESSION("shipselectoridx")%>" />
		<input type="hidden" name="shipselectoraction" value="<%=SESSION("shipselectoraction")%>" />
		<input type="hidden" name="altrates" value="<%=SESSION("altrates")%>" />
		<div class="cart4details">
			<div class="cart4header cartheader"><%=xxSorTrn%></div>
			<div class="copayresultrow">
				<div class="cobhl cobhl4"><%=xxTrnRes%></div><div class="cobll cobll4 ectwarning"><%=IIfVs(vsERRCODE<>"" AND debugmode, "(" & vsERRCODE & ") ") & vsRESPMSG%></div>
			</div>
			<div class="cobll cart2column cart4buttons"><%=imageorsubmit(imggoback,xxGoBack,"gobackbutton")%></div>
		</div>
	</form>
<%		end if
	end if
elseif checkoutmode="mailinglistsignup" then ' }{
	validsignup=TRUE
	if instr(getpost("mlsuemail"),"@")=0 OR NOT (is_numeric(getpost("mlsectgrp1")) AND is_numeric(getpost("mlsectgrp2"))) then
		validsignup=FALSE
	else
		suarr=split(getpost("mlsuemail"),"@")
		if len(suarr(0))<>int(getpost("mlsectgrp1")) OR len(suarr(1))<>int(getpost("mlsectgrp2")) then validsignup=FALSE
	end if
	if validsignup then call addtomailinglist(getpost("mlsuemail"),getpost("mlsuname"))
	print "<div style=""padding:24px;text-align:center;font-weight:bold"">&nbsp;<br>&nbsp;<br>" & IIfVr(validsignup,xxThkSub,"Sorry, there was a checksum error signing you up to the mailing list") & "</div>"
	if warncheckspamfolder=TRUE then print "<div class=""chkspamfolder ectwarning"">" & xxSpmWrn & "</div>"
	if getpost("rp")<>"" then thehref=htmlspecials(replace(replace(getpost("rp"),"""",""),"<","")) else thehref=storehomeurl
	print "<div style=""padding:24px;text-align:center;font-weight:bold"">" & imageorbutton(imgcontinueshopping,xxCntShp,"continueshopping",thehref,FALSE) & "<br>&nbsp;</div>"
	SESSION("MLSIGNEDUP")=TRUE
end if ' }
if getget("emailconf")<>"" OR getget("unsubscribe")<>"" then
	if getget("emailconf")<>"" then theemail=getget("emailconf") else theemail=getget("unsubscribe")
	sSQL="SELECT email,isconfirmed FROM mailinglist WHERE email='" & escape_string(theemail) & "'"
	rs.open sSQL,cnn,0,1
	foundemail=FALSE
	if NOT rs.EOF then
		foundemail=TRUE
		isconfirmed=(rs("isconfirmed")<>0)
	end if
	rs.close
	print "<div class=""cartemailconf""><div class=""cartemailconftitle"">" & xxMLConf & "</div><div class=""cartemailconfaction"">"
	if NOT foundemail then
		print xxEmNtFn
	elseif getget("unsubscribe")<>"" then
		ect_query("DELETE FROM mailinglist WHERE email='" & escape_string(theemail) & "'")
		print xxSucUns
	elseif isconfirmed then
		print xxAllSub
	else
		thecheck=left(calcmd5(uspsUser&upsUser&origZip&emailObject&checksumtext&":"&theemail), 10)
		if thecheck=getget("check") then
			ect_query("UPDATE mailinglist SET isconfirmed=1 WHERE email='" & escape_string(theemail) & "'")
			print xxSubAct
		else
			print xxSubNAc
		end if
	end if
	print "</div><div class=""cartemailconfcontinue""><a class=""ectlink"" href=""" & storehomeurl & """ onmouseover=""window.status='" & replace(xxCntShp,"'","\'") & "';return true"" onmouseout=""window.status='';return true"">" & xxCntShp & "</a></div>"
	print "</div>"
elseif getget("mode")="gw" then
	print "<form method=""post"" action=""cart"&extension&"?mode=gw"">" & whv("doupdate","1") & "<div class=""cartgiftwrapdiv"">"
	print "<div class=""giftwrapdetails_cntnr"">"
		print "<div class=""giftwrapdetails giftwrapid"">" & xxCODets & "</div>"
		print "<div class=""giftwrapdetails giftwrapname"">" & xxCOName & "</div>"
		print "<div class=""giftwrapdetails giftwrapquant"">" & xxQuant & "</div>"
		print "<div class=""giftwrapdetails giftwrapyes"">" & xxGifWra & "</div>"
	print "</div>"
	if getpost("doupdate")="1" then
		for each objItem in request.form
			if left(objItem,5)="gwset" then
				thecartid=right(objItem, len(objItem)-5)
				if is_numeric(thecartid) AND is_numeric(getpost(objItem)) then
					sSQL="UPDATE cart SET cartGiftWrap=" & getpost(objItem) & ",cartGiftMessage='" & escape_string(strip_tags2(getpost("gwmessage" & thecartid))) & "' WHERE cartID=" & thecartid & " AND " & getsessionsql()
					ect_query(sSQL)
				end if
			end if
		next
		call updategiftwrap()
		print "<div class=""giftwrapupdate"">"
		print "<meta http-equiv=""Refresh"" content=""2; URL=cart"&extension&""">"
		print "<div class=""giftwrapupdating"">" & xxGifUpd & "</div>"
		print "<div class=""giftwrappleasewait"">" & xxPlsWait & " <a class=""ectlink"" href=""cart"&extension&""">" & xxClkHere & "</a>.</div></div>"
	else
		sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pStaticPage,pDisplay,pGiftWrap,cartGiftWrap,cartGiftMessage FROM cart LEFT JOIN products ON cart.cartProdID=products.pID WHERE pGiftWrap<>0 AND cartCompleted=0 AND " & getsessionsql() & " ORDER BY cartID"
		rs.open sSQL,cnn,0,1
		if rs.EOF then
			print "<div class=""giftwrapupdate"">" & xxGifNop & "</div>"
		else
			do while NOT rs.EOF
				print "<div class=""giftwrapline"">"
					print "<div class=""giftwraplineid"">" & rs("cartProdID") & "</div>"
					print "<div class=""giftwraplinename"">" & rs("cartProdName") & "</div>"
					print "<div class=""giftwraplinequant"">" & rs("cartQuantity") & "</div>"
					print "<div class=""giftwraplineyes""><select size=""1"" name=""gwset" & rs("cartID") & """ onchange=""document.getElementById('gwmessage" & rs("cartID") & "').disabled=this[this.selectedIndex].value!='1';""><option value=""0"">" & xxNo & "</option><option value=""1""" & IIfVr(rs("cartGiftWrap")<>0, " selected=""selected""", "") & ">" & xxYes & "</option></select></div>"
					print "<div class=""giftwrapmessage""><div class=""giftwraptmessage"">" & xxGifMes & "</div><div class=""giftwraplinemessage""><textarea placeholder=""" & xxGwMeYe & """ class=""gwmessage" & rs("cartID") & """ name=""gwmessage" & rs("cartID") & """ id=""gwmessage" & rs("cartID") & """ rows=""3"" cols=""34""" & IIfVr(rs("cartGiftWrap")=0, " disabled=""disabled""", "") & ">" & htmlspecials(rs("cartGiftMessage")) & "</textarea></div></div>"
				print "</div>"
				rs.movenext
			loop
			print "<div class=""giftwrapbuttons""><input type=""submit"" value=""Update Selections"" class=""ectbutton giftwrapsubmit"" /> <input type=""button"" class=""ectbutton giftwrapcancel"" value=""" & xxCancel & """ onclick=""document.location='cart"&extension&"'"" /></div>"
		end if
		rs.close
	end if
	print "</div></form>"
elseif (left(getget("token"),2)<>"EC" OR checkoutmode="paypalcancel") AND (checkoutmode="dologin" OR checkoutmode="donewaccount" OR checkoutmode="update" OR checkoutmode="paypalcancel" OR checkoutmode="savecart" OR checkoutmode="movetocart" OR checkoutmode="") AND cartisincluded<>TRUE then ' {
	call getadminshippingparams()
	if SESSION("AmazonLoginTimeout")<>"" AND now()>=SESSION("AmazonLoginTimeout") then
		SESSION("AmazonLogin")=""
		SESSION("AmazonLoginTimeout")=""
	end if
	amazonpaycheckout=SESSION("AmazonLogin")<>"" AND SESSION("AmazonLoginTimeout")<>"" AND getget("amazonpay")="go"
	cartpath=storeurlssl & "cart" & extension
	loginerror=""
	if getget("mode")="logout" then
		SESSION("clientID")=empty
		SESSION("clientUser")=empty
		SESSION("clientActions")=empty
		SESSION("clientLoginLevel")=empty
		SESSION("clientPercentDiscount")=empty
		xxSryEmp=xxLOSuc
		call setacookie("WRITECLL","",-7)
		call setacookie("WRITECLP","",-7)
		if storeurlssl<>storeurl then print "<script src=""" & storeurlssl & "vsadmin/savecookie.asp?DELCLL=Y""></script>"
	end if
	addextrarows=0
	if shipType=0 then estimateshipping=FALSE
	wantstateselector=(FALSE OR forcestateselector OR defaultshipstate<>"") AND estimateshipping
	wantcountryselector=FALSE
	wantzipselector=FALSE
	shipcountry=origCountry
	if estimateshipping=TRUE then
		if commercialloc=2 then commercialloc_=TRUE
		if cartisincluded<>TRUE then
			if SESSION("clientID")<>"" AND getpost("country")="" AND SESSION("country")="" AND shipType>=1 then
				sSQL="SELECT addState,addCountry,addZip FROM address INNER JOIN countries ON address.addCountry=countries.countryName WHERE addCustID="&replace(SESSION("clientID"),"'","")&" ORDER BY addIsDefault DESC"
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then SESSION("country")=rs("addCountry") : SESSION("state")=rs("addState") : SESSION("zip")=rs("addZip")
				rs.close
			end if
			if getpost("state")<>"" then
				if is_numeric(getpost("state")) then shipstateid=getpost("state")
				SESSION("state")=getpost("state")
			elseif SESSION("state")<>"" then
				shipstateid=SESSION("state")
			else
				shipstate=defaultshipstate
			end if
			if getpost("country")<>"" then
				shipcountry=getcountryfromid(getpost("country"))
				SESSION("country")=shipcountry
			elseif SESSION("country")<>"" then
				shipcountry=SESSION("country")
			else
				shipCountryCode=origCountryCode
				shipcountry=origCountry
			end if
		end if
		sSQL="SELECT countryID,countryTax,countryCode,countryFreeShip FROM countries WHERE countryName='"&escape_string(shipcountry)&"'"
		for index=1 to 2
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if trim(SESSION("clientID"))<>"" AND (SESSION("clientActions") AND 2)=2 then countryTaxRate=0 else countryTaxRate=rs("countryTax")
				shipCountryID=rs("countryID")
				shipCountryCode=rs("countryCode")
				freeshipavailtodestination=(rs("countryFreeShip")=1)
				shiphomecountry=(rs("countryID")=origCountryID) OR ((rs("countryID")=1 OR rs("countryID")=2) AND usandcasplitzones)
				rs.close : exit for
			end if
			rs.close
			sSQL="SELECT countryID,countryTax,countryCode,countryFreeShip FROM admin INNER JOIN countries ON admin.adminCountry=countries.countryID WHERE adminID=1"
		next
		sSQL="SELECT shipInsurance"&IIfVr(shiphomecountry,"Dom","Int")&",insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Min,insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Percent,noCarrier"&IIfVr(shiphomecountry,"Dom","Int")&"Ins FROM admin WHERE adminID=1"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			addshippinginsurance=rs("shipInsurance"&IIfVr(shiphomecountry,"Dom","Int"))
			shipinsurancemin=rs("insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Min")
			shipinsurancepercent=rs("insurance"&IIfVr(shiphomecountry,"Dom","Int")&"Percent")
			nocarrierinsurancerates=rs("noCarrier"&IIfVr(shiphomecountry,"Dom","Int")&"Ins")<>0
		end if
		if addshippinginsurance=3 then forceinsuranceselection=TRUE : addshippinginsurance=2
		rs.close
		if cartisincluded<>TRUE then
			if getpost("zip")<>"" then
				destZip=getpost("zip")
				SESSION("zip")=getpost("zip")
			elseif SESSION("zip")<>"" then
				destZip=SESSION("zip")
			else
				if nodefaultzip<>TRUE AND (origCountryCode=shipCountryCode) then destZip=origZip
			end if
		end if
		if shipCountryID=1 OR shipCountryID=2 then shipStateAbbrev=getstateabbrev(IIfVr(is_numeric(shipstateid), shipstateid, shipstate))
		if shiphomecountry then
			sSQL="SELECT stateTax,stateAbbrev,stateFreeShip,stateName FROM states WHERE stateCountryID=" & shipCountryID & " AND " & IIfVr(is_numeric(shipstateid), "stateID=" & shipstateid, "(stateAbbrev='"&escape_string(shipstate)&"' OR stateName='"&escape_string(shipstate)&"')")
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then
				if shipCountryID=origCountryID OR ((shipCountryID=1 OR shipCountryID=2) AND usandcasplitzones) then stateTaxRate=rs("stateTax") else stateTaxRate=0
				freeshipavailtodestination=(freeshipavailtodestination AND (rs("stateFreeShip")=1))
				shipstate=rs("stateName")
			end if
			rs.close
		else
			shipstate=""
		end if
		ordState=shipstate
		shipType=getshiptype()
		addextrarows=1
		if shipType=2 OR shipType=5 then ' weight / price based
			wantcountryselector=TRUE
			if splitUSZones then wantstateselector=TRUE
		elseif shipType=3 OR shipType=4 OR shipType>=6 then
			wantzipselector=TRUE
			wantcountryselector=TRUE
		end if
		if shipType=4 AND upsnegdrates=TRUE then wantstateselector=TRUE
		if NOT nodiscounts AND NOT wantstateselector then
			sSQL="SELECT cpnID FROM coupons WHERE cpnCntry=1 AND cpnType=0 AND cpnNumAvail>0 AND cpnStartDate<=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND cpnEndDate>=" & vsusdate(DateAdd("h",dateadjust,Now()))&" AND (cpnIsCoupon=0 OR (cpnIsCoupon=1 AND cpnNumber='"&rgcpncode&"')) AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then statelimiteddiscount=TRUE else statelimiteddiscount=FALSE
			rs.close
			if statelimiteddiscount then
				sSQL="SELECT stateID FROM states WHERE stateFreeShip=0 AND stateEnabled<>0 AND stateCountryID=" & origCountryID
				rs.open sSQL,cnn,0,1
				if NOT rs.EOF then wantstateselector=TRUE
				rs.close
			end if
		end if
		if (adminAltRates=1 OR adminAltRates=2) AND (NOT wantzipselector OR NOT wantcountryselector) then
			sSQL="SELECT altrateid FROM alternaterates WHERE (usealtmethod<>0 OR usealtmethodintl<>0) AND altrateid IN (3,4,6,7,8,9)"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then wantzipselector=TRUE : wantcountryselector=TRUE
			rs.close
		end if
		if (adminAltRates=1 OR adminAltRates=2) AND NOT wantstateselector AND splitUSZones then
			sSQL="SELECT altrateid FROM alternaterates WHERE (usealtmethod<>0 OR usealtmethodintl<>0) AND altrateid IN (2,5)"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then wantstateselector=TRUE
			rs.close
		end if
		if (adminAltRates=1 OR adminAltRates=2) AND NOT wantcountryselector then
			sSQL="SELECT altrateid FROM alternaterates WHERE (usealtmethod<>0 OR usealtmethodintl<>0) AND altrateid>=2"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then wantcountryselector=TRUE
			rs.close
		end if
		stateSQL=""
		if wantstateselector then
			stateSQL="SELECT stateID,stateAbbrev,stateName,stateName2,stateName3,stateCountryID,countryName FROM states INNER JOIN countries ON states.stateCountryID=countries.countryID WHERE countryEnabled<>0 AND stateEnabled<>0 AND (stateCountryID=" & origCountryID & IIfVs((shipType=4 AND upsnegdrates=TRUE) OR origCountryID=1 OR origCountryID=2," OR stateCountryID=1 OR stateCountryID=2") & ") ORDER BY stateCountryID," & getlangid("stateName",1048576)
 			rs.open stateSQL,cnn,0,1
			if rs.EOF then wantstateselector=FALSE
			rs.close
		end if
		if wantstateselector then wantcountryselector=TRUE : addextrarows=addextrarows+1 else shipstate="" : shipStateAbbrev=""
		if wantcountryselector then addextrarows=addextrarows+1
		if zipisoptional(shipCountryID) then wantzipselector=FALSE
		if wantzipselector then addextrarows=addextrarows+1
	else
		SESSION("xsshipping")=empty
	end if
	initshippingmethods()
	loyaltypointsavailable=0
	redeempoints=TRUE
	SESSION("noredeempoints")=""
	if loyaltypoints<>"" AND SESSION("clientID")<>"" then
		if getget("redeempoints")="no" then
			SESSION("noredeempoints")=TRUE
			redeempoints=FALSE
		end if
		sSQL="SELECT loyaltyPoints FROM customerlogin WHERE clID=" & SESSION("clientID")
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then loyaltypointsavailable=rs("loyaltyPoints") : addextrarows=addextrarows+1
		rs.close
	end if
	stockalreadysubtracted=FALSE
	sSQL="SELECT ordID FROM orders WHERE ordStatus>1 AND ordAuthNumber='' AND ordAuthStatus<>'MODWARNOPEN' AND " & getordersessionsql()
	rs.open sSQL,cnn,0,1
	if NOT rs.EOF then stockalreadysubtracted=TRUE
	rs.close
	if showtaxinclusive<>0 then addextrarows=addextrarows+1
	if stockalreadysubtracted then stockwarning=FALSE else call do_stock_check(TRUE,backorder,stockwarning)
	alldata=""
	if getget("pla")<>"" then hideoptpricediffs=TRUE
	theqs=""
	for each objItem in request.querystring
		if objItem="pli" or objItem="pla" then theqs=theqs&"&"&urlencode(strip_tags2(objItem)) & "=" & urlencode(strip_tags2(getget(objItem)))
	next
%><script>/* <![CDATA[ */
var ectinwhichlist='';
var currqs='<%=jsescape(theqs)%>';
function checkchecked(){
	ischecked=false
	var inputs=document.getElementsByTagName('input');
	for(var i=0; i < inputs.length; i++)
		if(inputs[i].type=='checkbox'){
			if(inputs[i].checked&&inputs[i].name.substr(0,5)=='delet') ischecked=true;
		}
	if(! ischecked) alert("<%=xxNotSel%>");
	return(ischecked);
}
function dodelete(cid,inlist){
ectinwhichlist=inlist;
var ECinput=document.createElement("input");
ECinput.setAttribute("type", "hidden");
ECinput.setAttribute("name", "delet"+cid);
ECinput.setAttribute("value", "ON");
document.forms.checkoutform.appendChild(ECinput);
return doupdate(document.forms.checkoutform);
}
function setestimatorwarn(tobjt,showalert,tmsg){
	var tobj=document.getElementById(tobjt);
	if(tobj)ectaddclass(tobj,'ectwarning');
	if(showalert)alert(tmsg);
	return false;
}
function revealestimatorgo(){
if(document.getElementById('state')&&(document.getElementById('state').disabled||document.getElementById('state').selectedIndex!=0))ectremoveclass(document.getElementById('state'),'ectwarning');
if(document.getElementById('zip'))ectremoveclass(document.getElementById('zip'),'ectwarning');
if(document.getElementById('country'))ectremoveclass(document.getElementById('country'),'ectwarning');
document.getElementById('updateestimator').style.display='';
}
function updateestimator(inlist,showalert){
ectinwhichlist=inlist;
<%		if wantzipselector then %>
	if(document.getElementById('zip').value=='')
		return(setestimatorwarn('zip',showalert,'<%=jscheck(xxPlsZip)%>'));
<%		end if
		if wantstateselector then %>
	if(!(document.getElementById('state').disabled||document.getElementById('state').selectedIndex!=0))
		return(setestimatorwarn('state',showalert,'<%=jscheck(xxPlsSel)%>\n'+document.getElementById('statetxt').innerHTML));
<%		end if %>
	return doupdate(document.forms.checkoutform);
}
function doupdate(tform){
	tform.mode.value='update';
	tform.action='cart<%=extension%>'+(currqs+(ectinwhichlist!=''?'&lid='+ectinwhichlist+'#oc':'')).replace('&','?');
	tform.onsubmit='';
	tform.submit();
	return false;
}
function dosaveitem(lid){
	var ECinput=document.createElement("input");
	ECinput.setAttribute("type", "hidden");
	ECinput.setAttribute("name", "delet"+whichcartid);
	ECinput.setAttribute("value", "ON");
	document.forms.checkoutform.appendChild(ECinput);
	document.forms.checkoutform.mode.value=lid=='x'?'movetocart':'savecart';
	document.forms.checkoutform.listid.value=lid;
	document.forms.checkoutform.action='cart<%=extension%>'+(currqs+(ectinwhichlist!=''?'&lid='+ectinwhichlist+'#oc':'')).replace('&','?');
	document.forms.checkoutform.onsubmit='';
	document.forms.checkoutform.submit();
	return(false);
}
var cartoversldiv,whichcartid;
function cartdispsavelist(clid,inlist,el){
	ectinwhichlist=inlist;
	whichcartid=clid;
	cartoversldiv=false;
	var sld=document.getElementById('savecartlist');
	var parentdiv=el.parentNode;
	parentdiv.insertBefore(sld,parentdiv.firstChild);
	for(var sldindex=0;sldindex<sld.childNodes.length;sldindex++){
		if(sld.childNodes[sldindex].id=='savecartlist'+inlist)
			sld.childNodes[sldindex].style.display='none';
		else if(sld.childNodes[sldindex].style)
			sld.childNodes[sldindex].style.display='table-row';
	}
	sld.style.visibility="visible";
	setTimeout('cartchecksldiv()',2000);
	return(false);
}
function cartchecksldiv(){
	var sld=document.getElementById('savecartlist');
	if(! cartoversldiv)
		sld.style.visibility='hidden';
}
function selaltrate(id){
	document.forms.checkoutform.altrates.value=id;
	doupdate(document.forms.checkoutform);
}
<%			if (SESSION("clientActions") AND 64)=64 then
				cartidlist=""
				sSQL="SELECT cartID FROM cart WHERE cartCompleted=0 AND " & getsessionsql()
				rs.open sSQL,cnn,0,1
				do while NOT rs.EOF
					cartidlist=cartidlist&rs("cartID")&","
					rs.movenext
				loop
				rs.close
				if cartidlist<>"" then
					cartidlist=commaseplist(cartidlist) %>
function dosharecart(){
	document.getElementById('sharecartinput').value='<%
		securitykey=b64_hmac_sha256(adminSecret,thesessionid & "this is a saved cart:" & cartidlist)
		print storeurl & "cart" & extension & "?sharecart=" & thesessionid & "&key=" & urlencode(securitykey) & "&list=" & cartidlist
	%>';
	document.getElementById('sharecartbutton').style.display='none';
	document.getElementById('sharecartinput').style.display='';
	document.getElementById('sharecartinput').select();
	document.execCommand('copy');
}
<%				end if
			end if
	if (adminAltRates=2 AND SESSION("xsshipping")<>"") OR estimateshipping=FALSE then adminAltRates=0
	if adminAltRates=2 then
		sSQL="SELECT altrateid,"&getlangid("altratetext",65536)&" FROM alternaterates WHERE usealtmethod"&international&"<>0 ORDER BY altrateorder"&international&",altrateid"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			print "var shipservicetext=[];" & vbCrLf
			print "var extraship=["
			addcomma=""
			servicetext=""
			do while NOT rs.EOF
				servicetext=servicetext & "shipservicetext[" & rs("altrateid") & "]=""" & rs(getlangid("altratetext",65536)) & """;" & vbCrLf
				if rs("altrateid")<>shipType then
					print addcomma & rs("altrateid")
					addcomma=","
				end if
				rs.movenext
			loop
			print "];" & vbCrLf
			print servicetext & vbCrLf %>
function addCommas(ns,decs,thos){
if((dpos=ns.indexOf(decs))<0)dpos=ns.length;
dpos-=3;
while(dpos>0){
	ns=ns.substr(0,dpos)+thos+ns.substr(dpos);
	dpos-=3;
}
return(ns);
}
function formatestprice(i){
<%	tempStr=FormatEuroCurrency(0)
	print "var pTemplate='"&tempStr&"';" & vbCrLf
	if InStr(tempStr,",")<>0 OR InStr(tempStr,".")<>0 then %>
if(i==Math.round(i))i=i.toString()+".00";
else if(i*10.0==Math.round(i*10.0))i=i.toString()+"0";
else if(i*100.0==Math.round(i*100.0))i=i.toString();
<%	end if
	print "i=addCommas(i.toString()"&IIfVr(InStr(tempStr,",")<>0,".replace(/\./,','),',','.'",",'.',','")&");"
	print "pTemplate=pTemplate.toString().replace(/\d[,.]*\d*/,i.toString());"
	print "return(pTemplate);"
%>}
function acajaxcallback(){
	if(ajaxobj.readyState==4){
		var restxt=ajaxobj.responseText;
		var gssr=restxt.split('SHIPSELPARAM=');
		if(gssr[2]!='ERROR'&&parseFloat(gssr[1])<bestestimate){
			if(document.getElementById('estimatorcell')){
				document.getElementById('estimatorcell').colSpan='1';
				document.getElementById('estimatorcell').align='right';
				var newcell=document.getElementById('estimatorrow').insertCell(-1);
				newcell.className='cobll';
				newcell.innerHTML='&nbsp;';
				document.getElementById('estimatorcell').id='';
			}
			bestestimate=parseFloat(gssr[1]);
			bestcarrier=parseInt(gssr[4]);
			document.getElementById('estimatorspan').innerHTML=formatestprice(bestestimate);
			if(document.getElementById('shippingtotal_cntnr'))document.getElementById('shippingtotal_cntnr').style.display='';
			var discounts=0;
			if(document.getElementById('discountspan')){
				discounts=document.getElementById('discountspan').innerHTML.replace(/[^0-9\.\,]+/g,'');
				var testlatin=/\,\d\d$/;
				if(testlatin.test(discounts))
					discounts=parseFloat(discounts.replace(/\./g,'').replace(/,/g,'.'));
				else
					discounts=parseFloat(discounts.replace(/,/g,''));
			}
<%	if showtaxinclusive<>0 then print "var countrytax=parseFloat(gssr[3]);document.getElementById('countrytaxspan').innerHTML=formatestprice(countrytax);" & vbCrLf else print "var countrytax=0;" & vbCrLf %>
			document.getElementById('grandtotalspan').innerHTML=(formatestprice(Math.round((vstotalgoods+bestestimate+countrytax-discounts)*100)/100.0));
		}else if(gssr[2]=='ERROR'&&document.getElementById('estimatorerrors')){
			if(document.getElementById('estimatorerrors').innerHTML.indexOf(gssr[0])==-1){
				if(gssr[0]=='<%=jsescape(xxInvZip)%>'||gssr[0]=='<%=jsescape(xxPlsZip)%>'||gssr[0].indexOf('The postal code')>=0){
					var x=document.getElementsByClassName("cartzipselectortext");
					for(var aci=0;aci<x.length;aci++) if(x[aci].className.indexOf('ectwarning')==-1)x[aci].className+=' ectwarning';
					var x=document.getElementsByClassName("cartzipselector");
					for(var aci=0;aci<x.length;aci++) if(x[aci].className.indexOf('ectwarning')==-1)x[aci].className+=' ectwarning';

				}
				if(gssr[0]=='<%=jsescape(xxPlsSta)%>'){
					var x=document.getElementsByClassName("cartstateselectortext");
					for(var aci=0;aci<x.length;aci++) if(x[aci].className.indexOf('ectwarning')==-1)x[aci].className+=' ectwarning';
					var x=document.getElementsByClassName("cartstateselector");
					for(var aci=0;aci<x.length;aci++) if(x[aci].className.indexOf('ectwarning')==-1)x[aci].className+=' ectwarning';

				}
				document.getElementById('estimatorerrors').innerHTML+='<div class="estimatorerror ectwarning">'+gssr[0]+'</div>';
				document.getElementById('estimatorerrors').className='estimatorerrors';
			}
		}
		getalternatecarriers();
	}
}
function getalternatecarriers(){
	if(extraship.length>0){
		var shiptype=extraship.shift();
		if(document.getElementById('estimatorchecktext')){
			document.getElementById('estimatorchecktext').innerHTML='Checking carrier';
			document.getElementById('estimatorcheckcarrier').innerHTML=shipservicetext[shiptype];
		}else
			document.getElementById('checkaltspan').innerHTML='Checking ' + shipservicetext[shiptype] + ":";
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=acajaxcallback;
		ajaxobj.open("GET", "vsadmin/shipservice.asp?ratetype=estimator&best="+bestestimate+"&shiptype="+shiptype+"&sessionid=<%=urlencode(thesessionid)%>&destzip=<%=urlencode(destZip)%>&sc=<%=urlencode(shipCountryID)%>&sta=<%=urlencode(shipstateid)%>", true);
		ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		ajaxobj.send(null);
	}else{
		if(document.getElementById('estimatorchecktext')){
			document.getElementById('estimatorchecktext').innerHTML="<%=xxBesRaU%>";
			document.getElementById('estimatorcheckcarrier').innerHTML=shipservicetext[bestcarrier];
		}else
			document.getElementById('checkaltspan').innerHTML="<%=xxBesRaU%> " + shipservicetext[bestcarrier] + ":";
		document.forms.checkoutform.altrates.value=bestcarrier;
	}
}
<%		end if
		rs.close
	end if %>
function applycertcallback(){
	if(ajaxobj.readyState==4){
		var retvals=ajaxobj.responseText.split('&');
		alert(retvals[1]);
		if(retvals[0]=='success'){document.getElementById("cpncode").value='';document.location.reload();}
	}
}
function applycert(){
	var cpncode=document.getElementById("cpncode").value;
	if(cpncode!=""){
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=applycertcallback;
		ajaxobj.open("GET", "vsadmin/ajaxservice.asp?action=applycert&stg1=1&cpncode="+cpncode, true);
		ajaxobj.send(null);
	}
}
function removecert(cpncode){
	if(cpncode!=''){
		ajaxobj=window.XMLHttpRequest?new XMLHttpRequest():new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange=applycertcallback;
		ajaxobj.open("GET", "vsadmin/ajaxservice.asp?action=applycert&stg1=1&act=delete&cpncode="+cpncode, true);
		ajaxobj.send(null);
	}
}
function actionupdate(cid,tform,listid){
ectinwhichlist=listid;
document.getElementById('quant'+cid).name=document.getElementById('quant'+cid).id;
doupdate(tform);
}
function showupdatebutton(cid,cartlineunique,listid){
document.getElementById('cartlinetot'+cartlineunique).innerHTML='<input type="button" class="ectbutton cartlineupdate" value="<%=jscheck(xxUpdate)%>" onclick="actionupdate('+cid+',this.form,\''+listid+'\')" />';
}
/* ]]> */</script>
<%	cartliststring=""
	if SESSION("clientID")<>"" AND enablewishlists=TRUE then ' Wish List Popup
		cartliststring=cartliststring&"0,"
		sSQL="SELECT listID,listName FROM customerlists WHERE listOwner="&SESSION("clientID")
		rs.CursorLocation=3
		rs.open sSQL,cnn
		cartlistnumber=1
		print "<div id=""savecartlist"" class=""savecartlist"" style=""position:absolute;visibility:hidden;top:0px;left:0px;z-index:10000;display:table"" onmouseover=""cartoversldiv=true;"" onmouseout=""cartoversldiv=false;setTimeout('cartchecksldiv()',1000)"">"
		print "<div id=""savecartlist" & (rs.recordcount+1) & """ style=""display:table-row""><div onclick=""dosaveitem('x')"" style=""display:table-cell"">" & xxShoCar & "</div></div>" &vbCrLf
		print "<div style=""display:table-row"" class=""savecartdivider"">-</div>" & vbCrLf
		print "<div id=""savecartlist0"" style=""display:table-row""><div onclick=""dosaveitem('0')"" style=""display:table-cell"">" & xxMyWisL & "</div></div>" & vbCrLf
		do while NOT rs.EOF
			cartliststring=cartliststring&rs("listID")&","
			print "<div id=""savecartlist" & cartlistnumber & """ style=""display:table-row""><div onclick=""dosaveitem("&rs("listID")&")"" style=""display:table-cell"">" & htmlspecials(rs("listName")) & "</div></div>" & vbCrLf
			cartlistnumber=cartlistnumber+1
			rs.movenext
		loop
		rs.close
		rs.CursorLocation=2
		print "</div>"
	end if
	if xxCoStp1<>"" then
		sSQL="SELECT cartID FROM cart WHERE cartCompleted=0 AND "&getsessionsql()
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then print "<div class=""checkoutsteps"">" & xxCoStp1 & "</div>"
		rs.close
	end if
	print "<div class=""cart1details cartlists"">" ' {
	if enableclientlogin AND NOT noclientloginprompt AND NOT amazonpaycheckout then
		print "<div class=""cartlistdiv""><div class=""ectdivhead cartlistname cartlistlogin"">"
		if SESSION("clientID")<>"" then
			print "<div>&nbsp;</div><div class=""cartloggedin"">"
			print imageorbutton(imgcartlogout,xxLogout&" "& htmlspecials(SESSION("clientUser")),"ectlink","return dologoutaccount()",TRUE)
			if enablewishlists then print " <input type=""button"" class=""ectbutton"" onclick=""document.location='" & customeraccounturl&"#list';return false"" value="""&xxCreaGR&""" /> "
			if getget("warncheckspamfolder")="true" then print "<div class=""thanksubscribe"">" & xxThkSub & "</div><div class=""spamwarn"">" & xxSpmWrn & "</div>"
			if getget("cartchanged")="true" then print "<div class=""cartchanged"">" & xxCarCha & "</div>"
			print "</div><div>&nbsp;</div>"
		elseif noclientloginprompt<>TRUE then
			print "<div>&nbsp;</div><div class=""loginprompt""><div class=""logintoaccount"">" & imageorbutton(imgloginaccount,xxLogAcc,"logintoaccount","displayloginaccount()",TRUE)&"</div>"
			if allowclientregistration then print "<div class=""createaccount"">" & imageorbutton(imgcreateaccount,xxCreAcc,"createaccount","displaynewaccount()",TRUE)&"</div>"
			print "</div><div>&nbsp;</div>"
		end if
		print "</div></div>"
	end if
	haspubliclist=FALSE
	if is_numeric(getget("pli")) AND getget("pla")<>"" AND instr(","&cartliststring&",",","&getget("pli")&",")=0 then
		sSQL="SELECT listID,listName FROM customerlists WHERE listID="&getget("pli")&" AND listAccess='"&escape_string(getget("pla"))&"'"
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			haspubliclist=TRUE : cartliststring=cartliststring&getget("pli")&","
		else
			print "<div class=""ectwarning"">That public list was not found or has been deleted.</div>"
		end if
		rs.close
	end if
	if onvacation<>0 then
		cartliststring=""
		vacationmessage=""
		print "<div class=""ectwarning vacationmessage"" style=""text-align:center;font-size:1.3em;padding:16px"">"
		rs.open "SELECT vacationmessage FROM admin WHERE adminID=1"
		if NOT rs.EOF then vacationmessage=rs("vacationmessage") else vacationmessage="The store is currently offline."
		rs.close
		print vacationmessage
		print "</div>"
	end if
	opencartid=" id=""oc"""
	cartlineunique=0
	cartliststringarray=split(cartliststring&IIfVs(onvacation=0," "),",")
	saveshowtaxinclusive=showtaxinclusive
	saveestimateshipping=estimateshipping
	savestateTaxRate=stateTaxRate
	savecountryTaxRate=countryTaxRate
	savehandling=handling
	for cartlistnumber=0 to UBOUND(cartliststringarray)
		shipfreegoods=0 : totalgoods=0 : totalquantity=0 : shipping=0 : stateTax=0 : countryTax=0 : totaldiscounts=0 : freeshipamnt=0 : loyaltypointdiscount=0
		handling=savehandling
		ispubliclist=FALSE
		cartlistclass="cartlistgift"
		if listid="0" then cartlistclass="cartlistwish" else if listid="" then cartlistclass="cartlistshop"
		print "<div class=""cartlistdiv " & cartlistclass & """>" ' {
		if is_numeric(getget("lid")) then displaylistid=getget("lid") else displaylistid=""
		if is_numeric(getget("pli")) AND getget("pla")<>"" then displaylistid=getget("pli")
		listid=trim(cartliststringarray(cartlistnumber))
		checkoutsteptxt="&nbsp;"
		numitems=0
		if is_numeric(getget("pli")) AND cstr(listid)=getget("pli") AND haspubliclist then
			checkoutmode="savedcart"
			showtaxinclusive=0
			estimateshipping=FALSE
			sSQL="SELECT listID,listName FROM customerlists WHERE listID="&getget("pli")&" AND listAccess='"&escape_string(getget("pla"))&"'"
			rs.open sSQL,cnn,0,1
			if NOT rs.EOF then listname=rs("listName") : querystr="cartCompleted=3 AND cartListID="&listid : ispubliclist=TRUE
			rs.close
		elseif SESSION("clientID")<>"" AND listid<>"0" AND listid<>"" then
			checkoutmode="savedcart"
			showtaxinclusive=0
			estimateshipping=FALSE
			sSQL="SELECT listID,listName FROM customerlists WHERE listID="&listid&" AND listOwner="&SESSION("clientID")
			rs.open sSQL,cnn,0,1
			if rs.EOF then querystr="cartCompleted=0 AND "&getsessionsql() else listname=rs("listName") : querystr="cartCompleted>=0 AND cartListID="&listid
			rs.close
		elseif listid="0" then
			listname=xxMyWisL
			checkoutmode="savedcart"
			showtaxinclusive=0
			estimateshipping=FALSE
			querystr="cartCompleted=3 AND cartListID=0 AND "&getsessionsql()
		else
			listname=xxShoCar
			checkoutmode=""
			showtaxinclusive=saveshowtaxinclusive
			estimateshipping=saveestimateshipping
			stateTaxRate=savestateTaxRate
			countryTaxRate=savecountryTaxRate
			querystr="cartCompleted=0 AND "&getsessionsql()
			checkoutsteptxt="1"
		end if
		sSQL="SELECT COUNT(*) AS numitems FROM cart LEFT JOIN products ON cart.cartProdID=products.pID WHERE " & querystr
		rs.open sSQL,cnn,0,1
		if NOT rs.EOF then
			if NOT rs.EOF then if NOT isnull(rs("numitems")) then numitems=rs("numitems") else numitems=0
		end if
		rs.close
		if numitems<=0 then checkoutsteptxt="&nbsp;"
		print "<div class=""ectdivhead cartlistname"&IIfVs(listid<>""," cartlistgiftreg")&IIfVs(listid=""," cartlistcart")&"""" & IIfVs(cstr(cartlistnumber)=getget("lid")OR cstr(listid)=getpost("listid"),opencartid) & " style=""cursor:pointer"" onclick=""document.getElementById('cartlistid"&cartlistnumber&"').style.display==''?document.getElementById('cartlistimgid"&cartlistnumber&"').src='images/arrow-down.png':document.getElementById('cartlistimgid"&cartlistnumber&"').src='images/arrow-up.png';document.getElementById('cartlistid"&cartlistnumber&"').style.display=(document.getElementById('cartlistid"&cartlistnumber&"').style.display==''?'none':'')""><div class=""checkoutstep" & IIfVs(checkoutsteptxt<>"&nbsp;"," checkoutstepof3") & """>" & checkoutsteptxt & "</div><div class=""cartname"">" & htmlspecials(listname) & " (" & numitems & ")</div><div class=""cartlistimg""><img src=""images/arrow-" & IIfVr(displaycartclosed(),"down","up") & ".png"" id=""cartlistimgid"&cartlistnumber&""" alt=""Show / Hide Cart"" /></div></div>" & vbCrLf
		if cstr(cartlistnumber)=getget("lid") OR cstr(listid)=getpost("listid") then opencartid=""
		print "<div class=""cartlistcontents"" id=""cartlistid"&cartlistnumber&"""" & IIfVs(displaycartclosed()," style=""display:none""") & ">"
		sSQL="SELECT cartID,cartProdID,cartProdName,cartProdPrice,cartQuantity,pWeight,pShipping,pShipping2,pExemptions, pSection,topSection,pDims,pTax,pStaticPage,pDisplay,'' AS pImage,'' AS pLargeImage, cartCompleted,pGiftWrap,cartGiftWrap,pStaticURL,pMinQuant,pID,cartOrigProdID,"&getlangid("pDescription",2)&","&getlangid("pLongDescription",4)
		if mysqlserver=TRUE then
			sSQL=sSQL&" FROM cart LEFT JOIN products ON cart.cartProdID=products.pID LEFT OUTER JOIN sections ON products.pSection=sections.sectionID WHERE " & querystr & " ORDER BY cartID"
		else
			sSQL=sSQL&" FROM cart LEFT JOIN (products LEFT OUTER JOIN sections ON products.pSection=sections.sectionID) ON cart.cartProdID=products.pID WHERE " & querystr & " ORDER BY cartID"
		end if
		rs.open sSQL,cnn,0,1
		if NOT (rs.EOF OR rs.BOF) then alldata=rs.getrows else alldata=""
		rs.close
		if displaysoftlogindone="" then displaysoftlogindone=""
		if enableclientlogin then call displaysoftlogin()
		print "<form method=""post""" & IIfVs(checkoutmode=""," name=""checkoutform"" onkeydown=""return preventreturn(this,event)""") & " action=""" & cartpath & """" & IIfVs(isarray(alldata)," onsubmit=""return changechecker(this)""") & ">"
		print whv("mode","checkout") & whv("sessionid",getsessionid()) & whv("PARTNER",strip_tags2(trim(IIfVr(getget("PARTNER")<>"",getget("PARTNER"),request.cookies("PARTNER"))))) & whv("cart","") & whv("listid","")
		if SESSION("noredeempoints")=TRUE then call writehiddenidvar("noredeempoints", "1")
		if adminAltRates<>0 then print whv("altrates", getpost("altrates"))
		print "<div class=""cartcontentsdiv"">"
		if carterror<>"" then print "<div class=""carterror ectwarning"">" & carterror & "</div>"
		if noshowcart then
			' do nothing
		elseif isarray(alldata) then
			totaldiscounts=0
			changechecker=""
			call showcartlines(FALSE)
			if showtaxinclusive=0 then
				stateTaxRate=0
				countryTaxRate=0
			end if
			if checkoutmode="savedcart" OR nopriceanywhere then
				addextrarows=0
			else
				call calculatediscounts(totalgoods,FALSE,rgcpncode)
				if SESSION("giftcerts")<>"" then
					sSQL="SELECT gcID,gcRemaining FROM giftcertificate WHERE gcRemaining>0 AND gcAuthorized<>0 AND gcID IN ('" & replace(escape_string(SESSION("giftcerts"))," ","','") & "')"
					rs.open sSQL,cnn,0,1
					do while NOT rs.EOF
						giftcertsamount=giftcertsamount+rs("gcRemaining")
						rs.movenext
					loop
					rs.close
				end if
				if totaldiscounts>totalgoods+IIfVr(showtaxinclusive=3,countryTax,0) then totaldiscounts=totalgoods+IIfVr(showtaxinclusive=3,countryTax,0)
				if totaldiscounts=0 then
					SESSION("discounts")=empty
				else
					SESSION("discounts")=totaldiscounts
					addextrarows=addextrarows + 1
				end if
			end if
			if estimateshipping=TRUE then
				if SESSION("xsshipping")="" AND success then
					if calculateshipping() then
						insuranceandtaxaddedtoshipping()
						calculateshippingdiscounts(FALSE)
						calculatetaxandhandling()
						SESSION("xsshipping")=(shipping+handling)-freeshipamnt
					else
						freeshipmethodexists=TRUE
						calculateshippingdiscounts(FALSE)
						calculatetaxandhandling()
						handling=0
					end if
				elseif SESSION("xsshipping")="" AND NOT success then
					calculatetaxandhandling()
					handling=0
				else
					shipping=SESSION("xsshipping")
					countryTax=IIfVr(SESSION("xscountrytax")="",0,SESSION("xscountrytax"))
					handling=0
					calculatetaxandhandling()
				end if
			else
				freeshipmethodexists=TRUE
				calculateshippingdiscounts(FALSE)
				shipping=0 : handling=0 : freeshipamnt=0
				calculatetaxandhandling()
			end if
			if NOT amazonpaycheckout AND NOT nopriceanywhere then
				print "<div class=""cartshippingandtotals"">" ' { cartshippingandtotals
				if estimateshipping<>TRUE OR nohandlinginestimator=TRUE then handling=0 : handlingchargepercent=0
				print "<div class=""cartshippingdetails"">"
				if estimateshipping=TRUE then
					if instr(errormsg,xxInvZip)>0 OR instr(errormsg,xxPlsZip)>0 then invalidzip=TRUE else invalidzip=FALSE
					if instr(errormsg,xxPlsSta)>0 then invalidstate=TRUE else invalidstate=FALSE
					print IIfVs(xxEstimatorTitle<>"","<div class=""estimatortitle"">"&xxEstimatorTitle&"</div>")
					print "<div " & IIfVs(errormsg<>"","class=""estimatorerrors"" ") & "id=""estimatorerrors"">" & IIfVs(errormsg<>"","<div class=""estimatorerror ectwarning"">" & errormsg & "</div>") & "</div>"
					if adminAltRates>0 then
						print "<div class=""shipestimatemarkup_cntnr""><div class=""cartestimatortext"">"
						writeestimatormenu()
						print "</div></div>" & vbCrLf
					end if
					if wantstateselector then
						print "<div class=""cartstateselector_cntnr""><div class=""cartstateselectortext" & IIfVs(invalidstate," ectwarning") & """>"
						print "<label class=""ectlabel"" for=""state"" id=""statetxt"">"&xxState&"</label>"
						print "</div><div class=""cartstateselector" & IIfVs(invalidstate," ectwarning") & """>"
						print "<select name=""state"" id=""state"" class=""ectselectinput nofixw cartstateselector" & IIfVs(invalidstate," ectwarning") & """ size=""1"" onchange=""revealestimatorgo()"">" : show_states(shipstate) : print "</select>"
						print "</div></div>" & vbCrLf
					end if
					if wantcountryselector then
						print "<div class=""cartcountryselector_cntnr""><div class=""cartcountryselectortext"">"
						print labeltxt(xxCountry,"country")
						print "</div><div class=""cartcountryselector"">"
						print "<select name=""country"" id=""country"" class=""ectselectinput nofixw"" size=""1"""&IIfVs(mobilebrowser," style=""font-size:10px""")&" onchange="""&IIfVs(wantstateselector OR wantzipselector,"dynamiccountries(this,'');")&"updateestimator('"&cartlistnumber&"',false)"">"
						sSQL="SELECT countryID,countryName,countryCode,"&getlangid("countryName",8)&" AS cnameshow FROM countries WHERE countryEnabled=1 ORDER BY countryOrder DESC,"&getlangid("countryName",8)
						rs.open sSQL,cnn,0,1
						do while NOT rs.EOF
							print "<option value="""&rs("countryID")&""""
							if shipcountry=rs("countryName") then print " selected=""selected"""
							cnameshow=rs("cnameshow")
							if cnameshow="United States of America" AND mobilebrowser then cnameshow="USA"
							print ">" & cnameshow & "</option>"&vbCrLf
							rs.movenext
						loop
						rs.close
						print "</select>"
						print "</div></div>" & vbCrLf
					end if
					if wantzipselector then
						print "<div class=""cartzipselector_cntnr""><div class=""cartzipselectortext" & IIfVs(invalidzip," ectwarning") & """>"
						print "<label class=""ectlabel"" for=""zip"" id=""ziptxt"">"&xxZip&"</label>"
						print "</div><div class=""cartzipselector" & IIfVs(invalidzip," ectwarning") & """>"
						print "<input type=""text"" name=""zip"" id=""zip"" class=""ecttextinput nofixw cartzipselector" & IIfVs(invalidzip," ectwarning") & """ size=""8"" value=""" & htmlspecials(destZip) & """ onkeydown=""revealestimatorgo()"" autocapitalize=""characters"" />"
						print "</div></div>" & vbCrLf
					end if
					if wantstateselector OR wantcountryselector OR wantzipselector then
						print "<div id=""updateestimator"" class=""updateestimator"" style=""display:none"">" & imageorbutton(imgupdateestimator,xxUpdEst,"updateestimator","updateestimator('"&cartlistnumber&"',true)",TRUE) & "</div>"
					end if
				end if
				if xxEstimatorEnd<>"" AND checkoutmode<>"savedcart" then print "<div class=""cartestimatorend"">"&xxEstimatorEnd&"</div>"
				print "</div>"
				print "<div class=""carttotals"">" ' {
				print "<div class=""cartsubtotal_cntnr""><div class=""cartsubtotaltext"">" & xxSubTot & "</div><div class=""cartsubtotal"">" & FormatEuroCurrency(totalgoods) & "</div></div>" & vbCrLf
				if estimateshipping then
					print "<div class=""shippingtotal_cntnr"" id=""shippingtotal_cntnr"""&IIfVs(errormsg<>""," style=""display:none""")&"><div class=""shippingtotaltext"">" & IIfVr(handling=0 OR errormsg<>"",xxShpEst,xxShHaEs) & "</div><div class=""shippingtotal"" id=""estimatorspan"">"
					if errormsg<>"" then
						print errormsg
					elseif freeshipamnt=(shipping+handling) then
						print xxFree
					else
						print FormatEuroCurrency((shipping+handling)-freeshipamnt)
					end if
					print "</div></div>"
				end if
				if showtaxinclusive<>3 then call displaydiscounts()
				if showtaxinclusive<>0 then
					print "<div class=""cartcountrytax_cntnr""><div class=""cartcountrytaxtext"">"
					print xxCntTax
					print "</div><div class=""cartcountrytax"">"
					print "<span id=""countrytaxspan"">"&FormatEuroCurrency(countryTax)&"</span>"
					print "</div></div>"&vbCrLf
					if showtaxinclusive=3 then call displaydiscounts()
				else
					countryTax=0
				end if
				print "<div class=""cartgrandtotal_cntnr""><div class=""cartgrandtotaltext"">"
				if getget("pla")="" then
					print IIfVr(checkoutmode="savedcart",xxItmTot,xxGndTot)
					print "</div><div class=""cartgrandtotal"""&IIfVs(checkoutmode<>"savedcart"," id=""grandtotalspan""")&">"
					print FormatEuroCurrency((totalgoods+shipping+handling+countryTax)-(totaldiscounts+freeshipamnt+loyaltypointdiscount))
				end if
				print "</div></div>"
				if giftcertsamount<>0 then
					print "<div class=""cartgiftcert_cntnr""><div class=""cartgiftcerttext ectdscntt"">" & xxAppGC & "</div><div class=""cartgiftcert ectdscnt"">" & FormatEuroCurrency(vrmin(giftcertsamount,(totalgoods+shipping+handling+countryTax)-(totaldiscounts+freeshipamnt+loyaltypointdiscount))) & "</div></div>"
				end if
				if checkoutmode<>"savedcart" then
					sSQL="SELECT "&IIfVs(NOT mysqlserver,"TOP 1 ")&"cpnID FROM coupons WHERE cpnIsCoupon<>0 AND ((cpnLoginLevel>=0 AND cpnLoginLevel<="&minloglevel&") OR (cpnLoginLevel<0 AND -1-cpnLoginLevel="&minloglevel&"))"&IIfVs(mysqlserver," LIMIT 0,1")
					rs.open sSQL,cnn,0,1
					hasacoupon=NOT rs.EOF
					rs.close
					if hasacoupon then
						print "<div class=""cartcoupon_cntnr"">"
						print "<div class=""cartcoupontext"">" & labeltxt(xxGifNum,"cpncode") & "</div>"
						cpnarr=split(trim(SESSION("giftcerts")), " ")
						for index=0 to UBOUND(cpnarr)
							print "<div class=""cartcouponapplied"">" & imageorlink("",xxRemove&" : "&cpnarr(index),"applycoupon removecoupon1","removecert('"&cpnarr(index)&"')",TRUE) & "</div>"
						next
						cpnarr=split(trim(SESSION("cpncode")), " ")
						for index=0 to UBOUND(cpnarr)
							print "<div class=""cartcouponapplied"">" & imageorlink("",xxRemove&" : "&cpnarr(index),"applycoupon removecoupon1","removecert('"&cpnarr(index)&"')",TRUE) & "</div>"
						next
						print "<div class=""cartcoupon""><input type=""text"" name=""cpncode"" class=""ecttextinput cpncart1"" placeholder=""" & IIfVr(nogiftcertificate,xxGifNum,xxCoGfPl) & """ id=""cpncode"" autocomplete=""off"" /> " & imageorbutton(imgapplycoupon,xxApply,"applycoupon applycoupon1","applycert()",TRUE) & "</div>"
						print "</div>"
					end if
				end if
				print "</div></div>" ' carttotals cartshippingandtotals } }
				if checkoutmode<>"savedcart" then
					if SESSION("tofreeshipamount")<>"" then
						print "<div class=""tofreeshipping"">" & replace(xxToFSAm,"%s",FormatEuroCurrency(SESSION("tofreeshipamount"))) & "</div>"
					elseif SESSION("tofreeshipquant")<>"" then
						print "<div class=""tofreeshipping"">" & replace(xxToFSQu,"%s",SESSION("tofreeshipquant")) & "</div>"
					end if
				end if
			end if
			if amazonpaycheckout OR SESSION("AmazonLogin")<>"" then
				if getpayprovdetails(21,data1,data2,data3,demomode,ppmethod) then
					print "<script>window.onAmazonLoginReady=function(){amazon.Login.setClientId(""" & data1 & """);};</script>"
					print "<script src=""" & getamazonjsurl(demomode) & """></script>"
				end if
			end if
			if amazonpaycheckout then
				if now()>=SESSION("AmazonLoginTimeout") then
					SESSION("AmazonLogin")=""
					SESSION("AmazonLoginTimeout")=""
				else
					if getpayprovdetails(21,data1,data2,data3,demomode,ppmethod) then
						data2arr=split(data2,"&",2)
						if UBOUND(data2arr)>=0 then data2=data2arr(0)
						if UBOUND(data2arr)>0 then sellerid=data2arr(1)
						print "<div class=""amazoncontent"">"
							print "<div class=""amazonaddressandwallet"">"
								print "<div id=""addressBookWidgetDiv""></div>"
								print "<div id=""walletWidgetDiv""></div>"
							print "</div>"
						print "<div class=""amazonbuttons""><div class=""paynowamazon"">" & imageorbutton(imgamazonpaynow,"Click to Check Totals" & IIfVs(shipType<>0," / Select Shipping"),"amazonpaynow","amazonpaynow()",TRUE) & "</div><div class=""amazonlogout2"">" & imageorlink(imgamazonlogout,"Logout of your Amazon account","amazonlogout","return amazonlogout()",TRUE) & "</div></div>"
						print "</div>" %>
<script>
var amznorderreferenceid='';
var addressselected=false,paymentselected=false;
new OffAmazonPayments.Widgets.AddressBook({
  sellerId: '<%=sellerid%>',
  onOrderReferenceCreate: function(orderReference) {
    amznorderreferenceid=orderReference.getAmazonOrderReferenceId();
  },
  onAddressSelect: function(orderReference) {
	addressselected=true;
	paymentselected=false;
  },
  design: {
    designMode: 'responsive'
  },
  onError: function(error) {
	alert(error.getErrorMessage());
  }
}).bind("addressBookWidgetDiv");

new OffAmazonPayments.Widgets.Wallet({
  sellerId: '<%=sellerid%>',
  onPaymentSelect: function(orderReference) {
	paymentselected=true;
  },
  design: {
    designMode: 'responsive'
  },
  onError: function(error) {
  }
}).bind("walletWidgetDiv");
function amazonpaynow(){
	if(amznorderreferenceid!='')
		document.location='cart<%=extension%>?amzrefid='+amznorderreferenceid;
	else
		alert("Please select an address and payment method.");
}
</script>
<%					end if
				end if
			else ' cartcheckoutbuttons
				print "<div class=""cartcheckoutbuttons"">"
				if trim(SESSION("clientID"))<>"" then
					sequence=ip2long(REMOTE_ADDR)
					ect_query("DELETE FROM tmplogin WHERE tmplogindate < " & vsusdate(Date()-3) & " OR tmploginid='" & escape_string(thesessionid) & "'")
					ect_query("INSERT INTO tmplogin (tmploginid, tmploginname, tmploginchk,tmplogindate) VALUES ('" & escape_string(thesessionid) & "','" & replace(SESSION("clientID"),"'","") & "'," & sequence & "," & vsusdate(Date()) & ")")
					print whv("checktmplogin", sequence)
					if (SESSION("clientActions") AND 8)=8 OR (SESSION("clientActions") AND 16)=16 then
						if minwholesaleamount<>"" then minpurchaseamount=minwholesaleamount
						if minwholesalequantity<>"" then minpurchasequantity=minwholesalequantity
						if minwholesalemessage<>"" then minpurchasemessage=minwholesalemessage
					end if
				else
					print whv("checktmplogin","x")
				end if
				estimate=(totalgoods+shipping+handling+stateTax+countryTax)-(totaldiscounts+freeshipamnt+loyaltypointdiscount)
				if checkoutmode="savedcart" then
					' Do nothing
				elseif totalgoods<minpurchaseamount OR totalquantity<minpurchasequantity then
					print "<div class=""checkoutopts cominpurchase ectwarning"">" & minpurchasemessage & "</div>" & vbCrLf
				elseif forceclientlogin AND SESSION("clientID")="" then
					print "<div class=""checkoutopts coforcelogin"">" & xxBfChk & " <a class=""ectlink"" href=""#"" onclick=""return displayloginaccount()"">" & xxLogin & "</a>" & IIfVs(allowclientregistration," " & xxOr & " <a class=""ectlink"" href=""#"" onclick=""return displaynewaccount()"">" & xxCrAc & "</a>") & ".</div>"
				elseif stockwarning then
					' Do nothing
				else
					regularcheckoutshown=FALSE
					sSQL="SELECT payProvID,payProvData1,payProvData2,payProvDemo FROM payprovider WHERE payProvEnabled=1 AND payProvLevel<=" & minloglevel & IIfVs(paypalhostedsolution," AND payProvID<>18") & IIfVs(estimate<=0," AND payProvID<>19") & " ORDER BY payProvOrder"
					rs.open sSQL,cnn,0,1
					do while NOT rs.EOF
						if rs("payProvID")=21 then
							if getget("access_token")<>"" AND getget("token_type")="bearer" AND getget("expires_in")<>"" AND getget("scope")<>"" then
								if callxmlfunction("https://api." & IIfVs(rs("payProvDemo"),"sandbox.") & "amazon.com/auth/o2/tokeninfo?access_token=" & urlencode(getget("access_token")),"",res,"","WinHTTP.WinHTTPRequest.5.1",errormsg,FALSE) then
									if res<>"" then
										resarray=split(mid(trim(res),2,len(res)-2),",")
										for each objItem in resarray
											keypair=split(objItem,":")
											if keypair(0)="""aud""" AND keypair(1)="""" & rs("payProvData1") & """" then
												SESSION("AmazonLogin")=trim(res)
											elseif keypair(0)="""exp""" then
												SESSION("AmazonLoginTimeout")=dateadd("s",keypair(1),now())
											end if
										next
									end if
								end if
								if trim(SESSION("AmazonLogin"))<>"" AND trim(SESSION("AmazonLoginTimeout"))<>"" then response.redirect storeurlssl & "cart"&extension&"?amazonpay=go"
							else

		print "<div class=""checkoutopts coopt21"" id=""AmazonPayButton""></div>"
		print "<script>window.onAmazonLoginReady=function(){amazon.Login.setClientId(""" & rs("payProvData1") & """);};</script>"
		print "<script src=""" & getamazonjsurl(demomode) & """></script>"
%>					
<script>
var authRequest;
OffAmazonPayments.Button("AmazonPayButton", "<%=rs("payProvData2")%>", { // MERCHANT_ID
	type: "<%=IIfVr(SESSION("AmazonLogin")<>"","PwA","LwA")%>",
	authorization: function () {
		loginOptions = { scope: "profile postal_code payments:widget payments:shipping_address", popup: true };
		authRequest = amazon.Login.authorize(loginOptions, "<%=storeurlssl & "cart" & extension & "?amazonpay=go"%>");
	},
	onError: function (error) {
		alert("handle error function");
	}
});
function amazonlogout(){
	amazon.Login.logout();
	document.location='cart<%=extension%>?amazon=logout';
	return false;
}
</script>
<%

							end if
						end if
						if rs("payProvID")=19 then
							wantbillmelater=FALSE
							if rs("payProvData2")<>"" then
								data2arr=split(rs("payProvData2"),"&")
								if UBOUND(data2arr)>=1 then wantbillmelater=(data2arr(1)="1")
							end if
							if wantbillmelater then
				print "<div class=""checkoutopts checkoutbutton1 coopt" & rs("payProvID") & """><input type=""image"" src=""https://www.paypalobjects.com/webstatic/en_US/btn/btn_bml_SM.png"" onclick=""document.forms.checkoutform.cart.value='';document.forms.checkoutform.mode.value='billmelater';"" alt=""Bill Me Later"" title=""" & xxPPPBlu & """ /></div>" & vbCrLf
							end if
				print "<div class=""checkoutopts checkoutbutton1 coopt" & rs("payProvID") & """><input type=""image"" src=""https://www.paypal.com/en_US/i/btn/btn_xpressCheckoutsm.gif"" onclick=""document.forms.checkoutform.cart.value='';document.forms.checkoutform.mode.value='paypalexpress1';"" alt=""PayPal Express"" title=""" & xxPPPBlu & """ /></div>" & vbCrLf
						elseif NOT regularcheckoutshown then
							regularcheckoutshown=TRUE
				print "<div class=""checkoutopts checkoutbutton1 coopt" & rs("payProvID") & """>" & imageorsubmit(imgcheckoutbutton,IIfVr(xxCOTxt1<>"",xxCOTxt1,xxCOTxt)&""" onclick=""document.forms.checkoutform.action='"&cartpath&"';document.forms.checkoutform.cart.value='';document.forms.checkoutform.mode.value='checkout';"" title="""&xxPrsChk,"checkoutbutton checkoutbutton1") & "</div>" & vbCrLf
						end if
						rs.movenext
					loop
					rs.close
					if SESSION("AmazonLogin")<>"" then print "<div class=""checkoutopts checkoutbutton1 amazonlogout1"">" & imageorlink(imgamazonlogout,"Logout of your Amazon account","amazonlogout1","return amazonlogout()",TRUE) & "</div>"
				end if
				print "</div>" ' } cartcheckoutbuttons
				if checkoutmode<>"savedcart" AND googletagid<>"" then
					print "<script>gtag(""event"",""view_cart"",{ currency:'" & countryCurrency & "',value:" & totalgoods & ",items:[" & getcartforganalytics("") & "]});</script>" & vbLf
				end if
			end if %>
<script>
/* <![CDATA[ */
<%			if SESSION("AmazonLogin")<>"" then %>
function amazonlogout(){
	amazon.Login.logout();
	document.location='cart<%=extension%>?amazon=logout';
	return false;
}
<%			end if
			if wantstateselector OR wantzipselector then
				createdynamicstates(stateSQL)
				print "dynamiccountries(document.getElementById('country'),'');setinitialstate('');" & vbCrLf
			end if
			session.LCID=1033
			if adminAltRates=2 AND (((shipping+handling)-freeshipamnt)>0 OR errormsg<>"") then
				print "var bestcarrier="&shipType&";var bestestimate=" & (((shipping+handling)-freeshipamnt) + IIfVr(errormsg<>"",1000000,0)) & ";" & vbCrLf
				print "var vstotalgoods=" & totalgoods & ";" & vbCrLf & "getalternatecarriers();" & vbCrLf
			end if
			session.LCID=saveLCID %>
function changechecker(){
<%	if minquantityerror OR deleteditemerror then print "alert(""" & IIfVs(minquantityerror,jscheck(xxMinSom)) & IIfVs(deleteditemerror,IIfVs(minquantityerror,"\n")&jscheck(xxHasDel)) & """);return false;"
	if totalgoods<minpurchaseamount then print "if((document.forms.checkoutform.mode.value!='dologin')&&(document.forms.checkoutform.mode.value!='donewaccount'))return false;" & vbCrLf
%>dowarning=false;
<%=changechecker%>
if(document.getElementById("cpncode"))document.getElementById("cpncode").value="";
if(dowarning) return !confirm("<%=jscheck(xxWrnChQ)%>");
return true;
}
function preventreturn(tobj,evt){
if(evt.keyCode==13){
	evt.preventDefault();
	if(document.getElementById('zip')&&document.getElementById('zip')===document.activeElement)
		updateestimator('',true);
}
}
/* ]]> */</script>
<%			session.LCID=1033 %>
<input type="hidden" name="estimate" value="<%=FormatNumber(estimate,2,-1,0,0) %>" />
<%			session.LCID=saveLCID
			if (SESSION("clientActions") AND 64)=64 AND checkoutmode="" then %>
				<div class="sharecart" style="text-align:center">
					<input type="button" class="ectbutton widecheckout1" id="sharecartbutton" value="Share Cart" style="width:90%;padding:5px" onclick="dosharecart()" />
					<input type="text" id="sharecartinput" style="width:90%;display:none;padding:5px" />
				</div>
<%			end if
		else
			cartEmpty=TRUE
			print "<div class=""emptycart"">"
			if checkoutmode="savedcart" then
				print "<div class=""emptycartemptylist"">" & xxLisEmp & "</div>"
			else
				print "<div class=""sorrycartempty"">" & xxSryEmp & "</div>"
				print "<div class=""cartemptyclickhere"">" & IIfVr(getget("pli")<>"",xxEmpWis,xxGetSta & " <a class=""ectlink"" href="""&storehomeurl&"""><strong>"&xxClkHere&"</strong></a>") & "</div>"
			end if %>
<script>/* <![CDATA[ */
var ectvalue=Math.floor(Math.random()*10000 + 1);
document.cookie="ECTTESTCART=" + ectvalue + "; path=/<% print IIfVs(request.servervariables("HTTPS")="on","; secure")%>";
if((document.cookie+";").indexOf("ECTTESTCART=" + ectvalue + ";") < 0) document.write("<%=jscheck(xxNoCk & " " & xxSecWar)%>");
/* ]]> */</script>
<noscript><div class="ectwarning"><%=xxNoJS & " " & xxSecWar%></div></noscript>
<%			print "<div class=""emptycartcontinue"">"
			if thefrompage<>"" AND (actionaftercart=2 OR actionaftercart=3) then thehref=htmlspecials(thefrompage) else thehref=storehomeurl
			print imageorlink(imgcontinueshopping, xxCntShp, "", thehref, FALSE)
			print "</div></div>"
		end if
		print "</div>"
		print "</form>"

		print "</div></div>" & vbCrLf ' }
	next
	print "</div>" ' } cartlists
end if ' }
if cartisincluded<>TRUE then
	cnn.Close
	set rs=nothing
	set rs2=nothing
	set cnn=nothing
end if
%>