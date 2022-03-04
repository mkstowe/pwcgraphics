/*
 * This code is copyright (c) ViciSoft SL, all rights reserved.
 * The contents of this file are protected under law as the intellectual property
 * of ViciSoft SL. Any use, reproduction, disclosure or copying
 * of any kind without the express and written permission of ViciSoft SL is forbidden.
 * Author: Vince Reid, vincereid@gmail.com
 *
 * Last Modified: 2021-12-09
 */
var oversldiv;
var gtid;
function displaysavelist(el, evt, twin) {
	oversldiv = false;
	var theevnt = !evt ? twin.event : evt; //IE:FF
	var sld = document.getElementById("savelistdiv");
	var parentdiv = el.parentNode;
	parentdiv.insertBefore(sld, parentdiv.firstChild);
	sld.style.visibility = "visible";
	setTimeout("checksldiv()", 2000);
	return false;
}
function checksldiv() {
	var sld = document.getElementById("savelistdiv");
	if (!oversldiv) sld.style.visibility = "hidden";
}
var notifystockid;
var notifystocktid;
var notifystockoid;
var nsajaxobj;
function notifystockcallback() {
	if (nsajaxobj.readyState == 4) {
		var rstxt = nsajaxobj.responseText;
		if (rstxt != "SUCCESS") alert(rstxt);
		else alert(xxInStNo);
		closeinstock();
	}
}
function regnotifystock() {
	var regex = /[^@]+@[^@]+\.[a-z]{2,}$/i;
	var theemail = document.getElementById("nsemailadd");
	if (!regex.test(theemail.value)) {
		alert(xxValEm);
		theemail.focus();
		return false;
	} else {
		nsajaxobj = window.XMLHttpRequest
			? new XMLHttpRequest()
			: new ActiveXObject("MSXML2.XMLHTTP");
		nsajaxobj.onreadystatechange = notifystockcallback;
		nsajaxobj.open(
			"GET",
			"vsadmin/ajaxservice." +
				extensionabs +
				"?action=notifystock&pid=" +
				encodeURIComponent(notifystockid) +
				"&tpid=" +
				encodeURIComponent(notifystocktid) +
				"&oid=" +
				encodeURIComponent(notifystockoid) +
				"&email=" +
				encodeURIComponent(theemail.value),
			true
		);
		nsajaxobj.send(null);
	}
}
function closeinstock() {
	document.getElementById("notifyinstockcover").style.display = "none";
}
function notifyinstock(isoption, pid, tpid, oid) {
	if (globalquickbuyid !== "") closequickbuy(globalquickbuyid);
	notifystockid = pid;
	notifystocktid = tpid;
	notifystockoid = oid;
	document.getElementById("notifyinstockcover").style.display = "";
	return false;
}
var globallistname = "";
function subformid(tid, listid, listname) {
	if (document.getElementById("ectform" + tid).listid)
		document.getElementById("ectform" + tid).listid.value = listid;
	globallistname = listname;
	if (usehardaddtocart) {
		var tform = document.getElementById("ectform" + tid);
		if (tform.onsubmit()) tform.submit();
	} else ajaxaddcart(tid);
	return false;
}
var globalquickbuyid = "";
function displayquickbuy(qbid) {
	var quid;
	globalquickbuyid = qbid;
	document.getElementById("qbopaque" + qbid).style.display = "";
	if ((quid = document.getElementById("wqb" + qbid + "quant")))
		quid.name = "quant";
	if ((quid = document.getElementById("w" + qbid + "quant"))) quid.name = "";
	return false;
}
function closequickbuy(qbid) {
	var quid;
	if ((quid = document.getElementById("wqb" + qbid + "quant")))
		quid.name = "";
	if ((quid = document.getElementById("w" + qbid + "quant")))
		quid.name = "quant";
	document.getElementById("qbopaque" + qbid).style.display = "none";
	globalquickbuyid = "";
	return false;
}
function ajaxaddcartcb() {
	if (ajaxobj.readyState == 4) {
		var pparam,
			pname,
			pprice,
			pimage,
			optname,
			optvalue,
			retvals = ajaxobj.responseText.split("&");
		try {
			pimage = decodeURIComponent(retvals[6]);
		} catch (err) {
			pimage = "ERROR";
		}
		var schtml =
			'<div class="scart sccheckout">' +
			imgsoftcartcheckout +
			'</div><div class="scart scclose"><a href="#" onclick="document.getElementById(\'opaquediv\').style.display=\'none\';return false"><img src="images/close.gif" alt="' +
			xxClsWin +
			'" /></a></div>' +
			'<div class="scart scprodsadded">' +
			(globallistname != ""
				? xxAddWiL + " " + globallistname
				: xxSCAdOr) +
			"</div>";
		if (retvals[0] != "")
			schtml +=
				'<div class="scart scnostock">' +
				xxNotSto +
				": " +
				decodeURIComponent(retvals[0]) +
				"</div>";
		schtml +=
			'<div class="scimgproducts">' + // Image and products container
			'<div class="scart scimage"><img class="scimage" src="' +
			pimage +
			'" alt="" /></div>' +
			'<div class="scart scproducts">'; // start outer div for products
		var baseind = 7;
		for (var index = 0; index < retvals[3]; index++) {
			try {
				pname = decodeURIComponent(retvals[baseind + 1]);
			} catch (err) {
				pname = "ERROR";
			}
			try {
				pprice = decodeURIComponent(retvals[baseind + 3]);
			} catch (err) {
				pprice = "ERROR";
			}
			schtml +=
				'<div class="scart scproduct"><div class="scart scprodname"> ' +
				retvals[baseind + 2] +
				" " +
				pname +
				" " +
				xxHasAdd +
				"</div>";
			var prhtml =
				"<div " +
				(nopriceanywhere ? 'style="display:none" ' : "") +
				'class="scart scprice">' +
				xxPrice +
				(xxPrice != "" ? ":" : "") +
				pprice +
				"</div>";
			var numoptions = retvals[baseind + 5];
			baseind += 6;
			if (numoptions > 0) {
				schtml +=
					"<div " +
					(numoptions > 10
						? 'style="height:200px;overflow-y:scroll" '
						: "") +
					'class="scart scoptions">';
				for (var index2 = 0; index2 < numoptions; index2++) {
					try {
						optname = decodeURIComponent(retvals[baseind++]);
					} catch (err) {
						optname = "ERROR";
					}
					try {
						optvalue = decodeURIComponent(retvals[baseind++]);
					} catch (err) {
						optvalue = "ERROR";
					}
					schtml +=
						'<div class="scart scoption"><div class="scart scoptname">- ' +
						optname +
						':</div><div class="scart scoptvalue">' +
						optvalue +
						"</div></div>";
				}
				schtml += "</div>";
			}
			schtml += prhtml + "</div>";
		}
		schtml +=
			"</div>" + // end outer div for products
			"</div>" + // end image and products container
			"<div>";
		try {
			pprice = decodeURIComponent(retvals[5]);
		} catch (err) {
			pprice = "ERROR";
		}
		if (retvals[1] == 1)
			schtml += '<div class="scart scnostock">' + xxSCStkW + "</div>";
		if (retvals[2] == 1)
			schtml += '<div class="scart scbackorder">' + xxSCBakO + "</div>";
		schtml +=
			'<div class="scart sccartitems">' +
			xxCarCon +
			":" +
			retvals[4] +
			" " +
			xxSCItem +
			"</div>" +
			"<div " +
			(nopriceanywhere ? 'style="display:none" ' : "") +
			'class="scart sccarttotal"><span style="display:none" id="sccartdscnt" class="scart sccartdscnt">(' +
			xxDscnts +
			':<span id="sccartdscamnt" class="sccartdscamnt"></span>) / </span>' +
			(showtaxinclusive != 0
				? xxCntTax +
				  ':<span id="sccarttax" class="sccarttax"></span> / '
				: "") +
			xxSCCarT +
			":" +
			pprice +
			"</div>" +
			'<div class="scart sclinks"><a class="ectlink scclink" href="#" onclick="document.getElementById(\'opaquediv\').style.display=\'none\';return false">' +
			xxCntShp +
			'</a> | <a class="ectlink scclink" href="' +
			(cartpageonhttps ? storeurlssl : "") +
			"cart" +
			extension +
			"\" onclick=\"document.getElementById('opaquediv').style.display='none';return true\">" +
			xxEdiOrd +
			"</a></div>" +
			"</div>";
		if (softcartrelated) {
			screlated(retvals[7]);
			schtml += '<div id="softcartrelated" style="display:none"></div>';
		}
		document.getElementById("scdiv").innerHTML = schtml;
		if (document.getElementsByClassName) {
			var ectMCpm = document.getElementsByClassName("ectMCquant");
			for (var index = 0; index < ectMCpm.length; index++)
				ectMCpm[index].innerHTML = retvals[4];
			ectMCpm = document.getElementsByClassName("ectMCship");
			for (var index = 0; index < ectMCpm.length; index++)
				ectMCpm[index].innerHTML =
					'<a href="' +
					(cartpageonhttps ? storeurlssl : "") +
					"cart" +
					extension +
					'">' +
					xxClkHere +
					"</a>";
			ectMCpm = document.getElementsByClassName("ectMCtot");
			for (var index = 0; index < ectMCpm.length; index++)
				ectMCpm[index].innerHTML = pprice;
			if (retvals.length > baseind) {
				try {
					pparam = decodeURIComponent(retvals[baseind++]);
				} catch (err) {
					pparam = "-";
				}
				if ((ectMCpm = document.getElementById("sccarttax")))
					ectMCpm.innerHTML = pparam;
				try {
					pparam = decodeURIComponent(retvals[baseind++]);
				} catch (err) {
					pparam = "-";
				}
				var ectMCpm = document.getElementsByClassName("mcMCdsct");
				for (var index = 0; index < ectMCpm.length; index++)
					ectMCpm[index].innerHTML = pparam;
				document.getElementById("sccartdscamnt").innerHTML = pparam;
				try {
					pparam = decodeURIComponent(retvals[baseind++]);
				} catch (err) {
					pparam = "-";
				}
				var ectMCpm = document.getElementsByClassName("ecHidDsc");
				for (var index = 0; index < ectMCpm.length; index++)
					ectMCpm[index].style.display = pparam == "0" ? "none" : "";
				document.getElementById("sccartdscnt").style.display =
					pparam == "0" ? "none" : "";
				try {
					pparam = decodeURIComponent(retvals[baseind++]);
				} catch (err) {
					pparam = "-";
				}
				var ectMCpm = document.getElementsByClassName("mcLNitems");
				for (var index = 0; index < ectMCpm.length; index++)
					ectMCpm[index].innerHTML = pparam;
			}
		}
	}
}
var scrajaxobj;
function ajaxscrelatedcb() {
	if (scrajaxobj.readyState == 4) {
		if ((tid = document.getElementById("softcartrelated"))) {
			tid.innerHTML += scrajaxobj.responseText;
			tid.style.display = "";
		}
	}
}
function screlated(prodid) {
	scrajaxobj = window.XMLHttpRequest
		? new XMLHttpRequest()
		: new ActiveXObject("MSXML2.XMLHTTP");
	scrajaxobj.onreadystatechange = ajaxscrelatedcb;
	scrajaxobj.open(
		"POST",
		"vsadmin/ajaxservice." + extensionabs + "?action=screlated",
		true
	);
	scrajaxobj.setRequestHeader(
		"Content-type",
		"application/x-www-form-urlencoded"
	);
	scrajaxobj.send("prodid=" + encodeURIComponent(prodid));
}
function ajaxaddcart(frmid) {
	var elem = document.getElementById("ectform" + frmid).elements;
	var str = "";
	var postdata = "ajaxadd=true";
	eval(
		"var isvalidfm=formvalidator" +
			frmid +
			"(document.getElementById('ectform" +
			frmid +
			"'))"
	);
	if (isvalidfm) {
		for (var ecti = 0; ecti < elem.length; ecti++) {
			if (elem[ecti].disabled) {
			} else if (elem[ecti].type == "select-one") {
				if (elem[ecti].value != "")
					postdata += "&" + elem[ecti].name + "=" + elem[ecti].value;
			} else if (
				elem[ecti].type == "text" ||
				elem[ecti].type == "textarea" ||
				elem[ecti].type == "hidden"
			) {
				if (elem[ecti].value != "")
					postdata +=
						"&" +
						elem[ecti].name +
						"=" +
						encodeURIComponent(elem[ecti].value);
			} else if (
				elem[ecti].type == "radio" ||
				elem[ecti].type == "checkbox"
			) {
				if (elem[ecti].checked)
					postdata += "&" + elem[ecti].name + "=" + elem[ecti].value;
			}
		}
		if (document.getElementById("qbopaque" + frmid)) closequickbuy(frmid);
		ajaxobj = window.XMLHttpRequest
			? new XMLHttpRequest()
			: new ActiveXObject("MSXML2.XMLHTTP");
		ajaxobj.onreadystatechange = ajaxaddcartcb;
		ajaxobj.open(
			"POST",
			"vsadmin/shipservice." + extensionabs + "?action=addtocart",
			true
		);
		ajaxobj.setRequestHeader(
			"Content-type",
			"application/x-www-form-urlencoded"
		);
		ajaxobj.send(postdata);
		document.getElementById("opaquediv").innerHTML =
			'<div id="scdiv" class="scart scwrap"><img src="images/preloader.gif" alt="" style="margin:40px" /></div>';
		document.getElementById("opaquediv").style.display = "";
	}
}
var op = [], // Option Price Difference
	aIM = [],
	aIML = [], // Option Alternate Image
	dOP = [], // Dependant Options
	dIM = [], // Default Image
	pIM = [],
	pIML = [], // Product Image
	pIX = [], // Product Image Index
	ot = [], // Option Text
	pp = [], // Product Price
	pl = [], // List Price
	pi = [], // Alternate Product Image
	or = [], // Option Alt Id
	cp = [], // Current Price
	oos = [], // Option Out of Stock Id
	rid = [], // Resulting product Id
	otid = [], // Original product Id
	opttype = [],
	optperc = [],
	optmaxc = [],
	optacpc = [],
	fid = [],
	oS = [],
	ps = [];
function checkStock(x, i, backorder) {
	if (i != "" && (oS[i] > 0 || or[i])) return true;
	if (backorder && (globBakOrdChk || confirm(xxBakOpt)))
		return (globBakOrdChk = true);
	if (notifybackinstock)
		notifyinstock(true, x.form.id.value, x.form.id.value, i);
	else alert(xxOptOOS);
	x.focus();
	return false;
}
function dummyfunc() {}
function pricechecker(cnt, i) {
	if (i != "" && i in op && !isNaN(op[i])) return op[i];
	return 0; // Safari
}
function regchecker(cnt, i) {
	if (i != "") return or[i];
	return "";
}
function enterValue(x) {
	ectaddclass(x, "ectwarning");
	alert(xxPrdEnt);
	x.focus();
	return false;
}
function invalidChars(x) {
	alert(xxInvCha + " " + x);
	return false;
}
function enterDigits(x) {
	ectaddclass(x, "ectwarning");
	alert(xxDigits);
	x.focus();
	return false;
}
function removemultiwarning(pdiv) {
	var ti = pdiv.getElementsByClassName("ecttextinput");
	for (var i = 0; i < ti.length; i++) ectremoveclass(ti[i], "ectwarning");
}
function enterMultValue(pdiv) {
	var ti = pdiv.getElementsByClassName("ecttextinput");
	for (var i = 0; i < ti.length; i++) ectaddclass(ti[i], "ectwarning");
	alert(xxEntMul);
	return false;
}
function chooseOption(x) {
	if (x.type == "select-one") ectaddclass(x, "ectwarning");
	else {
		var ti = x.parentNode.parentNode.getElementsByTagName("INPUT");
		for (var i = 0; i < ti.length; i++) ectaddclass(ti[i], "ectwarning");
	}
	alert(xxPrdChs);
	x.focus();
	return false;
}
function dataLimit(x, numchars) {
	ectaddclass(x, "ectwarning");
	alert(xxPrd255.replace(255, numchars));
	x.focus();
	return false;
}
var hiddencurr = "";
function addCommas(ns, decs, thos) {
	ns = ns.toString().replace(/\./, decs);
	if ((dpos = ns.indexOf(decs)) < 0) dpos = ns.length;
	dpos -= 3;
	while (dpos > 0) {
		ns = ns.substr(0, dpos) + thos + ns.substr(dpos);
		dpos -= 3;
	}
	return ns;
}
function formatprice(i, currcode, currformat) {
	currcode = currcode || "";
	currformat = currformat || "";
	i = Math.round(i * 100) / 100;
	if (hiddencurr == "")
		hiddencurr = document.getElementById("hiddencurr").value;
	var pTemplate = hiddencurr;
	if (currcode != "")
		pTemplate =
			" " +
			zero2dps +
			(currcode != " " ? "<strong>" + currcode + "</strong>" : "");
	if (currcode == " JPY" || (!hasdecimals && currcode == ""))
		i = Math.round(i).toString();
	else if (hasdecimals) {
		if (i == Math.round(i)) i = i.toString() + ".00";
		else if (i * 10.0 == Math.round(i * 10.0)) i = i.toString() + "0";
		else if (i * 100.0 == Math.round(i * 100.0)) i = i.toString();
	}
	i = addCommas(i, currDecimalSep, currThousandsSep);
	if (currcode != "")
		pTemplate = currformat.toString().replace(/%s/, i.toString());
	else pTemplate = pTemplate.toString().replace(/\d[,.]*\d*/, i.toString());
	return pTemplate;
}
function vsdecimg(timg) {
	return decodeURIComponent(
		noencodeimages
			? timg
			: timg
					.replace("|", "prodimages/")
					.replace("<", ".gif")
					.replace(">", ".jpg")
					.replace("?", ".png")
	);
}
function updateprodimage(theitem, isnext) {
	return updateprodimage2(false, theitem, isnext);
}
function sz(szid, szprice, szlist, szimage, szstock) {
	if (usestockmanagement) ps[szid] = szstock;
	pp[szid] = szprice;
	pl[szid] = szlist;
	if (szimage != "") pi[szid] = szimage;
}
function gfid(tid) {
	if (tid in fid) return fid[tid];
	fid[tid] = document.getElementById(tid);
	return fid[tid];
}
function applyreg(arid, arreg) {
	if (arreg && arreg != "") {
		arreg = arreg.replace("%s", arid);
		if (arreg.indexOf(" ") > 0) {
			var ida = arreg.split(" ", 2);
			arid = arid.replace(ida[0], ida[1]);
		} else arid = arreg;
	}
	return arid;
}
function getaltid(theid, optns, prodnum, optnum, optitem, numoptions) {
	var thereg = "";
	for (var index = 0; index < numoptions; index++) {
		if (Math.abs(opttype[index]) == 4) {
			thereg = or[optitem];
		} else if (Math.abs(opttype[index]) == 2) {
			if (optnum == index) thereg = or[optns.options[optitem].value];
			else {
				var opt = gfid("optn" + prodnum + "x" + index);
				if (!opt.disabled)
					thereg = or[opt.options[opt.selectedIndex].value];
			}
		} else if (Math.abs(opttype[index]) == 1) {
			opt = document.getElementsByName("optn" + prodnum + "x" + index);
			if (optnum == index) {
				thereg = or[opt[optitem].value];
			} else {
				for (var y = 0; y < opt.length; y++)
					if (opt[y].checked && !opt[y].disabled)
						thereg = or[opt[y].value];
			}
		} else continue;
		theid = applyreg(theid, thereg);
	}
	return theid;
}
function getnonaltpricediff(
	optns,
	prodnum,
	optnum,
	optitem,
	numoptions,
	theoptprice
) {
	var nonaltdiff = 0;
	for (index = 0; index < numoptions; index++) {
		var optid = "";
		if (Math.abs(opttype[index]) == 4) {
			optid = optitem;
		} else if (Math.abs(opttype[index]) == 2) {
			if (optnum == index) optid = optns.options[optitem].value;
			else {
				var opt = gfid("optn" + prodnum + "x" + index);
				if (opt.style.display == "none" || opt.disabled) continue;
				optid = opt.options[opt.selectedIndex].value;
			}
		} else if (Math.abs(opttype[index]) == 1) {
			var opt = document.getElementsByName(
				"optn" + prodnum + "x" + index
			);
			if (optnum == index) optid = opt[optitem].value;
			else {
				for (var y = 0; y < opt.length; y++) {
					if (
						opt[y].checked &&
						opt[y].style.display != "none" &&
						!opt[y].disabled
					)
						optid = opt[y].value;
				}
			}
		} else continue;
		if (!or[optid] && optid in op && !isNaN(op[optid]))
			if (optperc[index])
				//isNaN for Safari Bug
				nonaltdiff += (op[optid] * theoptprice) / 100.0;
			else nonaltdiff += op[optid];
	}
	return nonaltdiff;
}
function ectaddclass(elem, classname) {
	elem.classList.add(classname);
}
function ectremoveclass(elem, classname) {
	elem.classList.remove(classname);
}
function ecttoggleclass(elem, classname) {
	if (elem.className.indexOf(classname) < 0) elem.classList.add(classname);
	else elem.classList.remove(classname);
}
function updateprice(
	numoptions,
	prodnum,
	prodprice,
	listprice,
	origid,
	thetax,
	stkbyopts,
	taxexmpt,
	backorder
) {
	var baseid = origid,
		origprice = prodprice,
		canresolve = true,
		hasiteminstock = false,
		hasaltid = false,
		hasmultioption = false,
		canresolve = true,
		allbutlastselected = true;
	oos[prodnum] = "";
	if (typeof stockdisplaythreshold === "undefined") stockdisplaythreshold = 0;
	for (cnt = 0; cnt < numoptions; cnt++) {
		if (Math.abs(opttype[cnt]) == 2) {
			optns = gfid("optn" + prodnum + "x" + cnt);
			if (!optns.disabled)
				baseid = applyreg(
					baseid,
					regchecker(
						prodnum,
						optns.options[optns.selectedIndex].value
					)
				);
			if (
				optns.options[optns.selectedIndex].value == "" &&
				cnt < numoptions - 1
			)
				allbutlastselected = false;
		} else if (Math.abs(opttype[cnt]) == 1) {
			optns = document.getElementsByName("optn" + prodnum + "x" + cnt);
			var hasonechecked = false;
			for (var i = 0; i < optns.length; i++) {
				if (optns[i].checked && !optns[i].disabled) {
					hasonechecked = true;
					baseid = applyreg(
						baseid,
						regchecker(prodnum, optns[i].value)
					);
				}
			}
			if (!hasonechecked && cnt < numoptions - 1)
				allbutlastselected = false;
		}
		if (baseid in pp) prodprice = pp[baseid];
		if (baseid in pl) listprice = pl[baseid];
	}
	var baseprice = prodprice;
	for (cnt = 0; cnt < numoptions; cnt++) {
		if (Math.abs(opttype[cnt]) == 2) {
			optns = gfid("optn" + prodnum + "x" + cnt);
			if (optns.disabled) continue;
			prodprice += optperc[cnt]
				? (baseprice *
						pricechecker(
							prodnum,
							optns.options[optns.selectedIndex].value
						)) /
				  100.0
				: pricechecker(
						prodnum,
						optns.options[optns.selectedIndex].value
				  );
		} else if (Math.abs(opttype[cnt]) == 1) {
			optns = document.getElementsByName("optn" + prodnum + "x" + cnt);
			for (var i = 0; i < optns.length; i++) {
				if (
					optns[i].checked &&
					optns[i].style.display != "none" &&
					!optns[i].disabled
				)
					prodprice += optperc[cnt]
						? (baseprice * pricechecker(prodnum, optns[i].value)) /
						  100.0
						: pricechecker(prodnum, optns[i].value);
			}
		}
	}
	var totalprice = prodprice;
	var prodtax =
		showtaxinclusive == 2
			? !taxexmpt
				? (prodprice * thetax) / 100.0
				: 0
			: 0;
	for (cnt = 0; cnt < numoptions; cnt++) {
		if (Math.abs(opttype[cnt]) == 2) {
			var optns = gfid("optn" + prodnum + "x" + cnt);
			for (var i = 0; i < optns.length; i++) {
				if (optns.options[i].value != "") {
					theid = origid;
					optns.options[i].text = ot[optns.options[i].value];
					theid = getaltid(theid, optns, prodnum, cnt, i, numoptions);
					theoptprice = theid in pp ? pp[theid] : origprice;
					if (
						pi[theid] &&
						pi[theid] != "" &&
						or[optns.options[i].value]
					) {
						aIM[optns.options[i].value] = pi[theid].split("*")[0];
						if (pi[theid].split("*")[1])
							aIML[optns.options[i].value] =
								pi[theid].split("*")[1];
					}
					if (usestockmanagement) {
						theoptstock =
							(theid in ps && or[optns.options[i].value]) ||
							!stkbyopts
								? ps[theid]
								: oS[optns.options[i].value];
						if (theoptstock > 0) hasiteminstock = true;
						if (theoptstock <= 0 && optns.selectedIndex == i) {
							oos[prodnum] = "optn" + prodnum + "x" + cnt;
							rid[prodnum] = theid;
							otid[prodnum] = origid;
						}
					}
					canresolve =
						!or[optns.options[i].value] || theid in pp
							? true
							: false;
					var staticpricediff = getnonaltpricediff(
						optns,
						prodnum,
						cnt,
						i,
						numoptions,
						theoptprice
					);
					theoptpricediff =
						theoptprice + staticpricediff - totalprice;
					if (!noprice && !hideoptpricediffs) {
						if (Math.round(theoptpricediff * 100) != 0)
							optns.options[i].text +=
								" (" +
								(absoptionpricediffs
									? ""
									: theoptpricediff > 0
									? "+"
									: "-") +
								formatprice(
									Math.abs(
										Math.round(
											((absoptionpricediffs
												? prodprice + prodtax
												: 0) +
												theoptpricediff +
												(showtaxinclusive == 2
													? !taxexmpt
														? (theoptpricediff *
																thetax) /
														  100.0
														: 0
													: 0)) *
												100
										) / 100.0
									)
								) +
								")";
					}
					if (
						usestockmanagement &&
						showinstock &&
						!noshowoptionsinstock &&
						(theoptstock < stockdisplaythreshold ||
							stockdisplaythreshold == 0)
					)
						if (stkbyopts && canresolve)
							optns.options[i].text += xxOpSkTx.replace(
								"%s",
								Math.max(theoptstock, 0)
							);
					if (
						usestockmanagement
							? theoptstock > 0 || !stkbyopts || !canresolve
							: true
					)
						ectremoveclass(optns.options[i], "oostock");
					else ectaddclass(optns.options[i], "oostock");
					if (
						allbutlastselected &&
						cnt == numoptions - 1 &&
						!canresolve
					)
						ectaddclass(optns.options[i], "oostock");
					if (or[optns.options[i].value]) hasaltid = true;
				}
			}
		} else if (Math.abs(opttype[cnt]) == 1) {
			optns = document.getElementsByName("optn" + prodnum + "x" + cnt);
			for (var i = 0; i < optns.length; i++) {
				theid = origid;
				optn = gfid("optn" + prodnum + "x" + cnt + "y" + i);
				optn.innerHTML = ot[optns[i].value];
				theid = getaltid(theid, optns, prodnum, cnt, i, numoptions);
				theoptprice = theid in pp ? pp[theid] : origprice;
				if (pi[theid] && pi[theid] != "" && or[optns[i].value]) {
					aIM[optns[i].value] = pi[theid].split("*")[0];
					if (pi[theid].split("*")[1])
						aIML[optns[i].value] = pi[theid].split("*")[1];
				}
				if (usestockmanagement) {
					theoptstock =
						(theid in ps && or[optns[i].value]) || !stkbyopts
							? ps[theid]
							: oS[optns[i].value];
					if (theoptstock > 0) hasiteminstock = true;
					if (theoptstock <= 0 && optns[i].checked) {
						oos[prodnum] = "optn" + prodnum + "x" + cnt + "y" + i;
						rid[prodnum] = theid;
						otid[prodnum] = origid;
					}
				}
				canresolve = !or[optns[i].value] || theid in pp ? true : false;
				var staticpricediff = getnonaltpricediff(
					optns,
					prodnum,
					cnt,
					i,
					numoptions,
					theoptprice
				);
				theoptpricediff = theoptprice + staticpricediff - totalprice;
				if (!noprice && !hideoptpricediffs) {
					if (Math.round(theoptpricediff * 100) != 0)
						optn.innerHTML +=
							" (" +
							(absoptionpricediffs
								? ""
								: theoptpricediff > 0
								? "+"
								: "-") +
							formatprice(
								Math.abs(
									Math.round(
										((absoptionpricediffs
											? prodprice + prodtax
											: 0) +
											theoptpricediff +
											(showtaxinclusive == 2
												? !taxexmpt
													? (theoptpricediff *
															thetax) /
													  100.0
													: 0
												: 0)) *
											100
									) / 100.0
								)
							) +
							")";
				}
				if (
					usestockmanagement &&
					showinstock &&
					!noshowoptionsinstock &&
					(theoptstock < stockdisplaythreshold ||
						stockdisplaythreshold == 0)
				)
					if (stkbyopts && canresolve)
						optn.innerHTML += xxOpSkTx.replace(
							"%s",
							Math.max(theoptstock, 0)
						);
				if (
					usestockmanagement
						? theoptstock > 0 || !stkbyopts || !canresolve
						: true
				)
					ectremoveclass(optn, "oostock");
				else ectaddclass(optn, "oostock");
				if (allbutlastselected && cnt == numoptions - 1 && !canresolve)
					ectaddclass(optn, "oostock");
				if (or[optns[i].value]) hasaltid = true;
			}
		} else if (Math.abs(opttype[cnt]) == 4) {
			var tstr = "optm" + prodnum + "x" + cnt + "y";
			var tlen = tstr.length;
			var optns = document.getElementsByTagName("input");
			hasmultioption = true;
			for (var i = 0; i < optns.length; i++) {
				if (optns[i].id.substr(0, tlen) == tstr) {
					theid = origid;
					var oid = optns[i].name.substr(4);
					var optn = optns[i];
					var optnt = gfid(optns[i].id.replace(/optm/, "optx"));
					optnt.innerHTML = "&nbsp;- " + ot[oid];
					theid = getaltid(
						theid,
						optns,
						prodnum,
						cnt,
						oid,
						numoptions
					);
					theoptprice = theid in pp ? pp[theid] : origprice;
					if (usestockmanagement) {
						theoptstock =
							(theid in ps && or[oid]) || !stkbyopts
								? ps[theid]
								: oS[oid];
						if (theoptstock > 0) hasiteminstock = true;
						if (theoptstock <= 0 && optns[i].checked) {
							oos[prodnum] =
								"optm" + prodnum + "x" + cnt + "y" + i;
							rid[prodnum] = theid;
							otid[prodnum] = origid;
						}
						canresolve =
							!or[oid] || applyreg(theid, or[oid]) in ps
								? true
								: false;
					}
					var staticpricediff = getnonaltpricediff(
						optns,
						prodnum,
						cnt,
						oid,
						numoptions,
						theoptprice
					);
					theoptpricediff =
						theoptprice + staticpricediff - totalprice;
					if (!noprice && !hideoptpricediffs) {
						if (Math.round(theoptpricediff * 100) != 0)
							optnt.innerHTML +=
								" (" +
								(absoptionpricediffs
									? ""
									: theoptpricediff > 0
									? "+"
									: "-") +
								formatprice(
									Math.abs(
										Math.round(
											((absoptionpricediffs
												? prodprice + prodtax
												: 0) +
												theoptpricediff +
												(showtaxinclusive == 2
													? !taxexmpt
														? (theoptpricediff *
																thetax) /
														  100.0
														: 0
													: 0)) *
												100
										) / 100.0
									)
								) +
								")";
					}
					if (
						usestockmanagement &&
						showinstock &&
						!noshowoptionsinstock &&
						(theoptstock < stockdisplaythreshold ||
							stockdisplaythreshold == 0)
					)
						if (
							stkbyopts &&
							canresolve &&
							!(or[oid] && theoptstock <= 0)
						)
							optnt.innerHTML += xxOpSkTx.replace(
								"%s",
								Math.max(theoptstock, 0)
							);
					if (usestockmanagement)
						if (
							theoptstock > 0 ||
							(or[oid] && !canresolve) ||
							backorder
						) {
							ectremoveclass(optn, "oostock");
							optn.disabled = false;
							optn.style.backgroundColor = "#FFFFFF";
						} else {
							ectaddclass(optn, "oostock");
							optn.disabled = true;
							optn.style.backgroundColor = "#EBEBE4";
						}
					if (or[oid]) hasaltid = true;
				}
			}
		}
	}
	if (hasmultioption) oos[prodnum] = "";
	if ((!cp[prodnum] || cp[prodnum] == 0) && prodprice == 0) return;
	cp[prodnum] = prodprice;
	var lpt = xxListPrice,
		yst = yousavetext;
	if (!noprice) {
		var qbprefix;
		for (var qbind = 0; qbind <= 1; qbind++) {
			qbprefix = qbind == 0 ? "" : "qb";
			if (document.getElementById(qbprefix + "taxmsg" + prodnum))
				document.getElementById(
					qbprefix + "taxmsg" + prodnum
				).style.display = "";
			if (!noupdateprice)
				if (document.getElementById(qbprefix + "pricediv" + prodnum))
					document.getElementById(
						qbprefix + "pricediv" + prodnum
					).innerHTML =
						pricezeromessage != "" && prodprice == 0
							? pricezeromessage
							: formatprice(
									prodprice +
										(showtaxinclusive == 2
											? !taxexmpt
												? (prodprice * thetax) / 100.0
												: 0
											: 0)
							  );
			if (showtaxinclusive == 1 || ectbody3layouttaxinc) {
				if (!taxexmpt && prodprice != 0) {
					if (
						document.getElementById(
							qbprefix + "pricedivti" + prodnum
						)
					)
						document.getElementById(
							qbprefix + "pricedivti" + prodnum
						).innerHTML = formatprice(
							prodprice + (prodprice * thetax) / 100.0
						);
				} else {
					if (document.getElementById(qbprefix + "taxmsg" + prodnum))
						document.getElementById(
							qbprefix + "taxmsg" + prodnum
						).style.display = "none";
				}
			}

			if (
				(currRate1 != 0 && currSymbol1 != "") ||
				(currRate2 != 0 && currSymbol2 != "") ||
				(currRate3 != 0 && currSymbol3 != "")
			) {
				if (
					document.getElementById(qbprefix + "pricedivec" + prodnum)
				) {
					document.getElementById(
						qbprefix + "pricedivec" + prodnum
					).innerHTML =
						prodprice == 0
							? ""
							: (currRate1 != 0 && currSymbol1 != ""
									? formatprice(
											prodprice * currRate1,
											currSymbol1,
											currFormat1
									  ) + currencyseparator
									: "") +
							  (currRate2 != 0 && currSymbol2 != ""
									? formatprice(
											prodprice * currRate2,
											currSymbol2,
											currFormat2
									  ) + currencyseparator
									: "") +
							  (currRate3 != 0 && currSymbol3 != ""
									? formatprice(
											prodprice * currRate3,
											currSymbol3,
											currFormat3
									  )
									: "");
				}
			}
			if (document.getElementById(qbprefix + "listdivec" + prodnum)) {
				var nlp =
					baseprice > 0
						? prodprice * (listprice / baseprice)
						: listprice;
				document.getElementById(
					qbprefix + "listdivec" + prodnum
				).style.display = nlp > prodprice ? "" : "none";
				if (showtaxinclusive == 2)
					nlp += !taxexmpt ? (nlp * thetax) / 100.0 : 0;
				var ysp =
					nlp -
					(prodprice +
						(showtaxinclusive == 2
							? !taxexmpt
								? (prodprice * thetax) / 100.0
								: 0
							: 0));
				document.getElementById(
					qbprefix + "listdivec" + prodnum
				).innerHTML =
					lpt.replace(/%s/, formatprice(nlp)) +
					(ysp > 0 ? yst.replace(/%s/, formatprice(ysp)) : "");
			}
		}
	}
	if (usestockmanagement && stkbyopts && hasaltid && !hasiteminstock) {
		var buttonid = document.getElementById("ectaddcart" + prodnum);
		if (buttonid && buttonid.type == "button") {
			if (notifybackinstock) buttonid.innerHTML = xxNotBaS;
			else {
				buttonid.innerHTML = xxOutStok;
				buttonid.disabled = true;
			}
		}
	}
}
function dependantopts(frmnum) {
	var objid,
		thisdep,
		depopt = "",
		grpid,
		alldeps = [];
	var allformelms = document.getElementById("ectform" + frmnum).elements;
	for (var iallelems = 0; iallelems < allformelms.length; iallelems++) {
		objid = allformelms[iallelems];
		thisdep = "";
		if (objid.type == "select-one") {
			thisdep = dOP[objid[objid.selectedIndex].value];
		} else if (objid.type == "checkbox" || objid.type == "radio") {
			if (objid.checked) thisdep = dOP[objid.value];
		}
		if (thisdep) alldeps = alldeps.concat(thisdep);
	}
	for (var iallelems = 0; iallelems < allformelms.length; iallelems++) {
		objid = allformelms[iallelems];
		if ((grpid = parseInt(objid.getAttribute("data-optgroup")))) {
			if (objid.getAttribute("data-isdep")) {
				var isdisabled = alldeps.indexOf(grpid) < 0,
					haschanged = isdisabled != objid.disabled;
				objid.disabled = isdisabled;
				var maxindex = 0,
					cobj = objid.parentNode;
				for (var cind = 0; cind < 3; cind++) {
					cobj = cobj.parentNode;
					if (cobj.id.substr(0, 4) == "divc") {
						cobj.style.display = isdisabled ? "none" : "";
						break;
					}
				}
				if (haschanged) {
					if (objid.onchange) objid.onchange();
					else if (objid.onclick) objid.onclick();
				}
			}
		}
	}
}
var globBakOrdChk;
function ectvalidate(theForm, numoptions, prodnum, stkbyopts, backorder) {
	(globBakOrdChk = false), (oneoutofstock = false);
	for (cnt = 0; cnt < numoptions; cnt++) {
		if (Math.abs(opttype[cnt]) == 4) {
			var validmulti = "";
			var cntnr = document.getElementById("divc" + prodnum + "x" + cnt);
			if (cntnr && cntnr.style.display == "none") continue;
			var intreg = /^(\d*)$/;
			var inputs = theForm.getElementsByTagName("input");
			var tt = "";
			for (var i = 0; i < inputs.length; i++) {
				if (
					inputs[i].type == "text" &&
					inputs[i].id.substr(0, 4) == "optm"
				) {
					ectremoveclass(inputs[i], "ectwarning");
					validmulti = inputs[i];
					if (!inputs[i].value.match(intreg))
						return enterDigits(inputs[i]);
					tt += inputs[i].value;
					if (usestockmanagement)
						if (
							inputs[i].value != "" &&
							oS[inputs[i].name.substr(4)] <= 0
						)
							oneoutofstock = true;
				}
			}
			if (tt == "")
				return enterMultValue(validmulti.parentNode.parentNode);
		} else if (Math.abs(opttype[cnt]) == 3 || Math.abs(opttype[cnt]) == 5) {
			var voptn = eval("theForm.voptn" + cnt);
			if (voptn.disabled) continue;
			if (optacpc[cnt].length > 0) {
				try {
					var re = new RegExp("[" + optacpc[cnt] + "]", "g");
				} catch (err) {
					alert(err.message);
				}
				if (voptn.value.replace(re, "") != "")
					return invalidChars(voptn.value.replace(re, ""));
			}
			if ((opttype[cnt] == 3 || opttype[cnt] == 5) && voptn.value == "")
				return enterValue(voptn);
			if (
				voptn.value.length >
				(optmaxc[cnt] > 0 ? optmaxc[cnt] : txtcollen)
			)
				return dataLimit(
					voptn,
					optmaxc[cnt] > 0 ? optmaxc[cnt] : txtcollen
				);
			ectremoveclass(voptn, "ectwarning");
		} else if (Math.abs(opttype[cnt]) == 2) {
			optn = document.getElementById("optn" + prodnum + "x" + cnt);
			if (optn.disabled) continue;
			if (opttype[cnt] == 2) {
				if (optn.selectedIndex == 0)
					return chooseOption(eval("theForm.optn" + cnt));
			}
			if (stkbyopts && optn.options[optn.selectedIndex].value != "") {
				if (
					!checkStock(
						optn,
						optn.options[optn.selectedIndex].value,
						backorder
					)
				)
					return false;
			}
		} else if (Math.abs(opttype[cnt]) == 1) {
			havefound = "";
			optns = document.getElementsByName("optn" + prodnum + "x" + cnt);
			if (optns[0].disabled) continue;
			if (opttype[cnt] == 1) {
				for (var i = 0; i < optns.length; i++)
					if (optns[i].checked) havefound = optns[i].value;
				if (havefound == "") return chooseOption(optns[0]);
			}
			if (stkbyopts) {
				if (havefound != "") {
					if (!checkStock(optns[0], havefound, backorder))
						return false;
				}
			}
		}
	}
	if (usestockmanagement) {
		if (backorder && oneoutofstock && !globBakOrdChk) {
			if (!confirm(xxBakOpt)) return false;
		}
	}
	if (oos[prodnum] && oos[prodnum] != "" && !backorder) {
		if (notifybackinstock)
			notifyinstock(true, otid[prodnum], rid[prodnum], 0);
		else alert(xxOptOOS);
		document.getElementById(oos[prodnum]).focus();
		return false;
	}
	return true;
}
function quantup(tobjid, qud) {
	tobj = document.getElementById("w" + tobjid + "quant");
	if (isNaN(parseInt(tobj.value))) tobj.value = 1;
	else if (qud == 1) tobj.value = parseInt(tobj.value) + 1;
	else tobj.value = Math.max(1, parseInt(tobj.value) - 1);
	if (document.getElementById("qnt" + tobjid + "x"))
		document.getElementById("qnt" + tobjid + "x").value = tobj.value;
}
function ectgocheck(tloc) {
	if (
		tloc.substr(0, 1).toLowerCase() == "/" ||
		tloc.substr(0, 5).toLowerCase() == "http:" ||
		tloc.substr(0, 6).toLowerCase() == "https:"
	)
		ectgoabs(tloc);
	else ectgonoabs(tloc);
}
function ectgoabs(tloc) {
	document.location = tloc;
}
function ectgonoabs(tloc) {
	document.location = (
		((ECTbh = document.getElementsByTagName("base")).length > 0 &&
		tloc.charAt(0) != "/"
			? ECTbh[0].href + "/"
			: "") + tloc
	).replace(/([^:]\/)\/+/g, "$1");
}
// ECT Slider
window.slidertimeout = [];
window.slide_index = [];
window.slide_repeat = [];
function changeectslider(n, sliderid) {
	ect_displayslider((window.slide_index[sliderid] += n), sliderid);
}
function ect_displayslider(n, sliderid) {
	var slides;
	if (sliderid != "")
		slides = document
			.getElementById(sliderid)
			.getElementsByClassName("sliderimages");
	else slides = document.getElementsByClassName("sliderimages");
	//document.getElementById('debugdiv').innerHTML+='Slider ID: ' + sliderid + "<br>";
	if (slides.length > 0) {
		if (n > slides.length) {
			window.slide_index[sliderid] = 1;
		}
		if (n < 1) {
			window.slide_index[sliderid] = slides.length;
		}
		for (var i = 0; i < slides.length; i++) {
			slides[i].style.opacity = 0;
			slides[i].style.zIndex = "1";
		}
		slides[window.slide_index[sliderid] - 1].style.opacity = 1;
		slides[window.slide_index[sliderid] - 1].style.zIndex = "2";
		clearTimeout(window.slidertimeout[sliderid]);
		window.slidertimeout[sliderid] = setTimeout(function () {
			changeectslider(1, sliderid);
		}, window.slide_repeat[sliderid]);
	}
}
function ect_slider(sliderepeat, sliderid) {
	if (sliderid === undefined) sliderid = "";
	document.addEventListener("DOMContentLoaded", function () {
		doect_slider(sliderepeat, sliderid);
	});
}
function doect_slider(sliderepeat, sliderid) {
	var slides;
	if (sliderid != "") slides = document.getElementById(sliderid);
	else slides = document.getElementsByClassName("slidercontainer")[0];
	slides.innerHTML +=
		'<a class="sliderarrow sliderleft" onclick="changeectslider(-1,\'' +
		sliderid +
		'\')">&#xab;</a><a class="sliderarrow sliderright" onclick="changeectslider(1,\'' +
		sliderid +
		"')\">&#xbb;</a>";
	window.slide_index[sliderid] = 1;
	window.slide_repeat[sliderid] = sliderepeat;
	ect_displayslider(1, sliderid);
}
// Mega Menu
function ect_megamenu(hamburger_color) {
	document.addEventListener("DOMContentLoaded", function () {
		doect_megamenu(hamburger_color);
	});
}
function doect_megamenu(hamburger_color) {
	if (hamburger_color === undefined) hamburger_color = "#000000";
	function fadeIn(el, time) {
		el.style.opacity = 0;
		el.style.display = "block";
		var tick = function () {
			el.style.opacity = parseFloat(el.style.opacity) + 1 / (time / 10);
			if (parseFloat(el.style.opacity) < 1) setTimeout(tick, 10);
		};
		tick();
	}
	function fadeOut(el, time) {
		el.style.opacity = 1;
		var tick = function () {
			el.style.opacity = parseFloat(el.style.opacity) - 1 / (time / 10);
			if (parseFloat(el.style.opacity) > 0) setTimeout(tick, 10);
			else el.style.display = "none";
		};
		tick();
	}
	function fadeToggle(el, time) {
		if (el.style.opacity > 0) fadeOut(el, time);
		else fadeIn(el, time);
	}
	function mouseovermega(event, isover) {
		if (document.documentElement.clientWidth > 950) {
			if (isover) event.target.classList.add("ectmega-is-open");
			else event.target.classList.remove("ectmega-is-open");
			var elems = event.target.querySelectorAll("ul");
			for (var i = 0, len = elems.length; i < len; ++i)
				if (isover) fadeIn(elems[i], 150);
				else fadeOut(elems[i], 150);
			event.preventDefault();
		}
	}
	function clickmegasub(event) {
		if (document.documentElement.clientWidth <= 950) {
			var elems = event.target.querySelectorAll("ul");
			for (var i = 0, len = elems.length; i < len; ++i)
				fadeToggle(elems[i], 150);
		}
	}
	function clickmegamain(event) {
		var elems = document.querySelectorAll(".ectmegamenu > ul");
		for (var i = 0, len = elems.length; i < len; ++i)
			ecttoggleclass(elems[i], "show-on-mobile");
		event.preventDefault();
	}
	var elems = document.querySelectorAll(".ectmegamenu > ul > li");
	for (var i = 0, len = elems.length; i < len; ++i) {
		if (elems[i].querySelector("ul")) {
			elems[i].classList.add("ectmega-has-dropdown");
			elems[i].addEventListener("mouseenter", function (e) {
				mouseovermega(e, true);
			});
			elems[i].addEventListener("mouseleave", function (e) {
				mouseovermega(e, false);
			});
			elems[i].addEventListener("click", clickmegasub);
		}
	}
	var elems = document.querySelectorAll(".ectmegamenu > ul > li > ul");
	for (var i = 0, len = elems.length; i < len; ++i) {
		if (!elems[i].querySelector("ul")) elems[i].classList.add("normal-sub");
	}
	var element = document.querySelector(".ectmegamenu > ul");
	var newDiv = document.createElement("div");
	var newEl = document.createElement("a");
	newDiv.className = "ectmegamobile";
	newEl.href = "#";
	newDiv.addEventListener("click", clickmegamain);
	newEl.innerHTML =
		'<span class="megamobiletext"></span><svg class="megahamburger" width="32px" viewBox="0 0 181 135"><g><path style="fill:' +
		hamburger_color +
		";fill-opacity:1;stroke:" +
		hamburger_color +
		';stroke-width:18.465;stroke-linecap:round;stroke-linejoin:miter;stroke-miterlimit:4;stroke-dasharray:none;stroke-opacity:1" d="M 30,27 H 165" /><path style="fill:' +
		hamburger_color +
		";fill-opacity:1;stroke:" +
		hamburger_color +
		';stroke-width:18.465;stroke-linecap:round;stroke-linejoin:miter;stroke-miterlimit:4;stroke-dasharray:none;stroke-opacity:1" d="M 30,66 H 165" /><path style="fill:' +
		hamburger_color +
		";fill-opacity:1;stroke:" +
		hamburger_color +
		';stroke-width:18.465;stroke-linecap:round;stroke-linejoin:miter;stroke-miterlimit:4;stroke-dasharray:none;stroke-opacity:1" d="M 30,105 H 165" /></g></svg>';
	newDiv.appendChild(newEl);
	element.parentNode.insertBefore(newDiv, element);
}
function ectexpandreview(reviewnum) {
	document.getElementById("extracomments" + reviewnum).style.display = "";
	document.getElementById("extracommentsdots" + reviewnum).style.display =
		"none";
}
// START: Auto Search Function
var ectAutoSearchTmr;
var ectAutoAjaxO;
var ectAutoSrchCSI = -1;
var ectAutoSrchExt = "";
function ectAutoSrchOnClick(oSelect, objid) {
	document.getElementById(objid).value = oSelect.innerText;
	document.getElementById(objid + "form").submit();
}
function ectAutoDoHideCombo(objid) {
	document.getElementById("select" + objid).style.display = "none";
}
function ectAutoHideCombo(oText) {
	setTimeout("ectAutoDoHideCombo('" + oText.id + "')", 300);
}
function ectAutoClrSrchClasses(objid) {
	var oSelect = document.getElementById("select" + objid);
	for (var sind = 0; sind < oSelect.childNodes.length; sind++) {
		oSelect.childNodes[sind].className = "";
	}
}
function ectAutoAjaxCB(objid) {
	if (ectAutoAjaxO.readyState == 4) {
		var resarr = ectAutoAjaxO.responseText.split("&");
		var index, newy;
		var oSelect = document.getElementById("select" + objid);
		oSelect.style.display = ectAutoAjaxO.responseText == "" ? "none" : "";
		var act = resarr[0].replace(/\d/g, "");
		for (index = 0; index < resarr.length - 1; index++) {
			if (index < oSelect.childNodes.length)
				newy = oSelect.childNodes[index];
			else newy = document.createElement("div");
			newy.onclick = function () {
				ectAutoSrchOnClick(this, objid);
			};
			newy.onmouseover = function () {
				ectAutoClrSrchClasses(objid);
			};
			newy.innerHTML = decodeURIComponent(resarr[index]);
			if (index >= oSelect.childNodes.length) oSelect.appendChild(newy);
		}
		if (oSelect) {
			for (var ii = oSelect.childNodes.length - 1; ii >= index; ii--)
				oSelect.removeChild(oSelect.childNodes[ii]);
		}
	}
}
function ectAutoSrchPopList(objid) {
	var stext = document.getElementById(objid).value.toLowerCase();
	var catobj = document.getElementById("scat");
	var cattxt = "";
	if (catobj) cattxt = "&listcat=" + catobj[catobj.selectedIndex].value;
	ectAutoAjaxO = window.XMLHttpRequest
		? new XMLHttpRequest()
		: new ActiveXObject("MSXML2.XMLHTTP");
	ectAutoAjaxO.onreadystatechange = function () {
		ectAutoAjaxCB(objid);
	};
	ectAutoAjaxO.open(
		"POST",
		"vsadmin/ajaxservice." + ectAutoSrchExt + "?action=autosearch",
		true
	);
	ectAutoAjaxO.setRequestHeader(
		"Content-type",
		"application/x-www-form-urlencoded"
	);
	ectAutoAjaxO.send("listtext=" + encodeURIComponent(stext) + cattxt);
}
function ectAutoSrchKeydown(oText, e, textn) {
	ectAutoSrchExt = textn;
	var objid = oText.id;
	var oSelect = document.getElementById("select" + objid);
	var keyCode = e.keyCode;
	if (keyCode == 40 || keyCode == 38) {
		// Up / down arrows
		var numelements = 0;
		for (var sind = 0; sind < oSelect.childNodes.length; sind++) {
			if (oSelect.childNodes[sind].nodeType == 1) numelements++;
		}
		if (numelements > 0) {
			oSelect.style.display = "";
			if (keyCode == 40) {
				ectAutoSrchCSI++;
				if (ectAutoSrchCSI >= numelements)
					ectAutoSrchCSI = numelements - 1;
			} else {
				ectAutoSrchCSI--;
				if (ectAutoSrchCSI < 0) ectAutoSrchCSI = 0;
			}
			var rowc = 0;
			for (var sind = 0; sind < oSelect.childNodes.length; sind++) {
				oSelect.childNodes[sind].className =
					rowc == ectAutoSrchCSI ? "autosearchselected" : "";
				rowc++;
			}
			if (oSelect.childNodes[ectAutoSrchCSI]) {
				document.getElementById(objid).value =
					oSelect.childNodes[ectAutoSrchCSI].innerText;
			}
		}
	} else if (keyCode == 13) {
		oText.form.submit();
		return false;
	} else if ((keyCode >= 32 || keyCode == 8) && keyCode != 36) {
		clearTimeout(ectAutoSearchTmr);
		ectAutoSearchTmr = setTimeout(
			"ectAutoSrchPopList('" + objid + "')",
			600
		);
	}
	return true;
}
// END: Auto Search Function
