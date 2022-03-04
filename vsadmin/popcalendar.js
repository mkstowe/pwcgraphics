//	written	by Tan Ling	Wee	on 2 Dec 2001
//	last updated 20 June 2003
//	email :	fuushikaden@yahoo.com

function languagetext(thelang){
	if(thelang=='fr'){
		todayString="Aujourd'hui:";
		monthName=new Array("Janvier","F&eacute;vrier","Mars","Avril","Mai","Juin","Juillet","Ao&ucirc;t","Septembre","Octobre","Novembre","D&eacute;cembre");
		monthName2=new Array("Janv.","F&eacute;v.","Mars","Avril","Mai","Juin","Juil.","Ao&ucirc;t","Sept.","Oct.","Nov.","D&eacute;c.");
		dayName=new Array("Lun.","Mar.","Mer.","Jeu.","Ven.","Sam.","Dim.");
	}else if(thelang=='de'){
		todayString="Heute:";
		monthName=new Array("Januar","Februar","M&auml;rz","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember");
		monthName2=new Array("Jan","Feb","Mar","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez");
		dayName=new Array("Mo","Di","Mi","Do","Fr","Sa","So");
	}else if(thelang=='es'){
		todayString="Hoy:";
		monthName=new Array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre");
		monthName2=new Array("Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic");
		dayName=new Array("Lun","Mar","Mi&eacute;","Jue","Vie","S&aacute;b","Dom");
	}else if(thelang=='nl'){
		todayString="Vandaag:";
		monthName=new Array("Januari","Februari","Maart","April","Mei","Juni","Juli","Augustus","September","Oktober ","November","December");
		monthName2=new Array("JAN","FEB","MAA","APR","MEI","JUN","JUL","AUG","SEP","OKT","NOV","DEC");
		dayName=new Array("Maa","Din","Woe","Don","Vri","Zat","Dom");
	}
}
var ectcalstartat=(typeof ectpopcalstartmonday!='undefined'?1:0); // 0 - sunday ; 1 - monday
var showToday=1;	// 0 - don't show; 1 - show
var ectcalimgdir=(typeof ectpopcalisproducts!='undefined'?'images/':'adminimages/');

var gotoString="Go To Current Month";
var todayString="Today is";

var	crossobj, crossMonthObj, crossYearObj, monthSelected, yearSelected, dateSelected, omonthSelected, oyearSelected, odateSelected, monthConstructed, yearConstructed, intervalID1, intervalID2, timeoutID1, timeoutID2, ctlToPlaceValue, ctlNow, dateFormat, nStartingYear;

var	bPageLoaded=false;
var	ectcalie=false;
var bua=navigator.userAgent.toLowerCase();
if(bua.indexOf("msie 8.")>0 || bua.indexOf("msie 7.")>0) ectcalie=true;

var	today=new Date();
var	dateNow=today.getDate();
var	monthNow=today.getMonth();
var	yearNow=today.getYear();
var	imgsrc=new Array("ectcaldrop1.png","ectcaldrop2.png","ectcalleft1.png","ectcalleft2.png","ectcalright1.png","ectcalright2.png");
var	img=new Array();

var bShow=false;

function hideElement(elmID, overDiv){
	if(ectcalie){
		for(i=0; i < document.all.tags(elmID).length; i++){
			obj=document.all.tags(elmID)[i];
			if(!obj || !obj.offsetParent) continue;

			// Find the element's offsetTop and offsetLeft relative to the BODY tag.
			objLeft=obj.offsetLeft;
			objTop=obj.offsetTop;
			objParent=obj.offsetParent;

			while(objParent.tagName.toUpperCase() != "BODY" && objParent.tagName.toUpperCase() != "HTML"){
				objLeft+=objParent.offsetLeft;
				objTop+=objParent.offsetTop;
				objParent=objParent.offsetParent;
			}
			objHeight=obj.offsetHeight;
			objWidth=obj.offsetWidth;
			if((overDiv.offsetLeft + overDiv.offsetWidth )<=objLeft);
			else if((overDiv.offsetTop + overDiv.offsetHeight) <= objTop);
			else if(overDiv.offsetTop>=(objTop + objHeight));
			else if(overDiv.offsetLeft>=(objLeft + objWidth));
			else
				obj.style.visibility="hidden";
		}
	}
}
function showElement(elmID){
	if(ectcalie){
		for( i=0; i < document.all.tags( elmID ).length; i++){
			obj=document.all.tags( elmID )[i];
			if( !obj || !obj.offsetParent) continue;
			obj.style.visibility="";
		}
	}
}
function HolidayRec(d,m,y,desc){
	this.d=d;
	this.m=m;
	this.y=y;
	this.desc=desc;
}

var HolidaysCounter=0;
var Holidays=new Array();

function addHoliday(d,m,y,desc){
	Holidays[HolidaysCounter++]=new HolidayRec(d,m,y,desc);
}

for(i=0;i<imgsrc.length;i++){
	img[i]=new Image;
	img[i].src=ectcalimgdir + imgsrc[i];
}
document.write(
"<div onclick='bShow=true' id='calendar' style='position:absolute;top:0px;left:0px;visibility:hidden;z-index:10001'>" +
	"<div class='ectcalendar'>" +
		"<table class='ectcalheader' id='ectcalheader'><tr><td id='caption'></td><td align=right><a href='javascript:hideCalendar()'><img src='"+ectcalimgdir+"ectcallogout.png' alt='Close Calendar' style='vertical-align:middle' /></a></td></tr></table>" +
		"<div id='ectcalcontent'></div>" +
(showToday==1 ? '<div class="ectcaltodaydate" style="cursor:pointer" id="lblToday" title="'+gotoString+'" onclick="monthSelected=monthNow;yearSelected=yearNow;constructCalendar()"></div>' : '') +
	"</div>" +
	"<div id='selectMonth' style='position:absolute;visibility:hidden;z-index:10002'></div><div id='selectYear' style='z-index:10002;position:absolute;visibility:hidden;z-index:10002'></div>" +
"</div>");

var	monthName=new Array("January","February","March","April","May","June","July","August","September","October","November","December");
var	monthName2=new Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec");
if(ectcalstartat==0)
	dayName=new Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat");
else
	dayName=new Array("Mon","Tue","Wed","Thu","Fri","Sat","Sun");
if(typeof ectpopcallang!='undefined') languagetext(ectpopcallang);

function swapImage(srcImg,destImg){
	document.getElementById(srcImg).src=ectcalimgdir+destImg;
}

function ectcalinit(){
	if(!ectcalie) yearNow+=1900;

	crossobj=document.getElementById("calendar").style;
	hideCalendar();
	crossMonthObj=document.getElementById("selectMonth").style;
	crossYearObj=document.getElementById("selectYear").style;

	monthConstructed=false;
	yearConstructed=false;

	if(showToday==1)
		document.getElementById("lblToday").innerHTML=todayString + " " + dayName[(today.getDay()-ectcalstartat==-1)?6:(today.getDay()-ectcalstartat)]+", " + dateNow + " " + monthName[monthNow].substring(0,3)	+ "	" +	yearNow;

	sHTML1="<span id='spanLeft'	style='cursor:pointer;padding:1px' onmouseover='swapImage(\"changeLeft\",\"ectcalleft2.png\")' onclick='javascript:decMonth()' onmouseout='clearInterval(intervalID1);swapImage(\"changeLeft\",\"ectcalleft1.png\")' onmousedown='clearTimeout(timeoutID1);timeoutID1=setTimeout(\"StartDecMonth()\",500)' onmouseup='clearTimeout(timeoutID1);clearInterval(intervalID1)'>&nbsp<img id='changeLeft' src='"+ectcalimgdir+"ectcalleft1.png' style='vertical-align:middle' />&nbsp</span>&nbsp;";
	sHTML1+="<span id='spanRight' style='cursor:pointer;padding:1px' onmouseover='swapImage(\"changeRight\",\"ectcalright2.png\")' onmouseout='clearInterval(intervalID1);swapImage(\"changeRight\",\"ectcalright1.png\")' onclick='incMonth()' onmousedown='clearTimeout(timeoutID1);timeoutID1=setTimeout(\"StartIncMonth()\",500)' onmouseup='clearTimeout(timeoutID1);clearInterval(intervalID1)'>&nbsp<img id='changeRight' src='"+ectcalimgdir+"ectcalright1.png' style='vertical-align:middle' />&nbsp</span>&nbsp";
	sHTML1+="<span id='spanMonth' style='cursor:pointer;padding:1px' onmouseover='swapImage(\"changeMonth\",\"ectcaldrop2.png\")' onmouseout='swapImage(\"changeMonth\",\"ectcaldrop1.png\")' onclick='popUpMonth()'></span>&nbsp;";
	sHTML1+="<span id='spanYear' style='cursor:pointer;padding:1px' onmouseover='swapImage(\"changeYear\",\"ectcaldrop2.png\")' onmouseout='swapImage(\"changeYear\",\"ectcaldrop1.png\")' onclick='popUpYear()'></span>&nbsp;";
	
	document.getElementById("caption").innerHTML=sHTML1

	bPageLoaded=true;
}

function hideCalendar(){
	crossobj.visibility="hidden";
	if(crossMonthObj!=null) crossMonthObj.visibility="hidden";
	if(crossYearObj!=null) crossYearObj.visibility="hidden";

	showElement('SELECT');
	showElement('APPLET');
}

function padZero(num){
	return (num	< 10)? '0' + num : num;
}

function constructDate(d,m,y){
	sTmp=dateFormat;
	sTmp=sTmp.replace("dd","<e>");
	sTmp=sTmp.replace("d","<d>");
	sTmp=sTmp.replace("<e>",padZero(d));
	sTmp=sTmp.replace("<d>",d);
	sTmp=sTmp.replace("mmmm","<p>");
	sTmp=sTmp.replace("mmm","<o>");
	sTmp=sTmp.replace("mm","<n>");
	sTmp=sTmp.replace("m","<m>");
	sTmp=sTmp.replace("<m>",m+1);
	sTmp=sTmp.replace("<n>",padZero(m+1));
	sTmp=sTmp.replace("<o>",monthName[m]);
	sTmp=sTmp.replace("<p>",monthName2[m]);
	sTmp=sTmp.replace("yyyy",y);
	return sTmp.replace("yy",padZero(y%100));
}

function closeCalendar(){
	var	sTmp;

	hideCalendar();
	ctlToPlaceValue.value=constructDate(dateSelected,monthSelected,yearSelected);
}

/*** Month Pulldown	***/
function StartDecMonth(){
	intervalID1=setInterval("decMonth()",80);
}

function StartIncMonth(){
	intervalID1=setInterval("incMonth()",80);
}

function incMonth(){
	monthSelected++;
	if(monthSelected>11){
		monthSelected=0;
		yearSelected++;
	}
	constructCalendar();
}

function decMonth(){
	monthSelected--;
	if(monthSelected<0){
		monthSelected=11;
		yearSelected--;
	}
	constructCalendar();
}

function constructMonth(){
	popDownYear();
	if(!monthConstructed){
		sHTML="";
		for(i=0; i<12; i++){
			sName=monthName[i];
			if(i==monthSelected){
				sName="<strong>" + sName + "</strong>";
			}
			sHTML+="<tr class='ectcalselector'><td class='ectcalselector' id='m" + i + "' onmouseover='this.style.backgroundColor=\"#FFCC99\"' onmouseout='this.style.backgroundColor=\"\"' style='cursor:pointer' onclick='monthConstructed=false;monthSelected=" + i + ";constructCalendar();popDownMonth();event.cancelBubble=true'>&nbsp;" + sName + "&nbsp;</td></tr>";
		}
		document.getElementById("selectMonth").innerHTML="<table class='ectcalselector' width='70' onmouseover='clearTimeout(timeoutID1)' onmouseout='clearTimeout(timeoutID1);timeoutID1=setTimeout(\"popDownMonth()\",100);event.cancelBubble=true'>" + sHTML + "</table>";
		monthConstructed=true;
	}
}

function popUpMonth(){
	constructMonth();
	crossMonthObj.visibility="visible";
	crossMonthObj.left=(parseInt(crossobj.left) + 50)+'px';
	crossMonthObj.top=(document.getElementById('ectcalheader').clientHeight-2)+'px';
	hideElement('SELECT', document.getElementById("selectMonth") );
	hideElement('APPLET', document.getElementById("selectMonth") );			
}

function popDownMonth(){
	crossMonthObj.visibility= "hidden";
}

/*** Year Pulldown ***/

function incYear(){
	for	(i=0; i<7; i++){
		newYear=(i+nStartingYear)+1;
		if(newYear==yearSelected)
		{ txtYear="&nbsp;<strong>"	+ newYear +	"</strong>&nbsp;"; }
		else
		{ txtYear="&nbsp;" + newYear + "&nbsp;"; }
		document.getElementById("y"+i).innerHTML=txtYear;
	}
	nStartingYear++;
	bShow=true;
}

function decYear(){
	for	(i=0; i<7; i++){
		newYear=(i+nStartingYear)-1
		if(newYear==yearSelected)
		{ txtYear="&nbsp;<strong>"	+ newYear +	"</strong>&nbsp;"; }
		else
		{ txtYear="&nbsp;" + newYear + "&nbsp;"; }
		document.getElementById("y"+i).innerHTML=txtYear;
	}
	nStartingYear --;
	bShow=true;
}

function selectYear(nYear){
	yearSelected=parseInt(nYear+nStartingYear);
	yearConstructed=false;
	constructCalendar();
	popDownYear();
}

function constructYear(){
	popDownMonth();
	sHTML="";
	if(!yearConstructed){
		sHTML="<tr class='ectcalselector'><td class='ectcalselector' onmouseover='this.style.backgroundColor=\"#FFCC99\"' onmouseout='clearInterval(intervalID1);this.style.backgroundColor=\"\"' style='cursor:pointer;text-align:center' onmousedown='clearInterval(intervalID1);intervalID1=setInterval(\"decYear()\",30)' onmouseup='clearInterval(intervalID1)'>-</td></tr>";

		j=0;
		nStartingYear=yearSelected-3;
		for(i=(yearSelected-3); i<=(yearSelected+3); i++){
			sName=i;
			if(i==yearSelected){
				sName="<strong>" + sName + "</strong>";
			}
			sHTML+="<tr class='ectcalselector'><td class='ectcalselector' id='y" + j + "' onmouseover='this.style.backgroundColor=\"#FFCC99\"' onmouseout='this.style.backgroundColor=\"\"' style='cursor:pointer' onclick='selectYear("+j+");event.cancelBubble=true'>&nbsp;" + sName + "&nbsp;</td></tr>";
			j++;
		}

		sHTML+="<tr class='ectcalselector'><td class='ectcalselector' onmouseover='this.style.backgroundColor=\"#FFCC99\"' onmouseout='clearInterval(intervalID2);this.style.backgroundColor=\"\"' style='cursor:pointer;text-align:center' onmousedown='clearInterval(intervalID2);intervalID2=setInterval(\"incYear()\",30)' onmouseup='clearInterval(intervalID2)'>+</td></tr>";

		document.getElementById("selectYear").innerHTML="<table class='ectcalselector' onmouseover='clearTimeout(timeoutID2)' onmouseout='clearTimeout(timeoutID2);timeoutID2=setTimeout(\"popDownYear()\",100)' cellspacing=0>" + sHTML + "</table>";

		yearConstructed=true;
	}
}

function popDownYear(){
	clearInterval(intervalID1);
	clearTimeout(timeoutID1);
	clearInterval(intervalID2);
	clearTimeout(timeoutID2);
	crossYearObj.visibility= "hidden";
}

function popUpYear(){
	var	leftOffset;
	constructYear();
	crossYearObj.visibility="visible";
	leftOffset=parseInt(crossobj.left) + document.getElementById("spanYear").offsetLeft;
	if(ectcalie) leftOffset+=6;
	crossYearObj.left=leftOffset+'px';
	crossYearObj.top=(document.getElementById('ectcalheader').clientHeight-2)+'px';
}

/*** calendar ***/
function WeekNbr(n) {
	// Algorithm used:
	// From Klaus Tondering's Calendar document (The Authority/Guru)
	// http://www.tondering.dk/claus/calendar.html
	// a=(14-month) / 12
	// y=year + 4800 - a
	// m=month + 12a - 3
	// J=day + (153m + 2) / 5 + 365y + y / 4 - y / 100 + y / 400 - 32045
	// d4=(J + 31741 - (J mod 7)) mod 146097 mod 36524 mod 1461
	// L=d4 / 1460
	// d1=((d4 - L) mod 365) + L
	// WeekNumber=d1 / 7 + 1

	year=n.getFullYear();
	month=n.getMonth() + 1;
	if(ectcalstartat==0)
		day=n.getDate() + 1;
	else
		day=n.getDate();

	a=Math.floor((14-month) / 12);
	y=year + 4800 - a;
	m=month + 12 * a - 3;
	b=Math.floor(y/4) - Math.floor(y/100) + Math.floor(y/400);
	J=day + Math.floor((153 * m + 2) / 5) + 365 * y + b - 32045;
	d4=(((J + 31741 - (J % 7)) % 146097) % 36524) % 1461;
	L=Math.floor(d4 / 1460);
	d1=((d4 - L) % 365) + L;
	week=Math.floor(d1/7) + 1;

	return week;
}

function constructCalendar(){
	var aNumDays=Array (31,0,31,30,31,30,31,31,30,31,30,31);
	var	startDate=new Date (yearSelected,monthSelected,1);
	var endDate;

	if(monthSelected==1){
		endDate=new Date (yearSelected,monthSelected+1,1);
		endDate=new Date (endDate-(24*60*60*1000));
		numDaysInMonth=endDate.getDate();
	}else
		numDaysInMonth=aNumDays[monthSelected];

	datePointer=0;
	dayPointer=startDate.getDay() - ectcalstartat;
	
	if(dayPointer<0)
		dayPointer=6;

	sHTML='<table class="ectcaldates"><tr>';

	for(i=0; i<7; i++)
		sHTML+="<td width='27' align='right'><strong>"+ dayName[i]+"</strong></td>";
	sHTML+="</tr><tr>";

	for(var i=1; i<=dayPointer;i++)
		sHTML+="<td>&nbsp;</td>";

	for(datePointer=1; datePointer<=numDaysInMonth; datePointer++){
		dayPointer++;
		sHTML+="<td class='ectcaldate'>";
		sStyle="";
		sHint="";
		for(k=0;k<HolidaysCounter;k++){
			if((parseInt(Holidays[k].d)==datePointer)&&(parseInt(Holidays[k].m)==(monthSelected+1))){
				if((parseInt(Holidays[k].y)==0)||((parseInt(Holidays[k].y)==yearSelected)&&(parseInt(Holidays[k].y)!=0))){
					sStyle+=" ectcaldatedisabled";
					sHint+=sHint==""?Holidays[k].desc:"\n"+Holidays[k].desc;
				}
			}
		}

		var regexp=/\"/g;
		sHint=sHint.replace(regexp,"&quot;");
		if(!(yearNow<yearSelected||(yearNow==yearSelected&&monthNow<monthSelected)||(yearNow==yearSelected&&monthNow==monthSelected&&dateNow<=datePointer)))
				sStyle+=' ectcalpastdate';
		if((datePointer==dateNow)&&(monthSelected==monthNow)&&(yearSelected==yearNow))
				sStyle+=' ectcaltoday';
		sStyle+=' ectcaldayno'+((dayPointer+ectcalstartat) % 7);
		sHTML+="<div class='ectcaldate"+sStyle+"' style='cursor:pointer' title=\"" + sHint + "\" onclick='dateSelected="+datePointer + ";closeCalendar();'>&nbsp;" + datePointer + "&nbsp;</div>";
		if((dayPointer+ectcalstartat) % 7 == ectcalstartat)
			sHTML+="</tr><tr>";
	}
	document.getElementById("ectcalcontent").innerHTML=sHTML;
	document.getElementById("spanMonth").innerHTML="&nbsp;" + monthName[monthSelected] + '&nbsp;<img id="changeMonth" src="'+ectcalimgdir+'ectcaldrop1.png" style="vertical-align:middle" />';
	document.getElementById("spanYear").innerHTML="&nbsp;" + yearSelected + '&nbsp;<img id="changeYear" src="'+ectcalimgdir+'ectcaldrop1.png" style="vertical-align:middle" />';
}

function popUpCalendar(ctl,	ctl2, format, shuffle){
	var	leftpos=0,toppos=0;

	if(bPageLoaded){
		if(crossobj.visibility=="hidden"){
			ctlToPlaceValue=ctl2;
			dateFormat=format;

			formatChar=" ";
			aFormat=dateFormat.split(formatChar);
			if(aFormat.length<3){
				formatChar="/";
				aFormat=dateFormat.split(formatChar);
				if(aFormat.length<3){
					formatChar=".";
					aFormat=dateFormat.split(formatChar);
					if(aFormat.length<3){
						formatChar="-";
						aFormat=dateFormat.split(formatChar);
						if(aFormat.length<3) formatChar=""; // invalid date format
					}
				}
			}
			tokensChanged=0;
			if(formatChar!=""){ // use user's date
				aData=ctl2.value.split(formatChar);

				for(i=0;i<3;i++){
					if((aFormat[i]=="d") || (aFormat[i]=="dd")){
						dateSelected=parseInt(aData[i], 10);
						tokensChanged++;
					}
					else if	((aFormat[i]=="m") || (aFormat[i]=="mm")){
						monthSelected=parseInt(aData[i], 10) - 1;
						tokensChanged++;
					}
					else if	(aFormat[i]=="yyyy"){
						yearSelected=parseInt(aData[i], 10);
						tokensChanged++;
					}
					else if	(aFormat[i]=="mmm"){
						for	(j=0; j<12;	j++){
							if(aData[i]==monthName[j]){
								monthSelected=j;
								tokensChanged++;
							}
						}
					}else if(aFormat[i]=="mmmm"){
						for	(j=0; j<12;	j++){
							if(aData[i]==monthName2[j]){
								monthSelected=j;
								tokensChanged++;
							}
						}
					}
				}
			}

			if((tokensChanged!=3)||isNaN(dateSelected)||isNaN(monthSelected)||isNaN(yearSelected)){
				dateSelected=dateNow;
				monthSelected=monthNow;
				yearSelected=yearNow;
			}

			odateSelected=dateSelected;
			omonthSelected=monthSelected;
			oyearSelected=yearSelected;
			
			var parentdiv=ctl2.parentNode;
			parentdiv.insertBefore(document.getElementById("calendar"),parentdiv.firstChild);
			crossobj.top=(ctl2.offsetHeight+2)+'px';

			constructCalendar(1, monthSelected, yearSelected);
			crossobj.visibility="visible";
			
			hideElement('SELECT', document.getElementById("calendar") );
			hideElement('APPLET', document.getElementById("calendar") );			

			bShow=true;
		}else{
			hideCalendar();
			if(ctlNow!=ctl) {popUpCalendar(ctl, ctl2, format);}
		}
		ctlNow=ctl;
	}
}
document.onclick=function hidecal2 () { 		
	if(!bShow) hideCalendar();
	bShow=false;
}
ectcalinit();

if(typeof ectpopcaldisabled!='undefined'){
	disabledarray=ectpopcaldisabled.split(';');
	for(var disindex=0;disindex<disabledarray.length;disindex++){
		disdate=disabledarray[disindex].split('-');
		if(disdate.length==2) addHoliday(disdate[1],disdate[0],0,'');
	}
}