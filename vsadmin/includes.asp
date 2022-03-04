<%
Dim maintablebg,innertablebg,headeralign,maintablewidth,innertablewidth,maintablespacing,innertablespacing,maintablepadding,innertablepadding,sqlserver,nobuyorcheckout,noprice
Dim mailhost,sortBy,pathtossl,taxShipping,pagebarattop,productcolumns,useproductbodyformat,usesearchbodyformat,usedetailbodyformat,useemailfriend,expireaffiliate,XXLicenseNumber

' For a description of these parameters and their useage, please open the following URL in your browser
' http://www.ecommercetemplates.com/help/ecommplus/parameters.asp

mailhost = "mail.pwcgraphics.com"
sortBy = 1

taxShipping=0
pagebarattop=1
productcolumns=1
useproductbodyformat=1
usesearchbodyformat=1
usedetailbodyformat=1
useemailfriend=true
nobuyorcheckout=false
noprice=false
expireaffiliate=0
sqlserver=false
disallowlogin=FALSE
notifyloginattempt=FALSE
usefirstlastname=TRUE
nochecksslserver=TRUE
IntlHandling=22



' ===================================================================
' Please do not edit anything below this line
' ===================================================================

maintablebg=""
innertablebg=""
maintablewidth="96%"
innertablewidth="100%"
maintablespacing="0"
innertablespacing="0"
maintablepadding="1"
innertablepadding="6"
headeralign="left"

Session.LCID = 1033

const maxprodopts=25
const helpbaseurl="http://www.ecommercetemplates.com/help/ecommplus/"

Function Max(a,b)
	if a > b then
		Max=a
	else
		Max=b
	end if
End function
Function Min(a,b)
	if a < b then
		Min=a
	else
		Min=b
	end if
End function
%>