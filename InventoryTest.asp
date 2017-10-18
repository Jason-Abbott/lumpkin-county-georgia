<%@ Language=VBScript%>
<%
'Create content for equipment ddlb

dim myEquipmentExpression


set rs=Server.CreateObject("ADODB.Recordset")
rs.cursorlocation=aduseclient 
rs.cachesize=5
SQL="SELECT DISTINCTROW [Type] FROM [tblEquipmentType]"


rs.open SQL, Conn


myEquipmentExpression="<form name=EquipForm><select name=EquipType onChange='filterByType();' >"

do while not rs.eof
	if request.querystring("stateName")="" or request.querystring("stateName")=0 then	
		myStatesExpression=myStatesExpression&"<option value="&rs(0)&">"&rs(1)
	else
			if CINT(request.querystring("stateName"))=CINT(rs(0)) then	
				myStatesExpression=myStatesExpression&"<option value="&rs(0)&" selected>"&rs(1)
			else
				myStatesExpression=myStatesExpression&"<option value="&rs(0)&">"&rs(1)
			end if
	end if
	rs.movenext
loop

myStatesExpression=myStatesExpression&"</select>"

rs.close

''''''''''''''''''''''''''''
'Create content for shipping points ddlb

dim myShipPtExpression


set rs=Server.CreateObject("ADODB.Recordset")
rs.cursorlocation=aduseclient 
rs.cachesize=5

if request.querystring("stateName")="" then
	SQL="SELECT DISTINCTROW [CommoditySource].[CommoditySourceID], [CommoditySource].[LocationName] FROM [CommoditySource]"
else
	SQL="SELECT DISTINCTROW [CommoditySource].[CommoditySourceID], [CommoditySource].[LocationName] FROM [CommoditySource] WHERE [CommoditySource].[SIDl]="&request.querystring("stateName")             
end if

rs.open SQL, Conn           '------------------------line 500

myShipPtExpression="<form name=agForm><select name=ShippingPt>"

do while not rs.eof
	myShipPtExpression=myShipPtExpression&"<option value="&rs(0)&">"&rs(1)
	rs.movenext
loop


myShipPtExpression=myShipPtExpression&"</select>"

rs.close
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page57</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var Text2 = null;
//-->
</SCRIPT>
<SCRIPT>
<!--
function filterByType(){
	location.href=("http://www.apicbt.com/LumpkinCounty/InventoryTest.asp?comFilter="+document.EquipForm.EquipType.value)
}
//-->
</SCRIPT>


<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:320;top:123;width:248;z-index:1 }
#Text2 { position:absolute;left:380;top:227;width:250;z-index:2 }
#EndOfPage { position:absolute;left:0;top:253;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkinsmall"><%=(myEquipmentExpression)%>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall"><%=(myEquipmentTypeExpression)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/InventoryTestInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
