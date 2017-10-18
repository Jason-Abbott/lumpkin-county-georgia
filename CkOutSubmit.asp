<%@ Language=VBScript%>
<%
duedate=Request("mm")&"/"&Request("dd")&"/"&Request("yy")

dim SQL, checkout

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment where Type="&request("equiptype")&" AND InventoryID="&request("equipdescrip")
rs.open SQL, conn, 3, 3

If rs.eof then
	oops="No equipment of this type currently exists in the inventory database."
else

available=CInt(rs("QuantityAvailable"))

If CInt(request("quantity")) <= CInt(rs("QuantityAvailable")) then

	checkout=CInt(available)-request("quantity")
	rs("QuantityAvailable")=checkout
	rs.update

rs.close
conn.close
set conn=nothing

	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open "lumpkin"
	set rs = Server.CreateObject("ADODB.Recordset")
	
	SQL = "Select * from tblCheckout where 1=2"
	
	rs.open SQL, conn, 3, 3

	rs.addnew

	rs("InventoryID")=request("equipdescrip")
	rs("Type")=request("equiptype")
	rs("CheckedOut")=Date
	rs("NumCheckedOut")=request("quantity")
	rs("DueDate")=duedate
	rs("ProgramID")=request("program")
	rs("FacilityID")=vbnull
	rs("PersonnelID")=request("personnel")

	rs.update

	rs.close
	conn.close
	set conn=nothing

	message = "Amount updated in the database."

else

message = "The quantity of equipment requested for checkout exceeds the quantity available."


end if
end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page56</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var Text2 = null;
//-->
</SCRIPT>

<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:193;top:143;width:538;z-index:1 }
#Text2 { position:absolute;left:195;top:177;width:530;z-index:2 }
#EndOfPage { position:absolute;left:0;top:203;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkinsmall"><%=(message)%>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall"><%=(oops)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/CkOutSubmitInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
