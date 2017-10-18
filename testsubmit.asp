<%@ Language=VBScript%>
<%
dim SQL, checkout

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment where Type="&request("equip")&" AND Available=True"

rs.open SQL, conn, 3, 3

available=CInt(rs("QuantityAvailable"))

response.write available

If CInt(request("quantity")) <= CInt(rs("QuantityAvailable")) then

	checkout=CInt(available)-request("quantity")
	rs("QuantityAvailable")=checkout
	rs.update
	message = "Amount updated in the database."

else

message = "The quantity of equipment requested for checkout exceeds the quantity available."

rs.close
conn.close
set conn=nothing

end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page56</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
//-->
</SCRIPT>

<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:193;top:143;width:538;z-index:1 }
#EndOfPage { position:absolute;left:0;top:169;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkinsmall"><%=(message)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/testsubmitInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
