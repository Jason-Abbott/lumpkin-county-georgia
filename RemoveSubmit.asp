<%@ Language=VBScript%>
<%
dim SQL

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "lumpkin"
set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment where InventoryID="&request("descrip")&" AND Type="&request("type")
rs.open SQL, Conn, 3,3

If rs.EOF then
	message = "There is no equipment of this type to delete from the inventory database."
else

rs("TotalQuantity")=rs("TotalQuantity")-CInt(request("quantity"))
rs("QuantityAvailable")=rs("QuantityAvailable")-CInt(request("quantity"))
rs.update
	message = request("quantity") &" items have been deleted.  The following table provides an updated view of the inventory for this item."


'write out new table reflecting inventory changes


Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "lumpkin"
set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment where InventoryID="&request("descrip")&" AND Type="&request("type")
rs.open SQL, Conn, 1,2

If rs.EOF then
	message2 = "There is no further equipment of this type located in the inventory database."
else

Response.Write "<DIV style='position:absolute;left:175;top:225;width:771;z-index:12'>"
Response.Write"<table border='1' width='100%'>"

  Response.Write"<tr>"
    Response.Write"<td width='20%'><strong>Equipment Type</strong></td>"
    Response.Write"<td width='20%'><strong>Description</strong></td>"
	Response.Write"<td width='20%'><strong>Purchase Date</strong></td>"
	Response.Write"<td width='20%'><strong>Total Quantity</strong></td>"
	Response.Write"<td width='20%'><strong>Available for Checkout</strong></td>"
  Response.Write"</tr>"

Do While Not rs.EOF

	Response.Write "<tr>"
	Response.Write "<td>"
		Set Conn1 = Server.CreateObject("ADODB.Connection")
		Conn1.Open "lumpkin"
		set rs1 = Server.CreateObject("ADODB.Recordset")
		SQL = "Select * from tblEquipmentType where InventoryTypeID="&rs("Type")
		rs1.open SQL, conn1, 1, 2
		Response.Write rs1("Type")
		rs1.close
		conn1.close
		set conn1=nothing
	Response.Write "</td><td>"
	Response.Write rs("Description")
	Response.Write "</td><td>"
	Response.Write rs("PurchaseDate")
	Response.Write "</td><td>"
	Response.Write rs("TotalQuantity")
	Response.Write "</td><td>"
	Response.Write rs("QuantityAvailable")
		Response.Write "</td>"
		rs.MoveNext

Loop 

	Response.Write "</tr></table>"
	Response.Write "</DIV>"

rs.close
conn.close
set conn = nothing

end if

end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page61</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text2 = null;
var Text1 = null;
//-->
</SCRIPT>

<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:3 }
#Text2 { position:absolute;left:191;top:160;width:682;z-index:2 }
#Text1 { position:absolute;left:192;top:180;width:682;z-index:1 }
#EndOfPage { position:absolute;left:0;top:206;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/inventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="Text2" Class="lumpkinsmall"><%=(message)%>
</DIV>
<DIV ID="Text1" Class="lumpkinsmall"><%=(message2)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/RemoveSubmitInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
