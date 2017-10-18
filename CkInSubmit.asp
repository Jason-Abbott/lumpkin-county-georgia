<%@ Language=VBScript%>
<%
dim SQL

'mark returned = yes

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "lumpkin"
set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblCheckout where InventoryID="&request("equiptype")&" AND Type="&request("type")&" AND ProgramID="&request("program")&" AND PersonnelID="&request("personnel")

rs.open SQL, Conn, 3,3

If rs.eof then
oops="There is no current checkout record for this equipment type, and this individual, in the inventory database."
else
numcheckedout=rs("NumCheckedOut")

If CInt(request("quantity")) > numcheckedout then
	message="The quantity marked as returned exceeds the original amount checked out."
else

	If CInt(request("quantity")) < numcheckedout then
     
	available=CInt(rs("NumCheckedOut"))

		checkbackin=CInt(available)-request("quantity")
		rs("NumCheckedOut")=checkbackin
		rs.update
	end if

	If CInt(request("quantity"))=numcheckedout then

		rs("Returned")=True
		rs("ReturnedDate")=Date
		rs.update
	end if

rs.close
conn.close
set conn=nothing
end if


'update equipment quantity in the db

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment where Type="&request("type")&" AND InventoryID="&request("equiptype")
rs.open SQL, conn, 3, 3

If rs.eof then
	oops="No equipment of this type currently exists in the inventory database."
else

available=CInt(rs("QuantityAvailable"))

	checkout=CInt(available)+request("quantity")
	rs("QuantityAvailable")=checkout
	rs.update

rs.close
conn.close
set conn=nothing

end if

message2="The inventory database has been updated to reflect the current changes.  Please view the following table for an updated inventory of this equipment type."

'write out new table reflecting inventory changes

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "lumpkin"
set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment where InventoryID="&request("equiptype")&" AND Type="&request("type")
rs.open SQL, Conn, 1,2

If rs.EOF then
	message = "There is no equipment of this type located in the inventory database."
else

Response.Write "<DIV style='position:absolute;left:175;top:250;width:771;z-index:12'>"
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

<TITLE>Page60</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text3 = null;
var Text2 = null;
var Text1 = null;
//-->
</SCRIPT>

<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:4 }
#Text3 { position:absolute;left:193;top:157;width:513;z-index:3 }
#Text2 { position:absolute;left:193;top:177;width:507;z-index:2 }
#Text1 { position:absolute;left:193;top:204;width:508;z-index:1 }
#EndOfPage { position:absolute;left:0;top:230;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/inventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="Text3" Class="lumpkinsmall"><%=(message2)%>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall"><%=(message)%>
</DIV>
<DIV ID="Text1" Class="lumpkinsmall"><%=(oops)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/CkInSubmitInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
