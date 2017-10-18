<%@ Language=VBScript%>
<%
dim SQL, programtype

programtype=request("searchtype")
equiptype=request("equiptype")

x=0

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "lumpkin"
set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment where InventoryID="&equiptype&" AND Type="&programtype

rs.open SQL, Conn, 1,2

If rs.EOF then
	message = "There is no equipment of this type located in the inventory database."
else

    message = "The following table represents the quantity of this type inventory in the database."
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

'write out second table

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "lumpkin"
set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblCheckout where InventoryID="&equiptype&" AND Type="&programtype&" AND Returned="&False

rs.open SQL, Conn, 1,2

If rs.EOF then
	message = "Currently, no equipment of this type has been checked out of inventory."
else
    table2="The following table contains equipment of this type which has been checked out of inventory."
Response.Write "<DIV style='position:absolute;left:175;top:410;width:771;z-index:12'>"
Response.Write"<table border='1' width='100%'>"

  Response.Write"<tr>"
    Response.Write"<td width='15%'><strong>Equipment Type</strong></td>"
    Response.Write"<td width='20%'><strong>Description</strong></td>"
	Response.Write"<td width='10%'><strong>Checked Out</strong></td>"
	Response.Write"<td width='10%'><strong>Quantity</strong></td>"
	Response.Write"<td width='15%'><strong>Released To</strong></td>"
	Response.Write"<td width='20%'><strong>Program</strong></td>"
	Response.Write"<td width='10%'><strong>Due</strong></td>"
  Response.Write"</tr>"

Do While Not rs.EOF

	Response.Write "<tr>"
	Response.Write "<td>"
		Set Conn1 = Server.CreateObject("ADODB.Connection")
		Conn1.Open "lumpkin"
		set rs1 = Server.CreateObject("ADODB.Recordset")
		SQL = "Select * from tblEquipmentType where InventoryTypeID="&programtype
		rs1.open SQL, conn1, 1, 2
		Response.Write rs1("Type")
		rs1.close
		conn1.close
		set conn1=nothing
	Response.Write "</td><td>"
		Set Conn2 = Server.CreateObject("ADODB.Connection")
		Conn2.Open "lumpkin"
		set rs2 = Server.CreateObject("ADODB.Recordset")
		SQL = "Select * from tblEquipment where Type="&programtype
		rs2.open SQL, conn2, 1, 2
		Response.Write rs2("Description")
		Response.Write "</td><td>"
		If rs("CheckedOut")<>"" then
			Response.Write rs("CheckedOut")
		else 
			Response.Write "Available"
		end if
		Response.Write "</td><td>"
		Response.Write rs("NumCheckedOut")
		Response.Write "</td><td>"
		
				Set Conn5 = Server.CreateObject("ADODB.Connection")
				Conn5.Open "lumpkin"
				set rs5 = Server.CreateObject("ADODB.Recordset")
				SQL = "Select * from tblPersonnel where PersonnelID="&rs("PersonnelID")
				rs5.open SQL, conn5, 1, 2
				Response.Write rs5("FirstName")&" "&rs5("LastName")
				Response.Write "</td><td>"
				rs5.close
				conn5.close
				set conn5=nothing
			
				programID=rs("ProgramID")
				Set Conn3 = Server.CreateObject("ADODB.Connection")
				Conn3.Open "lumpkin"
				set rs3 = Server.CreateObject("ADODB.Recordset")
				SQL = "Select * from tblPrograms where ProgramID="&programID
				rs3.open SQL, conn3, 1, 2
				Response.Write rs3("ProgramName")
				Response.Write "</td><td>"
				rs3.close
				conn3.close
				set conn3=nothing
	'	Response.Write "</td><td>"
		Response.Write rs("DueDate")
		Response.Write "</td>"
		rs.MoveNext

		rs2.close
		conn2.close
		set conn2=nothing
'	rs.MoveNext
	
Loop 

	Response.Write "</tr></table>"
	Response.Write "</DIV>"
rs.close
conn.close
set conn = nothing


end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page39</TITLE>
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
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:2 }
#Text1 { position:absolute;left:167;top:168;width:554;z-index:1 }
#Text2 { position:absolute;left:167;top:317;width:564;z-index:3 }
#EndOfPage { position:absolute;left:0;top:343;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/inventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="Text1" Class="lumpkinsmall"><%=(message)%>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall"><%=(table2)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/SearchByTypeInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
