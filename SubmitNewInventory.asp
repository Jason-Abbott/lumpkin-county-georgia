<%@ Language=VBScript%>
<%
'make changes to rs'Checked Out' field on this page!!

purchasedate=Request("mm")&"/"&Request("dd")&"/"&Request("yy")
description=request("newequipdescrip")

dim SQL
x=0

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment WHERE Type="&request("addnewdrop")&" AND Description='"&request("newequipdescrip")&"'"

rs.open SQL, conn, 3, 3

If rs.eof then

rs.addnew

rs("Type")=CInt(request("addnewdrop"))
rs("Description")=request("newequipdescrip")
rs("PurchaseDate")=purchasedate
rs("TotalQuantity")=CInt(request("quantity"))
rs("QuantityAvailable")=CInt(request("quantity"))
rs("Available")=True

rs.update

else

rs("PurchaseDate")=purchasedate
rs("TotalQuantity")=rs("TotalQuantity")+CInt(request("quantity"))
rs("QuantityAvailable")=rs("QuantityAvailable")+CInt(request("quantity"))
rs("Available")=True
rs.update


rs.close
conn.close
set conn=nothing

end if
message="A new inventory submission has been added to the database." 
set session("type")=nothing
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page41</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var ImageButton1 = null;
var Text1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:2 }
#DBStyleImageButton1 { position:absolute;left:15;top:170;width:146;z-index:3 }
#Text1 { position:absolute;left:197;top:189;width:546;z-index:1 }
#EndOfPage { position:absolute;left:0;top:216;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/inventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="Inventory.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/inventory.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="Text1" Class="lumpkinsmall"><%=(message)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/SubmitNewInventoryInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
