<%@ Language=VBScript%>
<%
dim SQL
x=0

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipmentType WHERE Type='"&request("type")&"'"

rs.open SQL, conn, 3, 3

If rs.eof then

rs.addnew

rs("Type")=request("type")
rs("Description")=request("description")

rs.update

rs.close
conn.close
set conn=nothing

message="A new inventory type, "&request("type")&" , has been added to the database."  
else
message="An inventory of type "&request("type")&" already exists in the database."
end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page38</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var ImageButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:184;top:150;width:546;z-index:1 }
#DBStyleImageButton1 { position:absolute;left:16;top:195;width:146;z-index:2 }
#EndOfPage { position:absolute;left:0;top:241;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkinsmall"><%=(message)%>
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="Inventory.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/inventory.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/NewTypeInventoryInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
