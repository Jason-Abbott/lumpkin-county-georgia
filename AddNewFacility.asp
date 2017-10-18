<%@ Language=VBScript%>
<%
dim SQL
x=0

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblFacilities where 1=2"

rs.open SQL, conn, 3, 3

rs.addnew

rs("Name")=request("facilityname")

rs.update

rs.close
conn.close
set conn=nothing

message="A new facility, "&request("facilityname")&" , has been added to the database."
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page45</TITLE>
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
#DBStyleImage1 { position:absolute;left:161;top:0;width:638;z-index:3 }
#DBStyleImageButton1 { position:absolute;left:16;top:174;width:146;z-index:2 }
#Text1 { position:absolute;left:188;top:183;width:564;z-index:1 }
#EndOfPage { position:absolute;left:0;top:220;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/facilitieshead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="Facilities.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/facilities.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="Text1" Class="lumpkinsmall"><%=(message)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/AddNewFacilityInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
