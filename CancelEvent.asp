<%@ Language=VBScript%>
<%
set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEvents WHERE EventID ="&session("programnum")

rs.open SQL,conn, 2, 2

name=rs("Title")
dates=rs("Date")
rs.delete

rs.close
conn.close
set conn = nothing

feedback = "The following event:  "&name&" has been cancelled on "&dates&"."
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page47</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var ImageButton4 = null;
var ImageButton1 = null;
var ImageButton6 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:178;top:135;width:565;z-index:1 }
#DBStyleImageButton4 { position:absolute;left:15;top:176;width:146;z-index:2 }
#DBStyleImageButton1 { position:absolute;left:15;top:233;width:146;z-index:3 }
#DBStyleImageButton6 { position:absolute;left:16;top:287;width:146;z-index:4 }
#EndOfPage { position:absolute;left:0;top:333;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1"><%=(feedback)%>
</DIV>
<DIV ID="DBStyleImageButton4">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton4" Name="ImageButton4" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="SearchEvents.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/searchevents.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton6">
<A HREF="NewEvent.asp"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton6" Name="ImageButton6" Class="Normal" SRC="media/newevent.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/CancelEventInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
