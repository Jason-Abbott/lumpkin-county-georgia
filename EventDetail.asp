<%@ Language=VBScript%>
<%
dim SQL

session("programnum")=request.querystring("id")

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEvents WHERE EventID="&session("programnum")
rs.open SQL,conn, 1,2

programname=rs("Title")
description=rs("Desc")
participants=rs("Group1")
participants2=rs("Group2")
dates=rs("Date")

rs.close
conn.close
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page17</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var Text2 = null;
var Text3 = null;
var ImageButton4 = null;
var Text4 = null;
var ImageButton5 = null;
var Text5 = null;
var Text6 = null;
var Text8 = null;
var Text7 = null;
var Text9 = null;
var Text10 = null;
var Text11 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:239;top:80;width:124;z-index:1 }
#Text2 { position:absolute;left:424;top:113;width:444;z-index:5 }
#Text3 { position:absolute;left:239;top:167;width:155;z-index:2 }
#DBStyleImageButton4 { position:absolute;left:15;top:172;width:146;z-index:12 }
#Text4 { position:absolute;left:425;top:224;width:451;z-index:6 }
#DBStyleImageButton5 { position:absolute;left:15;top:228;width:146;z-index:13 }
#Text5 { position:absolute;left:239;top:285;width:95;z-index:3 }
#Text6 { position:absolute;left:425;top:320;width:387;z-index:7 }
#Text8 { position:absolute;left:426;top:352;width:386;z-index:8 }
#Text7 { position:absolute;left:239;top:366;width:76;z-index:4 }
#Text9 { position:absolute;left:429;top:405;width:250;z-index:9 }
#Text10 { position:absolute;left:384;top:471;width:402;z-index:10 }
#Text11 { position:absolute;left:427;top:520;width:250;z-index:11 }
#EndOfPage { position:absolute;left:0;top:548;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1">Program Name:<BR>
</DIV>
<DIV ID="Text2" Class="lumpkin1"><%=(programname)%>
</DIV>
<DIV ID="Text3" Class="lumpkin1">Program Description:<BR>
</DIV>
<DIV ID="DBStyleImageButton4">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton4" Name="ImageButton4" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="Text4" Class="lumpkin1"><%=(description)%>
</DIV>
<DIV ID="DBStyleImageButton5">
<A HREF="SearchEvents.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton5" Name="ImageButton5" Class="Normal" SRC="media/searcheventshi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="Text5" Class="lumpkin1">Participants:<BR>
</DIV>
<DIV ID="Text6" Class="lumpkin1"><%=(participants)%>
</DIV>
<DIV ID="Text8" Class="lumpkin1"><%=(participants2)%>
</DIV>
<DIV ID="Text7" Class="lumpkin1">Dates:<BR>
</DIV>
<DIV ID="Text9" Class="lumpkin1"><%=(dates)%>
</DIV>
<DIV ID="Text10" Class="lumpkin1"><A HREF = "ManageEvents.asp">Click here to edit or reschedule this event.</A><BR>
</DIV>
<DIV ID="Text11" Class="lumpkin1"><A HREF = "CancelEvent.asp">Click here to cancel this event</A><BR>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/EventDetailInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
