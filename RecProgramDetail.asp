<%@ Language=VBScript%>
<%
set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblPrograms WHERE ProgramID="&request.querystring("id")

rs.open SQL,conn, 1,2

session("programnum")=rs("ProgramID")
session("programname")=rs("ProgramName")
description=rs("Description")
participants=rs("Participants")
session("dates")=rs("StartDate")

rs.close
conn.close

link="<a href='Register.asp'>Click here to register for this program. </a>"
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page20</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var Text2 = null;
var Text3 = null;
var Text4 = null;
var Text5 = null;
var Text6 = null;
var Text7 = null;
var Text8 = null;
var Text9 = null;
//-->
</SCRIPT>

<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:171;top:98;width:124;z-index:1 }
#Text2 { position:absolute;left:371;top:129;width:250;z-index:2 }
#Text3 { position:absolute;left:171;top:185;width:155;z-index:3 }
#Text4 { position:absolute;left:371;top:218;width:340;z-index:4 }
#Text5 { position:absolute;left:171;top:303;width:95;z-index:5 }
#Text6 { position:absolute;left:371;top:324;width:250;z-index:6 }
#Text7 { position:absolute;left:171;top:384;width:76;z-index:7 }
#Text8 { position:absolute;left:371;top:412;width:250;z-index:8 }
#Text9 { position:absolute;left:212;top:488;width:461;z-index:9 }
#EndOfPage { position:absolute;left:0;top:516;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1">Program Name:<BR>
</DIV>
<DIV ID="Text2" Class="lumpkin1"><%=(session("programname"))%>
</DIV>
<DIV ID="Text3" Class="lumpkin1">Program Description:<BR>
</DIV>
<DIV ID="Text4" Class="lumpkin1"><%=(description)%>
</DIV>
<DIV ID="Text5" Class="lumpkin1">Participants:<BR>
</DIV>
<DIV ID="Text6" Class="lumpkin1"><%=(participants)%>
</DIV>
<DIV ID="Text7" Class="lumpkin1">Dates:<BR>
</DIV>
<DIV ID="Text8" Class="lumpkin1"><%=(session("dates"))%>
</DIV>
<DIV ID="Text9" Class="lumpkin1"><%=(link)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/RecProgramDetailInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
