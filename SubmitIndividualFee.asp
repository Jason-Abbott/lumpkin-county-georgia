<%@ Language=VBScript%>
<%
dateofbirth=Request("mm")&"/"&Request("dd")&"/"&Request("yy")

'pull participant number

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblParticipants WHERE P_LastName='"&request("lastname")&"' AND P_FirstName='"&request("firstname")&"' AND DOB=#"&CDate("dateofbirth")&"#"

rs.open SQL,conn, 1,2

participantnum=rs("ParticipantID")

rs.close
conn.close
set conn=nothing

'change fee field in db

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblProgramsbyParticipant WHERE ParticipantID='"&participantnum

rs.open SQL,conn, 3,3

rs("FeesPaid")=True

rs.update

rs.close
conn.close
set conn=nothing

success= "Thank you.  The fee information for this individual has been updated in the database.  <br><br><a href='Fees.asp'> Please click here to submit fee information for additional individuals.</a>"
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page35</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var ImageButton3 = null;
var ImageButton2 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImageButton3 { position:absolute;left:15;top:177;width:146;z-index:1 }
#DBStyleImageButton2 { position:absolute;left:16;top:236;width:146;z-index:2 }
#EndOfPage { position:absolute;left:0;top:282;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImageButton3">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton3" Name="ImageButton3" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton2">
<A HREF="Fees.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton2" Name="ImageButton2" Class="Normal" SRC="media/feeshi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/SubmitIndividualFeeInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
