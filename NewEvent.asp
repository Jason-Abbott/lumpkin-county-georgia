<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set Recordset3 = CreateRS("Recordset3", "lumpkin", "admin", "", "SELECT tblFacilities.FacilityID, tblFacilities.Name FROM tblFacilities", "FacilityID", true, 2, 3, 3, "")
Call Recordset3.Open()
Call Recordset3.ProcessAction()
Call Recordset3.SetMessages("No Record","No Record")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page24</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var title = null;
var Text1 = null;
var ImageButton4 = null;
var desc = null;
var Text2 = null;
var ImageButton6 = null;
var mm = null;
var dd = null;
var yy = null;
var Text4 = null;
var r = null;
var ca = null;
var Text11 = null;
var Text12 = null;
var Text13 = null;
var recur = null;
var Text7 = null;
var Text8 = null;
var start = null;
var ampm1 = null;
var Text5 = null;
var end = null;
var ampm2 = null;
var Text6 = null;
var location = null;
var Text9 = null;
var group1 = null;
var Text3 = null;
var group2 = null;
var group3 = null;
var Text10 = null;
var notes = null;
var ImageButton9 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBChkBox.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/NewEvent.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:161;top:0;width:638;z-index:31 }
#DBStyletitle { position:absolute;left:347;top:163;width:328;z-index:1 }
#Text1 { position:absolute;left:202;top:165;width:89;z-index:6 }
#DBStyleImageButton4 { position:absolute;left:15;top:170;width:146;z-index:34 }
#DBStyledesc { position:absolute;left:347;top:202;width:328;z-index:2 }
#Text2 { position:absolute;left:202;top:208;width:136;z-index:7 }
#DBStyleImageButton6 { position:absolute;left:15;top:224;width:146;z-index:33 }
#DBStylemm { position:absolute;left:347;top:272;width:93;z-index:24 }
#DBStyledd { position:absolute;left:455;top:272;width:44;z-index:16 }
#DBStyleyy { position:absolute;left:513;top:272;width:58;z-index:17 }
#Text4 { position:absolute;left:202;top:275;width:56;z-index:9 }
#DBStyler { position:absolute;left:437;top:319;width:20;z-index:29 }
#DBStyleca { position:absolute;left:636;top:319;width:20;z-index:30 }
#Text11 { position:absolute;left:202;top:320;width:105;z-index:26 }
#Text12 { position:absolute;left:348;top:320;width:86;z-index:27 }
#Text13 { position:absolute;left:525;top:320;width:106;z-index:28 }
#DBStylerecur { position:absolute;left:447;top:354;width:33;z-index:4 }
#Text7 { position:absolute;left:204;top:359;width:233;z-index:12 }
#Text8 { position:absolute;left:500;top:359;width:60;z-index:13 }
#DBStylestart { position:absolute;left:326;top:396;width:94;z-index:3 }
#DBStyleampm1 { position:absolute;left:426;top:396;width:58;z-index:18 }
#Text5 { position:absolute;left:204;top:398;width:83;z-index:10 }
#DBStyleend { position:absolute;left:327;top:437;width:94;z-index:19 }
#DBStyleampm2 { position:absolute;left:427;top:437;width:58;z-index:20 }
#Text6 { position:absolute;left:205;top:438;width:82;z-index:11 }
#DBStylelocation { position:absolute;left:327;top:479;width:240;z-index:25 }
#Text9 { position:absolute;left:206;top:483;width:81;z-index:14 }
#DBStylegroup1 { position:absolute;left:364;top:514;width:314;z-index:5 }
#Text3 { position:absolute;left:206;top:516;width:128;z-index:8 }
#DBStylegroup2 { position:absolute;left:364;top:547;width:314;z-index:22 }
#DBStylegroup3 { position:absolute;left:364;top:580;width:314;z-index:21 }
#Text10 { position:absolute;left:208;top:617;width:54;z-index:15 }
#DBStylenotes { position:absolute;left:364;top:617;width:315;z-index:23 }
#DBStyleImageButton9 { position:absolute;left:686;top:665;width:135;z-index:32 }
#EndOfPage { position:absolute;left:0;top:719;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="NewEvent.asp">
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/neweventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<INPUT TYPE="text" ID="DBStyletitle" NAME="title" SIZE="39" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text1" Class="lumpkinsmall">Event Title:<BR>
</DIV>
<DIV ID="DBStyleImageButton4">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton4" Name="ImageButton4" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<TEXTAREA ID="DBStyledesc" NAME="desc" Class="FixedWidthFont" tabIndex="0" COLS="39" ROWS="3" WRAP=virtual >
</TEXTAREA>
<DIV ID="Text2" Class="lumpkinsmall">Event Description:<BR>
</DIV>
<DIV ID="DBStyleImageButton6">
<A HREF="NewEvent.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton6" Name="ImageButton6" Class="Normal" SRC="media/neweventhi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<SELECT  ID="DBStylemm" NAME="mm"  Class="lumpkinsmall" tabIndex="0" ><Option Value="01">January
<Option Value="02">February
<Option Value="03">March
<Option Value="04">April
<Option Value="05">May
<Option Value="06">June
<Option Value="07">July
<Option Value="08">August
<Option Value="09">September
<Option Value="10">October
<Option Value="11">November
<Option Value="12">December
</SELECT>
<SELECT  ID="DBStyledd" NAME="dd"  Class="lumpkinsmall" tabIndex="0" ><Option Value="1">1
<Option Value="2">2
<Option Value="3">3
<Option Value="4">4
<Option Value="5">5
<Option Value="6">6
<Option Value="7">7
<Option Value="8">8
<Option Value="9">9
<Option Value="10">10
<Option Value="11">11
<Option Value="12">12
<Option Value="13">13
<Option Value="14">14
<Option Value="15">15
<Option Value="16">16
<Option Value="17">17
<Option Value="18">18
<Option Value="19">19
<Option Value="20">20
<Option Value="21">21
<Option Value="22">22
<Option Value="23">23
<Option Value="24">24
<Option Value="25">25
<Option Value="26">26
<Option Value="27">27
<Option Value="28">28
<Option Value="29">29
<Option Value="30">30
<Option Value="31">31
</SELECT>
<SELECT  ID="DBStyleyy" NAME="yy"  Class="lumpkinsmall" tabIndex="0" ><Option Value="1999">1999
<Option Value="2000">2000
<Option Value="2001">2001
<Option Value="2002">2002
<Option Value="2003">2003
<Option Value="2004">2004
<Option Value="2005">2005
</SELECT>
<DIV ID="Text4" Class="lumpkinsmall">Date:<BR>
</DIV>
<INPUT TYPE="CheckBox"  ID="DBStyler" NAME="r" VALUE=""  tabIndex="0" >
<INPUT TYPE="CheckBox"  ID="DBStyleca" NAME="ca" VALUE=""  tabIndex="0" >
<DIV ID="Text11" Class="lumpkinsmall">Event Type:<BR>
</DIV>
<DIV ID="Text12" Class="lumpkinsmall">Recreation<BR>
</DIV>
<DIV ID="Text13" Class="lumpkinsmall">Cultural Arts<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylerecur" NAME="recur" SIZE="2" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text7" Class="lumpkinsmall">This event will reoccur weekly<BR>
</DIV>
<DIV ID="Text8" Class="lumpkinsmall">times.&nbsp; &nbsp; &nbsp; <BR>
</DIV>
<INPUT TYPE="text" ID="DBStylestart" NAME="start" SIZE="9" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<SELECT  ID="DBStyleampm1" NAME="ampm1"  Class="lumpkinsmall" tabIndex="0" ><Option Value="AM">A.M.
<Option Value="PM">P.M.
</SELECT>
<DIV ID="Text5" Class="lumpkinsmall">Start Time:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyleend" NAME="end" SIZE="9" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<SELECT  ID="DBStyleampm2" NAME="ampm2"  Class="lumpkinsmall" tabIndex="0" ><Option Value="AM">A.M.
<Option Value="PM">P.M.
</SELECT>
<DIV ID="Text6" Class="lumpkinsmall">End Time:<BR>
</DIV>
<SELECT  ID="DBStylelocation" NAME="location"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(Recordset3, "Name", "FacilityID", "", false, false, true))%></SELECT>
<DIV ID="Text9" Class="lumpkinsmall">Location:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylegroup1" NAME="group1" SIZE="37" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text3" Class="lumpkinsmall">Participant Groups:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylegroup2" NAME="group2" SIZE="37" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<INPUT TYPE="text" ID="DBStylegroup3" NAME="group3" SIZE="37" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text10" Class="lumpkinsmall">Notes:<BR>
</DIV>
<TEXTAREA ID="DBStylenotes" NAME="notes" Class="FixedWidthFont" tabIndex="0" COLS="37" ROWS="5" WRAP=virtual >
</TEXTAREA>
<DIV ID="DBStyleImageButton9">
<A HREF="JavaScript:_ImageButton9_onClick()"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton9" Name="ImageButton9" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="Recordset3_Action">
  <INPUT TYPE="Hidden" NAME="Recordset3_Position" VALUE="<%=(Recordset3.GetState())%>">
</FORM>
<SCRIPT SRC="media/NewEventInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
