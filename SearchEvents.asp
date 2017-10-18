<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set Recordset2 = CreateRS("Recordset2", "lumpkin", "admin", "", "SELECT tblFacilities.FacilityID, tblFacilities.Name FROM tblFacilities", "FacilityID", true, 2, 3, 3, "")
Call Recordset2.Open()
Call Recordset2.ProcessAction()
Call Recordset2.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page13</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var ImageButton4 = null;
var mm = null;
var dd = null;
var yy = null;
var location = null;
var ImageButton5 = null;
var ImageButton1 = null;
var ImageButton3 = null;
var ImageButton2 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/SearchEvents.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:4 }
#DBStyleImageButton4 { position:absolute;left:15;top:179;width:146;z-index:8 }
#DBStylemm { position:absolute;left:238;top:192;width:93;z-index:10 }
#DBStyledd { position:absolute;left:340;top:192;width:44;z-index:1 }
#DBStyleyy { position:absolute;left:390;top:192;width:58;z-index:2 }
#DBStylelocation { position:absolute;left:597;top:192;width:281;z-index:3 }
#DBStyleImageButton5 { position:absolute;left:15;top:235;width:146;z-index:9 }
#DBStyleImageButton1 { position:absolute;left:257;top:284;width:165;z-index:5 }
#DBStyleImageButton3 { position:absolute;left:587;top:284;width:165;z-index:7 }
#DBStyleImageButton2 { position:absolute;left:432;top:389;width:165;z-index:6 }
#EndOfPage { position:absolute;left:0;top:453;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="SearchEvents.asp">
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/searcheventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="DBStyleImageButton4">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton4" Name="ImageButton4" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
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
<SELECT  ID="DBStylelocation" NAME="location"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(Recordset2, "Name", "FacilityID", "", false, false, true))%></SELECT>
<DIV ID="DBStyleImageButton5">
<A HREF="#ImageButton5_Anchor"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton5" Name="ImageButton5" Class="Normal" SRC="media/searcheventshi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="JavaScript:_ImageButton1_onClick()"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/srchbydate.jpg" Width="165" Height="54" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton3">
<A HREF="JavaScript:_ImageButton3_onClick()"  onMouseOver="_B__onMouseOver( 3);"  onMouseOut="_B__onMouseOut( 3);">
<IMG ID="ImageButton3" Name="ImageButton3" Class="Normal" SRC="media/srchloc.jpg" Width="165" Height="54" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton2">
<A HREF="JavaScript:_ImageButton2_onClick()"  onMouseOver="_B__onMouseOver( 4);"  onMouseOut="_B__onMouseOut( 4);">
<IMG ID="ImageButton2" Name="ImageButton2" Class="Normal" SRC="media/srchdateloc.jpg" Width="165" Height="54" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="Recordset2_Action">
  <INPUT TYPE="Hidden" NAME="Recordset2_Position" VALUE="<%=(Recordset2.GetState())%>">
</FORM>
<SCRIPT SRC="media/SearchEventsInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
