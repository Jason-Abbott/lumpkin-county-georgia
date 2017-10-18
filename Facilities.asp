<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set FacilityRS = CreateRS("FacilityRS", "lumpkin", "admin", "", "SELECT tblFacilities.FacilityID, tblFacilities.Name FROM tblFacilities", "FacilityID", true, 2, 3, 3, "")
Call FacilityRS.Open()
Call FacilityRS.ProcessAction()
Call FacilityRS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page7</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text1 = null;
var ImageButton3 = null;
var ImageButton9 = null;
var location = null;
var ImageButton1 = null;
var Text2 = null;
var ImageButton6 = null;
var ImageButton10 = null;
var facilityname = null;
var Text3 = null;
var ImageButton2 = null;
var ImageButton4 = null;
var Text4 = null;
var ImageButton7 = null;
var ImageButton11 = null;
var deletefacility = null;
var Text5 = null;
var ImageButton5 = null;
var ImageButton8 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/Facilities.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:17 }
#Text1 { position:absolute;left:206;top:170;width:172;z-index:1 }
#DBStyleImageButton3 { position:absolute;left:15;top:191;width:146;z-index:9 }
#DBStyleImageButton9 { position:absolute;left:586;top:198;width:135;z-index:18 }
#DBStylelocation { position:absolute;left:301;top:210;width:240;z-index:3 }
#DBStyleImageButton1 { position:absolute;left:15;top:245;width:146;z-index:10 }
#Text2 { position:absolute;left:206;top:298;width:250;z-index:2 }
#DBStyleImageButton6 { position:absolute;left:16;top:299;width:146;z-index:11 }
#DBStyleImageButton10 { position:absolute;left:639;top:332;width:135;z-index:19 }
#DBStylefacilityname { position:absolute;left:301;top:342;width:250;z-index:4 }
#Text3 { position:absolute;left:206;top:345;width:149;z-index:5 }
#DBStyleImageButton2 { position:absolute;left:16;top:351;width:146;z-index:12 }
#DBStyleImageButton4 { position:absolute;left:16;top:405;width:146;z-index:13 }
#Text4 { position:absolute;left:206;top:437;width:250;z-index:6 }
#DBStyleImageButton7 { position:absolute;left:16;top:458;width:146;z-index:14 }
#DBStyleImageButton11 { position:absolute;left:589;top:469;width:135;z-index:20 }
#DBStyledeletefacility { position:absolute;left:301;top:481;width:240;z-index:8 }
#Text5 { position:absolute;left:206;top:482;width:149;z-index:7 }
#DBStyleImageButton5 { position:absolute;left:16;top:511;width:146;z-index:15 }
#DBStyleImageButton8 { position:absolute;left:16;top:563;width:146;z-index:16 }
#EndOfPage { position:absolute;left:0;top:609;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="Facilities.asp">
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/facilitieshead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="Text1" Class="lumpkin1">View events by facility:<BR>
</DIV>
<DIV ID="DBStyleImageButton3">
<A HREF="Home.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton3" Name="ImageButton3" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton9">
<A HREF="JavaScript:_ImageButton9_onClick()"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton9" Name="ImageButton9" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<SELECT  ID="DBStylelocation" NAME="location"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(FacilityRS, "Name", "FacilityID", "", false, false, true))%></SELECT>
<DIV ID="DBStyleImageButton1">
<A HREF="SearchEvents.asp"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/searchevents.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="Text2" Class="lumpkin1">Add a new facility:<BR>
</DIV>
<DIV ID="DBStyleImageButton6">
<A HREF="NewEvent.asp"  onMouseOver="_B__onMouseOver( 3);"  onMouseOut="_B__onMouseOut( 3);">
<IMG ID="ImageButton6" Name="ImageButton6" Class="Normal" SRC="media/newevent.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton10">
<A HREF="JavaScript:_ImageButton10_onClick()"  onMouseOver="_B__onMouseOver( 4);"  onMouseOut="_B__onMouseOut( 4);">
<IMG ID="ImageButton10" Name="ImageButton10" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<INPUT TYPE="text" ID="DBStylefacilityname" NAME="facilityname" SIZE="33" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text3" Class="lumpkinsmall">Facility Name:<BR>
</DIV>
<DIV ID="DBStyleImageButton2">
<A HREF="#ImageButton2_Anchor"  onMouseOver="_B__onMouseOver( 5);"  onMouseOut="_B__onMouseOut( 5);">
<IMG ID="ImageButton2" Name="ImageButton2" Class="Normal" SRC="media/facilitieshi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton4">
<A HREF="Personnel.asp"  onMouseOver="_B__onMouseOver( 6);"  onMouseOut="_B__onMouseOut( 6);">
<IMG ID="ImageButton4" Name="ImageButton4" Class="Normal" SRC="media/personnel.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="Text4" Class="lumpkin1">Delete a facility:<BR>
</DIV>
<DIV ID="DBStyleImageButton7">
<A HREF="Inventory.asp"  onMouseOver="_B__onMouseOver( 7);"  onMouseOut="_B__onMouseOut( 7);">
<IMG ID="ImageButton7" Name="ImageButton7" Class="Normal" SRC="media/inventory.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton11">
<A HREF="JavaScript:_ImageButton11_onClick()"  onMouseOver="_B__onMouseOver( 8);"  onMouseOut="_B__onMouseOut( 8);">
<IMG ID="ImageButton11" Name="ImageButton11" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<SELECT  ID="DBStyledeletefacility" NAME="deletefacility"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(FacilityRS, "Name", "FacilityID", "", false, false, true))%></SELECT>
<DIV ID="Text5" Class="lumpkinsmall">Facility Name:<BR>
</DIV>
<DIV ID="DBStyleImageButton5">
<A HREF="Fees.asp"  onMouseOver="_B__onMouseOver( 9);"  onMouseOut="_B__onMouseOut( 9);">
<IMG ID="ImageButton5" Name="ImageButton5" Class="Normal" SRC="media/fees.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton8">
<A HREF="#ImageButton8_Anchor"  onMouseOver="_B__onMouseOver( 10);"  onMouseOut="_B__onMouseOut( 10);">
<IMG ID="ImageButton8" Name="ImageButton8" Class="Normal" SRC="media/reports.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="FacilityRS_Action">
  <INPUT TYPE="Hidden" NAME="FacilityRS_Position" VALUE="<%=(FacilityRS.GetState())%>">
</FORM>
<SCRIPT SRC="media/FacilitiesInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
