<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set NewEquipRS = CreateRS("NewEquipRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true, 2, 3, 3, "")
Call NewEquipRS.Open()
Call NewEquipRS.ProcessAction()
Call NewEquipRS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page40</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text3 = null;
var Text7 = null;
var addnewdrop = null;
var Text8 = null;
var newequipdescrip = null;
var ImageButton3 = null;
var mm = null;
var dd = null;
var yy = null;
var Text9 = null;
var ImageButton9 = null;
var Text1 = null;
var quantity = null;
var Text4 = null;
var type = null;
var Text5 = null;
var description = null;
var Text6 = null;
var ImageButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/AddNew.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text3 { position:absolute;left:192;top:39;width:258;z-index:4 }
#Text7 { position:absolute;left:192;top:75;width:129;z-index:6 }
#DBStyleaddnewdrop { position:absolute;left:401;top:75;width:121;z-index:7 }
#Text8 { position:absolute;left:192;top:118;width:129;z-index:8 }
#DBStylenewequipdescrip { position:absolute;left:405;top:121;width:282;z-index:5 }
#DBStyleImageButton3 { position:absolute;left:16;top:172;width:146;z-index:19 }
#DBStylemm { position:absolute;left:352;top:237;width:93;z-index:3 }
#DBStyledd { position:absolute;left:460;top:237;width:44;z-index:1 }
#DBStyleyy { position:absolute;left:518;top:237;width:58;z-index:2 }
#Text9 { position:absolute;left:192;top:241;width:129;z-index:9 }
#DBStyleImageButton9 { position:absolute;left:561;top:272;width:135;z-index:18 }
#Text1 { position:absolute;left:190;top:280;width:129;z-index:16 }
#DBStylequantity { position:absolute;left:352;top:280;width:150;z-index:17 }
#DBStyleHorzRule1 { position:absolute;left:207;top:339;width:552;z-index:15 }
#Text4 { position:absolute;left:204;top:375;width:250;z-index:10 }
#DBStyletype { position:absolute;left:311;top:434;width:324;z-index:11 }
#Text5 { position:absolute;left:211;top:437;width:86;z-index:13 }
#DBStyledescription { position:absolute;left:311;top:503;width:325;z-index:12 }
#Text6 { position:absolute;left:211;top:508;width:86;z-index:14 }
#DBStyleImageButton1 { position:absolute;left:658;top:605;width:135;z-index:20 }
#EndOfPage { position:absolute;left:0;top:659;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="AddNew.asp">
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text3" Class="lumpkin1">Add new equipment to inventory:<BR>
</DIV>
<DIV ID="Text7" Class="lumpkinsmall">Equipment Type<BR>
</DIV>
<SELECT  ID="DBStyleaddnewdrop" NAME="addnewdrop"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(NewEquipRS, "Type", "InventoryTypeID", "", false, false, true))%></SELECT>
<DIV ID="Text8" Class="lumpkinsmall">Description<BR>
</DIV>
<TEXTAREA ID="DBStylenewequipdescrip" NAME="newequipdescrip" Class="FixedWidthFont" tabIndex="0" COLS="33" ROWS="5" WRAP=virtual >
</TEXTAREA>
<DIV ID="DBStyleImageButton3">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton3" Name="ImageButton3" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<SELECT  ID="DBStylemm" NAME="mm"  Class="lumpkinsmall" tabIndex="0" ><Option Value="1">January
<Option Value="2">February
<Option Value="3">March
<Option Value="4">April
<Option Value="5">May
<Option Value="6">June
<Option Value="7">July
<Option Value="8">August
<Option Value="9">September
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
<DIV ID="Text9" Class="lumpkinsmall">Purchase Date<BR>
</DIV>
<DIV ID="DBStyleImageButton9">
<A HREF="JavaScript:_ImageButton9_onClick()"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton9" Name="ImageButton9" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<DIV ID="Text1" Class="lumpkinsmall">Quantity<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylequantity" NAME="quantity" SIZE="16" tabIndex="0" maxLength=100  VALUE="">
<HR ID="DBStyleHorzRule1" Size="5" Width="552" >
<DIV ID="Text4" Class="lumpkin1">Add New Inventory Type:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyletype" NAME="type" SIZE="38" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text5" Class="lumpkinsmall">Type:<BR>
</DIV>
<TEXTAREA ID="DBStyledescription" NAME="description" Class="FixedWidthFont" tabIndex="0" COLS="38" ROWS="6" WRAP=virtual >
</TEXTAREA>
<DIV ID="Text6" Class="lumpkinsmall">Description:<BR>
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="JavaScript:_ImageButton1_onClick()"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="NewEquipRS_Action">
  <INPUT TYPE="Hidden" NAME="NewEquipRS_Position" VALUE="<%=(NewEquipRS.GetState())%>">
</FORM>
<SCRIPT SRC="media/AddNewInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
