<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set EquipCkOutRS = CreateRS("EquipCkOutRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true, 2, 3, 3, "")
Call EquipCkOutRS.Open()
Call EquipCkOutRS.ProcessAction()
Call EquipCkOutRS.SetMessages("No Records","No Records")
REM server-side recordset 
Set TypeforCkOutRS = CreateRS("TypeforCkOutRS", "lumpkin", "admin", "", "SELECT tblEquipment.InventoryID,  tblEquipment.Type,  tblEquipment.Description FROM tblEquipment", "InventoryID", true, 2, 3, 3, "")
Call TypeforCkOutRS.Open()
Call TypeforCkOutRS.ProcessAction()
Call TypeforCkOutRS.SetMessages("No Records","No Records")
REM server-side recordset 
Set Programs2RS = CreateRS("Programs2RS", "lumpkin", "admin", "", "SELECT tblPrograms.ProgramID,  tblPrograms.ProgramName FROM tblPrograms", "ProgramID", true, 2, 3, 3, "")
Call Programs2RS.Open()
Call Programs2RS.ProcessAction()
Call Programs2RS.SetMessages("No Records","No Records")
REM server-side recordset 
Set PersonnelRS = CreateRS("PersonnelRS", "lumpkin", "admin", "", "SELECT * FROM tblPersonnel", "PersonnelID", true, 2, 3, 3, "")
Call PersonnelRS.Open()
Call PersonnelRS.ProcessAction()
Call PersonnelRS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page42</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var ImageButton1 = null;
var Text2 = null;
var Text3 = null;
var equiptype = null;
var quantity = null;
var Text6 = null;
var equipdescrip = null;
var mm = null;
var dd = null;
var yy = null;
var Text1 = null;
var Text7 = null;
var personnel = null;
var program = null;
var ImageButton9 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/CheckoutInventory.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:3 }
#DBStyleImageButton1 { position:absolute;left:15;top:194;width:146;z-index:16 }
#Text2 { position:absolute;left:199;top:207;width:320;z-index:1 }
#Text3 { position:absolute;left:525;top:208;width:320;z-index:7 }
#DBStyleequiptype { position:absolute;left:199;top:237;width:121;z-index:2 }
#DBStylequantity { position:absolute;left:653;top:241;width:75;z-index:8 }
#Text6 { position:absolute;left:533;top:297;width:320;z-index:9 }
#DBStyleequipdescrip { position:absolute;left:199;top:306;width:155;z-index:5 }
#DBStylemm { position:absolute;left:594;top:326;width:93;z-index:12 }
#DBStyledd { position:absolute;left:696;top:326;width:44;z-index:10 }
#DBStyleyy { position:absolute;left:746;top:326;width:58;z-index:11 }
#Text1 { position:absolute;left:203;top:383;width:296;z-index:6 }
#Text7 { position:absolute;left:529;top:384;width:320;z-index:13 }
#DBStylepersonnel { position:absolute;left:199;top:442;width:111;z-index:15 }
#DBStyleprogram { position:absolute;left:594;top:442;width:193;z-index:14 }
#DBStyleImageButton9 { position:absolute;left:706;top:529;width:135;z-index:4 }
#EndOfPage { position:absolute;left:0;top:583;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="CheckoutInventory.asp">
  
  
  
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/inventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="Inventory.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/inventory.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall">Please choose the type of equipment for checkout:<BR>
</DIV>
<DIV ID="Text3" Class="lumpkinsmall">Please enter the quantity of equipment for checkout:<BR>
</DIV>
<SELECT  ID="DBStyleequiptype" NAME="equiptype"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(EquipCkOutRS, "Type", "InventoryTypeID", "", false, false, true))%></SELECT>
<INPUT TYPE="text" ID="DBStylequantity" NAME="quantity" SIZE="7" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text6" Class="lumpkinsmall">Please enter the due date to return this equipment:<BR>
</DIV>
<SELECT  ID="DBStyleequipdescrip" NAME="equipdescrip"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(TypeforCkOutRS, "Description", "InventoryID", TypeforCkOutRS.GetColumnValue("Description"), false, false, false))%></SELECT>
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
<DIV ID="Text1" Class="lumpkinsmall">Please select the name of the individual to whom <BR>the equipment will be checked out:<BR>
</DIV>
<DIV ID="Text7" Class="lumpkinsmall">Please select the program to which this equipment is checked out:<BR>
</DIV>
<SELECT  ID="DBStylepersonnel" NAME="personnel"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(PersonnelRS, "FullName", "PersonnelID", "", false, false, true))%></SELECT>
<SELECT  ID="DBStyleprogram" NAME="program"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(Programs2RS, "ProgramName", "ProgramID", "", false, false, true))%></SELECT>
<DIV ID="DBStyleImageButton9">
<A HREF="JavaScript:_ImageButton9_onClick()"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton9" Name="ImageButton9" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="EquipCkOutRS_Action">
  <INPUT TYPE="Hidden" NAME="EquipCkOutRS_Position" VALUE="<%=(EquipCkOutRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="TypeforCkOutRS_Action">
  <INPUT TYPE="Hidden" NAME="TypeforCkOutRS_Position" VALUE="<%=(TypeforCkOutRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="Programs2RS_Action">
  <INPUT TYPE="Hidden" NAME="Programs2RS_Position" VALUE="<%=(Programs2RS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="PersonnelRS_Action">
  <INPUT TYPE="Hidden" NAME="PersonnelRS_Position" VALUE="<%=(PersonnelRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="TypeforCkOutRS_Bindings" VALUE="Description!*+FRM!*+equipdescrip!*+Description!*+FRM!*+equipdescrip">
</FORM>
<SCRIPT SRC="media/CheckoutInventoryInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
