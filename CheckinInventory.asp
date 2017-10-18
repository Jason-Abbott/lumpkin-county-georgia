<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set EquipRS = CreateRS("EquipRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true, 2, 3, 3, "")
Call EquipRS.Open()
Call EquipRS.ProcessAction()
Call EquipRS.SetMessages("No Records","No Records")
REM server-side recordset 
Set EquipWDescripRS = CreateRS("EquipWDescripRS", "lumpkin", "admin", "", "SELECT tblEquipment.InventoryID,  tblEquipment.Type,  tblEquipment.Description FROM tblEquipment", "InventoryID", true, 2, 3, 3, "")
Call EquipWDescripRS.Open()
Call EquipWDescripRS.ProcessAction()
Call EquipWDescripRS.SetMessages("No Records","No Records")
REM server-side recordset 
Set PeopleRS = CreateRS("PeopleRS", "lumpkin", "admin", "", "SELECT * FROM tblPersonnel", "PersonnelID", true, 2, 3, 3, "")
Call PeopleRS.Open()
Call PeopleRS.ProcessAction()
Call PeopleRS.SetMessages("No Records","No Records")
REM server-side recordset 
Set Programs3RS = CreateRS("Programs3RS", "lumpkin", "admin", "", "SELECT tblPrograms.ProgramID,  tblPrograms.ProgramName FROM tblPrograms", "ProgramID", true, 2, 3, 3, "")
Call Programs3RS.Open()
Call Programs3RS.ProcessAction()
Call Programs3RS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page59</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text2 = null;
var type = null;
var equiptype = null;
var Text1 = null;
var personnel = null;
var Text4 = null;
var program = null;
var Text3 = null;
var ImageButton9 = null;
var quantity = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/CheckinInventory.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:2 }
#Text2 { position:absolute;left:200;top:167;width:320;z-index:1 }
#DBStyletype { position:absolute;left:200;top:201;width:121;z-index:3 }
#DBStyleequiptype { position:absolute;left:200;top:252;width:155;z-index:4 }
#Text1 { position:absolute;left:200;top:304;width:320;z-index:5 }
#DBStylepersonnel { position:absolute;left:200;top:367;width:111;z-index:6 }
#Text4 { position:absolute;left:201;top:417;width:339;z-index:10 }
#DBStyleprogram { position:absolute;left:201;top:462;width:193;z-index:11 }
#Text3 { position:absolute;left:200;top:507;width:339;z-index:8 }
#DBStyleImageButton9 { position:absolute;left:599;top:510;width:135;z-index:7 }
#DBStylequantity { position:absolute;left:200;top:542;width:107;z-index:9 }
#EndOfPage { position:absolute;left:0;top:578;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="CheckinInventory.asp">
  
  
  
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/inventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="Text2" Class="lumpkinsmall">Please choose the type of equipment for checkin:<BR>
</DIV>
<SELECT  ID="DBStyletype" NAME="type"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(EquipRS, "Type", "InventoryTypeID", "", false, false, true))%></SELECT>
<SELECT  ID="DBStyleequiptype" NAME="equiptype"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(EquipWDescripRS, "Description", "InventoryID", "", false, false, true))%></SELECT>
<DIV ID="Text1" Class="lumpkinsmall">Please choose the name of the personnel returning the equipment:<BR>
</DIV>
<SELECT  ID="DBStylepersonnel" NAME="personnel"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(PeopleRS, "FullName", "PersonnelID", "", false, false, true))%></SELECT>
<DIV ID="Text4" Class="lumpkinsmall">Please select the program to which the equipment was checked out:<BR>
</DIV>
<SELECT  ID="DBStyleprogram" NAME="program"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(Programs3RS, "ProgramName", "ProgramID", "", false, false, true))%></SELECT>
<DIV ID="Text3" Class="lumpkinsmall">Please enter the quantity of equipment being returned:<BR>
</DIV>
<DIV ID="DBStyleImageButton9">
<A HREF="JavaScript:_ImageButton9_onClick()"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton9" Name="ImageButton9" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<INPUT TYPE="text" ID="DBStylequantity" NAME="quantity" SIZE="11" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="EquipRS_Action">
  <INPUT TYPE="Hidden" NAME="EquipRS_Position" VALUE="<%=(EquipRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="EquipWDescripRS_Action">
  <INPUT TYPE="Hidden" NAME="EquipWDescripRS_Position" VALUE="<%=(EquipWDescripRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="PeopleRS_Action">
  <INPUT TYPE="Hidden" NAME="PeopleRS_Position" VALUE="<%=(PeopleRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="Programs3RS_Action">
  <INPUT TYPE="Hidden" NAME="Programs3RS_Position" VALUE="<%=(Programs3RS.GetState())%>">
</FORM>
<SCRIPT SRC="media/CheckinInventoryInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
