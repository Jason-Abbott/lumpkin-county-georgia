<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set EquipmentRS = CreateRS("EquipmentRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true, 2, 3, 3, "")
Call EquipmentRS.Open()
Call EquipmentRS.ProcessAction()
Call EquipmentRS.SetMessages("No Records","No Records")
REM server-side recordset 
Set EquipDescripRS = CreateRS("EquipDescripRS", "lumpkin", "admin", "", "SELECT tblEquipment.InventoryID,  tblEquipment.Type,  tblEquipment.Description FROM tblEquipment", "InventoryID", true, 2, 3, 3, "")
Call EquipDescripRS.Open()
Call EquipDescripRS.ProcessAction()
Call EquipDescripRS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page37</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text2 = null;
var ImageButton3 = null;
var ImageButton9 = null;
var searchtype = null;
var equiptype = null;
var ImageButton7 = null;
var ImageButton2 = null;
var ImageButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/Inventory.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:161;top:0;width:638;z-index:7 }
#Text2 { position:absolute;left:199;top:190;width:255;z-index:1 }
#DBStyleImageButton3 { position:absolute;left:14;top:191;width:146;z-index:3 }
#DBStyleImageButton9 { position:absolute;left:575;top:231;width:135;z-index:8 }
#DBStylesearchtype { position:absolute;left:199;top:243;width:121;z-index:2 }
#DBStyleequiptype { position:absolute;left:388;top:243;width:155;z-index:9 }
#DBStyleImageButton7 { position:absolute;left:15;top:246;width:146;z-index:4 }
#DBStyleImageButton2 { position:absolute;left:227;top:379;width:165;z-index:6 }
#DBStyleImageButton1 { position:absolute;left:489;top:379;width:165;z-index:5 }
#EndOfPage { position:absolute;left:0;top:443;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="Inventory.asp">
  
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/inventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="Text2" Class="lumpkinsmall">Search for existing equipment:<BR>
</DIV>
<DIV ID="DBStyleImageButton3">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton3" Name="ImageButton3" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton9">
<A HREF="JavaScript:_ImageButton9_onClick()"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton9" Name="ImageButton9" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<SELECT  ID="DBStylesearchtype" NAME="searchtype" onChange="return _searchtype__onChange()"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(EquipmentRS, "Type", "InventoryTypeID", "", false, false, true))%></SELECT>
<SELECT  ID="DBStyleequiptype" NAME="equiptype"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(EquipDescripRS, "Description", "InventoryID", EquipDescripRS.GetColumnValue("Description"), false, false, false))%></SELECT>
<DIV ID="DBStyleImageButton7">
<A HREF="Inventory.asp"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton7" Name="ImageButton7" Class="Normal" SRC="media/inventoryhi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton2">
<A HREF="JavaScript:_ImageButton2_onClick()"  onMouseOver="_B__onMouseOver( 3);"  onMouseOut="_B__onMouseOut( 3);">
<IMG ID="ImageButton2" Name="ImageButton2" Class="Normal" SRC="media/addinvent.jpg" Width="165" Height="54" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="JavaScript:_ImageButton1_onClick()"  onMouseOver="_B__onMouseOver( 4);"  onMouseOut="_B__onMouseOut( 4);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/ckoutinvent.jpg" Width="165" Height="54" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="EquipmentRS_Action">
  <INPUT TYPE="Hidden" NAME="EquipmentRS_Position" VALUE="<%=(EquipmentRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="EquipDescripRS_Action">
  <INPUT TYPE="Hidden" NAME="EquipDescripRS_Position" VALUE="<%=(EquipDescripRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="EquipDescripRS_Bindings" VALUE="Description!*+FRM!*+equiptype!*+Description!*+FRM!*+equiptype">
</FORM>
<SCRIPT SRC="media/InventoryInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
