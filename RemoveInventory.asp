<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set RemoveTypesRS = CreateRS("RemoveTypesRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type,  tblEquipmentType.Description FROM tblEquipmentType", "InventoryTypeID", true, 2, 3, 3, "")
Call RemoveTypesRS.Open()
Call RemoveTypesRS.ProcessAction()
Call RemoveTypesRS.SetMessages("No Records","No Records")
REM server-side recordset 
Set RemoveDescripRS = CreateRS("RemoveDescripRS", "lumpkin", "admin", "", "SELECT tblEquipment.InventoryID,  tblEquipment.Type,  tblEquipment.Description FROM tblEquipment", "InventoryID", true, 2, 3, 3, "")
Call RemoveDescripRS.Open()
Call RemoveDescripRS.ProcessAction()
Call RemoveDescripRS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page55</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text1 = null;
var ImageButton1 = null;
var type = null;
var descrip = null;
var ImageButton2 = null;
var quantity = null;
var Text2 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/RemoveInventory.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:162;top:0;width:638;z-index:1 }
#Text1 { position:absolute;left:223;top:177;width:477;z-index:3 }
#DBStyleImageButton1 { position:absolute;left:16;top:194;width:146;z-index:2 }
#DBStyletype { position:absolute;left:223;top:215;width:121;z-index:4 }
#DBStyledescrip { position:absolute;left:223;top:259;width:155;z-index:5 }
#DBStyleImageButton2 { position:absolute;left:495;top:314;width:135;z-index:8 }
#DBStylequantity { position:absolute;left:308;top:323;width:68;z-index:6 }
#Text2 { position:absolute;left:223;top:328;width:95;z-index:7 }
#EndOfPage { position:absolute;left:0;top:368;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="RemoveInventory.asp">
  
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/inventhead.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="Text1" Class="lumpkinsmall">Please select the type and quantity of inventory to remove from the database:<BR>
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="Inventory.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/inventory.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<SELECT  ID="DBStyletype" NAME="type"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(RemoveTypesRS, "Type", "InventoryTypeID", "", false, false, true))%></SELECT>
<SELECT  ID="DBStyledescrip" NAME="descrip"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(RemoveDescripRS, "Description", "InventoryID", "", false, false, true))%></SELECT>
<DIV ID="DBStyleImageButton2">
<A HREF="JavaScript:_ImageButton2_onClick()"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton2" Name="ImageButton2" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<INPUT TYPE="text" ID="DBStylequantity" NAME="quantity" SIZE="6" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text2" Class="lumpkinsmall">Quantity:<BR>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="RemoveTypesRS_Action">
  <INPUT TYPE="Hidden" NAME="RemoveTypesRS_Position" VALUE="<%=(RemoveTypesRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="RemoveDescripRS_Action">
  <INPUT TYPE="Hidden" NAME="RemoveDescripRS_Position" VALUE="<%=(RemoveDescripRS.GetState())%>">
</FORM>
<SCRIPT SRC="media/RemoveInventoryInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
