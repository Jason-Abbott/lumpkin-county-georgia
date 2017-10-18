<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set testinventRS = CreateRS("testinventRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true, 2, 3, 3, "")
Call testinventRS.Open()
Call testinventRS.ProcessAction()
Call testinventRS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page55</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var equip = null;
var Text1 = null;
var quantity = null;
var FormButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBFrmBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/InventTest.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleequip { position:absolute;left:193;top:8;width:121;z-index:3 }
#Text1 { position:absolute;left:384;top:58;width:250;z-index:1 }
#DBStylequantity { position:absolute;left:384;top:113;width:150;z-index:2 }
#DBStyleFormButton1 { position:absolute;left:384;top:189;width:93;z-index:4 }
#EndOfPage { position:absolute;left:0;top:223;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="InventTest.asp">
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<SELECT  ID="DBStyleequip" NAME="equip"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(testinventRS, "Type", "InventoryTypeID", "", false, false, true))%></SELECT>
<DIV ID="Text1" Class="lumpkinsmall">How many do you want to check out?<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylequantity" NAME="quantity" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<INPUT TYPE=Submit  ID="DBStyleFormButton1" NAME="FormButton1" value="Check Out" onClick="return _FormButton1__onClick()"  Class="lumpkinsmall" tabIndex="0">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="testinventRS_Action">
  <INPUT TYPE="Hidden" NAME="testinventRS_Position" VALUE="<%=(testinventRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="_Bindings" VALUE="">
</FORM>
<SCRIPT SRC="media/InventTestInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
