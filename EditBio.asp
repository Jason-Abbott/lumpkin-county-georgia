<%@ Language=VBScript%>
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page30</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text15 = null;
var Text16 = null;
var p_lastname = null;
var g_lastname = null;
var Text1 = null;
var Text8 = null;
var p_firstname = null;
var g_firstname = null;
var Text2 = null;
var Text9 = null;
var dob = null;
var g_address = null;
var Text17 = null;
var Text10 = null;
var p_address = null;
var Text3 = null;
var g_city = null;
var Text11 = null;
var p_city = null;
var g_state = null;
var Text4 = null;
var Text12 = null;
var p_state = null;
var g_zip = null;
var Text5 = null;
var Text13 = null;
var p_zip = null;
var g_hphone = null;
var Text6 = null;
var Text14 = null;
var p_phone = null;
var g_wphone = null;
var Text18 = null;
var Text7 = null;
var FormButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBFrmBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/EditBio.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text15 { position:absolute;left:295;top:10;width:329;z-index:30 }
#Text16 { position:absolute;left:352;top:37;width:250;z-index:31 }
#DBStyleplastname { position:absolute;left:258;top:115;z-index:1 }
#DBStyleglastname { position:absolute;left:556;top:115;z-index:8 }
#Text1 { position:absolute;left:171;top:119;width:80;z-index:15 }
#Text8 { position:absolute;left:462;top:120;width:80;z-index:22 }
#DBStylepfirstname { position:absolute;left:258;top:150;z-index:2 }
#DBStylegfirstname { position:absolute;left:556;top:150;z-index:9 }
#Text2 { position:absolute;left:171;top:155;width:88;z-index:16 }
#Text9 { position:absolute;left:462;top:156;width:88;z-index:23 }
#DBStyledob { position:absolute;left:258;top:181;z-index:32 }
#DBStylegaddress { position:absolute;left:556;top:183;z-index:10 }
#Text17 { position:absolute;left:170;top:187;width:88;z-index:33 }
#Text10 { position:absolute;left:462;top:196;width:64;z-index:24 }
#DBStylepaddress { position:absolute;left:258;top:213;z-index:3 }
#Text3 { position:absolute;left:171;top:225;width:64;z-index:17 }
#DBStylegcity { position:absolute;left:556;top:242;z-index:11 }
#Text11 { position:absolute;left:462;top:249;width:45;z-index:25 }
#DBStylepcity { position:absolute;left:258;top:272;z-index:4 }
#DBStylegstate { position:absolute;left:556;top:276;z-index:12 }
#Text4 { position:absolute;left:171;top:278;width:45;z-index:18 }
#Text12 { position:absolute;left:462;top:283;width:50;z-index:26 }
#DBStylepstate { position:absolute;left:258;top:306;z-index:5 }
#DBStylegzip { position:absolute;left:556;top:309;z-index:13 }
#Text5 { position:absolute;left:171;top:312;width:50;z-index:19 }
#Text13 { position:absolute;left:462;top:316;width:43;z-index:27 }
#DBStylepzip { position:absolute;left:258;top:339;z-index:6 }
#DBStyleghphone { position:absolute;left:556;top:343;z-index:14 }
#Text6 { position:absolute;left:171;top:345;width:43;z-index:20 }
#Text14 { position:absolute;left:462;top:349;width:103;z-index:28 }
#DBStylepphone { position:absolute;left:258;top:373;z-index:7 }
#DBStylegwphone { position:absolute;left:556;top:375;z-index:34 }
#Text18 { position:absolute;left:462;top:381;width:103;z-index:35 }
#Text7 { position:absolute;left:171;top:383;width:57;z-index:21 }
#DBStyleFormButton1 { position:absolute;left:611;top:420;width:133;z-index:29 }
#EndOfPage { position:absolute;left:0;top:457;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="EditBio.asp">
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text15" Class="lumpkin1">Lumpkin County Parks and Recreation<BR>
</DIV>
<DIV ID="Text16" Class="lumpkinsmall">Biographical Information Form<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyleplastname" NAME="p_lastname" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("p_lastname"))%>">
<INPUT TYPE="text" ID="DBStyleglastname" NAME="g_lastname" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("g_lastname"))%>">
<DIV ID="Text1" Class="lumpkinsmall">Last Name:<BR>
</DIV>
<DIV ID="Text8" Class="lumpkinsmall">Last Name:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepfirstname" NAME="p_firstname" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("p_firstname"))%>">
<INPUT TYPE="text" ID="DBStylegfirstname" NAME="g_firstname" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("g_firstname"))%>">
<DIV ID="Text2" Class="lumpkinsmall">First Name:<BR>
</DIV>
<DIV ID="Text9" Class="lumpkinsmall">First Name:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyledob" NAME="dob" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("dob"))%>">
<TEXTAREA ID="DBStylegaddress" NAME="g_address" Class="lumpkinsmall" tabIndex="0" COLS="20" ROWS="3" WRAP=virtual >
<%=(session("g_address"))%></TEXTAREA>
<DIV ID="Text17" Class="lumpkinsmall">Date of birth:<BR>
</DIV>
<DIV ID="Text10" Class="lumpkinsmall">Address:<BR>
</DIV>
<TEXTAREA ID="DBStylepaddress" NAME="p_address" Class="lumpkinsmall" tabIndex="0" COLS="20" ROWS="3" WRAP=virtual >
<%=(session("p_address"))%></TEXTAREA>
<DIV ID="Text3" Class="lumpkinsmall">Address:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylegcity" NAME="g_city" SIZE="19" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("g_city"))%>">
<DIV ID="Text11" Class="lumpkinsmall">City:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepcity" NAME="p_city" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("p_city"))%>">
<INPUT TYPE="text" ID="DBStylegstate" NAME="g_state" SIZE="19" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("g_state"))%>">
<DIV ID="Text4" Class="lumpkinsmall">City:<BR>
</DIV>
<DIV ID="Text12" Class="lumpkinsmall">State:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepstate" NAME="p_state" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("p_state"))%>">
<INPUT TYPE="text" ID="DBStylegzip" NAME="g_zip" SIZE="19" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("g_zip"))%>">
<DIV ID="Text5" Class="lumpkinsmall">State:<BR>
</DIV>
<DIV ID="Text13" Class="lumpkinsmall">Zip:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepzip" NAME="p_zip" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("p_zip"))%>">
<INPUT TYPE="text" ID="DBStyleghphone" NAME="g_hphone" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("g_hphone"))%>">
<DIV ID="Text6" Class="lumpkinsmall">Zip:<BR>
</DIV>
<DIV ID="Text14" Class="lumpkinsmall">Home Phone:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepphone" NAME="p_phone" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("p_phone"))%>">
<INPUT TYPE="text" ID="DBStylegwphone" NAME="g_wphone" SIZE="20" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(session("g_wphone"))%>">
<DIV ID="Text18" Class="lumpkinsmall">Work Phone:<BR>
</DIV>
<DIV ID="Text7" Class="lumpkinsmall">Phone:<BR>
</DIV>
<INPUT TYPE=Submit  ID="DBStyleFormButton1" NAME="FormButton1" value="Submit Edits" onClick="return _FormButton1__onClick()"  Class="lumpkin1" tabIndex="0">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
</FORM>
<SCRIPT SRC="media/EditBioInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
