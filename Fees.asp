<%@ Language=VBScript%>
<%
'set conn = Server.CreateObject("ADODB.Connection")

'conn.open "lumpkin"

'set rs = Server.CreateObject("ADODB.Recordset")

'SQL = "Select * from tblPrograms"

'rs.open SQL,conn, 1,2

'If rs.EOF then
'message= "Currently, there are no programs scheduled"

'else

'Response.Write "<DIV style='position:absolute;left:155;top:75;width:771;z-index:12'>"

'Response.Write"<table border='1' width='100%'>"

'  Response.Write"<tr>"
'   Response.Write"<td width='20%'><strong>Program Name</strong></td>"
'    Response.Write"<td width='40%'><strong>Description</strong></td>"
'	Response.Write"<td width='20%'><strong>Participants</strong></td>"
'	Response.Write"<td width='20%'><strong>Dates</strong></td>"
'  Response.Write"</tr>"

'Do While Not rs.EOF

'	Response.Write "<tr>"
'	Response.Write "<td>"
	
'	Response.Write "<a href='IndividualFee.asp?id=" & rs("ProgramID") &"'>"
'	Response.Write rs("ProgramName")
'	Response.Write "</a>"
'	Response.Write "</td><td>"
'	Response.Write rs("Description")
'	Response.Write "</td><td>"
'	Response.Write rs("Participants")
'	Response.Write "</td><td>"
'	Response.Write rs("Dates")
'	Response.Write "</td><td>"
	
'	rs.MoveNext
	
'Loop 

'	Response.Write "</tr></table>"
'	Response.Write "</DIV>"

'rs.close
'conn.close
'set conn = nothing

	
'end if
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set ProgramsRS = CreateRS("ProgramsRS", "lumpkin", "admin", "", "SELECT tblPrograms.ProgramID,  tblPrograms.ProgramName FROM tblPrograms", "ProgramID", true, 2, 3, 3, "")
Call ProgramsRS.Open()
Call ProgramsRS.ProcessAction()
Call ProgramsRS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page33</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var ImageButton3 = null;
var Text1 = null;
var ImageButton9 = null;
var programname = null;
var ImageButton2 = null;
var Text6 = null;
var lastname = null;
var Text2 = null;
var firstname = null;
var Text3 = null;
var ImageButton1 = null;
var mm = null;
var dd = null;
var yy = null;
var Text4 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/Fees.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImage1 { position:absolute;left:161;top:0;width:638;z-index:15 }
#DBStyleImageButton3 { position:absolute;left:15;top:178;width:146;z-index:16 }
#Text1 { position:absolute;left:195;top:183;width:293;z-index:9 }
#DBStyleImageButton9 { position:absolute;left:542;top:217;width:135;z-index:13 }
#DBStyleprogramname { position:absolute;left:345;top:229;width:193;z-index:11 }
#DBStyleImageButton2 { position:absolute;left:16;top:237;width:146;z-index:17 }
#DBStyleHorzRule1 { position:absolute;left:182;top:335;width:575;z-index:12 }
#Text6 { position:absolute;left:194;top:373;width:250;z-index:10 }
#DBStylelastname { position:absolute;left:345;top:413;width:150;z-index:1 }
#Text2 { position:absolute;left:194;top:417;width:328;z-index:2 }
#DBStylefirstname { position:absolute;left:345;top:457;width:150;z-index:3 }
#Text3 { position:absolute;left:194;top:461;width:330;z-index:4 }
#DBStyleImageButton1 { position:absolute;left:590;top:487;width:135;z-index:14 }
#DBStylemm { position:absolute;left:345;top:498;width:93;z-index:6 }
#DBStyledd { position:absolute;left:454;top:498;width:44;z-index:7 }
#DBStyleyy { position:absolute;left:516;top:498;width:58;z-index:8 }
#Text4 { position:absolute;left:194;top:502;width:150;z-index:5 }
#EndOfPage { position:absolute;left:0;top:541;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="Fees.asp">
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImage1"><IMG ID="Image1" Name="Image1" Class="Normal" SRC="media/feesheader.jpg" Width="638" Height="134" Border="0">
</DIV>
<DIV ID="DBStyleImageButton3">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton3" Name="ImageButton3" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="Text1" Class="lumpkin1">Search for Participants by Program Name<BR>
</DIV>
<DIV ID="DBStyleImageButton9">
<A HREF="JavaScript:_ImageButton9_onClick()"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton9" Name="ImageButton9" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<SELECT  ID="DBStyleprogramname" NAME="programname"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(ProgramsRS, "ProgramName", "ProgramID", "", false, false, true))%></SELECT>
<DIV ID="DBStyleImageButton2">
<A HREF="#ImageButton2_Anchor"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton2" Name="ImageButton2" Class="Normal" SRC="media/feeshi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<HR ID="DBStyleHorzRule1" Size="5" Width="575" >
<DIV ID="Text6" Class="lumpkin1">Search by Participant Name<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylelastname" NAME="lastname" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text2" Class="lumpkinsmall">Participant's last name:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylefirstname" NAME="firstname" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text3" Class="lumpkinsmall">Participant's first name:<BR>
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="JavaScript:_ImageButton1_onClick()"  onMouseOver="_B__onMouseOver( 3);"  onMouseOut="_B__onMouseOut( 3);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
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
<SELECT  ID="DBStyledd" NAME="dd"  Class="lumpkinsmall" tabIndex="0" ><Option Value="1">01
<Option Value="2">02
<Option Value="3">03
<Option Value="4">04
<Option Value="5">05
<Option Value="6">06
<Option Value="7">07
<Option Value="8">08
<Option Value="9">09
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
<SELECT  ID="DBStyleyy" NAME="yy"  Class="lumpkinsmall" tabIndex="0" ><Option Value="1996">1996
<Option Value="1995">1995
<Option Value="1994">1994
<Option Value="1993">1993
<Option Value="1992">1992
<Option Value="1991">1991
<Option Value="1990">1990
<Option Value="1989">1989
<Option Value="1988">1988
<Option Value="1987">1987
<Option Value="1986">1986
<Option Value="1985">1985
<Option Value="1984">1984
<Option Value="1983">1983
<Option Value="1982">1982
<Option Value="1981">1981
<Option Value="1980">1980
</SELECT>
<DIV ID="Text4" Class="lumpkinsmall">Date of Birth:<BR>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="ProgramsRS_Action">
  <INPUT TYPE="Hidden" NAME="ProgramsRS_Position" VALUE="<%=(ProgramsRS.GetState())%>">
  <INPUT TYPE="Hidden" NAME="_Bindings" VALUE="">
</FORM>
<SCRIPT SRC="media/FeesInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
