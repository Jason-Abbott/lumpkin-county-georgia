<%@ Language=VBScript%>
<%
'session("programID")= Rec or CA program number.
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page22</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text7 = null;
var Text8 = null;
var Text9 = null;
var Text20 = null;
var Text10 = null;
var Text17 = null;
var plname = null;
var glname = null;
var Text1 = null;
var Text11 = null;
var pfname = null;
var gfname = null;
var Text2 = null;
var Text12 = null;
var mm = null;
var dd = null;
var yy = null;
var gaddress = null;
var Text21 = null;
var Text13 = null;
var Radio1 = null;
var Radio2 = null;
var Text23 = null;
var Text24 = null;
var gcity = null;
var Text14 = null;
var paddress = null;
var Text3 = null;
var gstate = null;
var Text15 = null;
var pcity = null;
var Text4 = null;
var gzip = null;
var Text16 = null;
var pstate = null;
var Text5 = null;
var g_hphone = null;
var Text19 = null;
var pzip = null;
var Text6 = null;
var g_wphone = null;
var Text22 = null;
var pphone = null;
var Text18 = null;
var FormButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBFrmBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/DBRadBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/NewRegister.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text7 { position:absolute;left:478;top:6;width:308;z-index:13 }
#Text8 { position:absolute;left:458;top:48;width:535;z-index:14 }
#Text9 { position:absolute;left:458;top:78;width:351;z-index:15 }
#Text20 { position:absolute;left:246;top:126;width:734;z-index:36 }
#DBStyleHorzRule2 { position:absolute;left:159;top:180;width:805;z-index:37 }
#Text10 { position:absolute;left:244;top:204;width:250;z-index:16 }
#Text17 { position:absolute;left:680;top:204;width:291;z-index:29 }
#DBStyleplname { position:absolute;left:287;top:253;width:239;z-index:1 }
#DBStyleglname { position:absolute;left:729;top:255;width:239;z-index:18 }
#Text1 { position:absolute;left:159;top:257;width:107;z-index:7 }
#Text11 { position:absolute;left:601;top:259;width:107;z-index:23 }
#DBStylepfname { position:absolute;left:287;top:283;width:237;z-index:2 }
#DBStylegfname { position:absolute;left:729;top:285;width:239;z-index:19 }
#Text2 { position:absolute;left:159;top:286;width:107;z-index:8 }
#Text12 { position:absolute;left:601;top:288;width:107;z-index:24 }
#DBStylemm { position:absolute;left:288;top:313;width:98;z-index:39 }
#DBStyledd { position:absolute;left:404;top:313;width:48;z-index:40 }
#DBStyleyy { position:absolute;left:472;top:313;width:64;z-index:41 }
#DBStylegaddress { position:absolute;left:729;top:317;width:241;z-index:20 }
#Text21 { position:absolute;left:159;top:319;width:107;z-index:38 }
#Text13 { position:absolute;left:601;top:319;width:107;z-index:25 }
#DBStyleRadio1 { position:absolute;left:345;top:347;width:20;z-index:46 }
#DBStyleRadio2 { position:absolute;left:481;top:348;width:20;z-index:47 }
#Text23 { position:absolute;left:289;top:349;width:50;z-index:44 }
#Text24 { position:absolute;left:419;top:349;width:107;z-index:45 }
#DBStylegcity { position:absolute;left:729;top:374;width:240;z-index:21 }
#Text14 { position:absolute;left:601;top:375;width:107;z-index:26 }
#DBStylepaddress { position:absolute;left:287;top:376;width:239;z-index:3 }
#Text3 { position:absolute;left:159;top:378;width:107;z-index:9 }
#DBStylegstate { position:absolute;left:729;top:409;width:130;z-index:30 }
#Text15 { position:absolute;left:601;top:413;width:107;z-index:27 }
#DBStylepcity { position:absolute;left:287;top:433;width:237;z-index:4 }
#Text4 { position:absolute;left:159;top:434;width:107;z-index:10 }
#DBStylegzip { position:absolute;left:729;top:442;width:150;z-index:22 }
#Text16 { position:absolute;left:601;top:446;width:46;z-index:28 }
#DBStylepstate { position:absolute;left:287;top:468;width:130;z-index:17 }
#Text5 { position:absolute;left:159;top:472;width:107;z-index:11 }
#DBStyleghphone { position:absolute;left:729;top:480;width:150;z-index:31 }
#Text19 { position:absolute;left:601;top:484;width:105;z-index:33 }
#DBStylepzip { position:absolute;left:287;top:501;width:148;z-index:5 }
#Text6 { position:absolute;left:159;top:504;width:46;z-index:12 }
#DBStylegwphone { position:absolute;left:729;top:514;width:150;z-index:42 }
#Text22 { position:absolute;left:602;top:518;width:105;z-index:43 }
#DBStylepphone { position:absolute;left:287;top:538;width:150;z-index:6 }
#Text18 { position:absolute;left:159;top:542;width:55;z-index:32 }
#DBStyleHorzRule1 { position:absolute;left:159;top:572;width:805;z-index:34 }
#DBStyleFormButton1 { position:absolute;left:485;top:585;width:205;z-index:35 }
#EndOfPage { position:absolute;left:0;top:622;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff" onLoad="return document_onLoad()" >

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="NewRegister.asp">
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text7" Class="lumpkin1"><FONT SIZE="4" STYLE="font-size:14 pt"><B>Program Registration<BR></B></FONT>
</DIV>
<DIV ID="Text8" Class="lumpkin1"><%=(session("programname"))%>
</DIV>
<DIV ID="Text9" Class="lumpkin1"><%=(session("dates"))%>
</DIV>
<DIV ID="Text20" Class="lumpkin1"><FONT SIZE="2" STYLE="font-size:10 pt">To register your intent to participate in this program, please provide the following information.&nbsp; You will subsequently receive a registration confirmation letter, and request for payment of fees, from the Lumpkin County Parks and Recreation Service.<BR></FONT>
</DIV>
<HR ID="DBStyleHorzRule2" Size="5" Width="805" >
<DIV ID="Text10" Class="lumpkin1"><B>Participant Information<BR></B>
</DIV>
<DIV ID="Text17" Class="lumpkin1"><B>Guardian Information </B>(if appropriate)<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyleplname" NAME="plname" SIZE="27" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<INPUT TYPE="text" ID="DBStyleglname" NAME="glname" SIZE="27" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text1" Class="lumpkinsmall">Last Name: <BR>
</DIV>
<DIV ID="Text11" Class="lumpkinsmall">Last Name: <BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepfname" NAME="pfname" SIZE="27" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<INPUT TYPE="text" ID="DBStylegfname" NAME="gfname" SIZE="27" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text2" Class="lumpkinsmall">First Name: <BR>
</DIV>
<DIV ID="Text12" Class="lumpkinsmall">First Name: <BR>
</DIV>
<SELECT  ID="DBStylemm" NAME="mm"  Class="Normal" tabIndex="0" ><Option Value="01">January
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
<SELECT  ID="DBStyledd" NAME="dd"  Class="Normal" tabIndex="0" ><Option Value="1">01
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
<SELECT  ID="DBStyleyy" NAME="yy"  Class="Normal" tabIndex="0" ><Option Value="1996">1996
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
<TEXTAREA ID="DBStylegaddress" NAME="gaddress" Class="FixedWidthFont" tabIndex="0" COLS="28" ROWS="3" WRAP=virtual >
</TEXTAREA>
<DIV ID="Text21" Class="lumpkinsmall">Date of birth: <BR>
</DIV>
<DIV ID="Text13" Class="lumpkinsmall">Address: <BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio1" Name="gender" tabIndex="0" VALUE="male"  >
<INPUT TYPE="Radio"  ID="DBStyleRadio2" Name="gender" tabIndex="0" VALUE="female"  >
<DIV ID="Text23" Class="lumpkinsmall">Male:<BR>
</DIV>
<DIV ID="Text24" Class="lumpkinsmall">Female: <BR>
</DIV>
<INPUT TYPE="text" ID="DBStylegcity" NAME="gcity" SIZE="28" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text14" Class="lumpkinsmall">City: <BR>
</DIV>
<TEXTAREA ID="DBStylepaddress" NAME="paddress" Class="FixedWidthFont" tabIndex="0" COLS="27" ROWS="3" WRAP=virtual >
</TEXTAREA>
<DIV ID="Text3" Class="lumpkinsmall">Address: <BR>
</DIV>
<SELECT  ID="DBStylegstate" NAME="gstate" onBlur="return _gstate__onBlur()"  Class="Normal" tabIndex="0" ><Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------</SELECT>
<DIV ID="Text15" Class="lumpkinsmall">State: <BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepcity" NAME="pcity" SIZE="27" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text4" Class="lumpkinsmall">City: <BR>
</DIV>
<INPUT TYPE="text" ID="DBStylegzip" NAME="gzip" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text16" Class="lumpkinsmall">ZIP: <BR>
</DIV>
<SELECT  ID="DBStylepstate" NAME="pstate" onBlur="return _pstate__onBlur()"  Class="Normal" tabIndex="0" ><Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------<Option Value = "A" >------------------------------</SELECT>
<DIV ID="Text5" Class="lumpkinsmall">State: <BR>
</DIV>
<INPUT TYPE="text" ID="DBStyleghphone" NAME="g_hphone" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text19" Class="lumpkinsmall">Home Phone:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepzip" NAME="pzip" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text6" Class="lumpkinsmall">ZIP: <BR>
</DIV>
<INPUT TYPE="text" ID="DBStylegwphone" NAME="g_wphone" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text22" Class="lumpkinsmall">Work Phone:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylepphone" NAME="pphone" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text18" Class="lumpkinsmall">Phone:<BR>
</DIV>
<HR ID="DBStyleHorzRule1" Size="5" Width="805" >
<INPUT TYPE=Submit  ID="DBStyleFormButton1" NAME="FormButton1" value="Submit Registration" onClick="return _FormButton1__onClick()"  Class="lumpkin1" tabIndex="0">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
</FORM>
<SCRIPT SRC="media/NewRegisterInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
