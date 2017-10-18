<%@ Language=VBScript%>
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page27</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text1 = null;
var FormButton2 = null;
var lastname = null;
var Text2 = null;
var firstname = null;
var Text3 = null;
var mm = null;
var dd = null;
var yy = null;
var Text4 = null;
var FormButton1 = null;
var Text5 = null;
var Text9 = null;
var Text6 = null;
var Text7 = null;
var Text8 = null;
var Text13 = null;
var Text14 = null;
var Text10 = null;
var Text11 = null;
var Text12 = null;
var Text19 = null;
var Text15 = null;
var Text16 = null;
var Text17 = null;
var Text18 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBFrmBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/Register.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:167;top:87;width:573;z-index:1 }
#DBStyleFormButton2 { position:absolute;left:652;top:202;width:46;z-index:5 }
#DBStylelastname { position:absolute;left:491;top:258;width:150;z-index:2 }
#Text2 { position:absolute;left:167;top:262;width:328;z-index:3 }
#DBStylefirstname { position:absolute;left:491;top:302;width:150;z-index:6 }
#Text3 { position:absolute;left:167;top:306;width:330;z-index:7 }
#DBStylemm { position:absolute;left:370;top:343;width:98;z-index:9 }
#DBStyledd { position:absolute;left:486;top:343;width:48;z-index:10 }
#DBStyleyy { position:absolute;left:554;top:343;width:64;z-index:11 }
#Text4 { position:absolute;left:167;top:347;width:150;z-index:8 }
#DBStyleFormButton1 { position:absolute;left:610;top:404;width:75;z-index:4 }
#Text5 { position:absolute;left:180;top:477;width:153;z-index:12 }
#Text9 { position:absolute;left:663;top:497;width:195;z-index:16 }
#Text6 { position:absolute;left:220;top:510;width:78;z-index:13 }
#Text7 { position:absolute;left:368;top:510;width:78;z-index:14 }
#Text8 { position:absolute;left:536;top:510;width:78;z-index:15 }
#Text13 { position:absolute;left:658;top:546;width:250;z-index:20 }
#Text14 { position:absolute;left:182;top:554;width:21;z-index:21 }
#Text10 { position:absolute;left:224;top:554;width:59;z-index:17 }
#Text11 { position:absolute;left:374;top:554;width:59;z-index:18 }
#Text12 { position:absolute;left:548;top:554;width:61;z-index:19 }
#Text19 { position:absolute;left:181;top:591;width:21;z-index:26 }
#Text15 { position:absolute;left:223;top:593;width:59;z-index:22 }
#Text16 { position:absolute;left:376;top:593;width:59;z-index:23 }
#Text17 { position:absolute;left:553;top:593;width:61;z-index:24 }
#Text18 { position:absolute;left:658;top:593;width:250;z-index:25 }
#EndOfPage { position:absolute;left:0;top:619;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="Register.asp">
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1">Welcome to the Lumpkin County Parks and Recreation program registration page.&nbsp; If you have registered for prior county programs, please use the search function provided below to retrieve your biographical information.&nbsp; If you are new to Lumpkin County Parks and Recreation programs, please click the &quot;New&quot; button which follows.<BR>
</DIV>
<INPUT TYPE=Submit  ID="DBStyleFormButton2" NAME="FormButton2" value="New" onClick="return _FormButton2__onClick()"  Class="lumpkin1" tabIndex="0">
<INPUT TYPE="text" ID="DBStylelastname" NAME="lastname" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text2" Class="lumpkin1">Please enter the participant's last name:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylefirstname" NAME="firstname" SIZE="16" Class="FixedWidthFont" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text3" Class="lumpkin1">Please enter the participant's first name:<BR>
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
<DIV ID="Text4" Class="lumpkin1">Date of Birth:<BR>
</DIV>
<INPUT TYPE=Submit  ID="DBStyleFormButton1" NAME="FormButton1" value="Search" onClick="return _FormButton1__onClick()"  Class="lumpkin1" tabIndex="0">
<DIV ID="Text5" Class="lumpkinsmall">For testing purposes, try:<BR>
</DIV>
<DIV ID="Text9" Class="lumpkinsmall">What you will see after registering this individual:<BR>
</DIV>
<DIV ID="Text6" Class="lumpkinsmall">Last Name:<BR>
</DIV>
<DIV ID="Text7" Class="lumpkinsmall">First Name:<BR>
</DIV>
<DIV ID="Text8" Class="lumpkinsmall">Date of birth:<BR>
</DIV>
<DIV ID="Text13" Class="lumpkinsmall">Successful registration; please proceed all the way to the &quot;Liability Release&quot; form.<BR>
</DIV>
<DIV ID="Text14" Class="lumpkinsmall">1.<BR>
</DIV>
<DIV ID="Text10" Class="lumpkinsmall">Brown<BR>
</DIV>
<DIV ID="Text11" Class="lumpkinsmall">Tyler<BR>
</DIV>
<DIV ID="Text12" Class="lumpkinsmall"><FONT COLOR=#000000>1/1/90<BR></FONT>
</DIV>
<DIV ID="Text19" Class="lumpkinsmall">2.<BR>
</DIV>
<DIV ID="Text15" Class="lumpkinsmall">Sampson<BR>
</DIV>
<DIV ID="Text16" Class="lumpkinsmall">Seth<BR>
</DIV>
<DIV ID="Text17" Class="lumpkinsmall"><FONT COLOR=#000000>6/1/89<BR></FONT>
</DIV>
<DIV ID="Text18" Class="lumpkinsmall">Unable to register; fees are owed.<BR>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
</FORM>
<SCRIPT SRC="media/RegisterInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
