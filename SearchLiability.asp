<%@ Language=VBScript%>
<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page52</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var lastname = null;
var Text2 = null;
var firstname = null;
var Text3 = null;
var mm = null;
var dd = null;
var yy = null;
var Text4 = null;
var FormButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBFrmBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/SearchLiability.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStylelastname { position:absolute;left:550;top:153;width:150;z-index:1 }
#Text2 { position:absolute;left:226;top:157;width:328;z-index:2 }
#DBStylefirstname { position:absolute;left:550;top:197;width:150;z-index:3 }
#Text3 { position:absolute;left:226;top:201;width:330;z-index:4 }
#DBStylemm { position:absolute;left:429;top:238;width:93;z-index:6 }
#DBStyledd { position:absolute;left:545;top:238;width:44;z-index:7 }
#DBStyleyy { position:absolute;left:613;top:238;width:58;z-index:8 }
#Text4 { position:absolute;left:226;top:242;width:150;z-index:5 }
#DBStyleFormButton1 { position:absolute;left:652;top:333;width:75;z-index:9 }
#EndOfPage { position:absolute;left:0;top:370;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="SearchLiability.asp">
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<INPUT TYPE="text" ID="DBStylelastname" NAME="lastname" SIZE="19" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text2" Class="lumpkin1">Please enter the participant's last name:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylefirstname" NAME="firstname" SIZE="19" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="">
<DIV ID="Text3" Class="lumpkin1">Please enter the participant's first name:<BR>
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
<DIV ID="Text4" Class="lumpkin1">Date of Birth:<BR>
</DIV>
<INPUT TYPE=Submit  ID="DBStyleFormButton1" NAME="FormButton1" value="Search" onClick="return _FormButton1__onClick()"  Class="lumpkin1" tabIndex="0">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
</FORM>
<SCRIPT SRC="media/SearchLiabilityInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
