<%@ Language=VBScript%>
<%
'============pull bio info again for waiver

dim SQL

session("dateofbirth")=Request("mm")&"/"&Request("dd")&"/"&Request("yy")

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblParticipants WHERE P_LastName='"&request("lastname")&"' AND P_FirstName='"&request("firstname")&"' AND DOB=#"&CDate(session("dateofbirth"))&"#"

rs.open SQL,conn, 1,2

If rs.eof then 

Response.redirect "NoRecord.asp"

end if

If not rs.eof then

session("participantID")=rs("ParticipantID")
session("plname")=rs("P_LastName")
session("pfname")=rs("P_FirstName")
dob=rs("DOB")
gender=rs("gender")
paddress=rs("P_Address")
pcity=rs("P_City")
pstate=rs("P_State")
pZIP=rs("P_ZIP")
pphone=rs("P_Phone")
glname=rs("G_LastName")
gfname=rs("G_FirstName")
ghphone=rs("G_HomePhone")
gwphone=rs("G_WorkPhone")

end if 

rs.close
conn.close
set rs=nothing


'=====pull program name for liability waiver

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblProgramsbyParticipant WHERE ParticipantID="&session("participantID")
rs.open SQL,conn, 1,2

programid=rs("ProgramID")

rs.close
conn.close
set conn=nothing

	set conn = Server.CreateObject("ADODB.Connection")
	conn.open "lumpkin"
	set rs = Server.CreateObject("ADODB.Recordset")
	SQL = "Select * from tblPrograms WHERE ProgramID="&programid
	rs.open SQL,conn, 1,2
	
	nameofprogram=rs("ProgramName")

rs.close
conn.close
set conn=nothing
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page51</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text1 = null;
var Text2 = null;
var Text3 = null;
var Text4 = null;
var Text5 = null;
var Text6 = null;
var Text8 = null;
var FormButton1 = null;
var Text17 = null;
var Text16 = null;
var Text10 = null;
var Text9 = null;
var Text11 = null;
var Text12 = null;
var Text13 = null;
var Text18 = null;
var Text19 = null;
var Text20 = null;
var Text22 = null;
var Text14 = null;
var Text15 = null;
var Text24 = null;
var Text25 = null;
var Text26 = null;
var Text27 = null;
var Text28 = null;
var Text29 = null;
var Text30 = null;
var Radio1 = null;
var Radio2 = null;
var Text31 = null;
var Text32 = null;
var Text33 = null;
var insuranceco = null;
var Text34 = null;
var Radio3 = null;
var Radio4 = null;
var Text35 = null;
var Text36 = null;
var Text37 = null;
var limitations = null;
var Text38 = null;
var Text39 = null;
var medicine = null;
var Text41 = null;
var Radio5 = null;
var Radio11 = null;
var Text42 = null;
var Text43 = null;
var Radio12 = null;
var Radio6 = null;
var Text45 = null;
var Text44 = null;
var Radio13 = null;
var Radio7 = null;
var Text47 = null;
var Text46 = null;
var Radio14 = null;
var Text49 = null;
var Radio8 = null;
var Text48 = null;
var Radio15 = null;
var Text51 = null;
var Radio9 = null;
var Text50 = null;
var Radio16 = null;
var Text53 = null;
var Radio10 = null;
var Text52 = null;
var Text54 = null;
var Text55 = null;
var Text56 = null;
var Text57 = null;
var Text58 = null;
var Text59 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBFrmBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBRadBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/LiabilityForm.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:355;top:21;width:404;z-index:1 }
#Text2 { position:absolute;left:426;top:42;width:250;z-index:2 }
#Text3 { position:absolute;left:417;top:63;width:314;z-index:3 }
#Text4 { position:absolute;left:459;top:109;width:250;z-index:4 }
#Text5 { position:absolute;left:655;top:132;width:250;z-index:5 }
#Text6 { position:absolute;left:655;top:155;width:250;z-index:6 }
#Text8 { position:absolute;left:228;top:183;width:250;z-index:7 }
#DBStyleFormButton1 { position:absolute;left:2;top:204;width:153;z-index:82 }
#DBStyleHorzRule1 { position:absolute;left:234;top:212;width:739;z-index:21 }
#Text17 { position:absolute;left:724;top:237;width:250;z-index:16 }
#Text16 { position:absolute;left:644;top:238;width:68;z-index:15 }
#Text10 { position:absolute;left:353;top:239;width:250;z-index:9 }
#Text9 { position:absolute;left:228;top:240;width:115;z-index:8 }
#Text11 { position:absolute;left:353;top:261;width:250;z-index:10 }
#Text12 { position:absolute;left:230;top:283;width:115;z-index:11 }
#Text13 { position:absolute;left:353;top:283;width:250;z-index:12 }
#Text18 { position:absolute;left:643;top:284;width:68;z-index:17 }
#Text19 { position:absolute;left:724;top:284;width:250;z-index:18 }
#Text20 { position:absolute;left:643;top:310;width:68;z-index:19 }
#Text22 { position:absolute;left:724;top:310;width:132;z-index:20 }
#Text14 { position:absolute;left:229;top:312;width:115;z-index:13 }
#Text15 { position:absolute;left:353;top:312;width:250;z-index:14 }
#DBStyleHorzRule2 { position:absolute;left:234;top:351;width:739;z-index:22 }
#Text24 { position:absolute;left:228;top:376;width:143;z-index:23 }
#Text25 { position:absolute;left:388;top:376;width:250;z-index:24 }
#Text26 { position:absolute;left:388;top:399;width:250;z-index:25 }
#Text27 { position:absolute;left:227;top:430;width:143;z-index:26 }
#Text28 { position:absolute;left:388;top:430;width:250;z-index:27 }
#Text29 { position:absolute;left:227;top:456;width:143;z-index:28 }
#Text30 { position:absolute;left:388;top:456;width:250;z-index:29 }
#DBStyleHorzRule3 { position:absolute;left:234;top:497;width:739;z-index:30 }
#DBStyleRadio1 { position:absolute;left:506;top:526;width:20;z-index:32 }
#DBStyleRadio2 { position:absolute;left:582;top:526;width:20;z-index:33 }
#Text31 { position:absolute;left:227;top:528;width:250;z-index:31 }
#Text32 { position:absolute;left:471;top:528;width:30;z-index:34 }
#Text33 { position:absolute;left:545;top:528;width:30;z-index:35 }
#DBStyleinsuranceco { position:absolute;left:407;top:554;width:312;z-index:37 }
#Text34 { position:absolute;left:227;top:559;width:161;z-index:36 }
#DBStyleHorzRule4 { position:absolute;left:234;top:616;width:739;z-index:45 }
#DBStyleRadio3 { position:absolute;left:640;top:641;width:20;z-index:39 }
#DBStyleRadio4 { position:absolute;left:716;top:641;width:20;z-index:40 }
#Text35 { position:absolute;left:228;top:643;width:349;z-index:38 }
#Text36 { position:absolute;left:605;top:643;width:30;z-index:41 }
#Text37 { position:absolute;left:679;top:643;width:30;z-index:42 }
#DBStylelimitations { position:absolute;left:408;top:669;width:311;z-index:44 }
#Text38 { position:absolute;left:228;top:674;width:161;z-index:43 }
#DBStyleHorzRule5 { position:absolute;left:232;top:731;width:739;z-index:46 }
#Text39 { position:absolute;left:228;top:757;width:195;z-index:47 }
#DBStylemedicine { position:absolute;left:432;top:758;width:341;z-index:48 }
#DBStyleHorzRule6 { position:absolute;left:232;top:831;width:739;z-index:49 }
#Text41 { position:absolute;left:578;top:862;width:91;z-index:50 }
#DBStyleRadio5 { position:absolute;left:358;top:889;width:20;z-index:63 }
#DBStyleRadio11 { position:absolute;left:715;top:889;width:20;z-index:69 }
#Text42 { position:absolute;left:250;top:891;width:96;z-index:51 }
#Text43 { position:absolute;left:602;top:891;width:96;z-index:52 }
#DBStyleRadio12 { position:absolute;left:715;top:914;width:20;z-index:70 }
#DBStyleRadio6 { position:absolute;left:358;top:915;width:20;z-index:64 }
#Text45 { position:absolute;left:602;top:916;width:96;z-index:54 }
#Text44 { position:absolute;left:250;top:917;width:96;z-index:53 }
#DBStyleRadio13 { position:absolute;left:715;top:939;width:20;z-index:71 }
#DBStyleRadio7 { position:absolute;left:358;top:941;width:20;z-index:65 }
#Text47 { position:absolute;left:602;top:941;width:96;z-index:56 }
#Text46 { position:absolute;left:250;top:943;width:96;z-index:55 }
#DBStyleRadio14 { position:absolute;left:715;top:964;width:20;z-index:72 }
#Text49 { position:absolute;left:602;top:966;width:96;z-index:58 }
#DBStyleRadio8 { position:absolute;left:358;top:968;width:20;z-index:66 }
#Text48 { position:absolute;left:250;top:970;width:96;z-index:57 }
#DBStyleRadio15 { position:absolute;left:715;top:990;width:20;z-index:73 }
#Text51 { position:absolute;left:602;top:992;width:96;z-index:60 }
#DBStyleRadio9 { position:absolute;left:358;top:994;width:20;z-index:67 }
#Text50 { position:absolute;left:250;top:996;width:96;z-index:59 }
#DBStyleRadio16 { position:absolute;left:715;top:1015;width:20;z-index:74 }
#Text53 { position:absolute;left:602;top:1017;width:96;z-index:62 }
#DBStyleRadio10 { position:absolute;left:358;top:1020;width:20;z-index:68 }
#Text52 { position:absolute;left:250;top:1022;width:96;z-index:61 }
#DBStyleHorzRule7 { position:absolute;left:232;top:1064;width:739;z-index:75 }
#Text54 { position:absolute;left:475;top:1082;width:250;z-index:76 }
#Text55 { position:absolute;left:290;top:1114;width:627;z-index:77 }
#Text56 { position:absolute;left:237;top:1229;width:115;z-index:78 }
#Text57 { position:absolute;left:364;top:1229;width:250;z-index:79 }
#Text58 { position:absolute;left:654;top:1229;width:56;z-index:80 }
#Text59 { position:absolute;left:726;top:1229;width:250;z-index:81 }
#EndOfPage { position:absolute;left:0;top:1255;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="LiabilityForm.asp">
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1">Lumpkin County Parks and Recreation Department<BR>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall">365 Riley Road, Dahlonega, GA&nbsp; 30533<BR>
</DIV>
<DIV ID="Text3" Class="lumpkinsmall">Phone: (706) 864-3622&nbsp; Fax:&nbsp; (706) 864-9106<BR>
</DIV>
<DIV ID="Text4" Class="lumpkin1">Liability Release Form<BR>
</DIV>
<DIV ID="Text5" Class="lumpkinsmall">Check # ______________<BR>
</DIV>
<DIV ID="Text6" Class="lumpkinsmall">Receipt #________________________<BR>
</DIV>
<DIV ID="Text8" Class="lumpkinsmall"><%=(nameofprogram)%>
</DIV>
<INPUT TYPE=Submit  ID="DBStyleFormButton1" NAME="FormButton1" value="Check off Liability" onClick="return _FormButton1__onClick()"  Class="lumpkinsmall" tabIndex="0">
<HR ID="DBStyleHorzRule1" Size="5" Width="739" >
<DIV ID="Text17" Class="lumpkinsmall"><%=(paddress)%>
</DIV>
<DIV ID="Text16" Class="lumpkinsmall">Address:<BR>
</DIV>
<DIV ID="Text10" Class="lumpkinsmall"><%=(session("plname"))%>
</DIV>
<DIV ID="Text9" Class="lumpkinsmall">Participant's Name:<BR>
</DIV>
<DIV ID="Text11" Class="lumpkinsmall"><%=(session("pfname"))%>
</DIV>
<DIV ID="Text12" Class="lumpkinsmall">Gender:<BR>
</DIV>
<DIV ID="Text13" Class="lumpkinsmall"><%=(gender)%>
</DIV>
<DIV ID="Text18" Class="lumpkinsmall">City:<BR>
</DIV>
<DIV ID="Text19" Class="lumpkinsmall"><%=(pcity)%>
</DIV>
<DIV ID="Text20" Class="lumpkinsmall">State:<BR>
</DIV>
<DIV ID="Text22" Class="lumpkinsmall"><%=(pstate)%>
</DIV>
<DIV ID="Text14" Class="lumpkinsmall">Date of birth:<BR>
</DIV>
<DIV ID="Text15" Class="lumpkinsmall"><%=(dob)%>
</DIV>
<HR ID="DBStyleHorzRule2" Size="5" Width="739" >
<DIV ID="Text24" Class="lumpkinsmall">Parent/Guardian Name:<BR>
</DIV>
<DIV ID="Text25" Class="lumpkinsmall"><%=(glname)%>
</DIV>
<DIV ID="Text26" Class="lumpkinsmall"><%=(gfname)%>
</DIV>
<DIV ID="Text27" Class="lumpkinsmall">Home Phone:<BR>
</DIV>
<DIV ID="Text28" Class="lumpkinsmall"><%=(ghphone)%>
</DIV>
<DIV ID="Text29" Class="lumpkinsmall">Work Phone:<BR>
</DIV>
<DIV ID="Text30" Class="lumpkinsmall"><%=(gwphone)%>
</DIV>
<HR ID="DBStyleHorzRule3" Size="5" Width="739" >
<INPUT TYPE="Radio"  ID="DBStyleRadio1" Name="insured" tabIndex="0" VALUE="yes"  >
<INPUT TYPE="Radio"  ID="DBStyleRadio2" Name="insured" tabIndex="0" VALUE="no"  >
<DIV ID="Text31" Class="lumpkinsmall">Is the program participant insured?<BR>
</DIV>
<DIV ID="Text32" Class="lumpkinsmall">Yes<BR>
</DIV>
<DIV ID="Text33" Class="lumpkinsmall">No<BR>
</DIV>
<TEXTAREA ID="DBStyleinsuranceco" NAME="insuranceco" Class="FixedWidthFont" tabIndex="0" COLS="37" ROWS="3" WRAP=virtual >
</TEXTAREA>
<DIV ID="Text34" Class="lumpkinsmall">If yes, name of company:<BR>
</DIV>
<HR ID="DBStyleHorzRule4" Size="5" Width="739" >
<INPUT TYPE="Radio"  ID="DBStyleRadio3" Name="limitations" tabIndex="0" VALUE="yes"  >
<INPUT TYPE="Radio"  ID="DBStyleRadio4" Name="limitations" tabIndex="0" VALUE="no"  >
<DIV ID="Text35" Class="lumpkinsmall">Does the program participant have any physical limitations?<BR>
</DIV>
<DIV ID="Text36" Class="lumpkinsmall">Yes<BR>
</DIV>
<DIV ID="Text37" Class="lumpkinsmall">No<BR>
</DIV>
<TEXTAREA ID="DBStylelimitations" NAME="limitations" Class="FixedWidthFont" tabIndex="0" COLS="36" ROWS="3" WRAP=virtual >
</TEXTAREA>
<DIV ID="Text38" Class="lumpkinsmall">If yes, please list:<BR>
</DIV>
<HR ID="DBStyleHorzRule5" Size="5" Width="739" >
<DIV ID="Text39" Class="lumpkinsmall">Please list all medications used:<BR>
</DIV>
<TEXTAREA ID="DBStylemedicine" NAME="medicine" Class="FixedWidthFont" tabIndex="0" COLS="40" ROWS="3" WRAP=virtual >
</TEXTAREA>
<HR ID="DBStyleHorzRule6" Size="5" Width="739" >
<DIV ID="Text41" Class="lumpkinsmall">Pant Size:<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio5" Name="shirt" tabIndex="0" VALUE="ysmall"  >
<INPUT TYPE="Radio"  ID="DBStyleRadio11" Name="pant" tabIndex="0" VALUE="ysmall"  >
<DIV ID="Text42" Class="lumpkinsmall">Youth Small<BR>
</DIV>
<DIV ID="Text43" Class="lumpkinsmall">Youth Small<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio12" Name="pant" tabIndex="0" VALUE="ymedium"  >
<INPUT TYPE="Radio"  ID="DBStyleRadio6" Name="shirt" tabIndex="0" VALUE="ymedium"  >
<DIV ID="Text45" Class="lumpkinsmall">Youth Medium<BR>
</DIV>
<DIV ID="Text44" Class="lumpkinsmall">Youth Medium<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio13" Name="pant" tabIndex="0" VALUE="ylarge"  >
<INPUT TYPE="Radio"  ID="DBStyleRadio7" Name="shirt" tabIndex="0" VALUE="ylarge"  >
<DIV ID="Text47" Class="lumpkinsmall">Youth Large<BR>
</DIV>
<DIV ID="Text46" Class="lumpkinsmall">Youth Large<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio14" Name="pant" tabIndex="0" VALUE="asmall"  >
<DIV ID="Text49" Class="lumpkinsmall">Adult Small<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio8" Name="shirt" tabIndex="0" VALUE="asmall"  >
<DIV ID="Text48" Class="lumpkinsmall">Adult Small<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio15" Name="pant" tabIndex="0" VALUE="amedium"  >
<DIV ID="Text51" Class="lumpkinsmall">Adult Medium<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio9" Name="shirt" tabIndex="0" VALUE="amedium"  >
<DIV ID="Text50" Class="lumpkinsmall">Adult Medium<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio16" Name="pant" tabIndex="0" VALUE="alarge"  >
<DIV ID="Text53" Class="lumpkinsmall">Adult Large<BR>
</DIV>
<INPUT TYPE="Radio"  ID="DBStyleRadio10" Name="shirt" tabIndex="0" VALUE="alarge"  >
<DIV ID="Text52" Class="lumpkinsmall">Adult Large<BR>
</DIV>
<HR ID="DBStyleHorzRule7" Size="5" Width="739" >
<DIV ID="Text54" Class="lumpkin1">Liability Release<BR>
</DIV>
<DIV ID="Text55" Class="lumpkinsmall">In signing, participants, their parent/guardian agrees to assume all risks and hazards incidental to such participation including transportation to and from the activity, and waive, release, absolve, indemnify and agree to hold harmless the Lumpkin County Parks and Recreation Department, Lumpkin County, City of Dahlonega, the organizers, sponsors, supervisors and participants, for any claim arising from injury to the participant.<BR>
</DIV>
<DIV ID="Text56" Class="lumpkinsmall">Parent Signature:<BR>
</DIV>
<DIV ID="Text57" Class="lumpkinsmall">___________________________________<BR>
</DIV>
<DIV ID="Text58" Class="lumpkinsmall">Date:<BR>
</DIV>
<DIV ID="Text59" Class="lumpkinsmall">___________________________________<BR>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
</FORM>
<SCRIPT SRC="media/LiabilityFormInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
