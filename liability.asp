<%@ Language=VBScript%>
<%
'=====pull program name for liability waiver

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblPrograms WHERE ProgramID="&session("programnum")

rs.open SQL,conn, 1,2

nameofprogram=rs("ProgramName")

rs.close
conn.close
set conn=nothing

'============pull bio info again for waiver

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblParticipants WHERE P_LastName='"&session("participantlname")&"' AND P_FirstName='"&session("participantfname")&"'"

rs.open SQL,conn, 1,2

If not rs.eof then

plname=rs("P_LastName")
pfname=rs("P_FirstName")
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
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page32</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text1 = null;
var Text2 = null;
var Text3 = null;
var Text4 = null;
var Text5 = null;
var Text6 = null;
var Text7 = null;
var Text8 = null;
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
var Text21 = null;
var Text23 = null;
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
var Text40 = null;
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
<SCRIPT SRC="media/DBRadBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:298;top:5;width:404;z-index:1 }
#Text2 { position:absolute;left:369;top:26;width:250;z-index:2 }
#Text3 { position:absolute;left:360;top:47;width:314;z-index:3 }
#Text4 { position:absolute;left:402;top:93;width:250;z-index:4 }
#Text5 { position:absolute;left:598;top:116;width:250;z-index:5 }
#Text6 { position:absolute;left:598;top:139;width:250;z-index:6 }
#Text7 { position:absolute;left:171;top:167;width:95;z-index:7 }
#Text8 { position:absolute;left:273;top:167;width:250;z-index:8 }
#DBStyleHorzRule1 { position:absolute;left:177;top:196;width:739;z-index:24 }
#Text17 { position:absolute;left:667;top:221;width:250;z-index:17 }
#Text16 { position:absolute;left:587;top:222;width:68;z-index:16 }
#Text10 { position:absolute;left:296;top:223;width:250;z-index:10 }
#Text9 { position:absolute;left:171;top:224;width:115;z-index:9 }
#Text11 { position:absolute;left:296;top:245;width:250;z-index:11 }
#Text12 { position:absolute;left:173;top:267;width:115;z-index:12 }
#Text13 { position:absolute;left:296;top:267;width:250;z-index:13 }
#Text18 { position:absolute;left:586;top:268;width:68;z-index:18 }
#Text19 { position:absolute;left:667;top:268;width:250;z-index:19 }
#Text20 { position:absolute;left:586;top:294;width:68;z-index:20 }
#Text22 { position:absolute;left:667;top:294;width:132;z-index:22 }
#Text21 { position:absolute;left:806;top:294;width:26;z-index:21 }
#Text23 { position:absolute;left:840;top:294;width:250;z-index:23 }
#Text14 { position:absolute;left:172;top:296;width:115;z-index:14 }
#Text15 { position:absolute;left:296;top:296;width:250;z-index:15 }
#DBStyleHorzRule2 { position:absolute;left:177;top:335;width:739;z-index:25 }
#Text24 { position:absolute;left:171;top:360;width:143;z-index:26 }
#Text25 { position:absolute;left:331;top:360;width:250;z-index:27 }
#Text26 { position:absolute;left:331;top:383;width:250;z-index:28 }
#Text27 { position:absolute;left:170;top:414;width:143;z-index:29 }
#Text28 { position:absolute;left:331;top:414;width:250;z-index:30 }
#Text29 { position:absolute;left:170;top:440;width:143;z-index:31 }
#Text30 { position:absolute;left:331;top:440;width:250;z-index:32 }
#DBStyleHorzRule3 { position:absolute;left:177;top:481;width:739;z-index:33 }
#DBStyleRadio1 { position:absolute;left:449;top:510;width:20;z-index:35 }
#DBStyleRadio2 { position:absolute;left:525;top:510;width:20;z-index:36 }
#Text31 { position:absolute;left:170;top:512;width:250;z-index:34 }
#Text32 { position:absolute;left:414;top:512;width:30;z-index:37 }
#Text33 { position:absolute;left:488;top:512;width:30;z-index:38 }
#DBStyleinsuranceco { position:absolute;left:350;top:538;width:312;z-index:40 }
#Text34 { position:absolute;left:170;top:543;width:161;z-index:39 }
#DBStyleHorzRule4 { position:absolute;left:177;top:600;width:739;z-index:48 }
#DBStyleRadio3 { position:absolute;left:583;top:625;width:20;z-index:42 }
#DBStyleRadio4 { position:absolute;left:659;top:625;width:20;z-index:43 }
#Text35 { position:absolute;left:171;top:627;width:349;z-index:41 }
#Text36 { position:absolute;left:548;top:627;width:30;z-index:44 }
#Text37 { position:absolute;left:622;top:627;width:30;z-index:45 }
#DBStylelimitations { position:absolute;left:351;top:653;width:311;z-index:47 }
#Text38 { position:absolute;left:171;top:658;width:161;z-index:46 }
#DBStyleHorzRule5 { position:absolute;left:175;top:715;width:739;z-index:49 }
#Text39 { position:absolute;left:171;top:741;width:195;z-index:50 }
#DBStylemedicine { position:absolute;left:375;top:742;width:341;z-index:51 }
#DBStyleHorzRule6 { position:absolute;left:175;top:815;width:739;z-index:52 }
#Text40 { position:absolute;left:172;top:846;width:91;z-index:53 }
#Text41 { position:absolute;left:521;top:846;width:91;z-index:54 }
#DBStyleRadio5 { position:absolute;left:301;top:873;width:20;z-index:67 }
#DBStyleRadio11 { position:absolute;left:658;top:873;width:20;z-index:73 }
#Text42 { position:absolute;left:193;top:875;width:96;z-index:55 }
#Text43 { position:absolute;left:545;top:875;width:96;z-index:56 }
#DBStyleRadio12 { position:absolute;left:658;top:898;width:20;z-index:74 }
#DBStyleRadio6 { position:absolute;left:301;top:899;width:20;z-index:68 }
#Text45 { position:absolute;left:545;top:900;width:96;z-index:58 }
#Text44 { position:absolute;left:193;top:901;width:96;z-index:57 }
#DBStyleRadio13 { position:absolute;left:658;top:923;width:20;z-index:75 }
#DBStyleRadio7 { position:absolute;left:301;top:925;width:20;z-index:69 }
#Text47 { position:absolute;left:545;top:925;width:96;z-index:60 }
#Text46 { position:absolute;left:193;top:927;width:96;z-index:59 }
#DBStyleRadio14 { position:absolute;left:658;top:948;width:20;z-index:76 }
#Text49 { position:absolute;left:545;top:950;width:96;z-index:62 }
#DBStyleRadio8 { position:absolute;left:301;top:952;width:20;z-index:70 }
#Text48 { position:absolute;left:193;top:954;width:96;z-index:61 }
#DBStyleRadio15 { position:absolute;left:658;top:974;width:20;z-index:77 }
#Text51 { position:absolute;left:545;top:976;width:96;z-index:64 }
#DBStyleRadio9 { position:absolute;left:301;top:978;width:20;z-index:71 }
#Text50 { position:absolute;left:193;top:980;width:96;z-index:63 }
#DBStyleRadio16 { position:absolute;left:658;top:999;width:20;z-index:78 }
#Text53 { position:absolute;left:545;top:1001;width:96;z-index:66 }
#DBStyleRadio10 { position:absolute;left:301;top:1004;width:20;z-index:72 }
#Text52 { position:absolute;left:193;top:1006;width:96;z-index:65 }
#DBStyleHorzRule7 { position:absolute;left:175;top:1048;width:739;z-index:79 }
#Text54 { position:absolute;left:418;top:1066;width:250;z-index:80 }
#Text55 { position:absolute;left:233;top:1098;width:627;z-index:81 }
#Text56 { position:absolute;left:180;top:1213;width:115;z-index:82 }
#Text57 { position:absolute;left:307;top:1213;width:250;z-index:83 }
#Text58 { position:absolute;left:597;top:1213;width:56;z-index:84 }
#Text59 { position:absolute;left:669;top:1213;width:250;z-index:85 }
#EndOfPage { position:absolute;left:0;top:1239;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="liability.asp">
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
<DIV ID="Text7" Class="lumpkinsmall">Program Name:<BR>
</DIV>
<DIV ID="Text8" Class="lumpkinsmall"><%=(nameofprogram)%>
</DIV>
<HR ID="DBStyleHorzRule1" Size="5" Width="739" >
<DIV ID="Text17" Class="lumpkinsmall"><%=(paddress)%>
</DIV>
<DIV ID="Text16" Class="lumpkinsmall">Address:<BR>
</DIV>
<DIV ID="Text10" Class="lumpkinsmall"><%=(plname)%>
</DIV>
<DIV ID="Text9" Class="lumpkinsmall">Participant's Name:<BR>
</DIV>
<DIV ID="Text11" Class="lumpkinsmall"><%=(pfname)%>
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
<DIV ID="Text21" Class="lumpkinsmall">ZIP:<BR>
</DIV>
<DIV ID="Text23" Class="lumpkinsmall"><%=(pZIP)%>
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
<DIV ID="Text40" Class="lumpkinsmall">Shirt Size:<BR>
</DIV>
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
<SCRIPT SRC="media/liabilityInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
