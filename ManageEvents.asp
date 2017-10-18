<%@ Language=VBScript%>
<%
dim SQL

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEvents WHERE EventID="&session("programnum")
rs.open SQL,conn, 1,2

programname=rs("Title")
description=rs("Desc")
dates=rs("Date")
start=rs("Start")
endtime=rs("End")
location=rs("Location")
participants=rs("Group1")
participants2=rs("Group2")
participants3=rs("Group3")
notes=rs("Notes")

rs.close
conn.close
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<%
REM Connection Cache 
Set ConnectionCache = CreateConnectionCache()
REM server-side recordset 
Set locationchangeRS = CreateRS("locationchangeRS", "lumpkin", "admin", "", "SELECT tblFacilities.FacilityID, tblFacilities.Name FROM tblFacilities", "FacilityID", true, 2, 3, 3, "")
Call locationchangeRS.Open()
Call locationchangeRS.ProcessAction()
Call locationchangeRS.SetMessages("No Records","No Records")

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page46</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var eventname = null;
var Text2 = null;
var description = null;
var Text3 = null;
var ImageButton4 = null;
var eventdate = null;
var Text11 = null;
var starttime = null;
var Text4 = null;
var ImageButton1 = null;
var Text5 = null;
var endtime = null;
var ImageButton6 = null;
var location = null;
var Text6 = null;
var participants = null;
var Text7 = null;
var participants2 = null;
var Text8 = null;
var participants3 = null;
var Text9 = null;
var notes = null;
var Text10 = null;
var ImageButton9 = null;
var FormButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBEdtBox.js"></SCRIPT>
<SCRIPT SRC="media/DBFrmBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBLstBox.js"></SCRIPT>
<SCRIPT SRC="media/Client_Gen_3.0_Recordset.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/ManageEvents.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleeventname { position:absolute;left:357;top:90;z-index:1 }
#Text2 { position:absolute;left:199;top:92;width:156;z-index:10 }
#DBStyledescription { position:absolute;left:357;top:132;z-index:2 }
#Text3 { position:absolute;left:199;top:134;width:160;z-index:11 }
#DBStyleImageButton4 { position:absolute;left:16;top:176;width:146;z-index:23 }
#DBStyleeventdate { position:absolute;left:357;top:189;z-index:20 }
#Text11 { position:absolute;left:199;top:192;width:164;z-index:19 }
#DBStylestarttime { position:absolute;left:357;top:224;z-index:3 }
#Text4 { position:absolute;left:199;top:227;width:164;z-index:12 }
#DBStyleImageButton1 { position:absolute;left:16;top:233;width:146;z-index:24 }
#Text5 { position:absolute;left:199;top:264;width:159;z-index:13 }
#DBStyleendtime { position:absolute;left:357;top:264;z-index:4 }
#DBStyleImageButton6 { position:absolute;left:16;top:287;width:146;z-index:25 }
#DBStylelocation { position:absolute;left:357;top:309;width:281;z-index:5 }
#Text6 { position:absolute;left:199;top:312;width:150;z-index:14 }
#DBStyleparticipants { position:absolute;left:357;top:359;z-index:6 }
#Text7 { position:absolute;left:199;top:366;width:154;z-index:15 }
#DBStyleparticipants2 { position:absolute;left:357;top:407;z-index:7 }
#Text8 { position:absolute;left:199;top:413;width:160;z-index:16 }
#DBStyleparticipants3 { position:absolute;left:357;top:447;z-index:8 }
#Text9 { position:absolute;left:199;top:454;width:163;z-index:17 }
#DBStylenotes { position:absolute;left:357;top:491;z-index:9 }
#Text10 { position:absolute;left:199;top:499;width:156;z-index:18 }
#DBStyleImageButton9 { position:absolute;left:604;top:520;width:135;z-index:22 }
#DBStyleFormButton1 { position:absolute;left:553;top:593;width:133;z-index:21 }
#EndOfPage { position:absolute;left:0;top:630;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="ManageEvents.asp">
  
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<INPUT TYPE="text" ID="DBStyleeventname" NAME="eventname" SIZE="32" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(programname)%>">
<DIV ID="Text2" Class="lumpkinbold">Name of event:<BR>
</DIV>
<TEXTAREA ID="DBStyledescription" NAME="description" Class="lumpkinsmall" tabIndex="0" COLS="32" ROWS="3" WRAP=virtual >
<%=(description)%></TEXTAREA>
<DIV ID="Text3" Class="lumpkinbold">Event Description:<BR>
</DIV>
<DIV ID="DBStyleImageButton4">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton4" Name="ImageButton4" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<INPUT TYPE="text" ID="DBStyleeventdate" NAME="eventdate" SIZE="14" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(dates)%>">
<DIV ID="Text11" Class="lumpkinbold">Event Date:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStylestarttime" NAME="starttime" SIZE="14" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(start)%>">
<DIV ID="Text4" Class="lumpkinbold">Start Time:<BR>
</DIV>
<DIV ID="DBStyleImageButton1">
<A HREF="SearchEvents.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/searchevents.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="Text5" Class="lumpkinbold">End Time:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyleendtime" NAME="endtime" SIZE="14" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(endtime)%>">
<DIV ID="DBStyleImageButton6">
<A HREF="NewEvent.asp"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton6" Name="ImageButton6" Class="Normal" SRC="media/newevent.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<SELECT  ID="DBStylelocation" NAME="location"  Class="lumpkinsmall" tabIndex="0" ><%=(makeOptionFromRecordset(locationchangeRS, "Name", "FacilityID", location, false, false, false))%></SELECT>
<DIV ID="Text6" Class="lumpkinbold">Location:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyleparticipants" NAME="participants" SIZE="32" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(participants)%>">
<DIV ID="Text7" Class="lumpkinbold">Participants:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyleparticipants2" NAME="participants2" SIZE="32" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(participants2)%>">
<DIV ID="Text8" Class="lumpkinbold">Participants, Group Two:<BR>
</DIV>
<INPUT TYPE="text" ID="DBStyleparticipants3" NAME="participants3" SIZE="32" Class="lumpkinsmall" tabIndex="0" maxLength=100  VALUE="<%=(participants3)%>">
<DIV ID="Text9" Class="lumpkinbold">Participants, Group Three:<BR>
</DIV>
<TEXTAREA ID="DBStylenotes" NAME="notes" Class="lumpkinsmall" tabIndex="0" COLS="32" ROWS="4" WRAP=virtual >
<%=(notes)%></TEXTAREA>
<DIV ID="Text10" Class="lumpkinbold">Notes:<BR>
</DIV>
<DIV ID="DBStyleImageButton9">
<A HREF="JavaScript:_ImageButton9_onClick()"  onMouseOver="_B__onMouseOver( 3);"  onMouseOut="_B__onMouseOut( 3);">
<IMG ID="ImageButton9" Name="ImageButton9" Class="Normal" SRC="media/submit.jpg" Width="135" Height="44" Border="0"></A>
</DIV>
<INPUT TYPE=Submit  ID="DBStyleFormButton1" NAME="FormButton1" value="Submit Edits" onClick="return _FormButton1__onClick()"  Class="lumpkin1" tabIndex="0">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
  <INPUT TYPE="Hidden" NAME="locationchangeRS_Action">
  <INPUT TYPE="Hidden" NAME="locationchangeRS_Position" VALUE="<%=(locationchangeRS.GetState())%>">
</FORM>
<SCRIPT SRC="media/ManageEventsInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
