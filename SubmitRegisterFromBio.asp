<%@ Language=VBScript%>
<%
'============== retrieve participant number

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

'SQL = "Select * from tblParticipants WHERE P_LastName='"&session("p_lastname")&"' AND P_FirstName='"&session("p_firstname")&"'"
SQL = "Select * from tblParticipants WHERE P_LastName='"&session("p_lastname")&"' AND P_FirstName='"&session("p_firstname")&"' AND DOB=#"&CDate(session("dob"))&"#"

rs.open SQL,conn, 1,2

session("participantID")=rs("ParticipantID")
session("participantlname")=rs("P_LastName")
session("participantfname")=rs("P_FirstName")
name= rs("P_LastName")& "," &rs("P_FirstName")

rs.close
conn.close
set conn=nothing


'============= check to see if fees are paid, if so, register the participant in a new program

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblProgramsbyParticipant WHERE ParticipantID="&session("participantID")

rs.open SQL,conn, 1,2

If rs("FeesPaid")=True then

	rs.addnew

	rs("ProgramID")=session("programnum")
	rs("ParticipantID")=session("participantID")

	rs.update

	rs.close
	conn.close
	set rs=nothing
	set conn=nothing

	thanks= "Thank you.  Your registration has been submitted to the database.  <br><br><a href='liability.asp'> Please click here to download and print the liability release waiver for Lumpkin County Parks and Recreation programs.</a>  <br><br> This form needs to be signed and submitted to the parks and recreation office in order for your registration to be complete."

else

debtprogram=rs("ProgramID")

owe= "Sorry, our records indicate that you have not yet paid your fees for a prior Lumpkin County program.  Please contact the LCPR office at (706) 864-3622 for more information."

'======pull program name for attempted registration

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblPrograms WHERE ProgramID="&session("programnum")

rs.open SQL,conn, 1,2

notpaid=rs("ProgramName")

rs.close
conn.close
set conn=nothing

'========pull program name for unpaid program

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblPrograms WHERE ProgramID="&debtprogram

rs.open SQL,conn, 1,2

session("programname")=rs("ProgramName")

rs.close
conn.close
set conn=nothing

'==========generate notification email to Lumpkin County staff

'			Set JMail = Server.CreateObject("JMail.SMTPMail") 
' 			JMail.ServerAddress = "PDC"
'			JMail.Sender = "mgardner@apicbt.com"
'			JMail.Subject = "Failed registration attempt"
'			JMail.AddRecipient "bwhite@apicbt.com"
'			JMail.Body = date&vbcrlf&vbcrlf&name &" failed to register for the "&session("programname")&" program.  This individual has unpaid fees for the "&notpaid&" program."
'			JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
'			JMail.Execute
'			set JMail = nothing

	
end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page29</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var Text2 = null;
//-->
</SCRIPT>

<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:224;top:40;width:530;z-index:1 }
#Text2 { position:absolute;left:222;top:82;width:543;z-index:2 }
#EndOfPage { position:absolute;left:0;top:108;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1"><%=(thanks)%>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall"><%=(owe)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/SubmitRegisterFromBioInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
