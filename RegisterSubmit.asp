<%@ Language=VBScript%>
<%
response.write request("male")
response.write request("female")

session("dateofbirth")=Request("mm")&"/"&Request("dd")&"/"&Request("yy")

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblParticipants WHERE P_LastName='"&request("plname")&"' AND P_FirstName='"&request("pfname")&"' AND DOB=#"&CDate(session("dateofbirth"))&"#"

rs.open SQL,conn, 3,3

If rs.eof then

'individual has not previously participated in LCPR programs, so add new record

	rs.addnew

	rs("P_LastName")=request("plname")
	rs("P_FirstName")=request("pfname")
	rs("DOB")=session("dateofbirth")
		If request("gender")="male" then
			rs("Gender")="male"
		else rs("Gender")="female"
		end if
	rs("P_Address")=request("paddress")
	rs("P_City")=request("pcity")
	rs("P_State")=request("pstate")
	rs("P_ZIP")=request("pZIP")
	rs("P_Phone")=request("pphone")


	If request("glname") <>"" then
	rs("G_LastName")=request("glname")
	end if

	If request("gfname") <>"" then
	rs("G_FirstName")=request("gfname")
	end if

	If request("gaddress") <>"" then
	rs("G_Address")=request("gaddress")
	end if

	If request("gcity") <>"" then
	rs("G_City")=request("gcity")
	end if

	If request("gstate") <>"" then
	rs("G_State")=request("gstate")
	end if

	If request("gzip") <>"" then
	rs("G_ZIP")=request("gzip")
	end if

	If request("g_hphone") <>"" then
	rs("G_HomePhone")=request("g_hphone")
	end if

	If request("g_wphone") <>"" then
	rs("G_WorkPhone")=request("g_wphone")
	end if

	rs("DateStamp")=Date

	rs.update

	rs.close
	conn.close
	set rs=nothing
	set conn=nothing

'retrieve participant number for new participant

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblParticipants WHERE P_LastName='"&request("plname")&"' AND P_FirstName='"&request("pfname")&"'"

rs.open SQL,conn, 1,2

session("participantID")=rs("ParticipantID")
name= rs("P_LastName")& "," &rs("P_FirstName")
session("participantlname")=rs("P_LastName")
session("participantfname")=rs("P_FirstName")

rs.close
conn.close
set conn=nothing


else

'retrieve participant number for preexisting participant

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblParticipants WHERE P_LastName='"&request("plname")&"' AND P_FirstName='"&request("pfname")&"'"

rs.open SQL,conn, 1,2

session("participantID")=rs("ParticipantID")
name= rs("P_LastName")& "," &rs("P_FirstName")
session("participantlname")=rs("P_LastName")
session("participantfname")=rs("P_FirstName")

rs.close
conn.close
set conn=nothing

end if


'============= check to see if fees are paid, if so, register the participant in a new program

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblProgramsbyParticipant WHERE ParticipantID="&session("participantID")

rs.open SQL,conn, 1,2

If rs.eof then

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


elseif rs("FeesPaid")=False then

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

			Set JMail = Server.CreateObject("JMail.SMTPMail") 
			JMail.ServerAddress = "PDC"
			JMail.Sender = "FlagNotification@lckids.com"
			JMail.Subject = "Failed registration attempt"
			JMail.AddRecipient "bwhite@apicbt.com"
			JMail.Body = date&vbcrlf&vbcrlf&name &" failed to register for the "&session("programname")&" program.  This individual has unpaid fees for the "&notpaid&" program."
			JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
			JMail.Execute
			set JMail = nothing

	
	end if
end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page23</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var ImageButton1 = null;
var Text1 = null;
var Text2 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#DBStyleImageButton1 { position:absolute;left:0;top:30;width:140;z-index:2 }
#Text1 { position:absolute;left:200;top:34;width:509;z-index:1 }
#Text2 { position:absolute;left:205;top:103;width:501;z-index:3 }
#EndOfPage { position:absolute;left:0;top:129;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="DBStyleImageButton1">
<A HREF="Home.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/homeup.gif" Width="140" Height="44" Border="0"></A>
</DIV>
<DIV ID="Text1" Class="lumpkin1"><%=(thanks)%>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall"><%=(owe)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/RegisterSubmitInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
