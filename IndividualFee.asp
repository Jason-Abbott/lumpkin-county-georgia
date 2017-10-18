<%@ Language=VBScript%>
<%
dim SQL, a, x

session("dateofbirth")=Request("mm")&"/"&Request("dd")&"/"&Request("yy")

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblParticipants WHERE P_LastName='"&request("lastname")&"' AND P_FirstName='"&request("firstname")&"' AND DOB=#"&CDate(session("dateofbirth"))&"#"

rs.open SQL,conn, 1,2

If rs.eof then 

sorry = "We're sorry, but no records are found for "&request("lastname")&" , "&request("firstname")  
link = "<br><br><a href='Fees.asp'>  Click here to return and search again.</a>"

else 

	response.write "Got it"
	session("personnumforfee")=rs("ParticipantID")
	response.write session("personnumforfee")
rs.close
conn.close
set conn=nothing

end if

	a=1
	set conn = Server.CreateObject("ADODB.Connection")
	conn.open "lumpkin"
	set rs = Server.CreateObject("ADODB.Recordset")
	SQL = "Select * from tblProgramsbyParticipant WHERE ParticipantID="&session("personnumforfee")
	rs.open SQL,conn, 1,2
	
	Do While Not rs.eof

	Session("myprograms"&a)=rs("ProgramID")
		rs.movenext
		a=a+1

	loop

	rs.close
	conn.close
	set conn=nothing

For x=1 to a
	Response.write Session("myparticipants"&x)
	x=x+1
Next

'	Response.Write "<DIV style='position:absolute;left:155;top:75;width:771;z-index:12'>"

'	Response.Write"<table border='1' width='100%'>"
'	  Response.Write"<tr>"
'	    Response.Write"<td width='50%'><strong>Program Name</strong></td>"
'	    Response.Write"<td width='50%'><strong>Mark Fees as Paid</strong></td>"
'	 Response.Write"</tr>"

'	Do While Not rs.EOF

'		For a = 1 to x
'			Response.Write "<tr>"
'			Response.Write "<td>"
'			Response.Write Session("participantname"&a)
'			Response.Write "</td><td>"
'			Response.Write "<input type='checkbox' name='feespaid"&x&"' value='paid'>"
'			Response.Write "</td><td>"
'			a=a+1
'		Next
'				rs.MoveNext
'	Loop 

'			Response.Write "</tr></table>"
'			Response.Write "</DIV>"

'rs.close
'conn.close
'set conn = nothing
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page34</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var ImageButton3 = null;
var Text2 = null;
var ImageButton2 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:177;top:142;width:610;z-index:1 }
#DBStyleImageButton3 { position:absolute;left:15;top:165;width:146;z-index:3 }
#Text2 { position:absolute;left:176;top:185;width:611;z-index:2 }
#DBStyleImageButton2 { position:absolute;left:16;top:224;width:146;z-index:4 }
#EndOfPage { position:absolute;left:0;top:270;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1"><%=(sorry)%>
</DIV>
<DIV ID="DBStyleImageButton3">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton3" Name="ImageButton3" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="Text2" Class="lumpkin1"><%=(link)%>
</DIV>
<DIV ID="DBStyleImageButton2">
<A HREF="Fees.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton2" Name="ImageButton2" Class="Normal" SRC="media/feeshi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/IndividualFeeInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
