<%@ Language=VBScript%>
<%
location=CInt(Request("location"))
'response.write location

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEvents WHERE Location ="&location

rs.open SQL,conn, 1,2

If rs.EOF then

message= "There are no events scheduled for this location."

else

Response.Write "<P><BR><P>"

Response.Write "<DIV style='position:absolute;left:155;top:250;width:771;z-index:12'>"

Response.Write"<table border='1' width='100%'>"

  Response.Write"<tr>"
    Response.Write"<td width='20%'><strong>Scheduled Event</strong></td>"
    Response.Write"<td width='20%'><strong>Date</strong></td>"
	Response.Write"<td width='30%'><strong>Location</strong></td>"
	Response.Write"<td width='10%'><strong>Participant Group 1</strong></td>"
    Response.Write"<td width='10%'><strong>Participant Group 2</strong></td>"
	Response.Write" <td width='10%'><strong>Participant Group 3</strong></td>"
  Response.Write"</tr>"

Do While Not rs.EOF

	Response.Write "<tr>"
	Response.Write "<td>"
	
	Response.Write "<a href='EventDetail.asp?id=" & rs("EventID") &"'>"
	Response.Write rs("Title")
	Response.Write "</a>"
	Response.Write "</td><td>"
	Response.Write rs("Date")
	Response.Write "</td><td>"
		set conn2 = Server.CreateObject("ADODB.Connection")
		conn2.open "lumpkin"
		set rs2 = Server.CreateObject("ADODB.Recordset")
		SQL = "Select * from tblFacilities WHERE FacilityID ="&rs("Location")
		rs2.open SQL,conn2, 1,2
		Response.Write rs2("Name")
	Response.Write "</td><td>"
	Response.Write rs("Group1")
	Response.Write "</td><td>"
	Response.Write rs("Group2")
	Response.Write "</td><td>"
	Response.Write rs("Group3")
	Response.Write "</td><td>"
	
	rs.MoveNext
	
Loop 

	Response.Write "</tr></table>"

rs.close
conn.close
set conn = nothing

end if
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page15</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var ImageButton4 = null;
var ImageButton5 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:187;top:135;width:526;z-index:1 }
#DBStyleImageButton4 { position:absolute;left:16;top:177;width:146;z-index:2 }
#DBStyleImageButton5 { position:absolute;left:16;top:233;width:146;z-index:3 }
#EndOfPage { position:absolute;left:0;top:279;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1"><%=(message)%>
</DIV>
<DIV ID="DBStyleImageButton4">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton4" Name="ImageButton4" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton5">
<A HREF="SearchEvents.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton5" Name="ImageButton5" Class="Normal" SRC="media/searcheventshi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/ByLocationInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
