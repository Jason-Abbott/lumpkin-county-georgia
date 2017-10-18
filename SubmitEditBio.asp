<%@ Language=VBScript%>
<%
set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

'SQL = "Select * from tblParticipants WHERE P_LastName='"&request("p_lastname")&"' AND P_FirstName='"&request("p_firstname")&"'"
SQL = "Select * from tblParticipants WHERE P_LastName='"&request("p_lastname")&"' AND P_FirstName='"&request("P_firstname")&"' AND DOB=#"&CDate(session("dateofbirth"))&"#"

rs.open SQL,conn, 3,3

	rs("P_LastName")=request("p_lastname")
	rs("P_FirstName")=request("p_firstname")
	rs("DOB")=request("dob")
	rs("P_Address")=request("p_address")
	rs("P_City")=request("p_city")
	rs("P_State")=request("p_state")
	rs("P_Zip")=request("p_zip")
	rs("P_Phone")=request("p_phone")
	
		If request("g_lastname")<>"" then
		rs("G_LastName")=request("g_lastname")
		rs("G_FirstName")=request("g_firstname")
		rs("G_Address")=request("g_address")
		rs("G_City")=request("g_city")
		rs("G_State")=request("g_state")
		rs("G_Zip")=request("g_zip")
		rs("G_HomePhone")=request("g_hphone")
		rs("G_WorkPhone")=request("g_wphone")
		end if

rs.update

rs.close
conn.close
set conn=nothing
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page31</TITLE>

<SCRIPT SRC="media/DBRelPos.js" LANGUAGE="JavaScript"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#EndOfPage { position:absolute;left:0;top:10;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
