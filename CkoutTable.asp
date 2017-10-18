<%@ Language=VBScript%>
<%
dim SQL

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEquipment where Type="&request("type")&" AND Available=True"

rs.open SQL, conn, 1, 2

If rs.eof then
	message="Currently, there is no equipment of this inventory type available for checkout."
else

Response.Write "<DIV style='position:absolute;left:155;top:250;width:771;z-index:12'>"

Response.Write"<table border='1' width='100%'>"

  Response.Write"<tr>"
    Response.Write"<td width='40%'><strong>Equipment Type</strong></td>"
    Response.Write"<td width='40%'><strong>Description</strong></td>"
	Response.Write"<td width='20%'><strong>Check Out</strong></td>"
  Response.Write"</tr>"

Do While Not rs.EOF

	Response.Write "<tr>"
	Response.Write "<td>"
		Set Conn2 = Server.CreateObject("ADODB.Connection")
		Conn2.Open "lumpkin"
		set rs2 = Server.CreateObject("ADODB.Recordset")
		SQL = "Select * from tblEquipment where Type="&CInt(request("type"))
		rs2.open SQL, conn2, 1, 2
			Set Conn3 = Server.CreateObject("ADODB.Connection")
			Conn3.Open "lumpkin"
			set rs3 = Server.CreateObject("ADODB.Recordset")
			SQL = "Select * from tblEquipmentType where InventoryTypeID="&CInt(request("type"))
			rs3.open SQL, conn2, 1, 2
			Response.Write rs3("Type")
			Response.Write "</td><td>"
			Response.Write rs3("Description")
			rs3.close
			conn3.close
			set conn3=nothing
	Response.Write "</td><td>"
	Response.Write "<input type='checkbox' name='checkout"&x&"' value='paid'>"
	Response.Write "</td><td>"
	
	rs.MoveNext
		rs2.close
		conn2.close
		set conn2=nothing
Loop 

	Response.Write "</tr></table>"
	Response.Write "</DIV>"

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

<TITLE>Page43</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
//-->
</SCRIPT>

<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:186;top:155;width:530;z-index:1 }
#EndOfPage { position:absolute;left:0;top:183;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1"><%=(message)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/CkoutTableInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
