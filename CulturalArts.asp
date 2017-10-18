<%@ Language=VBScript%>
<%
set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblPrograms WHERE CAProgram=True"

rs.open SQL,conn, 1,2

If rs.EOF then
message= "Currently, there are no cultural arts programs scheduled"

else

Response.Write "<DIV style='position:absolute;left:155;top:250;width:771;z-index:12'>"
'Response.Write "<P><BR><P>"

Response.Write"<table border='1' width='100%'>"

  Response.Write"<tr>"
    Response.Write"<td width='20%'><strong>Cultural Arts Program</strong></td>"
    Response.Write"<td width='40%'><strong>Description</strong></td>"
	Response.Write"<td width='20%'><strong>Participants</strong></td>"
	Response.Write"<td width='20%'><strong>Dates</strong></td>"
  Response.Write"</tr>"

Do While Not rs.EOF

	Response.Write "<tr>"
	Response.Write "<td>"
	
	Response.Write "<a href='CAProgramDetail.asp?id=" & rs("ProgramID") &"'>"
	Response.Write rs("ProgramName")
	Response.Write "</a>"
	Response.Write "</td><td>"
	Response.Write rs("Description")
	Response.Write "</td><td>"
	Response.Write rs("Participants")
	Response.Write "</td><td>"
	Response.Write rs("Dates")
	Response.Write "</td><td>"
	
	rs.MoveNext
	
Loop 

	Response.Write "</tr></table>"
	Response.Write "</DIV>"

rs.close
conn.close
set conn = nothing

	
end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page10</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text3 = null;
var Text1 = null;
var Text2 = null;
//-->
</SCRIPT>

<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text3 { position:absolute;left:187;top:192;width:622;z-index:3 }
#Text1 { position:absolute;left:204;top:250;width:564;z-index:1 }
#Text2 { position:absolute;left:206;top:280;width:250;z-index:2 }
#EndOfPage { position:absolute;left:0;top:308;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text3" Class="lumpkin1">If you are interested in knowing more about a program, or would like to register for that program, click on a hyperlink below.<BR>
</DIV>
<DIV ID="Text1" Class="lumpkin1"><%=(message)%>
</DIV>
<DIV ID="Text2" Class="lumpkin1"><%=(Response.Write("<p>"))%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/CulturalArtsInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
