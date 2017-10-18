<%@ Language=VBScript%>
<%
dim SQL

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEvents WHERE EventID="&session("programnum")

rs.open SQL,conn, 3,3

If request("eventname") <>"" then
    rs("Title")=request("eventname")
end if

If request("description") <>"" then
    rs("Desc")=request("description") 
end if

If request("eventdate") <>"" then
   rs("Date")=request("eventdate")
end if

If request("starttime") <>"" then
    rs("Start")=request("starttime")
end if

If request("endtime") <>"" then
    rs("End")=request("endtime")
end if

If request("location") <>"" then
    rs("Location")=request("location")
end if

If request("participants") <>"" then
    rs("Group1")=request("participants")
end if

If request("participants2") <>"" then
    rs("Group2")=request("participants2")
end if

If request("participants3") <>"" then
    rs("Group3")=request("participants3")
end if

If request("notes") <>"" then
    rs("Notes")=request("notes")
end if

rs.update

rs.close
conn.close
set conn=nothing

message= "The edits for the scheduled event "&request("eventname")&" have been updated to the database. <br><br><a href='SearchEvents.asp'> Click here to return and search for, or edit, additional events. </a>"
%>

<!--#include file ="media/IIS_Gen_3.0_ConnectionPool.js"-->
<!--#include file ="media/IIS_Gen_3.0_Recordset.js"-->
<!--#include file ="media/MkOptnFmRecordset.js"-->
<!--#include file ="media/IIS_ServerHTMLEncode.js"-->
<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page48</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var ImageButton4 = null;
var ImageButton1 = null;
var ImageButton6 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:175;top:132;width:571;z-index:1 }
#DBStyleImageButton4 { position:absolute;left:16;top:179;width:146;z-index:2 }
#DBStyleImageButton1 { position:absolute;left:16;top:233;width:146;z-index:3 }
#DBStyleImageButton6 { position:absolute;left:16;top:287;width:146;z-index:4 }
#EndOfPage { position:absolute;left:0;top:333;width:0 }
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
<DIV ID="DBStyleImageButton1">
<A HREF="SearchEvents.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton1" Name="ImageButton1" Class="Normal" SRC="media/searchevents.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
<DIV ID="DBStyleImageButton6">
<A HREF="NewEvent.asp"  onMouseOver="_B__onMouseOver( 2);"  onMouseOut="_B__onMouseOut( 2);">
<IMG ID="ImageButton6" Name="ImageButton6" Class="Normal" SRC="media/newevent.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/EditEventInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
