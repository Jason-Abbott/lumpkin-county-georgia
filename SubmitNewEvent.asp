<%@ Language=VBScript%>
<%
response.buffer=true

eventdate=Request("mm")&"/"&Request("dd")&"/"&Request("yy")
starttime=Request("start")&Request("ampm1")
endtime=Request("end")&Request("ampm2")
'======================= Check for scheduling conflicts by location and time.===============
dim conflict, btb     
conflict=0
btb=0
fb2=""
fb3=""

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEvents WHERE Date = #"&CDate(eventdate)&"#"

rs.open SQL,conn, 1, 2

Do while not rs.EOF

	If rs("Date")=CDate(eventdate) then
	
		If rs("Location")=CINT(request("location")) then
			
			If rs("Start") >= CDate(starttime) and rs("Start") < CDate(endtime) then
				conflict=1
			end if

			If rs("Start") = CDate(endtime) then
			btb=1
			response.write "Back to back"
			End if

			If rs("End") <= CDate(endtime) and rs("End")>CDate(starttime) then
			conflict=1
			End if

			
			If rs("End")=CDate(starttime) then
			btb=2
			End if
		
        end if
   
	end if

	rs.movenext

loop

rs.close
conn.close
set conn = nothing

'==========This section defines the feedback that is displayed on screen.=============

'locconflict="No schedule conflicts"

If conflict = 1 then locconflict="This event location conflicts with a previously-scheduled event.  Please choose a new location, or a new time for this location."
if btb=2 then fb2="Warning:  This event begins at the same time that a previously-scheduled event ends."
if btb=1 then fb2="Warning:  This event ends at the same time that a previously-scheduled event begins."

If conflict=0 then fb3="This event has been successfully updated to the database."



'====================================Standard event submission==========================

If conflict=0 then

Set Conn=Server.CreateObject("ADODB.Connection")

Conn.Open"lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblEvents where 1=2"

rs.open SQL,conn, 3, 3

	rs.addnew
	rs("Title")=request("title")
	rs("Desc")=request("desc")
	rs("Date")=eventdate
		If request("recur")<>"" then
		rs("Recur")=request("recur")
		end if
	rs("Start")=starttime
	rs("End")=endtime
	rs("Location")=request("location")
	rs("Group1")=request("group1")
		If request("group2") <>"" then
		rs("Group2")=request("group2")
		end if
		If request("group3")<>"" then
		rs("Group3")=request("group3")
		end if
		If request("notes")<>"" then
		rs("Notes")=request("notes")
		end if
		If request("r")<>"" then
		rs("Type")="R"
		end if
		If request("ca")<>"" then
		rs("Type")="C"
		end if

	
rs.update

rs.close
conn.close
set conn = nothing

end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page25</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Text1 = null;
var Text2 = null;
var ImageButton4 = null;
var Text3 = null;
var ImageButton6 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBImgBtn.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:175;top:133;width:625;z-index:1 }
#Text2 { position:absolute;left:176;top:178;width:691;z-index:2 }
#DBStyleImageButton4 { position:absolute;left:16;top:179;width:146;z-index:4 }
#Text3 { position:absolute;left:176;top:219;width:671;z-index:3 }
#DBStyleImageButton6 { position:absolute;left:16;top:235;width:146;z-index:5 }
#EndOfPage { position:absolute;left:0;top:281;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkinsmall"><%=(locconflict)%>
</DIV>
<DIV ID="Text2" Class="lumpkinsmall"><%=(fb2)%>
</DIV>
<DIV ID="DBStyleImageButton4">
<A HREF="StaffHome.asp"  onMouseOver="_B__onMouseOver( 0);"  onMouseOut="_B__onMouseOut( 0);">
<IMG ID="ImageButton4" Name="ImageButton4" Class="Normal" SRC="media/mainhome.jpg" Width="146" Height="38" Border="0"></A>
</DIV>
<DIV ID="Text3" Class="lumpkinsmall"><%=(fb3)%>
</DIV>
<DIV ID="DBStyleImageButton6">
<A HREF="NewEvent.asp"  onMouseOver="_B__onMouseOver( 1);"  onMouseOut="_B__onMouseOut( 1);">
<IMG ID="ImageButton6" Name="ImageButton6" Class="Normal" SRC="media/neweventhi.jpg" Width="146" Height="36" Border="0"></A>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/SubmitNewEventInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
