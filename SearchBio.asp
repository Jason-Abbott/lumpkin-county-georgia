<%@ Language=VBScript%>
<%
'set prior variable =""
session("p_lastname")=""
session("p_firstname")=""
session("p_address")=""
session("p_city")=""
session("p_state")=""
session("p_zip")=""
session("p_phone")=""	
session("g_lastname")=""
session("g_firstname")=""
session("g_address")=""
session("g_city")=""
session("g_state")=""
session("g_zip")=""
session("g_hphone")=""
session("g_wphone")=""

session("dateofbirth")=Request("mm")&"/"&Request("dd")&"/"&Request("yy")

set conn = Server.CreateObject("ADODB.Connection")

conn.open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblParticipants WHERE P_LastName='"&request("lastname")&"' AND P_FirstName='"&request("firstname")&"' AND DOB=#"&CDate(session("dateofbirth"))&"#"

rs.open SQL,conn, 1,2

If rs.eof then 

sorry = "We're sorry, but no records are found for "&request("lastname")&" , "&request("firstname")  
link = "<br><br><a href='Register.asp'>  Click here to return and search again.</a>"

else

	If not rs.eof then

	p="Participant:"
	session("p_lastname")=rs("P_LastName")
	session("p_firstname")=rs("P_FirstName")
	session("p_address")=rs("P_Address")
	session("dob")=rs("DOB")
	session("p_city")=rs("P_City")
	session("p_state")=rs("P_State")
	session("p_zip")=rs("P_Zip")
	session("p_phone")=rs("P_Phone")	
	
		If rs("G_LastName")<>"" then
		g="Guardian:"
		session("g_lastname")=rs("G_LastName")
		session("g_firstname")=rs("G_FirstName")
		session("g_address")=rs("G_Address")
		session("g_city")=rs("G_City")
		session("g_state")=rs("G_State")
		session("g_zip")=rs("G_Zip")
		session("g_hphone")=rs("G_HomePhone")
		session("g_wphone")=rs("G_WorkPhone")
		end if

	end if
end if


rs.close
conn.close
set conn=nothing
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page28</TITLE>
<SCRIPT LANGUAGE="JavaScript">
<!--
var Form1 = null;
var Text1 = null;
var Text20 = null;
var Text2 = null;
var Text3 = null;
var Text4 = null;
var Text11 = null;
var Text5 = null;
var Text12 = null;
var Text18 = null;
var Text13 = null;
var Text6 = null;
var Text14 = null;
var Text7 = null;
var Text15 = null;
var Text8 = null;
var Text16 = null;
var Text9 = null;
var Text17 = null;
var Text10 = null;
var Text19 = null;
var FormButton2 = null;
var FormButton1 = null;
//-->
</SCRIPT>
<SCRIPT SRC="media/DBCommon.js"></SCRIPT>
<SCRIPT SRC="media/DBFrmBtn.js"></SCRIPT>
<SCRIPT SRC="media/DBForm.js"></SCRIPT>
<SCRIPT SRC="media/SearchBio.js"></SCRIPT>

<SCRIPT SRC="media/DBRelPos.js"></SCRIPT>

<!-- CSS styles and positioning for IE 4 browsers. -->
<LINK REL="StyleSheet" TYPE="text/css" HREF="Lumpkin_IE4.css">
<STYLE ID="IE4_ABSOLUTE_POSITIONING_STYLE_SHEET">
#Text1 { position:absolute;left:168;top:32;width:583;z-index:1 }
#Text20 { position:absolute;left:166;top:73;width:582;z-index:22 }
#Text2 { position:absolute;left:165;top:80;width:192;z-index:4 }
#Text3 { position:absolute;left:503;top:80;width:222;z-index:5 }
#Text4 { position:absolute;left:165;top:115;width:250;z-index:6 }
#Text11 { position:absolute;left:503;top:115;width:250;z-index:13 }
#Text5 { position:absolute;left:165;top:141;width:250;z-index:7 }
#Text12 { position:absolute;left:503;top:141;width:250;z-index:14 }
#Text18 { position:absolute;left:166;top:166;width:242;z-index:20 }
#Text13 { position:absolute;left:503;top:168;width:250;z-index:15 }
#Text6 { position:absolute;left:165;top:192;width:250;z-index:8 }
#Text14 { position:absolute;left:503;top:209;width:250;z-index:16 }
#Text7 { position:absolute;left:165;top:233;width:250;z-index:9 }
#Text15 { position:absolute;left:503;top:233;width:250;z-index:17 }
#Text8 { position:absolute;left:165;top:257;width:250;z-index:10 }
#Text16 { position:absolute;left:503;top:257;width:250;z-index:18 }
#Text9 { position:absolute;left:165;top:281;width:250;z-index:11 }
#Text17 { position:absolute;left:503;top:281;width:250;z-index:19 }
#Text10 { position:absolute;left:165;top:305;width:250;z-index:12 }
#Text19 { position:absolute;left:503;top:306;width:250;z-index:21 }
#DBStyleFormButton2 { position:absolute;left:599;top:353;width:88;z-index:3 }
#DBStyleFormButton1 { position:absolute;left:597;top:395;width:124;z-index:2 }
#EndOfPage { position:absolute;left:0;top:432;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>



<FORM  NAME="Form1" METHOD="Post" ACTION="SearchBio.asp">
<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="lumpkin1"><%=(sorry)%>
</DIV>
<DIV ID="Text20" Class="lumpkin1"><%=(link)%>
</DIV>
<DIV ID="Text2" Class="lumpkin1"><%=(p)%>
</DIV>
<DIV ID="Text3" Class="lumpkin1"><%=(g)%>
</DIV>
<DIV ID="Text4" Class="lumpkin1"><%=(session("p_lastname"))%>
</DIV>
<DIV ID="Text11" Class="lumpkin1"><%=(session("g_lastname"))%>
</DIV>
<DIV ID="Text5" Class="lumpkin1"><%=(session("p_firstname"))%>
</DIV>
<DIV ID="Text12" Class="lumpkin1"><%=(session("g_firstname"))%>
</DIV>
<DIV ID="Text18" Class="lumpkin1"><%=(session("dob"))%>
</DIV>
<DIV ID="Text13" Class="lumpkin1"><%=(session("g_address"))%>
</DIV>
<DIV ID="Text6" Class="lumpkin1"><%=(session("p_address"))%>
</DIV>
<DIV ID="Text14" Class="lumpkin1"><%=(session("g_city"))%>
</DIV>
<DIV ID="Text7" Class="lumpkin1"><%=(session("p_city"))%>
</DIV>
<DIV ID="Text15" Class="lumpkin1"><%=(session("g_state"))%>
</DIV>
<DIV ID="Text8" Class="lumpkin1"><%=(session("p_state"))%>
</DIV>
<DIV ID="Text16" Class="lumpkin1"><%=(session("g_zip"))%>
</DIV>
<DIV ID="Text9" Class="lumpkin1"><%=(session("p_zip"))%>
</DIV>
<DIV ID="Text17" Class="lumpkin1"><%=(session("g_hphone"))%>
</DIV>
<DIV ID="Text10" Class="lumpkin1"><%=(session("p_phone"))%>
</DIV>
<DIV ID="Text19" Class="lumpkin1"><%=(session("g_wphone"))%>
</DIV>
<INPUT TYPE=Submit  ID="DBStyleFormButton2" NAME="FormButton2" value="Register" onClick="return _FormButton2__onClick()"  Class="lumpkin1" tabIndex="0">
<INPUT TYPE=Submit  ID="DBStyleFormButton1" NAME="FormButton1" value="Edit Record" onClick="return _FormButton1__onClick()"  Class="lumpkin1" tabIndex="0">
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
</FORM>
<SCRIPT SRC="media/SearchBioInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
