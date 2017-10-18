<%@ Language=VBScript%>
<%
dim SQL
x=0

Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open "lumpkin"

set rs = Server.CreateObject("ADODB.Recordset")

SQL = "Select * from tblPersonnel where Password='"&Request("pword")&"'"

rs.open SQL, conn, 1, 2
    
Session("StaffName")=rs("FirstName")

    If rs("UserName")=Request("staffname") Then  
            x=(x+1)
    End if
 
	If rs("Password")=Request("pword") Then
			x=(x+1)
	End if

   If x=2 Then  
		Response.Redirect "StaffHome.asp"
	else
		failed="Sorry, the name or password you entered is not correct.  <br><br><a href='StaffEntry.asp'> Click here to try again. </a>" 
   End if

    

rs.close
conn.close
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" CONTENT="Elemental Software Drumbeat 2000 3.01.369">

<TITLE>Page4</TITLE>
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
#Text1 { position:absolute;left:152;top:41;width:580;z-index:1 }
#EndOfPage { position:absolute;left:0;top:69;width:0 }
</STYLE>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0   BACKGROUND="media/greenback.jpg" BGColor="#ffffff">

<A NAME="StartOfPage"></A>


<DIV ID=Section0 CLASS=section STYLE="position:relative; top:0; height:0">
<DIV ID="Text1" Class="DBCUSTOMSTYLE1"><%=(failed)%>
</DIV>
</DIV><SCRIPT>__calcSectionHeight(Section0, 0)</SCRIPT>
<SCRIPT SRC="media/EntryCheckInline.js"></SCRIPT>

<DIV ID="EndOfPage"><A NAME="EndOfPage"></A></DIV>

</BODY></HTML>
