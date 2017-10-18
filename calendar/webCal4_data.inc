<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/22/1999

dim strDSN, strDelim

' Access uses # to delimit dates while SQL uses '
' so switch as necessary

strDelim = "#"

' strDSN= "DSN=****;UID=****;PWD=****"

strDSN = "DSN=lumpkin"

'strDSN = "DRIVER={Microsoft Access Driver (*.mdb)}; " _
'	& "DBQ=" & Server.Mappath("../lumpkin.mdb")

%>