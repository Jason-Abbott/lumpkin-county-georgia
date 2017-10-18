<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function ServerHTMLEncode(EncodeMe)
{
	if(EncodeMe == null)
	{
		EncodeMe = "";
	}
	return Server.HTMLEncode(EncodeMe);
}
</SCRIPT>

