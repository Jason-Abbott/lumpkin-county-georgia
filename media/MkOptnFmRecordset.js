<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function makeOptionFromRecordset(Recordset, DisplayColumn, ValueColumn, SelectedValue, EncodeDisplay, EncodeValue, EncodeSelected)
{
	var Option = "";
	if(Recordset.RS.state == 1/*adStateOpen*/)
	{
		/* save current position */
		var Bookmark;
		var Abs;
		var SupportsBookmark = false;
		if(Recordset.RS.Supports(0x00002000/*adBookmark*/))
		{
			Bookmark = Recordset.RS.Bookmark;
			SupportsBookmark = true;
		}
		else
		{
			Abs = Recordset.RS.AbsolutePosition;
		}

		/* if there is no value column, use display column for value */
		if(ValueColumn == null || ValueColumn == "")
		{
			ValueColumn = DisplayColumn;
		}

		/* initialize SelectedValue */
		if(SelectedValue == null)
		{
			SelectedValue = "";
		}
		else
		{
			SelectedValue = String(SelectedValue);
		}

		/* build the option string */
		Recordset.MoveFirst();
		while(!Recordset.IsEOF())
		{
			var Display = String(Recordset.GetColumnValue(DisplayColumn));
			var Value = String(Recordset.GetColumnValue(ValueColumn));
			var IsSelected = (Value != "" && Value == SelectedValue) ? true : false;
			Option += "<Option Value=\"";
			if(EncodeValue == true)
			{
				Value = ServerHTMLEncode(Value);
			}
			Option += Value + "\"";
			if(IsSelected == true)
			{
				Option += " SELECTED";
			}
			Option += ">";
			if(EncodeDisplay == true)
			{
				Display = ServerHTMLEncode(Display);
			}
			Option += Display + "\n";
			Recordset.MoveNext();
		}

		/* restore current position */
		if(SupportsBookmark)
		{
			Recordset.RS.Bookmark = Bookmark;
		}
		else
		{
			Recordset.RS.Move(Abs);
		}
	}
	return Option;
}
</SCRIPT>
