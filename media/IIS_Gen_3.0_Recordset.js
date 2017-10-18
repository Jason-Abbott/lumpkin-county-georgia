<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function Recordset(Name, DSN, UID, PWD, SQL, UniqueKeyCols, Updateable, CursorType, CursorLocation, LockType, TableNames)
{
	RSReconnect(this);
	this.Name = Name;
	this.DSN = DSN;
	this.UID = UID;
	this.PWD = PWD;
	this.SQL = SQL;
	this.OrgSQL = SQL;	
	this.Updateable = Updateable;
	this.CursorType = CursorType;
	this.CursorLocation = CursorLocation;
	this.LockType = LockType;
	this.NoRecordMessage = "";
	this.NoFoundMessage = "";
	this.RSConnection = "";
	this.DTFormat = 0;
	this.FilterCriteria = "";
	this.OrderByCol = "";
	this.Found = true;
	this.SupportsApproxPosition = false;
	this.SupportsMove = false;
	this.SupportsCount = false;
	this.ApproxPosition = 1;
	this.Separator = "!*+";
	this.SpecialChars = "~@#$%^&*-+=\\}{\"';:?/><,()_\f\n\r\t ";
	this.TrimCRLF = true;
	this.bPreserveInputState = false;
	this.Action = "" + Request(this.Name+"_Action");

	// create bindings array
	if("" + Request(this.Name + "_Bindings") != "undefined")
	{
		this.BindingsArray = new Array();
		var Bindings = "" + Request(this.Name + "_Bindings");
		var nBinding = 0;
		for(;;)
		{
			var BindIndex = Bindings.indexOf(this.Separator);
			if(BindIndex >= 0)
			{
				var BoundColumn = Bindings.substring(0, BindIndex);
				Bindings = Bindings.substring(BindIndex + this.Separator.length, Bindings.length);
				BindIndex = Bindings.indexOf(this.Separator);
				if(BindIndex >= 0)
				{
					// get the location of the value
					var BindingType = Bindings.substring(0, BindIndex);
					Bindings = Bindings.substring(BindIndex + this.Separator.length, Bindings.length);

					// get the value
					BindIndex = Bindings.indexOf(this.Separator);
					if(BindIndex < 0)
					{
						BindIndex = Bindings.length;
					}
					var BoundValue = Bindings.substring(0, BindIndex);

					var Binding = new Object();
					Binding["Name"] = BoundColumn;
					Binding["Type"] = BindingType;
					Binding["Value"] = BoundValue;
					this.BindingsArray[nBinding] = Binding;
					nBinding++;

					// next binding
					if(BindIndex == Bindings.length)
					{
						BindIndex = -1;
					}
					else
					{
						Bindings = Bindings.substring(BindIndex + this.Separator.length, Bindings.length);
					}
				}
			}
			if(BindIndex < 0)
			{
				break;
			}
		}
	}

	// set or restore parameters
	this.State = "" + Request(this.Name + "_Position");
	if(this.State != "undefined")
	{
		var State = this.State;
		if(State != "")
		{
			var FilterIndex = State.indexOf("FIL:");
			var OrdIndex = State.indexOf("ORD:");
			var AbsIndex = State.indexOf("ABS:");
			var ParamIndex = State.indexOf("PAR:");
			if(FilterIndex >= 0 && OrdIndex >= 0)
			{
				this.FilterCriteria = State.substring(FilterIndex + 4, OrdIndex);
			}
			if(OrdIndex >= 0 && AbsIndex >= 0)
			{
				this.OrderByCol = State.substring(OrdIndex + 4, AbsIndex);
			}
			if(ParamIndex >= 0)
			{
				this.RestoreParams(State.substring(ParamIndex + 4, State.length));
			}
		}
	}
	if(this.ParamArray == null && Recordset.arguments.length > Recordset.length)
	{
		// set default parameters (variable arguments to this function)
		this.ParamArray = new Array();
		var nArgs = Recordset.arguments.length;
		var nArg = Recordset.length;
		var nParam = 0;
		for(; nArg < nArgs; nArg++, nParam++)
		{
			var ParmName = "" + Recordset.arguments[nArg];
			nArg++;
			if(nArg >= nArgs)
			{
				break;
			}
			var ParmValue = "" + Recordset.arguments[nArg];
			var Pair = new Object();
			Pair["Name"] = ParmName;
			Pair["Value"] = ParmValue;
			this.ParamArray[nParam] = Pair;
		}
	}

	// create unique key cols array
	if(UniqueKeyCols != "")
	{
		this.UniqueKeyColArray = new Array();
		var Temp = UniqueKeyCols;
		for(i = 0;;i++)
		{
			var Index = Temp.indexOf(",");
			if(Index == -1)
			{
				Index = Temp.length;
			}
			this.UniqueKeyColArray[i] = Temp.substring(0, Index);
			if(Index == Temp.length)
			{
				break;
			}
			Temp = Temp.substring(Index + 1, Temp.length);
		}
	}

	// create table names array
	if(TableNames != "")
	{
		this.TableNamesArray = new Array();
		var nTableName = 0;
		for(;;)
		{
			var TableNameIndex = TableNames.indexOf(this.Separator);
			if(TableNameIndex >= 0)
			{
				// get the column index
				var nCol = parseInt(TableNames.substring(0, TableNameIndex));
				TableNames = TableNames.substring(TableNameIndex + this.Separator.length, TableNames.length);
				// get the column name
				TableNameIndex = TableNames.indexOf(this.Separator);
				if(TableNameIndex < 0)
				{
					TableNameIndex = TableNames.length;
				}
				var TableName = TableNames.substring(0, TableNameIndex);
				var TableNameObj = new Object();
				TableNameObj["nCol"] = nCol;
				TableNameObj["TableName"] = TableName;
				this.TableNamesArray[nTableName] = TableNameObj;
				nTableName++;

				// next TableName
				if(TableNameIndex == TableNames.length)
				{
					TableNameIndex = -1;
				}
				else
				{
					TableNames = TableNames.substring(TableNameIndex + this.Separator.length, TableNames.length);
				}
			}
			if(TableNameIndex < 0)
			{
				break;
			}
		}
	}
	return this;
}

function RSReconnect(Object)
{
	Object.FindConnection = RSFindConnection;
	Object.AddConnection = RSAddConnection;
	Object.Open = RSOpen;
	Object.GetColumnValue = RSGetColumnValue;
	Object.SetColumnValue = RSSetColumnValue;
	Object.ProcessAction = RSProcessAction;
	Object.MoveFirst = RSMoveFirst;
	Object.MoveNext = RSMoveNext;
	Object.MovePrevious = RSMovePrevious;
	Object.MoveLast = RSMoveLast;
	Object.Insert = RSInsert;
	Object.Update = RSUpdate;
	Object.Delete = RSDelete;
	Object.IsBOF = RSIsBOF;
	Object.IsEOF = RSIsEOF;
	Object.BeginTrans = RSBeginTrans;
	Object.CommitTrans = RSCommitTrans;
	Object.RollbackTrans = RSRollbackTrans;
	Object.Close = RSClose;
	Object.Filter = RSFilter;
	Object.Find = RSFind;
	Object.GetState = RSGetState;
	Object.RestoreState = RSRestoreState;
	Object.GetKeyVals = RSGetKeyVals;
	Object.GetAbsPos = RSGetAbsPos;
	Object.OrderBy = RSOrderBy;
	Object.SetMessages = RSSetMessages;
	Object.SetAction = RSSetAction;
	Object.SetOrderByCol = RSSetOrderByCol;
	Object.GetOrderbyCol = RSGetOrderByCol;
	Object.Move = RSMove;
	Object.SetParameter = RSSetParam;
	Object.GetParameter = RSGetParam;
	Object.GetParams = RSGetParams;
	Object.RestoreParams = RSRestoreParams;
	Object.SubstituteParams = RSSubstituteParams;
	Object.SetMaxRecords = RSSetMaxRecords;
	Object.GetMaxRecords = RSGetMaxRecords;
	Object.FindBinding = RSFindBinding;
	Object.AddBinding = RSAddBinding;
	Object.RemoveBinding = RSRemoveBinding;
	Object.GetBindingValue = RSGetBindingValue;
	Object.SetBindingValue = RSSetBindingValue;
	Object.GetBindingType = RSGetBindingType;
	Object.SetBindingType = RSSetBindingType;
	Object.GetBoundColumnName = RSGetBoundColumnName;
	Object.GetBoundColumnNameByIndex = RSGetBoundColumnNameByIndex;
	Object.SetRS = RSSetRS;
	Object.GetField = RSGetField;
}

function RSSetRS(RS)
{
	if(RS)
	{
		this.RS = RS;
	}
}

function RSGetField(ColName)
{
	var Field = null;
	if(this.TableNamesArray == null)
	{
		Field = this.RS.Fields.Item(ColName);
	}
	else
	{
		// column not found in Fields collection, so look for table.column
		var nDot = ColName.indexOf(".");
		if(nDot > 0)
		{
			var RSTable = ColName.substring(0, nDot);
			var RSColumn = ColName.substring(nDot + 1, ColName.length);
			for(var i = 0, n = this.TableNamesArray.length; i < n; i++)
			{
				var TObj = this.TableNamesArray[i];
				var TableName = TObj["TableName"];
				if(TableName == RSTable)
				{
					var nCol = TObj["nCol"];
					if(this.RS.Fields.Item(nCol).Name == RSColumn)
					{
						Field = this.RS.Fields.Item(nCol);
						break;
					}
				}
			}
		}
		else
		{
			Field = this.RS.Fields.Item(ColName);
		}
	}
	return Field;
}

function RSFindConnection(DSN, UID, PWD)
{
	var Connections = ConnectionCache.Connections;
	for(var i = 0, n = Connections.length; i < n; i++)
	{
		TestPair = Connections[i];
		TestKey = TestPair["Key"];
    	if(TestKey["DSN"] != DSN || TestKey["UID"] != UID || TestKey["PWD"] != PWD)
		{
			continue;
		}
		return TestPair["Value"];
	}
	return null;
}

function RSAddConnection(DSN, UID, PWD, Connection)
{
	var Pair = new Object();
	var Key = new Object();
	Key["DSN"] = DSN;
	Key["UID"] = UID;
	Key["PWD"] = PWD;
	Pair["Key"] = Key;
	Pair["Value"] = Connection;
	var Connections = ConnectionCache.Connections;
	Connections[Connections.length] = Pair;
}

function RSOpen()
{
	this.RSConnection = this.FindConnection(this.DSN, this.UID, this.PWD);
	if(this.RSConnection == null)
	{
		this.RSConnection = Server.CreateObject("ADODB.Connection");
		this.AddConnection(this.DSN, this.UID, this.PWD, this.RSConnection);
	}
	if(this.RSConnection)
	{
		if(!(this.RSConnection.State & 1/*adStateOpen*/))
		{
			this.RSConnection.Open("dsn=" + this.DSN + ";uid=" + this.UID + ";pwd=" + this.PWD);
		}
		this.RS = Server.CreateObject("ADODB.Recordset");
		if(this.RS)
		{
			var version = parseFloat(this.RSConnection.Version);
			if (version > 1.0)
			{
				this.RS.CursorLocation = this.CursorLocation;
			}

			var SQL = this.SQL;
			if(this.OrderByCol != "")
			{
				SQL += " ORDER BY " + this.OrderByCol;
			}
			if(this.ParamArray != null && this.ParamArray.length > 0)
			{
				// substitute parameters
				SQL = this.SubstituteParams(SQL);
			}

			this.RS.Open(SQL, this.RSConnection, this.CursorType, this.LockType);
			if (this.RS.State == 1/*adStateOpen*/)
			{
				this.SupportsApproxPosition = this.RS.Supports(0x4000/*adApproxPosition*/);
				this.SupportsMove = this.RS.Supports(0x200/*adMovePrevious*/);
				this.SupportsCount = this.RS.RecordCount != -1;
				this.RestoreState();
			}
			else
			{
				Response.Write("<BR>Error: Unable to open Recordset!<BR>");
			}
		}
	}
}

function RSSetAction(action)
{
	this.Action = "" + action;
}

function RSSetOrderByCol(OrderByCol)
{
	this.OrderByCol = OrderByCol;
}

function RSGetOrderByCol()
{
	return this.OrderByCol;
}

function RSRestoreState()
{
	var Action = this.Action;
	if(Action != 'Insert' && Action != 'Close' && Action.indexOf('Filter(') < 0)
	{
		if(this.State != "undefined")
		{
			var Position = this.State;
			var FilterIndex = Position.indexOf("FIL:");
			var OrdIndex = Position.indexOf("ORD:");
			var AbsIndex = Position.indexOf("ABS:");
			var KeyIndex = Position.indexOf("KEY:");
			var ParamIndex = Position.indexOf("PAR:");
			if(ParamIndex < 0)
			{
				ParamIndex = Position.length;
			}

			// reapply filter
			if(this.FilterCriteria != "") 
			{
				this.Filter(this.FilterCriteria);
			}

			//if the filter produces an empty recordset , return
			if (this.IsEOF())
			{
				return -1;
			}

			// move to previous position
			if(KeyIndex >= 0 && AbsIndex >= 0)
			{
				var Abs = Position.substring(AbsIndex + 4, KeyIndex);
				var Key	= "";
				if((KeyIndex + 4) < ParamIndex)
				{
					Key = Position.substring(KeyIndex + 4, ParamIndex);
				}

				// build array of unique keys
				var KeyArray = null;
				if(this.UniqueKeyColArray != null && Key != "")
				{
					KeyArray = new Array();
					for(i = 0;; i++)
					{
						var Index = Key.indexOf(this.Separator);
						if(Index == -1)
						{
							Index = Key.length;
						}
						KeyArray[i] = Key.substring(0, Index);
						if(Index == Key.length)
						{
							break;
						}
						Key = Key.substring(Index + 3, Key.length);
					}
				}

				if(this.SupportsApproxPosition == true)
				{
					if(Abs != "")
					{
						if(Abs == "BOF")
						{
							this.RS.MovePrevious();
						}
						else if(Abs == "EOF")
						{
							this.RS.MoveLast();
							this.RS.MoveNext();
						}
						else
						{
							if(Abs != "MTY") // was empty
							{
								this.RS.Move(parseInt(Abs) - 1);
							}
						}
					}
					if(Key != "")
					{
						// fan out from current position looking for key
						if(this.UniqueKeyColArray.length == KeyArray.length)
						{
							var GotBOF = false;
							var GotEOF = false;
							var Direction = -1;
							var Count = 1;
							var n = this.UniqueKeyColArray.length;
							for(;;)
							{
								// test for match to key
								this.Found = true;
								for(i = 0; i < n; i++)
								{
									if(String(this.GetField(this.UniqueKeyColArray[i]).Value) != String(KeyArray[i]))
									{
										this.Found = false;
										break;
									}
								}
								if(this.Found == false)
								{
									// key mismatch
									if(Direction < 0)
									{
										if(GotBOF == false)
										{
											this.RS.Move(Count * Direction);
											if(GotEOF == false)
											{
												Direction *= -1;
											}
											if(this.RS.BOF)
											{
												GotBOF = true;
												if(!GotEOF)
												{
													this.RS.MoveNext();
													this.RS.Move(Count);
													Count = 1;
													if(this.RS.EOF)
													{
														GotEOF = true;
													}
												}
											}
											else if(GotEOF == false)
											{
												Count++;
											}
										}
									}
									else
									{
										if(GotEOF == false)
										{
											this.RS.Move(Count * Direction);
											if(GotBOF == false)
											{
												Direction *= -1;
											}
											if(this.RS.EOF)
											{
												GotEOF = true;
												if(!GotBOF)
												{
													this.RS.MovePrevious();
													this.RS.Move(-Count);
													Count = 1;
													if(this.RS.BOF)
													{
														GotBOF = true;
													}
												}
											}
											else if(GotBOF == false)
											{
												Count++;
											}
										}
									}
								}
								if(GotBOF == true && GotEOF == true)
								{
									// key not found so move to abs position
									if(this.RS.BOF)
									{
										this.RS.MoveNext();
									}
									else
									{
										this.RS.MoveFirst();
									}
									if(Abs != "")
									{
										if(Abs == "BOF")
										{
											this.RS.MovePrevious();
										}
										else if(Abs == "EOF")
										{
											this.RS.MoveLast();
											this.RS.MoveNext();
										}
										else
										{
											var Increment = 1;
											var Action = this.Action;
											if(Action == 'Previous')
											{
												if(Direction ==	1)
												{
													Increment--;
												}
											}
											else if(Action == 'Next')
											{
												if(Direction == -1)
												{
													Increment++;
												}
											}

											this.RS.Move(parseInt(Abs) - Increment);
										}
									}
									break;
								}
								if(this.Found == true)
								{
									break;
								}
							}
						}
					}
				}
				else // SupportsApproxPosition == false
				{
					if(KeyArray != null)
					{
						// move forward through recordset until key is found
						while(!this.IsEOF())
						{
							var n = this.UniqueKeyColArray.length;
							this.Found = true;
							for(i = 0; i < n; i++)
							{
								if(String(this.GetField(this.UniqueKeyColArray[i]).Value) != String(KeyArray[i]))
								{
									this.Found = false;
									break;
								}
							}
							if(this.Found)
							{
								break;
							}
							else
							{
								this.MoveNext();
							}
						}

					}
					else if(Abs != "" && Abs != "MTY")
					{
						// move forward through recordset until approx position is found
						if(this.SupportsMove)
						{
							this.RS.Move(parseInt(Abs) - 1);
							this.ApproxPosition = Abs;
						}
						else
						{
							Abs = parseInt(Abs);
							this.ApproxPosition = 1;
							while(this.ApproxPosition != Abs)
							{
								this.RS.MoveNext();
								this.ApproxPosition++;
							}
						}
					}
				}
			}
		}
	}
}

function RSGetState()
{
	if (this.bPreserveInputState)
	{
		return this.State;
	}
	
	return "FIL:" + this.FilterCriteria + "ORD:" + this.OrderByCol + "ABS:" + this.GetAbsPos() + "KEY:" + this.GetKeyVals() + "PAR:" + this.GetParams();
}


function RSGetMaxRecords()
{
	return this.RS.MaxRecords;
}

function RSSetMaxRecords(lMaxRecords)
{
	this.RS.MaxRecords = lMaxRecords;
}

function RSGetKeyVals()
{
	var KeyVals = "";
	if(this.UniqueKeyColArray != null)
	{
		var n = this.UniqueKeyColArray.length;
		if(n > 0)
		{
			var First = true;
			for(i = 0; i < n; i++)
			{
				if(First == true)
				{
					First = false;
				}
				else
				{
					KeyVals += this.Separator;
				}
				if(!this.IsBOF() && !this.IsEOF())
				{
					var Field = this.GetField(this.UniqueKeyColArray[i]);
					if(Field != null)
					{
						KeyVals += String(Field.Value);
					}
				}
			}
		}
	}
	return KeyVals;
}

function RSGetAbsPos()
{
	var AbsPos = "";
	if(this.IsBOF() && this.IsEOF())
	{
		AbsPos = "MTY"; // empty result set
	}
	else
	{
		if(this.SupportsApproxPosition)
		{
			var nAbsPos = this.RS.AbsolutePosition;
			if(nAbsPos == -2/*adPosBOF*/)
			{
				AbsPos += "BOF";
			}
			else if(nAbsPos == -3)
			{
				AbsPos += "EOF";
			}
			else
			{
				AbsPos += nAbsPos;
			}
		}
		else
		{
			AbsPos += this.ApproxPosition;
		}
	}
	return AbsPos;
}

function RSGetParams()
{
	var Params = "";
	if(this.ParamArray != null)
	{
		var n = this.ParamArray.length;
		if(n > 0)
		{
			var First = true;
			for(i = 0; i < n; i++)
			{
				if(First == true)
				{
					First = false;
				}
				else
				{
					Params += this.Separator;
				}
				Pair = this.ParamArray[i];
				Params += Pair["Name"];
				Params += "=";
				Params += Pair["Value"];
			}
		}
	}
	return Params;
}

function RSGetColumnValue(ColName)
{
	if(this.IsBOF() || this.IsEOF())
	{
		var Action = "" + this.Action;
		if ((Action == "Previous") || (Action == "Next"))
		{
			return this.NoRecordMessage;
		}
		else
		{
			return this.NoFoundMessage
		}
	}
	var Field = this.GetField(ColName);
	if(Field != null)
	{
		if(this.DTFormat >= 0)
		{
			var ColType = Field.Type;
			switch(ColType)
			{
				case 7:		/*adDate*/
				case 133:	/*adDBDate*/
				case 134:	/*adDBTime*/
				case 135:	/*adDBTimeStamp*/
					return RSFormatDateTime(Field.Value, this.DTFormat);
				default:
					break;
			}
		}
		return Field.Value;
	}
	return "";
}

function RSSetColumnValue(ColName, Val)
{
	if(!this.IsBOF() && !this.IsEOF())
	{
		var Field = this.GetField(ColName);
		if(Field != null)
		{
			Field.Value = Val;
		}
	}
}

function RSProcessAction()
{
	var Action = this.Action;
	if(Action.indexOf('Find(') == 0)
	{
		eval(this.Name + '.' + Action);
	}
	else if(Action.indexOf('Filter(') == 0)
	{
		eval(this.Name + '.' + Action);
	}
	else
	{
		Action = "" + this.Action;
		if(Action == 'First')
		{
			this.MoveFirst();
		}
		else if (Action == 'Previous')
		{
			this.MovePrevious();
		}
		else if (Action == 'Next')
		{
			this.MoveNext();
		}
		else if (Action == 'Last')
		{
			this.MoveLast();
		}
		else if (Action == 'Insert')
		{
			this.Insert();
		}
		else if(Action == 'Update')
		{
			this.Update();
		}
		else if(Action == 'Delete')
		{
			this.Delete();
		}
		else if(Action == 'Close')
		{
			this.Close();
		}
		else if(Action.indexOf('Sort') == 0)
		{
			this.OrderBy();	
		}
		else if(Action.indexOf('OrderBy(') == 0)
		{
			eval(this.Name + '.' + Action);
		}
	}
}

function RSSetMessages(NoRecordMessage, NoFoundMessage)
{
	this.NoRecordMessage = NoRecordMessage;
	this.NoFoundMessage = NoFoundMessage;
}

function RSMoveFirst()
{
	if(this.RS.BOF)
	{
		if(!(this.RS.EOF))
		{
			this.RS.MoveNext();
		}
	}
	else
	{
		if(this.SupportsMove)
		{
			this.RS.MoveFirst();
		}
		else
		{
			this.RS.Requery();
		}
	}
	this.ApproxPosition = 1;
}

function RSMoveNext()
{
	if(!(this.RS.EOF))
	{
		this.RS.MoveNext();
		this.ApproxPosition++;
	}
}

function RSMovePrevious()
{
	if(!(this.RS.BOF))
	{
		if(this.SupportsMove)
		{
			this.RS.MovePrevious();
			this.ApproxPosition--;
		}
		else
		{
			var ApproxPosition = this.ApproxPosition - 1;
			this.MoveFirst();
			while(this.ApproxPosition < ApproxPosition)
			{
				this.RS.MoveNext();
				this.ApproxPosition++;
			}
		}
	}
}

function RSMoveLast()
{
	if(this.RS.EOF)
	{
		this.MovePrevious();
	}
	else
	{
		if(this.SupportsMove && this.SupportsCount)
		{
    		this.RS.MoveLast();
			this.ApproxPosition = this.RS.RecordCount;
		}
		else
		{
			if(this.SupportsCount)
			{
				var RecordCount = this.RS.RecordCount;
				if(RecordCount > 0)
				{
					while(this.ApproxPosition < RecordCount)
					{
						this.RS.MoveNext();
						this.ApproxPosition++;
					}
				}
			}
			else
			{
				var LastRecord = 1;
				this.MoveFirst();
				while(!this.RS.EOF)
				{
					this.RS.MoveNext();
					LastRecord++;
				}
				LastRecord--;
				this.MoveFirst();
				while(this.ApproxPosition < LastRecord)
				{
					this.MoveNext();
				}
			}
		}
	}
}


function RSMove(offset)
{
	if(this.SupportsApproxPosition)
	{
		this.RS.Move(offset);
		this.ApproxPosition += offset;
	}
	else
	{
		if(offset > 0)
		{
			while(offset > 0)
			{
				this.RS.MoveNext();
				this.ApproxPosition++;
				offset--;
				if(this.RS.EOF)
				{
					break;
				}
			}
		}
		else
		{
			var ApproxPosition = this.ApproxPosition - offset;
			if(ApproxPosition < 1)
			{
				ApproxPosition = 1;
			}
			this.MoveFirst();
			while(this.ApproxPosition < ApproxPosition)
			{
				this.RS.MoveNext();
				this.ApproxPosition++;
			}
		}
	}
}

function RSInsert()
{
	if(this.Updateable != false)
	{
		this.RS.AddNew();
		this.Update();
	}
}

function RSUpdate()
{
	if(this.Updateable == true && this.Found == true)
	{
		if(this.BindingsArray != null)
		{
			var nBindings = this.BindingsArray.length;
			for(var nBinding = 0; nBinding < nBindings; nBinding++)
			{
				var Binding = this.BindingsArray[nBinding];
				var BoundColumn = Binding["Name"];
				var BindingType = Binding["Type"];
				var FormElement = null;
				var BoundValue = Binding["Value"];
				var DoUpdate = true;

				if(BindingType == "FRM") // form element name
				{
					FormElement = BoundValue;
					if(String(Request(FormElement)) != "undefined")
					{
						BoundValue = Request(FormElement);
					}
					else
					{
						BoundValue = "";	
					}
				}
				else if(BindingType == "VAR") // javascript variable
				{
					BoundValue = eval(BoundValue);
				}
				else if(BindingType != "VAL") // literal value
				{
					DoUpdate = false;
				}

				// check whether new value is different from old value
				var Field = this.GetField(BoundColumn);
				if(Field != null && DoUpdate == true && String(Field.Value) != String(BoundValue))
				{
					// put the value into the column
					if((String(BoundValue) == "") && (Field.Attributes & 0x00000020/*adFldIsNullable*/))
					{
						Field.Value = null;
					}
					else if(FormElement != null && String(Request(FormElement)) == "undefined")
					{
						// handle the unchecked state for checkbox control
						if(Field.Type == 11/*adBoolean*/)
						{
							Field.Value = "0";
						}
					}
					else
					{
						if(this.TrimCRLF == true)
						{
							var ColType = Field.Type;
							if(ColType == 129/*adChar*/ ||
								ColType == 200/*adVarChar*/ ||
								ColType == 201/*adLongVarChar*/ ||
								ColType == 8/*adBSTR*/)
							{
								BoundValue = RSTrimCRLF(BoundValue);
							}
						}
						Field.Value = String(BoundValue);
					}
				}
			}
		}
		this.RS.Update();
	}
}

function RSDelete()
{
	if(this.Updateable == true && this.Found == true && !this.IsEOF() && !this.IsBOF())
	{
		this.RS.Delete();
		this.MoveNext();
		if(this.IsEOF())
		{
			this.MovePrevious();
		}
	}
}

function RSIsBOF()
{
	return this.RS.BOF;
}

function RSIsEOF()
{
	return this.RS.EOF;
}

function RSBeginTrans()
{
	this.RSConnection.BeginTrans();
}

function RSCommitTrans()
{
	this.RSConnection.CommitTrans();
}

function RSRollbackTrans()
{
	this.RSConnection.RollbackTrans();
}

function RSClose()
{
	this.RS.Close();
}

function RSFilter(Where)
{
	if(Where == null)
	{
		Where = "";
	}

	this.FilterCriteria = Where;
	this.RS.Filter = Where;		
}

function RSFind(ColumnName, Key)
{

	var nArgs = RSFind.arguments.length;
	while(!this.IsEOF())
	{
		var nArg = 0;
		var KeyValueFound = true;
		for(; nArg < nArgs; nArg++)
		{
			ColumnName = "" + RSFind.arguments[nArg];
			nArg++;
			var Key = "" + RSFind.arguments[nArg];
			if(String(this.GetField(ColumnName).Value) != String(Key))
			{
				KeyValueFound = false;
				break;
			}
		}
		if(KeyValueFound)
		{
			break;
		}
		else
		{
			this.MoveNext();
		}
	}
}

function RSOrderBy(OrderByCol)
{
	if(OrderByCol != "")
	{
		this.SetOrderByCol(OrderByCol);
		this.Close();
		this.Open();
	}
}

function RSSetParam(Name, Value)
{
	var Pair = null;
	if(this.ParamArray == null)
	{
		this.ParamArray = new Array();
	}
	else
	{
		// find name
		for(i = 0, n = this.ParamArray.length; i < n; i++)
		{
			Target = this.ParamArray[i];
			if(Target["Name"] == Name)
			{
				Pair = Target;
				break;
			}
		}
	}
	if(Pair == null)
	{
		// create new
		Pair = new Object();
    	Pair["Name"] = Name;
		this.ParamArray[this.ParamArray.length] = Pair;
	}
	Pair["Value"] = Value;
}

function RSRestoreParams(Params)
{
	if(Params != null && Params != "")
	{
		this.ParamArray = new Array();
		for(i = 0;; i++)
		{
			var Index = Params.indexOf(this.Separator);
			if(Index == -1)
			{
				Index = Params.length;
			}
			var NameValue = Params.substring(0, Index);
			var Pair = new Object();
			var Eq = NameValue.indexOf("=");
			Pair["Name"] = NameValue.substring(0, Eq);
			Pair["Value"] = NameValue.substring(Eq + 1, NameValue.length);
			this.ParamArray[i] = Pair;
			if(Index == Params.length)
			{
				break;
			}
			Params = Params.substring(Index + 3, Params.length);
		}
	}
}

function RSGetParam(Name)
{
	if(this.ParamArray == null)
	{
		return null;
	}

	// find name
	for(i = 0, n = this.ParamArray.length; i < n; i++)
	{
		Binding = this.ParamArray[i];
		if(Target["Name"] == Name)
		{
			return Target["Value"];
		}
	}
}

function RSSubstituteParams(SQL)
{
	for(i = 0, n = this.ParamArray.length; i < n; i++)
	{
		var Param = this.ParamArray[i];
		var ParamName = Param["Name"];
		var nIndex = -1;
		do
		{
			nIndex = SQL.indexOf(ParamName, nIndex + 1);
			if(nIndex > 0)
			{
				// found an occurrance of the param name
				// now see whether it is a token or not by
				// testing the character before and after
				var C1 = SQL.charAt(nIndex - 1);
				var C2 = '~';
				if((nIndex + ParamName.length) < SQL.length)
				{
					C2 = SQL.charAt(nIndex + ParamName.length);
				}
				if((this.SpecialChars.indexOf(C1) != -1) && (this.SpecialChars.indexOf(C2) != -1))
				{
					var NewSQL = SQL.slice(0, nIndex);
					NewSQL += Param["Value"];
					NewSQL += SQL.slice(nIndex + ParamName.length, SQL.length);
    				SQL = NewSQL;
					break;
				}
			}
			nIndex++;
		} while(nIndex > 0 && nIndex < SQL.length);
	}
	return SQL;
}

function RSFindBinding(ColumnName)
{
	if(this.BindingsArray != null)
	{
		for(i = 0, n = this.BindingsArray.length; i < n; i++)
		{
			BindingObj = this.BindingsArray[i];
			if(BindingObj["Name"] == ColumnName)
			{
				return BindingObj;
			}
		}
	}
	return null;
}

function RSAddBinding(ColumnName, BoundValue, BindingType)
{
	if(this.BindingsArray == null)
	{
		this.BindingsArray = new Array();
	}
	BindingObj = this.FindBinding(ColumnName);
	if(BindingObj == null)
	{
		BindingObj = new Object();
		BindingObj["Name"] = ColumnName;
		this.BindingsArray[this.BindingsArray.length] = BindingObj;
	}
	BindingObj["Type"] = BindingType;
	BindingObj["Value"] = BoundValue;
}

function RSRemoveBinding(ColumnName)
{
	if(this.BindingsArray != null && this.BindingsArray.length)
	{
		BindingsArray = new Array();
		for(i = 0, n = this.BindingsArray.length; i < n; i++)
		{
			BindingObj = this.BindingsArray[i];
			if(BindingObj["Name"] != ColumnName)
			{
				BindingsArray[BindingsArray.length] = BindingObj;
			}
		}
		this.BindingsArray = null;
		this.BindingsArray = new Array();
		if(BindingsArray.length)
		{
			for(i = 0, n = BindingsArray.length; i < n; i++)
			{
				this.BindingsArray[i] = BindingsArray[i];
			}
		}
		BindingsArray = null;
	}
}

function RSGetBindingValue(ColumnName)
{
	BindingObj = this.FindBinding(ColumnName);
	if(BindingObj == null)
	{
		return "";
	}
	return BindingObj["Value"];
}

function RSSetBindingValue(ColumnName, BoundValue)
{
	BindingObj = this.FindBinding(ColumnName);
	if(BindingObj != null)
	{
		BindingObj["Value"] = BoundValue;
	}
}

function RSGetBindingType(ColumnName)
{
	BindingObj = this.FindBinding(ColumnName);
	if(BindingObj == null)
	{
		return "";
	}
	return BindingObj["Type"];
}

function RSSetBindingType(ColumnName, BindingType)
{
	BindingObj = this.FindBinding(ColumnName);
	if(BindingObj != null)
	{
		BindingObj["Type"] = BindingType;
	}
}

function RSGetBoundColumnName(BoundValue, BindingType)
{
	if(this.BindingsArray != null && this.BindingsArray.length)
	{
		for(i = 0, n = this.BindingsArray.length; i < n; i++)
		{
			BindingObj = this.BindingsArray[i];
			if(BindingObj["Value"] == BoundValue && BindingObj["Type"] == BindingType)
			{
				return BindingObj["Name"];
			}
		}
	}
	return "";
}

function RSGetBoundColumnNameByIndex(Index)
{
	if(this.BindingsArray != null && Index >= 0 && Index < this.BindingsArray.length)
	{
		BindingObj = this.BindingsArray[Index];
		return BindingObj["Name"];
	}
	return "";
}


function CreateRS(Name, DSN, UID, PWD, SQL, UniqueKeyCols, Updateable, CursorType, CursorLocation, LockType, TableNames)
{
	var Object;
	if(CreateRS.arguments.length > CreateRS.length)
	{
		// variable args
		var EvalMe = "Object = new Recordset(Name, DSN, UID, PWD, SQL, UniqueKeyCols, Updateable, CursorType, CursorLocation, LockType, TableNames";
		var nArgs = CreateRS.arguments.length;
		var nArg = CreateRS.length;
		for(; nArg < nArgs; nArg++)
		{
			EvalMe += ", CreateRS.arguments[" + nArg + "]";
		}
		EvalMe += ");"
		eval(EvalMe);
	}
	else
	{
		Object = new Recordset(Name, DSN, UID, PWD, SQL, UniqueKeyCols, Updateable, CursorType, CursorLocation, LockType, TableNames);
	}
	return Object;
}

function iif(cond, X, Y)
{
	if(cond)
	{
		return X;
	}
	return Y;
}

function iifWithHTMLEncode(cond, X, Y)
{
	if(cond)
	{
		if(X == null)
		{
			X = "";
		}
		return Server.HTMLEncode(X);
	}
	return Y;
}


function RSTrimCRLF(Value)
{
	TrimValue = "" + Value;
	CRLFIndex = TrimValue.length;
	CRLFIndex = CRLFIndex -1;
	if(CRLFIndex >= 0)
	{
		for (;;)
		{
			if ((TrimValue.charAt(CRLFIndex) != '\n') && (TrimValue.charAt(CRLFIndex) != '\r'))
			{
				break;
			}
			else
			{
				CRLFIndex--;
			}
		}

		if (CRLFIndex < 0)
		{
			CRLFIndex =0;
		}

		TrimValue = TrimValue.substring(0,CRLFIndex+1);
	}
	
	return TrimValue;
}

function ReplaceSingleQuoteWithTwo(Value)
{
	var re = /'/i;
	var RValue = String(Value).replace(re, "''");
	return RValue;
}

</SCRIPT>

<SCRIPT LANGUAGE=VBScript RUNAT=Server>
function RSFormatDateTime(Date, NamedFormat)
	if Date then
		RSFormatDateTime = FormatDateTime(Date, NamedFormat)
	else
		RSFormatDateTime = Date
	end if
end function
</SCRIPT>

