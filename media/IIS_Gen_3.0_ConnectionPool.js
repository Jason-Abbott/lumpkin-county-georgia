<SCRIPT LANGUAGE=JavaScript RUNAT=Server>
function ConnectionCache()
{
	CPReconnect(this);

	// create connections array
	if(!this.Connections) 
	{
		this.Connections = new Array();
	}

	return this;
}

function CPReconnect(Object)
{
	Object.FindConnection = CPFindConnection;
	Object.AddConnection = CPAddConnection;
}

function CPFindConnection(DSN, UID, PWD)
{
	var Connections = this.Connections;
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

function CPAddConnection(DSN, UID, PWD, Connection)
{
	var Pair = new Object();
	var Key = new Object();
	Key["DSN"] = DSN;
	Key["UID"] = UID;
	Key["PWD"] = PWD;
	Pair["Key"] = Key;
	Pair["Value"] = Connection;
	var Connections = this.Connections;
	Connections[Connections.length] = Pair;
}

function CreateConnectionCache()
{
	return new ConnectionCache();
}

</SCRIPT>

