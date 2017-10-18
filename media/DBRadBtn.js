function radioButtonDef(name,value,onClick) {
	this.name = name;
	this.type = "Radio";
	this.value = value;
	this.onClick = onClick;
	this.select = radS;
	this.getState = radGS;
	this.getValue = radGV;
	this.setValue = radSV;
	this.getElementID = getElementID;
	this.elementResolved = elementResolved;
	this.elementID = null;
}
 
function radS() { if (this.elementResolved()) this.elementID.checked = true; }
function radGS() { return (this.elementResolved() ? this.elementID.checked : false); }
function radGV() { return (this.elementResolved() ? this.elementID.value : ""); }

function radSV(value) {
	this.value = value;
	if (this.elementResolved()) this.elementID.value = value;
}

