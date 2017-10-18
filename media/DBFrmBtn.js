function fButtonDef(name,onClick) {
	this.name = name;
	this.type = "Form Element";
	this.onClick = onClick;
	this.click = fBC;
	this.getValue = fBGV;
	this.setValue = fBSV;
	this.getElementID = getElementID;
	this.elementResolved = elementResolved;
	this.elementID = null;
}

function fBC() { return (this.elementResolved() ? this.elementID.click() : false); }
function fBGV() { return (this.elementResolved() ? this.elementID.value : ""); }
function fBSV(text) { if (this.elementResolved()) this.elementID.value = text; }

