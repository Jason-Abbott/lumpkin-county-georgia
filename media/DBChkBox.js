function checkBoxDef(name,onClick) {
	this.name = name;
	this.type = "Form Element";
	this.onClick = onClick;
	this.setState = CBSS;
	this.getState = CBGS;
	this.getValue = CBGV;
	this.setValue = CBSV;
	this.toggle = CBT;
	this.getElementID = getElementID;
	this.elementResolved = elementResolved;
	this.elementID = null;
}
 
function CBSS(state) { if (this.elementResolved()) this.elementID.checked = state; }
function CBGS() { return (this.elementResolved() ? this.elementID.checked : false); }
function CBGV() { return (this.elementResolved() ? this.elementID.value : ""); }
function CBSV(text) { if (this.elementResolved()) this.elementID.value = text; }
function CBT() { if (this.elementResolved()) this.elementID.click(); }

