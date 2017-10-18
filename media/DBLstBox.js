function listDef(name,select,dataFld,onChange,onBlur,onFocus) {
	this.name = name;
	this.type = "Form Element";
	this.select = select;
	this.dataFld = dataFld;
	this.onChange = onChange;
	this.onBlur = onBlur;
	this.onFocus = onFocus;
	this.clear = listCl;
	this.blur = listBl;
	this.focus = listFo;
	this.getCount = listGCo;
	this.addOption = addOpt;
	this.getSelectedPosition = listGetSP;
	this.getSelectedText = listGetST;
	this.getSelectedValue = listGetSV;
	this.setSelectedByPosition = listSetSBP;
	this.setSelectedByText = listSetSBT;
	this.setSelectedByValue = listSetSBV;
	this.SelAll = listSelAll;
	this.setSelectedAll = listSetSA;
	this.setSelectedNone = listSetSN;
	this.stuffListFromQuery = listSFQ;
	this.deleteOptionByPosition = listDOBP;
	this.deleteOptionByText = listDOBT;
	this.getElementID = getElementID;
	this.elementResolved = elementResolved;
	this.elementID = null;
}

function listCl() {
	with(this) {
		if (!elementResolved()) return;
		while (elementID.options.length != 0) {
			var a = elementID.options.length - 1;
			elementID.options[a] = null;
		}
	}
}

function listGetSP() {
	with(this) {
		if (!elementResolved()) return -1;
		if (select == "select-one"){
			var i = elementID.selectedIndex;
			return i;
		} else {
			var r = "";
			var b = false;
			for (var i = 0; i < elementID.length; i++) {
				if (elementID.options[i].selected == true) {
					if (b) r += ",";
					else   b = true;
					r += i;
				}
			}
			return r;
		}
	}
}

function listGetST() {
	with(this) {
		if (!elementResolved()) return "";
		if (select == "select-one"){
			var i = elementID.selectedIndex;
			var t = elementID.options[i].text;
			return t;
		} else {
			var r = "";
			var b = false;
			for (var i = 0; i < elementID.length; i++) {
				if (elementID.options[i].selected == true) {
					if (b) r += ",";
					else   b = true;
					r += elementID.options[i].text;
				}
			}
			return r;
		}
	}
}

function listGetSV() {
	with(this) {
		if (!elementResolved()) return "";
		if (select == "select-one"){
			var i = elementID.selectedIndex;
			var t = elementID.options[i].value;
			return t;
		} else {
			var r = "";
			var b = false;
			for (var i = 0; i < elementID.length; i++) {
				if (elementID.options[i].selected == true) {
					if (b) r += ",";
					else   b = true;
					r += elementID.options[i].value;
				}
			}
			return r;
		}
	}
}

function listSetSBP(position) {
	with(this) {
		if (!elementResolved()) return;
		if (position < 0 || position >= elementID.length) return;
		elementID.options[position].selected = true;
	} 
}

function listSetSBT(text) {
	with(this) {
		if (!elementResolved()) return;
		for (var i = 0; i < elementID.length; i++) {
			if (text == elementID.options[i].text) {
				elementID.options[i].selected = true;
				return;
			}
		} 
	} 
}

function listSetSBV(value) {
	with(this) {
		if (!elementResolved()) return;
		for (var i = 0; i < elementID.length; i++) {
			if (value == elementID.options[i].value) {
				elementID.options[i].selected = true;
				return;
			}
		} 
	} 
}

function listSelAll(state) {
	with(this) {
		if (!elementResolved()) return;
		if (select != "select-multiple") return;
		for (var i = 0; i < elementID.length; i++) {
			elementID.options[i].selected = state;
		} 
	} 
}

function listDOBP(position) {
	with(this) {
		if (elementResolved() && elementID.options &&
			position >= 0 && position < elementID.length &&
			elementID.options[position])
				elementID.options[position] = null;
	} 
}

function listDOBT(text) {
	with(this) {
		if (elementResolved() && elementID.options){
			for (var i = 0; i < elementID.length; i++) {
				if (elementID.options[i] && elementID.options[i].text &&
					text == elementID.options[i].text)
						elementID.options[i] = null;
			} 
		}
	}
}

function addOpt(text, position, value) {
	with(this) {
		if (!elementResolved()) return;
		var a = new Option(text);
		if (value == "") value = text;
		a.value = value;
		var newUpperBound = elementID.length;
		if (position >= 0 && position < newUpperBound) {
			for (var i = newUpperBound; i > position; i-- ) {
				elementID.options[i] = elementID.options[i-1];
			}
			elementID.options[position] = a;
		} else {
			elementID.options[newUpperBound] = a;
		}
	}
}

function listSFQ(queryName) {
	with(this) {
		clear();
		for (var i=0; i < queryName.getRowCount(); i++ ) {
			s = queryName.getString(0);
			addOption(s);
			queryName.next();
		}
		setSelectedByPosition(0);
	}
}

function listGCo() { return (this.elementResolved() ? this.elementID.length : -1); }
function listBl() { if (this.elementResolved()) this.elementID.blur(); }
function listFo() { if (this.elementResolved()) this.elementID.focus(); }
function listSetSA() { this.SelAll(true); }
function listSetSN() { this.SelAll(false); }

