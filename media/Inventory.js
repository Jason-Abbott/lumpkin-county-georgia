function searchtype__onChange() {
var action_len = document.forms[0].elements.EquipDescripRS_Action.value.length;
temp_action = document.forms[0].elements.EquipDescripRS_Action.value;
var text_or_value = "VALUE";
var filter = "";

//filter on selectedText or selectedValue depending on parameter
if (text_or_value == "CONTENT")
  filter = searchtype.getSelectedText();
else
  filter = searchtype.getSelectedValue();


//check to see if there is something to filter on

//nothing to filter on
if(("" + filter) == "")
{
	//check and see if the filter action has already been set by another field
 if (action_len < 1)
		document.forms[0].elements.EquipDescripRS_Action.value = 'Filter("")';
}

//something to filter on
else
{
	//check and see if the filter action has already been set by another field
	var building; 
	
	//if it has append and AND or OR filter depending on the parameter.
    if ((action_len > 10) && (temp_action.indexOf("Filter") != -1)) {
		      building = document.forms[0].elements.EquipDescripRS_Action.value.substring(0, (action_len-2));
        building += " " + "AND" + " " + "Type" + ' LIKE \'' + filter + '*\'")';
        document.forms[0].elements.EquipDescripRS_Action.value = building;
		  }

	// Filter already set-up but no AND/OR
	else if (temp_action.indexOf("Filter") != -1) {
		      building = document.forms[0].elements.EquipDescripRS_Action.value.substring(0, (action_len-2));
        building += "Type" + ' LIKE \'' + filter + '*\'")';
      		document.forms[0].elements.EquipDescripRS_Action.value = building;
		}

	//Filter Not Set Up!
	else {
		      building = 'Filter("' + "Type" + ' LIKE \'' + filter + '*\'")';
		      document.forms[0].elements.EquipDescripRS_Action.value = building;
		}
}		
document.forms[0].method='POST';
 }
function _searchtype__onChange() { if (searchtype) return searchtype.onChange(); }
function ImageButton1_onClick() {
document.forms[0].action="CheckoutInventory.asp";
document.forms[0].method="POST";
document.forms[0].submit();
 }
function _ImageButton1_onClick() { if (ImageButton1) return ImageButton1.onClick(); }
function ImageButton2_onClick() {
document.forms[0].action="AddNew.asp";
document.forms[0].method="POST";
document.forms[0].submit();
 }
function _ImageButton2_onClick() { if (ImageButton2) return ImageButton2.onClick(); }
function ImageButton9_onClick() {
document.forms[0].action="SearchByType.asp";
document.forms[0].method="POST";
document.forms[0].submit();
 }
function _ImageButton9_onClick() { if (ImageButton9) return ImageButton9.onClick(); }

