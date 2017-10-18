// Support Script (49)
function PopulateWith(TheList, Items)
{
  var iCount = TheList.getCount();
  for (var i=0; i<iCount; i++)
    if (TheList.elementID)
      TheList.elementID.options[0] = null;
    else
      alert("bad or wrong DE for state population");

  i = 0;
  iCount = Items.length;
 	for (i=0; i<iCount; i++)
    TheList.addOption(Items[i], i, "");

  TheList.setSelectedByPosition(0);
}
// Support Script (50)
function ValidateDropDown()
{
  msg = "";
  var iPos = this.getSelectedPosition();
  if (iPos<=0)
  {
    msg = "Please make a selection."
  }

  return msg;
}
function document_onLoad() {
var i = 0;
s = new Array(51);
s[i++] = "Please select a State"
s[i++] = "AK"
s[i++] = "AL"
s[i++] = "AZ"
s[i++] = "AR"
s[i++] = "CA"
s[i++] = "CO"
s[i++] = "CT"
s[i++] = "DE"
s[i++] = "FL"
s[i++] = "GA"
s[i++] = "HI"
s[i++] = "ID"
s[i++] = "IL"
s[i++] = "IN"
s[i++] = "IA"
s[i++] = "KS"
s[i++] = "KY"
s[i++] = "LA"
s[i++] = "ME"
s[i++] = "MD"
s[i++] = "MA"
s[i++] = "MI"
s[i++] = "MN"
s[i++] = "MS"
s[i++] = "MO"
s[i++] = "MT"
s[i++] = "NE"
s[i++] = "NV"
s[i++] = "NH"
s[i++] = "NJ"
s[i++] = "NM"
s[i++] = "NY"
s[i++] = "NC"
s[i++] = "ND"
s[i++] = "OH"
s[i++] = "OK"
s[i++] = "OR"
s[i++] = "PA"
s[i++] = "RI"
s[i++] = "SC"
s[i++] = "SD"
s[i++] = "TN"
s[i++] = "TX"
s[i++] = "UT"
s[i++] = "VT"
s[i++] = "VA"
s[i++] = "WA"
s[i++] = "WV"
s[i++] = "WI"
s[i++] = "WY"

PopulateWith(pstate, s);
pstate.Validate=ValidateDropDown;
var i = 0;
s = new Array(51);
s[i++] = "Please select a State"
s[i++] = "AK"
s[i++] = "AL"
s[i++] = "AZ"
s[i++] = "AR"
s[i++] = "CA"
s[i++] = "CO"
s[i++] = "CT"
s[i++] = "DE"
s[i++] = "FL"
s[i++] = "GA"
s[i++] = "HI"
s[i++] = "ID"
s[i++] = "IL"
s[i++] = "IN"
s[i++] = "IA"
s[i++] = "KS"
s[i++] = "KY"
s[i++] = "LA"
s[i++] = "ME"
s[i++] = "MD"
s[i++] = "MA"
s[i++] = "MI"
s[i++] = "MN"
s[i++] = "MS"
s[i++] = "MO"
s[i++] = "MT"
s[i++] = "NE"
s[i++] = "NV"
s[i++] = "NH"
s[i++] = "NJ"
s[i++] = "NM"
s[i++] = "NY"
s[i++] = "NC"
s[i++] = "ND"
s[i++] = "OH"
s[i++] = "OK"
s[i++] = "OR"
s[i++] = "PA"
s[i++] = "RI"
s[i++] = "SC"
s[i++] = "SD"
s[i++] = "TN"
s[i++] = "TX"
s[i++] = "UT"
s[i++] = "VT"
s[i++] = "VA"
s[i++] = "WA"
s[i++] = "WV"
s[i++] = "WI"
s[i++] = "WY"

PopulateWith(gstate, s);
gstate.Validate=ValidateDropDown;
 }
function pstate__onBlur() {
ErrorMsg = this.Validate();
if (ErrorMsg != "")
{
  alert(ErrorMsg);
}
 }
function _pstate__onBlur() { if (pstate) return pstate.onBlur(); }
function gstate__onBlur() {
ErrorMsg = this.Validate();
if (ErrorMsg != "")
{
  alert(ErrorMsg);
}
 }
function _gstate__onBlur() { if (gstate) return gstate.onBlur(); }
function FormButton1__onClick() {
document.forms[0].action="RegisterSubmit.asp";
document.forms[0].method="POST";
document.forms[0].submit();
 }
function _FormButton1__onClick() { if (FormButton1) return FormButton1.onClick(); }

