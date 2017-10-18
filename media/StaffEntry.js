function document_onLoad() {
staffname.focus();
 }
function ImageButton9_onClick() {
document.forms[0].action="EntryCheck.asp";
document.forms[0].method="POST";
document.forms[0].submit();
 }
function _ImageButton9_onClick() { if (ImageButton9) return ImageButton9.onClick(); }

