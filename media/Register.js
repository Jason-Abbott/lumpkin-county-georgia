function FormButton1__onClick() {
document.forms[0].action="SearchBio.asp";
document.forms[0].method="POST";
document.forms[0].submit();
 }
function _FormButton1__onClick() { if (FormButton1) return FormButton1.onClick(); }
function FormButton2__onClick() {
document.forms[0].action="NewRegister.asp";
document.forms[0].method="POST";
document.forms[0].submit();
 }
function _FormButton2__onClick() { if (FormButton2) return FormButton2.onClick(); }

