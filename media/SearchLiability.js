function FormButton1__onClick() {
document.forms[0].action="LiabilityForm.asp";
document.forms[0].method="POST";
document.forms[0].submit();
 }
function _FormButton1__onClick() { if (FormButton1) return FormButton1.onClick(); }

