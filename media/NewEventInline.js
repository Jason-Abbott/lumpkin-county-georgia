var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var Recordset3 = new Recordset("Recordset3", "lumpkin", "admin", "", "SELECT tblFacilities.FacilityID, tblFacilities.Name FROM tblFacilities", "FacilityID", true);
var title= new editDef('title',39,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text1 = document.all.Text1;
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton4 = new makeButton('media/mainhome.jpg,media/mainhilite.jpg',
	'StaffHome.asp','ImageButton4','NewEvent','ImageButton4',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton4;
var desc= new editDef('desc',39,3,nullFunc,nullFunc,nullFunc);
if (document.all) var Text2 = document.all.Text2;
var ImageButton6 = new makeButton('media/neweventhi.jpg',
	'NewEvent.asp','ImageButton6','NewEvent','ImageButton6',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton6;
var mm= new listDef('mm','select-one','Column 243',nullFunc,nullFunc,nullFunc);
var dd= new listDef('dd','select-one','Column 227',nullFunc,nullFunc,nullFunc);
var yy= new listDef('yy','select-one','Column 230',nullFunc,nullFunc,nullFunc);
if (document.all) var Text4 = document.all.Text4;
var r= new checkBoxDef('r',nullFunc);
var ca= new checkBoxDef('ca',nullFunc);
if (document.all) var Text11 = document.all.Text11;
if (document.all) var Text12 = document.all.Text12;
if (document.all) var Text13 = document.all.Text13;
var recur= new editDef('recur',2,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text7 = document.all.Text7;
if (document.all) var Text8 = document.all.Text8;
var start= new editDef('start',9,1,nullFunc,nullFunc,nullFunc);
var ampm1= new listDef('ampm1','select-one','Column 233',nullFunc,nullFunc,nullFunc);
if (document.all) var Text5 = document.all.Text5;
var end= new editDef('end',9,1,nullFunc,nullFunc,nullFunc);
var ampm2= new listDef('ampm2','select-one','Column 237',nullFunc,nullFunc,nullFunc);
if (document.all) var Text6 = document.all.Text6;
var location= new listDef('location','select-one','Name',nullFunc,nullFunc,nullFunc);
if (document.all) var Text9 = document.all.Text9;
var group1= new editDef('group1',37,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text3 = document.all.Text3;
var group2= new editDef('group2',37,1,nullFunc,nullFunc,nullFunc);
var group3= new editDef('group3',37,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text10 = document.all.Text10;
var notes= new editDef('notes',37,5,nullFunc,nullFunc,nullFunc);
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','NewEvent','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 2] = ImageButton9;

