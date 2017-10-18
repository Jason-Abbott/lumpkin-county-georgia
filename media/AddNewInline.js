var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var NewEquipRS = new Recordset("NewEquipRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true);
if (document.all) var Text3 = document.all.Text3;
if (document.all) var Text7 = document.all.Text7;
var addnewdrop= new listDef('addnewdrop','select-one','Type',nullFunc,nullFunc,nullFunc);
if (document.all) var Text8 = document.all.Text8;
var newequipdescrip= new editDef('newequipdescrip',33,5,nullFunc,nullFunc,nullFunc);
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton3 = new makeButton('media/mainhome.jpg,media/mainhilite.jpg',
	'StaffHome.asp','ImageButton3','AddNew','ImageButton3',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton3;
var mm= new listDef('mm','select-one','Column 528',nullFunc,nullFunc,nullFunc);
var dd= new listDef('dd','select-one','Column 522',nullFunc,nullFunc,nullFunc);
var yy= new listDef('yy','select-one','Column 525',nullFunc,nullFunc,nullFunc);
if (document.all) var Text9 = document.all.Text9;
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','AddNew','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton9;
if (document.all) var Text1 = document.all.Text1;
var quantity= new editDef('quantity',16,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text4 = document.all.Text4;
var type= new editDef('type',38,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text5 = document.all.Text5;
var description= new editDef('description',38,6,nullFunc,nullFunc,nullFunc);
if (document.all) var Text6 = document.all.Text6;
var ImageButton1 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton1_onClick()','ImageButton1','AddNew','ImageButton1',ImageButton1_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 2] = ImageButton1;

