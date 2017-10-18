var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var EquipCkOutRS = new Recordset("EquipCkOutRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true);
/* client-side recordset */
var TypeforCkOutRS = new Recordset("TypeforCkOutRS", "lumpkin", "admin", "", "SELECT tblEquipment.InventoryID,  tblEquipment.Type,  tblEquipment.Description FROM tblEquipment", "InventoryID", true);
/* client-side recordset */
var Programs2RS = new Recordset("Programs2RS", "lumpkin", "admin", "", "SELECT tblPrograms.ProgramID,  tblPrograms.ProgramName FROM tblPrograms", "ProgramID", true);
/* client-side recordset */
var PersonnelRS = new Recordset("PersonnelRS", "lumpkin", "admin", "", "SELECT * FROM tblPersonnel", "PersonnelID", true);
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton1 = new makeButton('media/inventory.jpg,media/inventoryhi.jpg',
	'Inventory.asp','ImageButton1','CheckoutInventory','ImageButton1',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton1;
if (document.all) var Text2 = document.all.Text2;
if (document.all) var Text3 = document.all.Text3;
var equiptype= new listDef('equiptype','select-one','Type',nullFunc,nullFunc,nullFunc);
var quantity= new editDef('quantity',7,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text6 = document.all.Text6;
var equipdescrip= new listDef('equipdescrip','select-one','Description',nullFunc,nullFunc,nullFunc);
var mm= new listDef('mm','select-one','Column 832',nullFunc,nullFunc,nullFunc);
var dd= new listDef('dd','select-one','Column 826',nullFunc,nullFunc,nullFunc);
var yy= new listDef('yy','select-one','Column 829',nullFunc,nullFunc,nullFunc);
if (document.all) var Text1 = document.all.Text1;
if (document.all) var Text7 = document.all.Text7;
var personnel= new listDef('personnel','select-one','FullName',nullFunc,nullFunc,nullFunc);
var program= new listDef('program','select-one','ProgramName',nullFunc,nullFunc,nullFunc);
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','CheckoutInventory','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton9;

