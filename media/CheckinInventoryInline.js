var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var EquipRS = new Recordset("EquipRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true);
/* client-side recordset */
var EquipWDescripRS = new Recordset("EquipWDescripRS", "lumpkin", "admin", "", "SELECT tblEquipment.InventoryID,  tblEquipment.Type,  tblEquipment.Description FROM tblEquipment", "InventoryID", true);
/* client-side recordset */
var PeopleRS = new Recordset("PeopleRS", "lumpkin", "admin", "", "SELECT * FROM tblPersonnel", "PersonnelID", true);
/* client-side recordset */
var Programs3RS = new Recordset("Programs3RS", "lumpkin", "admin", "", "SELECT tblPrograms.ProgramID,  tblPrograms.ProgramName FROM tblPrograms", "ProgramID", true);
if (document.all) var Text2 = document.all.Text2;
var type= new listDef('type','select-one','Type',nullFunc,nullFunc,nullFunc);
var equiptype= new listDef('equiptype','select-one','Description',nullFunc,nullFunc,nullFunc);
if (document.all) var Text1 = document.all.Text1;
var personnel= new listDef('personnel','select-one','FullName',nullFunc,nullFunc,nullFunc);
if (document.all) var Text4 = document.all.Text4;
var program= new listDef('program','select-one','ProgramName',nullFunc,nullFunc,nullFunc);
if (document.all) var Text3 = document.all.Text3;
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','CheckinInventory','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton9;
var quantity= new editDef('quantity',11,1,nullFunc,nullFunc,nullFunc);

