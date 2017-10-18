var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var EquipmentRS = new Recordset("EquipmentRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true);
/* client-side recordset */
var EquipDescripRS = new Recordset("EquipDescripRS", "lumpkin", "admin", "", "SELECT tblEquipment.InventoryID,  tblEquipment.Type,  tblEquipment.Description FROM tblEquipment", "InventoryID", true);
if (document.all) var Text2 = document.all.Text2;
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton3 = new makeButton('media/mainhome.jpg,media/mainhilite.jpg',
	'StaffHome.asp','ImageButton3','Inventory','ImageButton3',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton3;
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','Inventory','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton9;
var searchtype= new listDef('searchtype','select-one','Type',searchtype__onChange,nullFunc,nullFunc);
var equiptype= new listDef('equiptype','select-one','Description',nullFunc,nullFunc,nullFunc);
var ImageButton7 = new makeButton('media/inventoryhi.jpg',
	'Inventory.asp','ImageButton7','Inventory','ImageButton7',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 2] = ImageButton7;
var ImageButton2 = new makeButton('media/addinvent.jpg,media/addinventhi.jpg',
	'JavaScript:_ImageButton2_onClick()','ImageButton2','Inventory','ImageButton2',ImageButton2_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 3] = ImageButton2;
var ImageButton1 = new makeButton('media/ckoutinvent.jpg,media/ckoutinventhi.jpg',
	'JavaScript:_ImageButton1_onClick()','ImageButton1','Inventory','ImageButton1',ImageButton1_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 4] = ImageButton1;

