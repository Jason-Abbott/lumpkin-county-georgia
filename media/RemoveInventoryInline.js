var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var RemoveTypesRS = new Recordset("RemoveTypesRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type,  tblEquipmentType.Description FROM tblEquipmentType", "InventoryTypeID", true);
/* client-side recordset */
var RemoveDescripRS = new Recordset("RemoveDescripRS", "lumpkin", "admin", "", "SELECT tblEquipment.InventoryID,  tblEquipment.Type,  tblEquipment.Description FROM tblEquipment", "InventoryID", true);
if (document.all) var Text1 = document.all.Text1;
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton1 = new makeButton('media/inventory.jpg,media/inventoryhi.jpg',
	'Inventory.asp','ImageButton1','RemoveInventory','ImageButton1',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton1;
var type= new listDef('type','select-one','Type',nullFunc,nullFunc,nullFunc);
var descrip= new listDef('descrip','select-one','Description',nullFunc,nullFunc,nullFunc);
var ImageButton2 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton2_onClick()','ImageButton2','RemoveInventory','ImageButton2',ImageButton2_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton2;
var quantity= new editDef('quantity',6,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text2 = document.all.Text2;

