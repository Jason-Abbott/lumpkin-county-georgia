var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var testinventRS = new Recordset("testinventRS", "lumpkin", "admin", "", "SELECT tblEquipmentType.InventoryTypeID,  tblEquipmentType.Type FROM tblEquipmentType", "InventoryTypeID", true);
var equip= new listDef('equip','select-one','Type',nullFunc,nullFunc,nullFunc);
if (document.all) var Text1 = document.all.Text1;
var quantity= new editDef('quantity',16,1,nullFunc,nullFunc,nullFunc);
var FormButton1= new fButtonDef('FormButton1',FormButton1__onClick);

