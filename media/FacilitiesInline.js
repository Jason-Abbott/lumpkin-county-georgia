var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var FacilityRS = new Recordset("FacilityRS", "lumpkin", "admin", "", "SELECT tblFacilities.FacilityID, tblFacilities.Name FROM tblFacilities", "FacilityID", true);
if (document.all) var Text1 = document.all.Text1;
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton3 = new makeButton('media/mainhome.jpg,media/mainhilite.jpg',
	'Home.asp','ImageButton3','Facilities','ImageButton3',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton3;
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','Facilities','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton9;
var location= new listDef('location','select-one','Name',nullFunc,nullFunc,nullFunc);
var ImageButton1 = new makeButton('media/searchevents.jpg,media/searcheventshi.jpg',
	'SearchEvents.asp','ImageButton1','Facilities','ImageButton1',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 2] = ImageButton1;
if (document.all) var Text2 = document.all.Text2;
var ImageButton6 = new makeButton('media/newevent.jpg,media/neweventhi.jpg',
	'NewEvent.asp','ImageButton6','Facilities','ImageButton6',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 3] = ImageButton6;
var ImageButton10 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton10_onClick()','ImageButton10','Facilities','ImageButton10',ImageButton10_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 4] = ImageButton10;
var facilityname= new editDef('facilityname',33,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text3 = document.all.Text3;
var ImageButton2 = new makeButton('media/facilitieshi.jpg',
	'#ImageButton2_Anchor','ImageButton2','Facilities','ImageButton2',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 5] = ImageButton2;
var ImageButton4 = new makeButton('media/personnel.jpg,media/personnelhi.jpg',
	'Personnel.asp','ImageButton4','Facilities','ImageButton4',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 6] = ImageButton4;
if (document.all) var Text4 = document.all.Text4;
var ImageButton7 = new makeButton('media/inventory.jpg,media/inventoryhi.jpg',
	'Inventory.asp','ImageButton7','Facilities','ImageButton7',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 7] = ImageButton7;
var ImageButton11 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton11_onClick()','ImageButton11','Facilities','ImageButton11',ImageButton11_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 8] = ImageButton11;
var deletefacility= new listDef('deletefacility','select-one','Name',nullFunc,nullFunc,nullFunc);
if (document.all) var Text5 = document.all.Text5;
var ImageButton5 = new makeButton('media/fees.jpg,media/feeshi.jpg',
	'Fees.asp','ImageButton5','Facilities','ImageButton5',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 9] = ImageButton5;
var ImageButton8 = new makeButton('media/reports.jpg,media/reportshi.jpg',
	'#ImageButton8_Anchor','ImageButton8','Facilities','ImageButton8',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 10] = ImageButton8;

