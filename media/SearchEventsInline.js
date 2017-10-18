var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var Recordset2 = new Recordset("Recordset2", "lumpkin", "admin", "", "SELECT tblFacilities.FacilityID, tblFacilities.Name FROM tblFacilities", "FacilityID", true);
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton4 = new makeButton('media/mainhome.jpg,media/mainhilite.jpg',
	'StaffHome.asp','ImageButton4','SearchEvents','ImageButton4',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton4;
var mm= new listDef('mm','select-one','Column 694',nullFunc,nullFunc,nullFunc);
var dd= new listDef('dd','select-one','Column 68',nullFunc,nullFunc,nullFunc);
var yy= new listDef('yy','select-one','Column 71',nullFunc,nullFunc,nullFunc);
var location= new listDef('location','select-one','Name',nullFunc,nullFunc,nullFunc);
var ImageButton5 = new makeButton('media/searcheventshi.jpg',
	'#ImageButton5_Anchor','ImageButton5','SearchEvents','ImageButton5',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton5;
var ImageButton1 = new makeButton('media/srchbydate.jpg,media/srchbydatehi.jpg',
	'JavaScript:_ImageButton1_onClick()','ImageButton1','SearchEvents','ImageButton1',ImageButton1_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 2] = ImageButton1;
var ImageButton3 = new makeButton('media/srchloc.jpg,media/srchlochi.jpg',
	'JavaScript:_ImageButton3_onClick()','ImageButton3','SearchEvents','ImageButton3',ImageButton3_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 3] = ImageButton3;
var ImageButton2 = new makeButton('media/srchdateloc.jpg,media/srchdatelochi.jpg',
	'JavaScript:_ImageButton2_onClick()','ImageButton2','SearchEvents','ImageButton2',ImageButton2_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 4] = ImageButton2;

