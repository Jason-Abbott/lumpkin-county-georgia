var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var ProgramsRS = new Recordset("ProgramsRS", "lumpkin", "admin", "", "SELECT tblPrograms.ProgramID,  tblPrograms.ProgramName FROM tblPrograms", "ProgramID", true);
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton3 = new makeButton('media/mainhome.jpg,media/mainhilite.jpg',
	'StaffHome.asp','ImageButton3','Fees','ImageButton3',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton3;
if (document.all) var Text1 = document.all.Text1;
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','Fees','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton9;
var programname= new listDef('programname','select-one','ProgramName',nullFunc,nullFunc,nullFunc);
var ImageButton2 = new makeButton('media/feeshi.jpg',
	'#ImageButton2_Anchor','ImageButton2','Fees','ImageButton2',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 2] = ImageButton2;
if (document.all) var Text6 = document.all.Text6;
var lastname= new editDef('lastname',16,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text2 = document.all.Text2;
var firstname= new editDef('firstname',16,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text3 = document.all.Text3;
var ImageButton1 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton1_onClick()','ImageButton1','Fees','ImageButton1',ImageButton1_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 3] = ImageButton1;
var mm= new listDef('mm','select-one','Column 464',nullFunc,nullFunc,nullFunc);
var dd= new listDef('dd','select-one','Column 467',nullFunc,nullFunc,nullFunc);
var yy= new listDef('yy','select-one','Column 470',nullFunc,nullFunc,nullFunc);
if (document.all) var Text4 = document.all.Text4;

