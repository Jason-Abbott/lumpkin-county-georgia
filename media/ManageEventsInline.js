var Form1= new formDef('Form1',nullFunc,nullFunc);
/* client-side recordset */
var locationchangeRS = new Recordset("locationchangeRS", "lumpkin", "admin", "", "SELECT tblFacilities.FacilityID, tblFacilities.Name FROM tblFacilities", "FacilityID", true);
var eventname= new editDef('eventname',32,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text2 = document.all.Text2;
var description= new editDef('description',32,3,nullFunc,nullFunc,nullFunc);
if (document.all) var Text3 = document.all.Text3;
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton4 = new makeButton('media/mainhome.jpg,media/mainhilite.jpg',
	'StaffHome.asp','ImageButton4','ManageEvents','ImageButton4',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton4;
var eventdate= new editDef('eventdate',14,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text11 = document.all.Text11;
var starttime= new editDef('starttime',14,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text4 = document.all.Text4;
var ImageButton1 = new makeButton('media/searchevents.jpg,media/searcheventshi.jpg',
	'SearchEvents.asp','ImageButton1','ManageEvents','ImageButton1',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 1] = ImageButton1;
if (document.all) var Text5 = document.all.Text5;
var endtime= new editDef('endtime',14,1,nullFunc,nullFunc,nullFunc);
var ImageButton6 = new makeButton('media/newevent.jpg,media/neweventhi.jpg',
	'NewEvent.asp','ImageButton6','ManageEvents','ImageButton6',nullFunc,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 2] = ImageButton6;
var location= new listDef('location','select-one','Name',nullFunc,nullFunc,nullFunc);
if (document.all) var Text6 = document.all.Text6;
var participants= new editDef('participants',32,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text7 = document.all.Text7;
var participants2= new editDef('participants2',32,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text8 = document.all.Text8;
var participants3= new editDef('participants3',32,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text9 = document.all.Text9;
var notes= new editDef('notes',32,4,nullFunc,nullFunc,nullFunc);
if (document.all) var Text10 = document.all.Text10;
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','ManageEvents','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 3] = ImageButton9;
var FormButton1= new fButtonDef('FormButton1',FormButton1__onClick);

