var Form1= new formDef('Form1',nullFunc,nullFunc);
var staffname= new editDef('staffname',14,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text1 = document.all.Text1;
var pword= new editDef('pword',14,1,nullFunc,nullFunc,nullFunc);
if (document.all) var Text2 = document.all.Text2;
moreImage = 'media/_more_.gif'
transImage= 'media/_trans_.gif'
var ImageButton9 = new makeButton('media/submit.jpg,media/submithi.jpg',
	'JavaScript:_ImageButton9_onClick()','ImageButton9','StaffEntry','ImageButton9',ImageButton9_onClick,nullFunc,nullFunc,nullFunc,nullFunc);
BTNArray[ 0] = ImageButton9;

