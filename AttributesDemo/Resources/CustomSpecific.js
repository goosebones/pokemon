var removedElement = new Object();
var countOfEleId = 0;
function extractLabelFromName(name){
	var i = name.indexOf('_');
	var label = name.substring(i+1);
	return label;
}
function optionSelect(name,value){
	var label = extractLabelFromName(name);
	var id = "eBayItemSpecificName_"+label;
	var element = document.getElementById(id);
	element.value = value;
}
function remove(name){
	var label = extractLabelFromName(name);
	var id = "SpecificLayer_"+label;
	var element = document.getElementById(id);
	removedElement[id] = element;
	id = "eBayItemSpecificNameCache_"+label;
	var nameElement = document.getElementById(id);
	var eName = nameElement.value;
	var root = document.getElementById("CustomItemSpecificGroup");
	root.removeChild(element);
	buildElement(label,eName);
}
function buildElement(label,eName){
	var newSpan = document.createElement("span");
	newSpan.setAttribute("id","Suggest_"+label);
	newSpan.innerHTML = '<span style=\"margin: 10px 10px 20px 0px;\" ><a href=\"javascript:void(0);\">'+
		'<span id=\"reAddBtn_'+label+'\" onclick=\"reAddElement(this.id);\"><img hspace=\"1\" border=\"0\" align=\"absmiddle\" src=\"http://pics.qa.ebaystatic.com/aw/pics/buttons/btnOptionAdd.gif\" />'+
		eName+'</span></a></span>';
	var root = document.getElementById("SuggestedSectionLyr");
	root.appendChild(newSpan);
}
function reAddElement(name){
	var label = extractLabelFromName(name);
	var id = "SpecificLayer_"+label;
	var element = removedElement[id];
	var root = document.getElementById("CustomItemSpecificGroup");
	root.appendChild(element);
	id = "Suggest_"+label;
	element = document.getElementById(id);
	root = document.getElementById("SuggestedSectionLyr");
	root.removeChild(element);
}
function addNewSpecific(){
	var newElement = document.createElement("span");
	newElement.setAttribute("id","Custom_"+countOfEleId);
	newElement.innerHTML = '<div style=\'padding-top: 10px;\'>'+
		'<input id=\'itemSpecificName_'+countOfEleId+'\' type=\'text\' style=\'color: gray;font: italic 900 11px verdana;\' maxlength=\'30\' value=\'\' name=\'itemSpecificName_'+countOfEleId+'\'/>'+
		'<div style=\'margin-top: 5px;\'>'+'<input id=\'itemSpecificValue_'+countOfEleId+'\' '+
		'type=\'text\' style=\'color: gray; font: italic 900 11px verdana;\' maxlength=\'50\' value=\'\' name=\'itemSpecificValue_'+countOfEleId+'\'/>'+
		'<span id=\'CusRemove_'+countOfEleId+'\'onclick=\'removeNewSpecific(this.id);\'><a id=\'CusRemove_'+countOfEleId+'\' class=\'navigation\' href=\'javascript:void(0)\'>Remove</a></span>';
	countOfEleId++;
	var root = document.getElementById("NewCustomItemSpecific");
	root.appendChild(newElement);
}
function removeNewSpecific(name){
	var label = extractLabelFromName(name);
	var id = "Custom_"+label;
	var element = document.getElementById(id);
	var root = document.getElementById("NewCustomItemSpecific");
	root.removeChild(element);
}
