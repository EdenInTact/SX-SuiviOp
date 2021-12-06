function onOpen() {
	AutoCRA.onOpen();
}

function onChange(e) {
	AutoCRA.onChange(e);
}

function trigger() {
	return AutoCRA.trigger();
}

function doGet() {
	var ss = SpreadsheetApp.getActive();
	AutoCRA.doGet();
}

// function setActivity() {
// 	return AutoCRA.setActivity();
// }

// function addDev() {
// 	return AutoCRA.addDev();
// }

// function addUx() {
// 	return AutoCRA.addUx();
// }

// function addSprint() {
// 	return AutoCRA.addSprint();
// }

// function addRecette() {
// 	return AutoCRA.addRecette();
// }

// function addWeek() {
// 	return AutoCRA.addWeek();
// }

// function getSelectedRange(textFieldId) {
// 	return AutoCRA.getSelectedRange(textFieldId);
// }

// function generateNamedRanges() {
// 	return AutoCRA.generateNamedRanges();
// }
