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

function addDev() {
	return AutoCRA.addDev();
}

function addUx() {
	return AutoCRA.addUx();
}

function addSprint() {
	return AutoCRA.addSprint();
}

function addRecette() {
	return AutoCRA.addRecette();
}

function addWeek() {
	return AutoCRA.addWeek();
}
function cleanTable() {
	return AutoCRA.cleanTable();
}
