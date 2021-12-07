import Main from "./Actions";
import Menu from "./Menu";
import SprintTable from "./Module/tables/SprintTable";
import UxWeekTable from "./Module/tables/UxWeekTable";
import ProjectProperties from "./Module/utils/ProjectProperties";

let ss = SpreadsheetApp.getActive();

//************** Trigger on spreadSheet open *********//
//************** => use to creat custom menu *********//

export function atOpen(): void {
	SpreadsheetApp.getUi()
		.createMenu("Automatisation")
		.addItem("Initialiser l'automatisation du suivi opérationnel", "doGet")
		.addItem("Démarrer l'automatisation !", "trigger")
		.addToUi();
}

//******************* Trigger on spreadSheet change **************//
//************** => use to catch suivi operationelle change *********//

export function atChange(e): void {
	const main = new Main(e);

	main.setProperties(e);

	let lastChange = e?.range?.getA1Notation();
	let sprintDev = main?.spreadsheet?.spreadsheet
		.getRangeByName("sprintDev")
		.getA1Notation();
	let semaineUX = main?.spreadsheet?.spreadsheet
		.getRangeByName("semaineUX")
		.getA1Notation();

	if (lastChange && (lastChange === sprintDev || lastChange === semaineUX)) {
		//Si ont change de semaine ou de sprint, MAJ des valeurs ...
		SpreadsheetApp.getActive().toast("Chargement des valeurs de SX...", "", -1);

		main.openSXNomenc();
		main.spreadsheet.spreadsheet.toast("Terminé.");
	}
}

//******************* Initiate the SideBar **************//

export function doGet() {
	const main = new Main();
	main.setProperties();

	// 1. Delete all previous information from property
	PropertiesService.getScriptProperties().setProperty("activity", "undefined");
	PropertiesService.getScriptProperties().setProperty(
		"projectrow",
		"undefined"
	);
	// 2. Delete all previous information from sheet
	ss.getRange("F3").setValue(null);
	ss.getRange("F4").setValue(null);

	// 3. Open Side bar
	var html = HtmlService.createTemplateFromFile("src/Module/Page")
		.evaluate()
		.setTitle("Initier l'automatisation");

	SpreadsheetApp.getUi().showSidebar(html);
}

//******************* Automatisation **************//

export function trigger() {
	let ss = SpreadsheetApp.getActive();
	let triggers = ScriptApp.getProjectTriggers();
	let exist: boolean = false;

	triggers.forEach((trigger) => {
		if (trigger.getHandlerFunction() === "atChange") {
			exist = true;
		}
	});
	if (!exist) {
		ScriptApp.newTrigger("atChange").forSpreadsheet(ss).onEdit().create();
	}
}

//******************* Set Activity from Google Sheet **************//

export function setPhases(phaseIndex, phaseText) {
	const main = new Main();
	const phases = JSON.parse(main.properties._phases);

	// 1. Write phases on sheet
	ss.getRange("F3").setValue(phaseText);

	// 2. set property activity compared of phase chosen
	return PropertiesService.getScriptProperties().setProperty(
		"activity",
		JSON.stringify(phases[phaseIndex])
	);
}

export function setActivity(activityIndex, ActivityText) {
	const main = new Main();

	// 1. write activity on sheet
	const activity = JSON.parse(main.properties._activity);
	ss.getRange("F4").setValue(ActivityText);

	// 2. set property projectrow compared of activity chosen
	let projectrow = JSON.stringify(
		activity.projectActivity[activityIndex].projectRow
	);
	return PropertiesService.getScriptProperties().setProperty(
		"projectrow",
		projectrow
	);
}

export function setActivityHTML() {
	const main = new Main();

	// Write activity dropdown

	if (main.properties._activity) {
		const activity = JSON.parse(main.properties._activity);
		let activityArray = activity.projectActivity;
		let string = [];

		for (var y = 0; y < activityArray.length; y++) {
			string.push(
				"<option value='" +
					y.toString() +
					"'>" +
					activityArray[y].label.toString() +
					"</option>"
			);
		}
		return (
			'<div class="btn mb-3" > <label for="activity"> <b>Choisir l\'activité</b></label><select class="form-select form-select-lg mb-3" name="activity" id="activity" onChange="getSelectedValue(\'activity\')"><option value={null}>Selectionner une activite</option>' +
			string.join() +
			"</select></div>"
		);
	}
}

export function getNewHtml() {
	const main = new Main();

	if (ss.getRange("F4").getValue()) {
		// 1. close sidebar
		var html2 = HtmlService.createHtmlOutput(
			"<script>google.script.host.close();</script>"
		);
		SpreadsheetApp.getUi().showSidebar(html2);

		// 2. set user list dropdown
		return main.setUserTimesheetRow();
	} else {
		// 1. reload side bar
		var html = HtmlService.createTemplateFromFile("src/Module/Page")
			.evaluate()
			.getContent();
		return html;
	}
}

//******************* Pick selected range from Google Sheet **************//

// export function getSelectedRange(textFieldId) {
// 	var selected = SpreadsheetApp.getActiveSheet().getActiveRange(); // Gets the selected range
// 	var rangeString = selected.getA1Notation(); // converts it to the A1 type notation
// 	SpreadsheetApp.getActive().setNamedRange(textFieldId, selected);
// 	return rangeString;
// }
// //******************* Pick named range from Google Sheet **************//

// export function generateNamedRanges() {
// 	let main: Main = new Main();
// 	// let spreadSheet = main.spreadsheet;
// 	// let properties = main.properties;
// 	// let sprintTable: SprintTable = main.sprintTable;
// 	// let uxTable: UxWeekTable = main.weekTable;
// 	// sprintTable.initialize("dev"); //TODO What append do i need this ?
// 	// uxTable.initialize("ux"); //TODO What append do i need this ?
// }
