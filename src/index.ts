import Main from "./Actions";
import Menu from "./Menu";
import API from "./API/request";

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
	console.log("Initialisation");

	// 1. Delete all previous information from property
	PropertiesService.getScriptProperties().setProperty("activity", "undefined");
	PropertiesService.getScriptProperties().setProperty(
		"projectrow",
		"undefined"
	);
	// 2. Delete all previous information from sheet
	ss.getRange("F3").setValue(null);
	ss.getRange("F4").setValue(null);
	main.weekTable.setNameDropdown("");
	main.sprintTable.setNameDropdown("");

	//3. Delete all time information of the sheet
	main.cleanTable();

	main.setProperties();
	if (ss.getRange("H2").isChecked()) {
		let api = new API();
		let spreadsheet = main.spreadsheet;
		let firstRow = spreadsheet.suiviSheet
			.createTextFinder("CONSOMMÉ EN JOURS DEV")
			.findNext()
			.getRow();

		let project_code = ss.getRange("F2").getValue();
		let date_start = ss.getRange(`B${firstRow + 1}`).getValue();
		var date = new Date(date_start),
			mnth = ("0" + (date.getMonth() + 1)).slice(-2),
			day = ("0" + date.getDate()).slice(-2);
		// let start_date = [date.getFullYear(), mnth, day].join("-");
		// let end_date = [date.getFullYear() + 5, mnth, day].join("-");

		//4. Launch search
		// let resultAPI = api.getAllTimeSheetRow(project_code, start_date, end_date);
		let resultUserAPI: any = api.getUsersByProject(project_code);

		let responseUser = resultUserAPI.userArray.users;
		let usersArrFrmted = Object.keys(responseUser).map((key) => {
			return `${responseUser[key].name} - ${responseUser[key].function}`;
		});

		main.weekTable.setNameDropdown(usersArrFrmted);
		main.sprintTable.setNameDropdown(usersArrFrmted);
		// console.log("resultAPI", JSON.stringify(resultAPI, null, 2));

		// let responseTimesheetRow = resultAPI?.["hydra:member"];

		// return PropertiesService.getScriptProperties().setProperty(
		// 	"projectrow",
		// 	JSON.stringify(responseTimesheetRow)
		// );
	} else {
		// 4. Open Side bar
		var html = HtmlService.createTemplateFromFile("src/Module/Page")
			.evaluate()
			.setTitle("Initier l'automatisation");
		SpreadsheetApp.getUi().showSidebar(html);
	}
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
	let projectrow = JSON.stringify(activity.projectActivity[activityIndex].id);

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

		activityArray.forEach((element, index) => {
			string.push(
				"<option value='" +
					index.toString() +
					"'>" +
					element.label.toString() +
					"</option>"
			);
		});
		return (
			'<div class="btn mb-3" > <label for="activity" class="mb-2"> <b>Choisir l\'activité</b></label><select class="form-select form-select-lg mb-3" name="activity" id="activity" onChange="getSelectedValue(\'activity\')"><option value={null}>Selectionner une activite</option>' +
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
