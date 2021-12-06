import Main from "../Actions";
import { columnToLetter, letterToColumn } from "../Module/utils/utils";
import GSpreadsheet = GoogleAppsScript.Spreadsheet;

export default class Menu {
	private ui: GoogleAppsScript.Base.Ui;
	private _periodMenu: GoogleAppsScript.Base.Menu;
	private _workerMenu: GoogleAppsScript.Base.Menu;

	constructor() {
		this.ui = SpreadsheetApp.getUi();

		this.periodMenu = this.ui
			.createMenu("Ajouter Semaine/Sprint")
			.addItem("Ajouter une semaine UX", "addWeek")
			.addItem("Ajouter un sprint", "addSprint")
			.addItem("Ajouter une recette", "addRecette");

		this.workerMenu = this.ui
			.createMenu("Ajouter Dev/UX")
			.addItem("Ajouter un UX", "addUx")
			.addItem("Ajouter un développeur", "addDev");
	}

	init() {
		this._periodMenu.addToUi();
		this.workerMenu.addToUi();
	}

	get periodMenu(): GoogleAppsScript.Base.Menu {
		return this._periodMenu;
	}

	set periodMenu(value: GoogleAppsScript.Base.Menu) {
		this._periodMenu = value;
	}

	get workerMenu(): GoogleAppsScript.Base.Menu {
		return this._workerMenu;
	}

	set workerMenu(value: GoogleAppsScript.Base.Menu) {
		this._workerMenu = value;
	}

	addSprint() {}
	addWeek() {}
}

export function addDev() {
	let ss: GSpreadsheet.Spreadsheet = SpreadsheetApp.getActive();
	let devNames = ss.getRangeByName("devsNames");
	let lastRow = devNames.getLastRow();
	let firstRow = devNames.getRow();
	console.log("devNames", devNames);

	ss.getActiveSheet().insertRowAfter(lastRow);
	let col = ss.getRangeByName("sprintDev").getColumn();

	let newDevNames = devNames.offset(0, 0, lastRow - firstRow + 2, 1);
	ss.setNamedRange("devsNames", newDevNames);
	ss.getSheetByName("Suivi opérationnel")
		.getRange(lastRow, col)
		.copyTo(ss.getSheetByName("Suivi opérationnel").getRange(lastRow + 1, col));

	let sprintCount = 1;
	ss.getNamedRanges().forEach((range) => {
		if (
			range.getName().match("Sprint") &&
			range.getName().match("Count") &&
			ss.getRangeByName(range.getName())
		) {
			let newSprintCount = ss
				.getRangeByName(range.getName())
				.offset(0, 0, lastRow - firstRow + 2, 1);
			ss.setNamedRange(range.getName(), newSprintCount);
			sprintCount++;
		}
	});

	let newRecette = ss
		.getRangeByName("recetteCount")
		.offset(0, 0, lastRow - firstRow + 2, 1);
	ss.setNamedRange("recetteCount", newRecette);

	let total = ss
		.getSheetByName("Suivi opérationnel")
		.getRange(firstRow, col, lastRow - firstRow + 2);

	ss.getSheetByName("Suivi opérationnel")
		.getRange(lastRow + 2, col)
		.setFormula(`=SUM(${total.getA1Notation()})`);
}

export function addUx() {
	let ss: GSpreadsheet.Spreadsheet = SpreadsheetApp.getActive();
	let uxNames = ss.getRangeByName("uxNames");
	let lastRow = uxNames.getLastRow();
	let firstRow = uxNames.getRow();
	let col = ss.getRangeByName("semaineUX").getColumn();
	ss.getActiveSheet().insertRowAfter(lastRow);

	let newDevNames = uxNames.offset(0, 0, lastRow - firstRow + 2, 1);
	ss.setNamedRange("uxNames", newDevNames);
	ss.getSheetByName("Suivi opérationnel")
		.getRange(lastRow, col)
		.copyTo(ss.getSheetByName("Suivi opérationnel").getRange(lastRow + 1, col));

	let weekCount = 0;
	ss.getNamedRanges().forEach((range) => {
		if (range.getName().match("Semaine") && range.getName().match("Count")) {
			let newWeekCount = ss
				.getRangeByName(range.getName())
				.offset(0, 0, lastRow - firstRow + 2, 1);
			ss.setNamedRange(range.getName(), newWeekCount);
			weekCount++;
		}
	});

	let total = ss
		.getSheetByName("Suivi opérationnel")
		.getRange(firstRow, col, lastRow - firstRow + 2);

	ss.getSheetByName("Suivi opérationnel")
		.getRange(lastRow + 2, col)
		.setFormula(`=SUM(${total.getA1Notation()})`);
}

function addDevPeriod(isSprint: boolean) {
	let main = new Main();
	let sprintTable = main.sprintTable;
	let spreadsheet = main.spreadsheet;
	let ss: GSpreadsheet.Spreadsheet = SpreadsheetApp.getActive();

	let lastColBacklog = spreadsheet.backlogSheet.getLastColumn();
	spreadsheet.backlogSheet.insertColumns(lastColBacklog, 3);
	let SprintToCopy = spreadsheet.backlogSheet.getRange(
		2,
		lastColBacklog - 6,
		spreadsheet.backlogSheet.getLastRow() - 2,
		3
	);
	let prevSprint = spreadsheet.backlogSheet.getRange(2, lastColBacklog - 1);
	let iterations = spreadsheet.backlogSheet.getRange(
		2,
		1,
		5,
		spreadsheet.backlogSheet.getMaxColumns()
	);
	let sprintNumbers = iterations
		.getValues()[0]
		.filter((it) => typeof it === "number");
	let recettes = iterations
		.getValues()[0]
		.filter((it) => typeof it === "string" && it.match("recette"));
	let dates = iterations.getValues()[2].filter((it) => it instanceof Date);
	SprintToCopy.copyTo(
		spreadsheet.backlogSheet.getRange(2, lastColBacklog, 5, 3)
	);
	let rangeToChange = spreadsheet.backlogSheet.getRange(
		2,
		lastColBacklog,
		2,
		1
	);
	let formulasToChange = rangeToChange.getFormulas();
	let sprintCol = formulasToChange[0][0].substr(1, 2);
	let ColNum = letterToColumn(sprintCol);
	ColNum--;
	let newSprint = formulasToChange[0][0].replace(
		sprintCol,
		columnToLetter(ColNum)
	);
	formulasToChange[0][0] = newSprint;
	let dateCol = formulasToChange[1][0].substr(1, 2);
	ColNum = letterToColumn(dateCol);
	ColNum += 2;
	let newDate = formulasToChange[1][0].replace(dateCol, columnToLetter(ColNum));
	formulasToChange[1][0] = newDate;
	rangeToChange.setFormulas(formulasToChange);
	let prevSprintWidth = spreadsheet.backlogSheet.getColumnWidth(
		lastColBacklog - 4
	);
	for (let i = 0; i < 3; i++) {
		spreadsheet.backlogSheet.setColumnWidth(
			lastColBacklog + i,
			prevSprintWidth
		);
	}
	let startIt =
		spreadsheet.backlogSheet
			.createTextFinder("itération")
			.findNext()
			.getColumn() + 1;
	if (isSprint) {
		spreadsheet.backlogSheet.getRange(2, lastColBacklog).setFormula(
			`${spreadsheet.backlogSheet
				.getRange(
					2,
					startIt,
					5,
					spreadsheet.backlogSheet.getLastColumn() - startIt
				)
				.createTextFinder(sprintNumbers[sprintNumbers.length - 1])
				.findNext()
				.getA1Notation()} + 1`
		);
	} else {
		spreadsheet.backlogSheet
			.getRange(2, lastColBacklog)
			.setValue("recette " + (recettes.length + 1));
	}
	let lastDate = dates[dates.length - 1];
	let lastDateCol;
	iterations.getValues()[2].forEach((date, index) => {
		if (Date.parse(lastDate) === Date.parse(date)) {
			lastDateCol = index + 1;
		}
	});
	spreadsheet.backlogSheet
		.getRange(3, lastColBacklog)
		.setFormula(
			`${spreadsheet.backlogSheet.getRange(4, lastDateCol).getA1Notation()} + 1`
		);
	spreadsheet.backlogSheet
		.getRange(7, lastColBacklog + 1, spreadsheet.backlogSheet.getLastRow(), 2)
		.setValue("");
	let bilanCol = spreadsheet.backlogSheet.getLastColumn();
	let bilan = spreadsheet.backlogSheet.getRange(
		7,
		bilanCol,
		spreadsheet.backlogSheet.getLastRow()
	);
	let bilanFormulas = bilan.getFormulas();
	bilanFormulas.forEach((formula, index) => {
		if (formula[0] && formula[0].length > 0) {
			formula[0] += `+${spreadsheet.backlogSheet
				.getRange(7 + index, bilanCol - 1)
				.getA1Notation()}`;
		}
	});
	bilan.setFormulas(bilanFormulas);

	let lastColData = spreadsheet.dataSheet.getLastColumn();
	spreadsheet.dataSheet.insertColumnAfter(lastColData);
	let formulas = spreadsheet.dataSheet
		.getRange(1, lastColData, spreadsheet.dataSheet.getLastRow(), 1)
		.getFormulas();
	formulas = updateFromBacklog(formulas, 3);
	spreadsheet.dataSheet
		.getRange(1, lastColData + 1, spreadsheet.dataSheet.getLastRow(), 1)
		.setFormulas(formulas);
	formulas.forEach((formula, i) => {
		if (!formula) {
			spreadsheet.dataSheet
				.getRange(i + 1, lastColData + 1, 1, 1)
				.setValue(
					spreadsheet.dataSheet
						.getRange(i + 1, lastColData + 1, 1, 1)
						.getValue()
				);
		}
	});
	spreadsheet.dataSheet
		.getRange(7, spreadsheet.dataSheet.getLastColumn())
		.setFormula(
			spreadsheet.dataSheet
				.getRange(8, spreadsheet.dataSheet.getLastColumn() - 1)
				.getA1Notation()
		);

	updateTable(spreadsheet.suiviSheet, "CONSOMMÉ EN JOURS DEV", "TOTAL");
	updateTable(
		spreadsheet.suiviSheet,
		"REALISE EN POINTS",
		"CIBLE Ratio - Nb jours DEV+TU pour 1 point"
	);
	updateTable(spreadsheet.suiviSheet, "NB ANOMALIES", "Anomalies restantes");
	let firstRow = spreadsheet.suiviSheet
		.createTextFinder("CONSOMMÉ EN JOURS DEV")
		.findNext()
		.getRow();
	let lastCol = spreadsheet.suiviSheet
		.getRange(firstRow, 1, 1, spreadsheet.suiviSheet.getLastColumn())
		.createTextFinder("TOTAL")
		.matchCase(true)
		.findNext()
		.getColumn();
	let lastRow = spreadsheet.suiviSheet
		.getRange(firstRow, 1, 100, 1)
		.createTextFinder("TOTAL")
		.findNext()
		.getRow();
	let sprintNameRange = spreadsheet.suiviSheet.getRange(
		firstRow,
		lastCol - 1,
		1,
		1
	);
	let sprintCountRange = spreadsheet.suiviSheet.getRange(
		firstRow + 3,
		lastCol - 1,
		lastRow - firstRow - 4,
		1
	);
	let sprintName = sprintNameRange.getValue().replace(/\s/g, "");
	ss.setNamedRange(sprintName, sprintNameRange);
	ss.setNamedRange(sprintName + "Count", sprintCountRange);
	spreadsheet.suiviSheet
		.getRange(lastRow, lastCol - 1, 1, 1)
		.setFormula(`=SUM(${sprintName}Count)`);
	sprintTable.props.table = sprintTable.props.table.offset(
		0,
		0,
		sprintTable.props.content.length,
		sprintTable.props.content[0].length + 1
	);
	sprintTable.props.content = sprintTable.props.table.getValues();
	sprintTable.updateCache();
	main.runDev_();
}

function addSprint() {
	addDevPeriod(true);
}

function addRecette() {
	addDevPeriod(false);
}

function addWeek() {
	let main = new Main();
	let spreadsheet = main.spreadsheet;
	let properties = main.properties;
	let weekTable = main.weekTable;

	let prevSprint = weekTable.props.table.offset(
		0,
		weekTable.props.table.getLastColumn() - 2,
		weekTable.props.lastRow - weekTable.props.firstRow + 1
	);

	prevSprint.copyTo(
		spreadsheet.suiviSheet.getRange(
			weekTable.props.firstRow,
			weekTable.props.lastCol,
			weekTable.props.lastRow - weekTable.props.firstRow + 1,
			1
		)
	);
	let added = spreadsheet.suiviSheet.getRange(
		weekTable.props.firstRow,
		weekTable.props.lastCol,
		weekTable.props.lastRow - weekTable.props.firstRow + 1,
		1
	);
	let prevWeekNumber = parseInt(
		prevSprint.offset(0, 0, 1).getValue().match(/(\d+)/)
	);
	let addedName = added.offset(0, 0, 1).getValue();
	let addedNumber = parseInt(addedName.match(/(\d+)/));
	addedName = addedName.replace(
		String(prevWeekNumber),
		String(addedNumber + 1)
	);
	added.offset(0, 0, 1).setValue(addedName);
	let weekNameRange = spreadsheet.suiviSheet.getRange(
		weekTable.props.firstRow,
		weekTable.props.lastCol,
		1,
		1
	);
	let weekCountRange = spreadsheet.suiviSheet.getRange(
		weekTable.props.firstRow + 3,
		weekTable.props.lastCol,
		weekTable.props.lastRow - weekTable.props.firstRow - 6,
		1
	);
	let weekName = weekNameRange.getValue().replace(/\s/g, "");
	spreadsheet.spreadsheet.setNamedRange(weekName + "Count", weekCountRange);
	spreadsheet.suiviSheet
		.getRange(weekTable.props.totalRow, weekTable.props.lastCol, 1, 1)
		.setFormula(`=SUM(${weekName}Count)`);
	spreadsheet.spreadsheet.setNamedRange(
		"semaineUX",
		spreadsheet.spreadsheet.getRangeByName("semaineUX").offset(0, 1)
	);
	weekTable.props.table = weekTable.props.table.offset(
		0,
		0,
		weekTable.props.content.length,
		weekTable.props.content[0].length + 1
	);
	weekTable.props.content = weekTable.props.table.getValues();
	weekTable.updateCache();
	main.runUX_();
}

function updateFromBacklog(formulas: string[][], increment) {
	formulas.forEach((formula, index) => {
		if (formula[0].match("'Product Backlog'!")) {
			let startRef = formula[0].indexOf("'Product Backlog'!");
			let length = "'Product Backlog'!".length;
			let firstLetter = startRef + length;
			while (!formula[0].charAt(length).match(/\d/)) {
				length++;
			}
			let lastLetter = length - 1;
			let colRef = formula[0].substr(firstLetter, lastLetter - firstLetter + 1);
			let colNum = letterToColumn(colRef);
			colNum += increment;
			let newCol = columnToLetter(colNum);
			formula[0] = formula[0].replace(
				"'Product Backlog'!" + colRef,
				"'Product Backlog'!" + newCol
			);
		}
	});
	return formulas;
}

function updateTable(sheet, textFinder, delimiter, recette?) {
	let firstRow = sheet.createTextFinder(textFinder).findNext().getRow();
	let lastCol = sheet
		.getRange(firstRow, 1, 1, sheet.getLastColumn())
		.createTextFinder("TOTAL")
		.matchCase(true)
		.findNext()
		.getColumn();
	let lastRow = sheet
		.getRange(firstRow, 1, 100, 1)
		.createTextFinder(delimiter)
		.findNext()
		.getRow();
	sheet
		.getRange(firstRow - 1, lastCol, lastRow - firstRow + 2, 1)
		.insertCells(SpreadsheetApp.Dimension.COLUMNS);
	if (
		!sheet
			.getRange(firstRow, lastCol - 1)
			.getValue()
			.match("recette")
	) {
		sheet
			.getRange(firstRow - 1, lastCol - 1, lastRow - firstRow + 2, 1)
			.copyTo(sheet.getRange(firstRow - 1, lastCol, lastRow - firstRow + 2, 1));
		let values = sheet
			.getRange(firstRow - 1, lastCol - 1, lastRow - firstRow + 2, 1)
			.getValues();
		let formulas = sheet
			.getRange(firstRow - 1, lastCol - 1, lastRow - firstRow + 2, 1)
			.getFormulas();
		if (textFinder === "CONSOMMÉ EN JOURS DEV") {
			formulas = updateFromBacklog(formulas, 3);
			sheet
				.getRange(firstRow - 1, lastCol, lastRow - firstRow + 2, 1)
				.setFormulas(formulas);
			let sprintName = sheet.getRange(firstRow, lastCol).getValue();
			if (typeof sprintName === "string" && sprintName.match("recette")) {
				let prevFormula = sheet.getRange(firstRow, lastCol).getFormula();
				let nextFormula = prevFormula.replace('"Sprint " & ', "");
				sheet.getRange(firstRow, lastCol).setFormula(nextFormula);
			}
		}
		formulas.forEach((formula, i) => {
			if (!formula[0]) {
				sheet
					.getRange(firstRow - 1 + i, lastCol, 1, 1)
					.setValue(values[i][0].toString());
			}
		});
	} else {
		sheet
			.getRange(firstRow - 1, lastCol - 2, lastRow - firstRow + 2, 1)
			.copyTo(sheet.getRange(firstRow - 1, lastCol, lastRow - firstRow + 2, 1));
		let values = sheet
			.getRange(firstRow - 1, lastCol - 2, lastRow - firstRow + 2, 1)
			.getValues();
		let formulas = sheet
			.getRange(firstRow - 1, lastCol - 2, lastRow - firstRow + 2, 1)
			.getFormulas();
		if (textFinder === "CONSOMMÉ EN JOURS DEV") {
			updateFromBacklog(formulas, 6);
			sheet
				.getRange(firstRow - 1, lastCol, lastRow - firstRow + 2, 1)
				.setFormulas(formulas);
			let sprintName = sheet.getRange(firstRow, lastCol).getValue();
			if (typeof sprintName === "string" && sprintName.match("recette"))
				sheet.getRange(firstRow, lastCol).setValue("recette");
		}
		formulas.forEach((formula, i) => {
			if (!formula[0]) {
				sheet
					.getRange(firstRow - 1 + i, lastCol, 1, 1)
					.setValue(values[i][0].toString());
			}
		});
	}
}
