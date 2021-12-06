import Spreadsheet = GoogleAppsScript.Spreadsheet;
import Base = GoogleAppsScript.Base;
import SprintTable from "../Module/tables/SprintTable";
import UxWeekTable from "../Module/tables/UxWeekTable";
import SpreadSheet from "../Module/utils/SpreadSheet";
import ProjectProperties from "../Module/utils/ProjectProperties";
import API from "../API/request";

var ss = SpreadsheetApp.getActive();

export default class Main {
	// ****************************************************************************************** \\
	//                                           PROPERTIES                                       \\
	// ****************************************************************************************** \\
	private _ui: Base.Ui;
	public _properties: ProjectProperties;
	private _sprintTable: SprintTable;
	private _weekTable: UxWeekTable;
	private _isSprint: boolean;
	private _spreadsheet: SpreadSheet;

	// ****************************************************************************************** \\
	//                                          CONSTRUCTOR                                       \\
	// ****************************************************************************************** \\
	constructor(e?: any) {
		this.spreadsheet = new SpreadSheet();
		this._properties = new ProjectProperties(this.spreadsheet);
		this._ui = SpreadsheetApp.getUi();
		this._sprintTable = new SprintTable(this);
		this._weekTable = new UxWeekTable(this);
	}

	// ****************************************************************************************** \\
	//                                             ACCESSORS                                      \\
	// ****************************************************************************************** \\
	get spreadsheet(): SpreadSheet {
		return this._spreadsheet;
	}
	set spreadsheet(value: SpreadSheet) {
		this._spreadsheet = value;
	}
	get ui(): GoogleAppsScript.Base.Ui {
		return this._ui;
	}
	set ui(value: GoogleAppsScript.Base.Ui) {
		this._ui = value;
	}
	get properties(): ProjectProperties {
		return this._properties;
	}
	get sprintTable(): SprintTable {
		return this._sprintTable;
	}
	get weekTable(): UxWeekTable {
		return this._weekTable;
	}
	get isSprint(): boolean {
		return this._isSprint;
	}

	runUX_(): void {
		let range: string[] = this.weekTable.props.content[0].slice(1, -1);

		const activeCell: Spreadsheet.Range =
			this._spreadsheet.spreadsheet.getRangeByName("semaineUX");

		const validationRule: Spreadsheet.DataValidation =
			SpreadsheetApp.newDataValidation().requireValueInList(range).build();
		activeCell.setDataValidation(validationRule);
	}

	/**
	 * set the sprint dropdown in "Suivi Opérationelle" Sheet dynamically based on "itération" Row in "Data" Sheet
	 * */
	runDev_(): void {
		let range: string[] = this.sprintTable.props.content[0].slice(1, -1);
		const activeCell: Spreadsheet.Range =
			this._spreadsheet.spreadsheet.getRangeByName("sprintDev");
		let numRecette = 0;
		range.forEach((value, index) => {
			if (value === "recette") {
				numRecette++;
				if (numRecette > 1) {
					value = "recette" + numRecette;
					range[index] = value;
				}
			}
		});
		const validationRule: Spreadsheet.DataValidation =
			SpreadsheetApp.newDataValidation().requireValueInList(range).build();
		activeCell.setDataValidation(validationRule);
	}

	setProperties(e?: { range: Spreadsheet.Range }) {
		let api = new API();
		let sprintcellEnd: number | string;
		let sprintcellStart: number | string;
		if (!e) {
			// 1. get project info by his code
			let resultAPI = api.getProjectByCode(this._properties.keyDev);

			// 2. set phase and project property
			let project = resultAPI["hydra:member"] && resultAPI["hydra:member"][0];
			let phases = project?.projectPhases;

			this._properties._project = project;
			this._properties._phases = phases;

			PropertiesService.getScriptProperties().setProperty(
				"project",
				JSON.stringify(project)
			);
			PropertiesService.getScriptProperties().setProperty(
				"phases",
				JSON.stringify(phases)
			);

			return;
		} else if (e) {
			// 1. set sprint and week time start and end
			let lastChange = e.range.getA1Notation();
			let sprintDev = this.spreadsheet.spreadsheet
				.getRangeByName("sprintDev")
				.getA1Notation();
			let semaineUX = this.spreadsheet.spreadsheet
				.getRangeByName("semaineUX")
				.getA1Notation();

			console.log(lastChange, sprintDev, semaineUX);

			// 2. MAJ sprint ou semaine
			if (lastChange === sprintDev) {
				sprintcellStart = ss.getRange("B22").getValue();
				console.log("sprintcellStart === >", sprintcellStart);

				sprintcellEnd = ss.getRange("G23").getValue();
				console.log("sprintcellEnd === >", sprintcellEnd);
			}

			//TODO a faire pour les semaines UX
			// else if (lastChange === semaineUX) {
			// 	this._weekTable.selectedWeek = e.range;
			// 	this._isSprint = false;
			// 	let col = 0;
			// 	this.weekTable.props.iterations.list.forEach((week, index) => {
			// 		if (
			// 			e.range.getValue() ===
			// 			this.weekTable.props.content[week.name.start.row][
			// 				week.name.start.col
			// 			]
			// 		) {
			// 			col = index;
			// 		}
			// 	});

			// 	sprintcellEnd =
			// 		this._weekTable.props.content[
			// 			this.weekTable.props.iterations.list[col].end.start.row
			// 		][this.weekTable.props.iterations.list[col].end.start.col];

			// 	sprintcellStart =
			// 		this._weekTable.props.content[
			// 			this.weekTable.props.iterations.list[col].start.start.row
			// 		][this.weekTable.props.iterations.list[col].start.start.col];
			// }

			//*********** Set Property **********//
			this._properties.dropdown = e.range.getA1Notation();

			this._properties.time = {
				end: { day: 0, full: "", month: 0, year: 0 },
				start: { day: 0, full: "", month: 0, year: 0 },
			};

			this._properties.dateStart = new Date(sprintcellStart);
			this._properties.dateEnd = new Date(sprintcellEnd);
		}
	}

	setUser() {
		let api = new API();
		let main = new Main();

		let resultAPI: any = api.getUsers(JSON.parse(this._properties.projectrow));
		let users = [];

		console.log("result APi", JSON.stringify(resultAPI, null, 2));
		resultAPI.map((each) => {
			each.timesheetRow.map((item) => {
				users.push(item.user.toString());
			});
		});

		let filteredUsersArray = users.filter(
			(ele, pos) => users.indexOf(ele) === pos
		);
		main.weekTable.setNameDropdown(filteredUsersArray, 13);
		main.sprintTable.setNameDropdown(filteredUsersArray, 24);
	}

	/**
	 * Retrieve corresponding SX in order to estimate days passed on project.
	 * */
	openSXNomenc() {
		// retrieve year and month from properties
		const timeStart: string = (
			this._properties.time.start.year + this._properties.time.start.month
		).toString();

		// if the sprint is on two different months, retrieve the second CRA
		let timeEnd: string;
		if (this._properties.time.end.month && this._properties.time.end.year) {
			timeEnd = (
				this._properties.time.end.year + this._properties.time.end.month
			).toString();
		}
		console.log("timeStart", timeStart);
		console.log("timeEnd", timeEnd);

		// let timesheetRow =
		// if (this._isSprint) {
		// 	this._sprintTable.setWorkDays(craDateID, timeStart, timeEnd);
		// } else {
		// 	this._weekTable.setWorkDays(craDateID, timeStart, timeEnd);
		// }
	}

	// openCraNomenc() {
	//retrieve year and month from properties
	// const timeStart: string =
	// 	this._properties.time.start.year + "" + this._properties.time.start.month;
	// // if the sprint is on two different months, retrieve the second CRA
	// let timeEnd: string;
	// if (this._properties.time.end.month && this._properties.time.end.year) {
	// 	timeEnd =
	// 		this._properties.time.end.year + "" + this._properties.time.end.month;
	// }
	// console.log("code", timeStart);
	// console.log("code2", timeEnd);
	// //The nomenclature file ID listed all CRA files -> ID is 19dNWRaX3ycWAYm-ftmjuRIJ6uQ1UUZ7bx_nJYhQHTIA
	// const ss: Spreadsheet.Spreadsheet = SpreadsheetApp.openById(
	// 	"19dNWRaX3ycWAYm-ftmjuRIJ6uQ1UUZ7bx_nJYhQHTIA"
	// );
	// const range: Spreadsheet.Range = ss.getActiveSheet().getDataRange();
	// const craDateID: string[][] = range.getValues();
	// // console.log("cells", JSON.stringify(craDateID, null, 3));
	// if (this._isSprint) {
	// 	this._sprintTable.setWorkDays(craDateID, timeStart, timeEnd);
	// } else {
	// 	this._weekTable.setWorkDays(craDateID, timeStart, timeEnd);
	// }
	// }
}
