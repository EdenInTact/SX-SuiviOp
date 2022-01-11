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
			let project = resultAPI.code;
			let phases = resultAPI.projectPhases;

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
		} else if (e) {
			// 1. set sprint and week time start and end
			let lastChange = e.range.getA1Notation();
			let sprintDev = this.spreadsheet.spreadsheet
				.getRangeByName("sprintDev")
				.getA1Notation();
			let semaineUX = this.spreadsheet.spreadsheet
				.getRangeByName("semaineUX")
				.getA1Notation();

			// 2. MAJ sprint ou semaine
			if (lastChange === sprintDev) {
				this._isSprint = true;
				let col = 0;
				this.sprintTable.props.iterations.list.forEach((sprint, index) => {
					if (
						e.range.getValue() ===
						this.sprintTable.props.content[sprint.name.start.row][
							sprint.name.start.col
						]
					) {
						col = index;
					}
				});
				sprintcellEnd =
					this._sprintTable.props.content[
						this.sprintTable.props.iterations.list[col].end.start.row
					][this.sprintTable.props.iterations.list[col].end.start.col];
				sprintcellStart =
					this._sprintTable.props.content[
						this.sprintTable.props.iterations.list[col].start.start.row
					][this.sprintTable.props.iterations.list[col].start.start.col];
			} else if (lastChange === semaineUX) {
				this._weekTable.selectedWeek = e.range;
				this._isSprint = false;
				let col = 0;
				this.weekTable.props.iterations.list.forEach((week, index) => {
					if (
						e.range.getValue() ===
						this.weekTable.props.content[week.name.start.row][
							week.name.start.col
						]
					) {
						col = index;
					}
				});
				sprintcellEnd =
					this._weekTable.props.content[
						this.weekTable.props.iterations.list[col].end.start.row
					][this.weekTable.props.iterations.list[col].end.start.col];

				sprintcellStart =
					this._weekTable.props.content[
						this.weekTable.props.iterations.list[col].start.start.row
					][this.weekTable.props.iterations.list[col].start.start.col];
			}

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

	setUserTimesheetRow() {
		let api = new API();
		let main = new Main();

		let resultAPI: any = api.getUsersTimesheetRow(
			JSON.parse(this._properties.projectrow)
		);

		//1.  set users dropdown
		let responseUser = resultAPI.userArray.users;
		let usersArrFrmted = Object.keys(responseUser).map((key) => {
			return `${responseUser[key].name} - ${responseUser[key].function}`;
		});

		main.weekTable.setNameDropdown(usersArrFrmted);
		main.sprintTable.setNameDropdown(usersArrFrmted);

		//2. merge and save time sheet row on script proprieties
		let responseTimesheetRow = resultAPI?.timesheetRow["hydra:member"];
		
		let timesheetRowArr = [];
		responseTimesheetRow.map((each) => {
			each.timesheetRow.map((item) => {
				timesheetRowArr.push(item);
			});
		});

		return PropertiesService.getScriptProperties().setProperty(
			"projectrow",
			JSON.stringify(timesheetRowArr)
		);
	}

	/**
	 * Retrieve corresponding SX in order to estimate days passed on project.
	 * */
	openSXNomenc() {
		let main = new Main();

		// retrieve year and month from properties
		const timeStart =
			this._properties.time.start.year +
			"" +
			("0" + this._properties.time.start.month).slice(-2) +
			"" +
			("0" + this._properties.time.start.day).slice(-2);

		let timeEnd =
			this._properties.time.end.year +
			"" +
			("0" + this._properties.time.end.month).slice(-2) +
			"" +
			("0" + this._properties.time.end.day).slice(-2);
		// }
		let timesheetRow = main._properties.projectrow;

		if (this._isSprint) {
			this._sprintTable.setWorkDays(timesheetRow, timeStart, timeEnd);
		} else {
			this._weekTable.setWorkDays(timesheetRow, timeStart, timeEnd);
		}
	}

	 cleanTable(){
		let main = new Main();
		let spreadsheet = main.spreadsheet;
	
		let firstRow = spreadsheet.suiviSheet
		.createTextFinder("CONSOMMÉ EN JOURS DEV")
		.findNext()
		.getRow();

		let lastRow = spreadsheet.suiviSheet
		.getRange(firstRow, 1, 100, 1)
		.createTextFinder("TOTAL")
		.findNext()
		.getRow();
	
		console.log('lastRow', lastRow)
		let lastColData = spreadsheet.dataSheet.getLastColumn();
	
		let range = spreadsheet.suiviSheet.getRange(
			firstRow + 3,
			2,
			lastRow -firstRow - 4,
			lastColData-2
		);
		range.setValue(null)
	}
}
