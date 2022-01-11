import Spreadsheet = GoogleAppsScript.Spreadsheet;
import SpreadSheet from "../utils/SpreadSheet";
import ProjectProperties from "../utils/ProjectProperties";
import Main from "../../Actions";
import { Consultant, Iteration, Props } from "../../Types";
let ss = SpreadsheetApp.getActive();

export default abstract class AbstractTable {
	// ****************************************************************************************** \\
	//                                           PROPERTIES                                       \\
	// ****************************************************************************************** \\

	private _main: Main;
	private _spreadsheet: SpreadSheet;
	private _properties: ProjectProperties;
	private _cache: Props;
	private _props: Props;
	private _tableName: string;

	// ****************************************************************************************** \\
	//                                          CONSTRUCTOR                                       \\
	// ****************************************************************************************** \\

	protected constructor(main: Main, cache: Props, tableName: string) {
		this._main = main;
		this._spreadsheet = main.spreadsheet;
		this._properties = main.properties;
		this._tableName = tableName;
		this.props = {
			consultants: {
				length: 0,
				content: { start: { col: 0, row: 0 }, end: { col: 0, row: 0 } },
				list: [],
			},
			content: [],
			formulas: [],
			corner: {
				top: { left: undefined, right: undefined },
				bottom: { left: undefined, right: undefined },
			},
			firstCol: 0,
			firstRow: 0,
			iterations: {
				length: 0,
				content: { start: { col: 0, row: 0 }, end: { col: 0, row: 0 } },
				list: [],
			},
			key: "",
			lastCol: 0,
			lastRow: 0,
			selection: undefined,
			table: undefined,
			totalRow: 0,
		};
	}

	// ****************************************************************************************** \\
	//                                             ACCESSORS                                      \\
	// ****************************************************************************************** \\

	get cache(): Props {
		return this._cache;
	}

	get props(): Props {
		return this._props;
	}

	set props(value: Props) {
		this._props = value;
		this.updateCache();
	}

	get tableName(): string {
		return this._tableName;
	}

	set tableName(value: string) {
		this._tableName = value;
	}

	get main(): Main {
		return this._main;
	}

	get spreadsheet(): SpreadSheet {
		return this._spreadsheet;
	}

	get properties(): ProjectProperties {
		return this._properties;
	}

	// ****************************************************************************************** \\
	//                                            METHODS                                         \\
	// ****************************************************************************************** \\

	updateCache() {
		CacheService.getScriptCache().put(
			this._tableName,
			JSON.stringify(this.props)
		);
	}

	init(topLeft: string, bottomLeft: string, selection: string) {
		this.props.corner.top.left = this._spreadsheet.spreadsheet
			.createTextFinder(topLeft)
			.findNext();
		this.props.corner.top.right = this.props.corner.top.left
			.offset(0, 0, 1, this._spreadsheet.suiviSheet.getLastColumn())
			.createTextFinder("TOTAL")
			.findNext();
		this.props.corner.bottom.left = this._spreadsheet.suiviSheet
			.getRange(
				this.props.corner.top.left.getRow(),
				this.props.corner.top.left.getColumn(),
				this._spreadsheet.suiviSheet.getLastRow(),
				1
			)
			.createTextFinder(bottomLeft)
			.findNext();
		this.props.corner.bottom.right = this._spreadsheet.suiviSheet.getRange(
			this.props.corner.bottom.left.getRow(),
			this.props.corner.top.right.getColumn()
		);
		this.props.firstRow = this.props.corner.top.left.getRow();
		this.props.lastRow = this.props.corner.bottom.left.getRow();
		this.props.totalRow = this._spreadsheet.suiviSheet
			.getRange(
				this.props.corner.top.left.getRow(),
				1,
				this.props.lastRow - this.props.firstRow + 1
			)
			.createTextFinder("TOTAL")
			.findNext()
			.getRow();

		this.props.firstCol = 1;
		this.props.lastCol = this.props.corner.top.right.getColumn();

		this.props.table = this.spreadsheet.suiviSheet.getRange(
			this.props.firstRow,
			this.props.firstCol,
			this.props.totalRow - this.props.firstRow + 1,
			this.props.lastCol
		);
		this.props.content = this.props.table.getValues();
		this.props.formulas = this.props.table.getFormulas();
		this.props.selection =
			this._spreadsheet.spreadsheet.getRangeByName(selection);

		this.props.consultants.length =
			this.props.totalRow - (this.props.firstRow + 3);
		this.props.consultants.content = {
			start: { col: 0, row: 3 },
			end: { col: -1, row: this.props.totalRow - this.props.firstRow },
		};

		//retrieve coord of name of consultant dropdow list
		for (let i = 0; i < this.props.consultants.length; i++) {
			let consultant: Consultant = {
				name: undefined,
				count: undefined,
				total: undefined,
			};
			consultant.name = { col: 1, row: 3 + i };
			consultant.count = {
				start: { col: 1, row: 3 + i },
				end: { col: this.props.content[0].length - 2, row: 3 + i },
			};
			consultant.total = {
				start: { col: this.props.content[0].length - 1, row: 3 + i },
				end: { col: this.props.content[0].length - 1, row: 3 + i },
			};
			this.props.consultants.list.push(consultant);
		}
		this.props.iterations.length = this.props.lastCol - 2;
		this.props.iterations.content = {
			start: { col: 1, row: 0 },
			end: { col: -2, row: this.props.totalRow - this.props.firstRow },
		};

		//retrieve coord of sprint in table
		for (let i = 0; i < this.props.iterations.length; i++) {
			let iteration: Iteration = {
				name: undefined,
				start: undefined,
				end: undefined,
				count: undefined,
				total: undefined,
			};
			iteration.name = {
				start: { col: 1 + i, row: 0 },
				end: { col: 1 + i, row: 0 },
			};
			iteration.start = {
				start: { col: 1 + i, row: 1 },
				end: { col: 1 + i, row: 1 },
			};
			iteration.end = {
				start: { col: 1 + i, row: 2 },
				end: { col: 1 + i, row: 2 },
			};
			iteration.count = {
				start: { col: 1 + i, row: 3 },
				end: { col: 1 + i, row: 3 + this.props.consultants.list.length - 1 },
			};
			iteration.total = {
				start: { col: 1 + i, row: 3 + this.props.consultants.list.length },
				end: { col: 1 + i, row: 3 + this.props.consultants.list.length },
			};
			this.props.iterations.list.push(iteration);
		}
		this._cache = this.props;
		this.updateCache();
	}

	setNameDropdown(usersArray: string[] | "") {
		let namesRange = this.spreadsheet.suiviSheet.getRange(
			this.props.firstRow + 3,
			1,
			this.props.consultants.length
		);
		if (usersArray !== "") {
			//Set name of worker in dropdown
			const validationRule: Spreadsheet.DataValidation =
				SpreadsheetApp.newDataValidation()
					.requireValueInList(usersArray)
					.build();
			namesRange.setDataValidation(validationRule);
		} else {
			namesRange.clearContent();
			namesRange.setDataValidation(null);
		}
	}

	setWorkDays(timesheetRow: string, timeStart: string, timeEnd: string) {
		let consultantArray = [];
		//1. get row and name of each consultant
		this.props.consultants.list.map((each) => {
			var eachConsultant = this.spreadsheet.suiviSheet
				.getRange(
					this.props.corner.top.left.getRow() + each.name.row,
					each.name.col
				)
				.getValue();
			if (eachConsultant) {
				consultantArray.push({
					user: eachConsultant,
					row: this.props.corner.top.left.getRow() + each.name.row,
				});
			}
		});

		//2. get column of sprint
		let range = this._spreadsheet.suiviSheet.getRange(
			this.props.corner.top.left.getRow(),
			1,
			1,
			this.props.corner.top.right.getColumn()
		);

		let sprintCol = range
			.createTextFinder(this.props.selection.getValue())
			.findNext()
			.getColumn();

		//3. filter timesheet row to keep only who's between start and end time
		let workerTimeArray = [];
		let timesheeRowArr = JSON.parse(timesheetRow);

		let timesheetFltr = timesheeRowArr.filter((item) => {
			let dateTime = item.dateTime.split("-").join("");
			return dateTime <= timeEnd && dateTime >= timeStart;
		});

		//4. create a array with row, user name and qty worked
		consultantArray.forEach((consultant) => {
			let userName = consultant.user.split("-")[0].toString();
			let workertimesheet = [];
			timesheetFltr.map((timesheet) => {
				if (timesheet.userName.indexOf(userName)) {
					workertimesheet.push({
						row: consultant.row,
						qty: timesheet.qty,
						user: userName,
						date: timesheet.dateTime,
					});
				}
			});

			workertimesheet.length > 0
				? workerTimeArray.push(workertimesheet)
				: workerTimeArray.push([
						{ row: consultant.row, qty: 0, user: userName },
				  ]);
		});


		//5. setvalue of day worked for each 
		workerTimeArray.forEach((workerInfo) => {
			let qty = 0;
			let rowWorker;
			workerInfo.map((item) => {
				qty += item.qty;
				rowWorker = item.row;
			});
			let dayVal = qty / 8;

			ss.getSheetByName("Suivi opérationnel")
				.getRange(rowWorker, sprintCol)
				.setValue(dayVal);
		});
	}

	//*************** Mathematique function ***********/
	getRow(i: number) {
		return this.props.content[i];
	}

	getColumn(i: number) {
		let col: any[] = [];
		for (let j = 0; j < this.props.content.length - 3; j++) {
			col.push(this.props.content[j][i]);
		}
		return col;
	}

	getCell(row: number, col: number) {
		return this.props.content[row][col];
	}

	setCell(row: number, col: number, value: any) {
		this.props.content[row][col] = value;
	}

	getRange(startRow: number, startCol: number, endRow: number, endCol: number) {
		let range: any[][] = [];
		for (let i = startRow - 1; i <= endRow - 1; i++) {
			let row: any[] = [];
			for (let j = startCol - 1; j <= endCol - 1; j++) {
				row.push(this.props.content[i][j]);
			}
			range.push(row);
		}
		return range;
	}
}
