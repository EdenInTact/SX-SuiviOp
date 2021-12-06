import Spreadsheet = GoogleAppsScript.Spreadsheet;
import CRA from "../utils/CRA";
import SpreadSheet from "../utils/SpreadSheet";
import ProjectProperties from "../utils/ProjectProperties";
import Main from "../../Actions";
import { Consultant, Iteration, Props } from "../../Types";

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
		for (let i = 0; i < this.props.consultants.length; i++) {
			let consultant: Consultant = {
				name: undefined,
				count: undefined,
				total: undefined,
			};
			consultant.name = {
				start: { col: 0, row: 3 + i },
				end: { col: 0, row: 3 + i },
			};
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

	//********* VALIDE ********/
	setNameDropdown(usersArray: string[], firstRow: number) {
		//Set name of worker in dropdown
		let namesRange = this.spreadsheet.suiviSheet.getRange(firstRow, 1, 3);
		const validationRule: Spreadsheet.DataValidation =
			SpreadsheetApp.newDataValidation().requireValueInList(usersArray).build();
		namesRange.setDataValidation(validationRule);
	}
	//***************************/
	setWorkDays(cells, code, code2) {
		// console.log("cells", cells, "code", code, "code2", code2);
		let range = this._spreadsheet.suiviSheet.getRange(
			this.props.corner.top.left.getRow(),
			1,
			1,
			this.props.corner.top.right.getColumn()
		);

		// console.log("range", range);
		let sprintCol = range
			.createTextFinder(this.props.selection.getValue())
			.findNext()
			.getColumn();

		console.log("sprintCol", sprintCol);

		this.props.consultants.list.forEach((dev) => {
			this.props.iterations.list.forEach((it) => {
				this.props.content[dev.count.start.row][sprintCol - 1] = 0;
			});
		});
		// for (let n = 0; n < cells.length; ++n) {
		// 	//if the month and year
		// 	if (cells[n][2] === code || cells[n][2] === code2) {
		// 		CRA.openCra(cells, n, this.props.key, this);
		// 	}
		// }
		let values = this.getRange(
			4,
			2,
			this.props.totalRow - this.props.firstRow,
			this.props.lastCol - 1
		);

		console.log("values", values);

		this.props.table
			.offset(3, 1, this.props.consultants.length, this.props.lastCol - 2)
			.setValues(values);
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
