import Spreadsheet = GoogleAppsScript.Spreadsheet;
import AbstractTable from "../tables/AbstractTable";

export default class CRA {
	/**
	 * Open the Corresponding CRA and read datas for the dev
	 *
	 * @param { string[][] } cells , the datas from de CRA Nomenclature
	 * @param { number } row, the current row with the CRA values
	 * @param { string[] } keyDev, the dev Id
	 *
	 * */
	static openCra(
		cells: string[][],
		row: number,
		keyDev: string,
		table: AbstractTable
	) {
		//on recupere l'ID du CRA
		let ssCRA: Spreadsheet.Spreadsheet;
		let properties = table.properties;
		let spreadSheet = table.spreadsheet;
		let craValues: any[][];
		if (cells[row][3] != properties.cra_id) {
			properties.cra_id = cells[row][3];
		}
		ssCRA = SpreadsheetApp.openById(cells[row][3]);
		let name: string;
		let colWorker: number;
		const days: string[][] = ssCRA
			.getActiveSheet()
			.getRange("C1:C1000")
			.getValues();
		table.props.consultants.list.forEach((worker) => {
			if (table.props.content[worker.name.start.row][worker.name.start.col]) {
				name = table.props.content[worker.name.start.row][worker.name.start.col]
					.toString()
					.trim()
					.replace(/\s+/g, " ");
				// try {
				// 	colWorker = ssCRA.createTextFinder(name).findNext().getColumn();
				// } catch (e) {
				// 	SpreadsheetApp.getUi().alert(
				// 		`${name} introuvable dans le CRA,
				//     êtes vous sûr.e d'avoir écrit le nom tel qu'écrit dans le CRA ?`
				// 	);
				// 	return false;
				// }
				let dateStart = new Date(
					new Date(properties.time.start.full).setHours(0)
				);
				let dateEnd = new Date(new Date(properties.time.end.full).setHours(0));
				let startRow: number;
				let endRow: number;
				if (dateStart.getMonth() === dateEnd.getMonth()) {
					days.forEach((day, index) => {
						if (new Date(day[0]).toString() == dateStart.toString())
							startRow = index;
						if (new Date(day[0]).toString() == dateEnd.toString())
							endRow = index + 2;
					});
				} else if (
					dateStart.getMonth() + 1 < parseInt(cells[row][1]) ||
					dateStart.getFullYear() < parseInt(cells[row][0])
				) {
					startRow = 1;
					days.forEach((day, index) => {
						if (new Date(day[0]).toString() == dateEnd.toString())
							endRow = index + 2;
					});
				} else if (
					dateEnd.getMonth() + 1 > parseInt(cells[row][1]) ||
					dateEnd.getFullYear() > parseInt(cells[row][0])
				) {
					days.forEach((day, index) => {
						if (new Date(day[0]).toString() == dateStart.toString())
							startRow = index;
					});
					endRow = days.length;
				} else {
					Logger.log("ERROR");
				}
				let column: Spreadsheet.Range;
				column = ssCRA
					.getActiveSheet()
					.getRange(startRow, colWorker, endRow - startRow, 1); //toutes la colonne assossiée à la personne
				let count: number = 0;
				let i = 0;
				for (const columnValue of column.getValues()) {
					//foreach CRA key => retrieve the wanted one
					let dateValue = new Date(days[i][0]);
					if (columnValue.toString() == keyDev) {
						count++;
					}
					i++;
					if (dateValue > dateEnd) {
						break;
					}
				}
				let range = spreadSheet.suiviSheet.getRange(
					table.props.corner.top.left.getRow(),
					1,
					1,
					table.props.corner.top.right.getColumn()
				);
				let sprintCol = range
					.createTextFinder(table.props.selection.getValue())
					.findNext()
					.getColumn();
				table.props.consultants.list.forEach((dev, index) => {
					if (
						table.props.content[dev.name.start.row][dev.name.start.col] === name
					) {
						table.props.content[dev.count.start.row][sprintCol - 1] +=
							count / 2;
					}
				});
			}
		});
	}

	private static formatDateFull(date: Date): string {
		return `${date.getDate() < 10 ? "0" + date.getDate() : date.getDate()}/${
			date.getMonth() + 1 < 10
				? "0" + (date.getMonth() + 1)
				: date.getMonth() + 1
		}/${date.getFullYear()}`;
	}
	private static formatDateMonth(date: Date): string {
		return `${date.getDate()}/${date.getMonth() + 1}`;
	}
}
