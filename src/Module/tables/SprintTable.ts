import AbstractTable from "./AbstractTable";
import { ConsultantList, Props } from "../../Types";
import Main from "../../Actions";
import Spreadsheet = GoogleAppsScript.Spreadsheet;

export default class SprintTable extends AbstractTable {
	// ****************************************************************************************** \\
	//                                          CONSTRUCTOR                                       \\
	// ****************************************************************************************** \\

	constructor(main: Main) {
		let cache: Props;
		super(main, cache, "sprintTable");
		this.init("CONSOMMÉ EN JOURS DEV", "TOTAL", "sprintDev");
		this.props.key = main.properties.keyDev as string;
		CacheService.getScriptCache().put(
			"sprintTable",
			JSON.stringify(this.cache)
		);
	}

	// ****************************************************************************************** \\
	//                                             ACCESSORS                                      \\
	// ****************************************************************************************** \\

	get selectedSprint(): Spreadsheet.Range {
		return this.props.selection;
	}

	set selectedSprint(value: Spreadsheet.Range) {
		this.props.selection = value;
	}

	get dev(): ConsultantList {
		return this.props.consultants;
	}

	set dev(value: ConsultantList) {
		this.props.consultants = value;
	}

	setWorkDays(cells, code, code2) {
		super.setWorkDays(cells, code, code2);
	}
}
