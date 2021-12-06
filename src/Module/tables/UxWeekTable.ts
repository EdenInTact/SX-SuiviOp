import AbstractTable from "./AbstractTable";
import { ConsultantList, Props } from "../../Types";
import Main from "../../Actions";
import Spreadsheet = GoogleAppsScript.Spreadsheet;

export default class UxWeekTable extends AbstractTable {
	// ****************************************************************************************** \\
	//                                          CONSTRUCTOR                                       \\
	// ****************************************************************************************** \\

	constructor(main: Main) {
		let cache: Props;
		super(main, cache, "uxTable");
		// this.init("CONSOMMÉ EN JOURS UX/UI", "Nb de jours restants", "semaineUX");
		this.props.key = main.properties.keyUX as string;
		CacheService.getScriptCache().put("uxTable", JSON.stringify(this.props));
	}

	// ****************************************************************************************** \\
	//                                             ACCESSORS                                      \\
	// ****************************************************************************************** \\

	get selectedWeek(): Spreadsheet.Range {
		return this.props.selection;
	}

	set selectedWeek(value: Spreadsheet.Range) {
		this.props.selection = value;
	}

	get ux(): ConsultantList {
		return this.props.consultants;
	}

	set ux(value: ConsultantList) {
		this.props.consultants = value;
	}

	setWorkDays(cells, code, code2) {
		super.setWorkDays(cells, code, code2);
	}
}
