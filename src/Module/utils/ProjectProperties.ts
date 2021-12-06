import GProperties = GoogleAppsScript.Properties;
import SpreadSheet from "./SpreadSheet";
import { Property, TimeProps } from "../../Types";
import API from "../../API/request";

export default class ProjectProperties {
	// ****************************************************************************************** \\
	//                                           PROPERTIES                                       \\
	// ****************************************************************************************** \\

	private _script: GProperties.Properties;
	private _spreadsheet: SpreadSheet;

	private _keyDev: Property;
	private _keyUX: Property;

	public _project: Property;
	public _phases: Property;
	public _activity: Property;
	public _projectrow: Property;

	private _time: TimeProps;

	private _dropdown: Property;

	private _cra_id: Property;
	private _startRow: Property;
	private _endRow: Property;

	// ****************************************************************************************** \\
	//                                          CONSTRUCTOR                                       \\
	// ****************************************************************************************** \\

	constructor(spreadsheet: SpreadSheet) {
		this._script = PropertiesService.getScriptProperties();
		this._spreadsheet = spreadsheet;

		this._time = JSON.parse(this._script.getProperty("time"));

		this.keyDev = this._spreadsheet.spreadsheet
			.getRangeByName("clefCraDev")
			.getValue();
		this.keyUX = this._spreadsheet.spreadsheet
			.getRangeByName("clefCraUX")
			.getValue();
		this._dropdown = this._script.getProperty("dropdown");
		this.cra_id = this._script.getProperty("cra_id");
		this.startRow = this._script.getProperty("startRow");
		this.endRow = this._script.getProperty("endRow");
		this._phases = this._script.getProperty("phases");
		this._project = this._script.getProperty("project");
		this._activity = this._script.getProperty("activity");
		this._projectrow = this._script.getProperty("projectrow");
	}

	// ****************************************************************************************** \\
	//                                             ACCESSORS                                      \\
	// ****************************************************************************************** \\

	set projectrow(value: Property) {
		this._projectrow = value;
		if (value) this._script.setProperty("projectrow", value.toString());
	}

	get projectrow(): Property {
		return this._projectrow;
	}

	set activity(value: Property) {
		this._activity = value;
		if (value) this._script.setProperty("projectrow", value.toString());
	}

	get activity(): Property {
		return this._activity;
	}

	set phases(value: Property) {
		this._phases = value;
		if (value) this._script.setProperty("phases", value.toString());
	}

	get phases(): Property {
		return this._phases;
	}

	set project(value: Property) {
		this._project = value;
		if (value) this._script.setProperty("project", value.toString());
	}

	get project(): Property {
		return this._project;
	}

	get keyDev(): Property {
		return this._keyDev;
	}

	set keyDev(value: Property) {
		this._keyDev = value;
		if (value) this._script.setProperty("keyDev", value.toString());
	}

	get keyUX(): Property {
		return this._keyUX;
	}

	set keyUX(value: Property) {
		this._keyUX = value;
		if (value) this._script.setProperty("keyUX", value.toString());
	}

	get dropdown(): Property {
		return this._dropdown;
	}

	set dropdown(value: Property) {
		this._dropdown = value;
		if (value) this._script.setProperty("dropdown", value.toString());
	}

	deleteKeyDev() {
		this.keyDev = null;
		this._script.deleteProperty("keyDev");
	}

	get time(): TimeProps {
		return this._time;
	}

	set time(value: TimeProps) {
		this._time = value;
		if (value) this._script.setProperty("time", JSON.stringify(value));
	}

	get cra_id() {
		return this._cra_id;
	}

	set cra_id(value) {
		this._script.deleteProperty("cra_id");
		this._cra_id = value;
		if (value) this._script.setProperty("cra_id", JSON.stringify(value));
	}

	get startRow(): Property {
		return this._startRow;
	}

	set startRow(value: Property) {
		this._startRow = value;
		if (value) this._script.setProperty("startRow", value.toString());
	}

	get endRow(): Property {
		return this._endRow;
	}

	set endRow(value: Property) {
		this._endRow = value;
		if (value) this._script.setProperty("endRow", value.toString());
	}

	set dateStart(date: Date) {
		this.time.start = {
			day: date.getDate(),
			month: date.getMonth() + 1,
			year: date.getFullYear(),
			full: date.toString(),
		};
	}

	set dateEnd(date: Date) {
		this.time.end = {
			day: date.getDate(),
			month: date.getMonth() + 1,
			year: date.getFullYear(),
			full: date.toString(),
		};
	}
}
