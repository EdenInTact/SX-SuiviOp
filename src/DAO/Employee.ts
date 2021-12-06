export default class Employee {
	// ****************************************************************************************** \\
	//                                           PROPERTIES                                       \\
	// ****************************************************************************************** \\
	private _name: string;
	private _halfDays: number[] = [];

	// ****************************************************************************************** \\
	//                                          CONSTRUCTOR                                       \\
	// ****************************************************************************************** \\
	constructor(name: string = null, halfDays: number[] = []) {
		this.name = name;
		this.halfDays.forEach((sprint, index) => {
			if (sprint) {
				this.halfDays[index] = sprint;
			}
		});
	}

	// ****************************************************************************************** \\
	//                                             ACCESSORS                                      \\
	// ****************************************************************************************** \\
	get name(): string {
		return this._name;
	}

	set name(value: string) {
		this._name = value;
	}

	get halfDays(): number[] {
		return this._halfDays;
	}

	set halfDays(value: number[]) {
		this._halfDays = value;
	}

	setSprintHalfDays(value: number, index: number) {
		this._halfDays[index] = value;
	}
}
