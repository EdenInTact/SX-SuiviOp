export interface Coordinates {
	start: {
		col: number;
		row: number;
	};
	end: {
		col: number;
		row: number;
	};
}

export interface ConsultantList {
	list: Consultant[];
	content: Coordinates;
	length: number;
}

export interface Consultant {
	name: Coordinates;
	count: Coordinates;
	total: Coordinates;
}

export interface IterationList {
	list: Iteration[];
	length: number;
	content: Coordinates;
}

export interface Iteration {
	name: Coordinates;
	start: Coordinates;
	end: Coordinates;
	count: Coordinates;
	total: Coordinates;
}

export type Props = {
	corner: {
		top: {
			left: Spreadsheet.Range;
			right: Spreadsheet.Range;
		};
		bottom: {
			left: Spreadsheet.Range;
			right: Spreadsheet.Range;
		};
	};
	firstRow: number;
	lastRow: number;
	totalRow: number;
	firstCol: number;
	lastCol: number;
	table: Spreadsheet.Range;
	content: any[][];
	formulas: any[][];
	consultants: ConsultantList;
	iterations: IterationList;
	selection: Spreadsheet.Range;
	key: string;
};

export enum PROPS {
	corner_top_left,
}

export type Property = string | number | null | object;

export type TimeProps = {
	start: {
		day: number;
		month: number;
		year: number;
		full: string;
	};
	end: {
		day: number;
		month: number;
		year: number;
		full: string;
	};
};
