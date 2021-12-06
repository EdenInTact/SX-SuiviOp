export default class API {
	private _username: string = "app_script";
	private _password: string = "2*fKv4q8sp^rxgRz&FEBf9otD%vL##";
	private _apiPath: string = "https://sx-intact.atecna.fr/";
	private _token: string;

	authenticate(): string {
		let payload = { username: this._username, password: this._password };
		let options: object = {
			method: "post",
			contentType: "application/json",
			payload: JSON.stringify(payload),
			muteHttpExceptions: true,
		};

		let url = `${this._apiPath}authentication_token`;
		let response = UrlFetchApp.fetch(url, options);
		this._token = JSON.parse(response.getContentText()).token;
		return this._token;
	}

	getUsers(arrayTimerow): object {
		this.authenticate();
		let options: any = {
			method: "get",
			headers: { Authorization: `Bearer ${this._token}` },
		};
		let usersArray = [];

		arrayTimerow.map((each) => {
			let resp = UrlFetchApp.fetch(
				`${this._apiPath}api/project_rows/${each.id}`,
				options
			);
			usersArray.push(JSON.parse(resp));
		});

		return usersArray;
	}

	getProjectByCode(codeProject): any {
		this.authenticate();
		let options: any = {
			method: "get",
			headers: { Authorization: `Bearer ${this._token}` },
		};
		let response = UrlFetchApp.fetch(
			`${this._apiPath}api/projects?code=${codeProject}`,
			options
		);

		let result = JSON.parse(response.getContentText());

		return result;
	}
}
