export default class API {
  private _username: string = "app_script";
  private _password: string = "seE@^@EXvvSziqNfwD*8K!FbYh7rzN";
  private _apiPath: string = "https://sxintact.atecna.fr/";
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

  getUsersTimesheetRow(activityId): object {
    this.authenticate();
    let options: any = {
      method: "get",
      headers: { Authorization: `Bearer ${this._token}` },
    };

    // let resp = UrlFetchApp.fetch(
    // 	`${this._apiPath}api/project_activities/${activityId}/project_rows`,
    // 	options
    // );

    let userArray = UrlFetchApp.fetch(
      `${this._apiPath}api/project_activities/${activityId}/users`,
      options
    );

    return {
      // timesheetRow: JSON.parse(resp.toString()),
      userArray: JSON.parse(userArray.toString()),
    };
  }

  getUsersByProject(projectCode): object {
    this.authenticate();
    let options: any = {
      method: "get",
      headers: { Authorization: `Bearer ${this._token}` },
    };

    let userArray = UrlFetchApp.fetch(
      `${this._apiPath}api/projects/code/${projectCode}/users`,
      options
    );

    return { userArray: JSON.parse(userArray.toString()) };
  }

  getProjectByCode(codeProject): any {
    this.authenticate();
    let options: any = {
      method: "get",
      headers: { Authorization: `Bearer ${this._token}` },
    };
    let response = UrlFetchApp.fetch(
      `${this._apiPath}api/projects/code/${codeProject}`,
      options
    );
    let result = JSON.parse(response.getContentText());
    return result;
  }

  getAllTimeSheetRow(
    codeProject: string,
    date_debut: string,
    date_fin: string
  ) {
    this.authenticate();
    let options: any = {
      method: "get",
      headers: { Authorization: `Bearer ${this._token}` },
    };
    let response = UrlFetchApp.fetch(
      `${this._apiPath}api/projects/code/${codeProject}/timesheet_rows?dateTime%5Bbefore%5D=${date_fin}&dateTime%5Bafter%5D=${date_debut}`,
      options
    );

    let result = JSON.parse(response.getContentText());

    return result;
  }
}
