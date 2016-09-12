import { IPerson } from '../interfaces/IPerson';
import { ServiceScope, HttpClient, IODataBatchOptions, ODataBatch } from '@microsoft/sp-client-base';

export interface IUserProfileService {
  getPropertiesForCurrentUSer: () => Promise<IPerson>;
  getPropertiesForUsers(userLoginNames: string[]): Promise<IPerson[]>
}

export class UserProfileService {

  private httpClient: HttpClient;
  private serviceScope: ServiceScope;

  constructor(serviceScope: ServiceScope) {
    this.httpClient = new HttpClient(serviceScope);
    this.serviceScope = serviceScope;
  }

  private getPropertiesForCurrentUSer(): Promise<IPerson> {
    return this.httpClient.get(
      `/_api/SP.UserProfiles.PeopleManager/GetMyProperties?$select=DisplayName,Title,PersonalUrl,PictureUrl,DirectReports,ExtendedManagers`)
      .then((response: Response) => {
        return response.json();
      });
  }

  private getPropertiesForUsers(userLoginNames: string[]): Promise<IPerson[]> {
    return new Promise<IPerson[]>((resolve, reject) => {

      let arrayPersons: IPerson[] = [];

      const batchOpts: IODataBatchOptions = {};

      const odataBatch: ODataBatch = new ODataBatch(this.serviceScope, batchOpts);

      let userResponses: Promise<Response>[] = [];

      for (let userLoginName of userLoginNames) {
        let getUserProps: Promise<Response> = odataBatch.get(`/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(userLoginName)}'
        &$select=DisplayName,Title,PersonalUrl,PictureUrl,DirectReports,ExtendedManagers`);
        userResponses.push(getUserProps);
      }

      // Make the batch request
      odataBatch.execute().then(() => {

        userResponses.forEach((item, index) => {
          item.then((response: Response) => {

            response.json().then((responseJSON: IPerson) => {

              arrayPersons.push(responseJSON);

              if (index == (userResponses.length) - 1) {
                resolve(arrayPersons);
              }
            });
          });
        });
      });
    });
  }
}