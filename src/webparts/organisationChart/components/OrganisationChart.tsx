import * as React from 'react';

import { IOrganisationChartWebPartProps } from '../IOrganisationChartWebPartProps';

import styles from '../OrganisationChart.module.scss';

import { ServiceScope, ServiceKey } from '@microsoft/sp-client-base';
import { IUserProfileService, UserProfileService } from '../services/UserProfileService';
import { IPerson } from '../interfaces/IPerson';

export interface IOrganisationChartWebPartState {
  managers?: IPerson[];
  user?: IPerson;
  reports?: IPerson[];
}

export interface IOrganisationChartProps extends IOrganisationChartWebPartProps {
}

export default class OrganisationChart extends React.Component<IOrganisationChartProps, IOrganisationChartWebPartState> {


  constructor(props: IOrganisationChartProps) {
    super(props);

    this.state = {
      managers:[],
      user: {

      },
      reports:[],
    };
  }

  public render(): JSX.Element {
    return (
      <div className={styles['ms-OrgChart']}>
        <div className="ms-OrgChart-group">
          <div className="ms-OrgChart-groupTitle">Managers</div>
          <ul className={styles['ms-OrgChart-list']}>
          {this.state.managers.map((manager, index) => (
            <li key={index} className={styles['ms-OrgChart-listItem']}>
              <button className={styles['ms-OrgChart-listItemBtn']} onClick={()=> this.onProfileLinkClick(manager.PersonalUrl)}>
                <div className="ms-Persona">
                  <div className="ms-Persona-imageArea">
                    <i className="ms-Persona-placeholder ms-Icon ms-Icon--person"></i>
                    <img className="ms-Persona-image" src={manager.PictureUrl}></img>
                  </div>
                  <div className="ms-Persona-details">
                    <div className="ms-Persona-primaryText">{manager.DisplayName}</div>
                    <div className="ms-Persona-secondaryText">{manager.Title}</div>
                  </div>
                </div>
              </button>
            </li>))}
          </ul>
        </div>
        <div className="ms-OrgChart-group">
        <div className="ms-OrgChart-groupTitle">You</div>
          <ul className={styles['ms-OrgChart-list']}>
            <li className={styles['ms-OrgChart-listItem']}>
              <button className={styles['ms-OrgChart-listItemBtn']} onClick={()=> this.onProfileLinkClick(this.state.user.PersonalUrl)}>
                <div className="ms-Persona">
                  <div className="ms-Persona-imageArea">
                    <i className="ms-Persona-placeholder ms-Icon ms-Icon--person"></i>
                    <img className="ms-Persona-image" src={this.state.user.PictureUrl}></img>
                  </div>
                  <div className="ms-Persona-details">
                    <div className="ms-Persona-primaryText">{this.state.user.DisplayName}</div>
                    <div className="ms-Persona-secondaryText">{this.state.user.Title}</div>
                  </div>
                </div>
              </button>
            </li>
          </ul>
        </div>
        <div className="ms-OrgChart-group">
          <div className="ms-OrgChart-groupTitle">Reports</div>
          <ul className={styles['ms-OrgChart-list']}>
          {this.state.reports.map((report, index) => (
            <li key={index} className={styles['ms-OrgChart-listItem']}>
              <button className={styles['ms-OrgChart-listItemBtn']} onClick={()=> this.onProfileLinkClick(report.PersonalUrl)}>
                <div className="ms-Persona">
                  <div className="ms-Persona-imageArea">
                    <i className="ms-Persona-placeholder ms-Icon ms-Icon--person"></i>
                    <img className="ms-Persona-image" src={report.PictureUrl}></img>
                  </div>
                  <div className="ms-Persona-details">
                    <div className="ms-Persona-primaryText">{report.DisplayName}</div>
                    <div className="ms-Persona-secondaryText">{report.Title}</div>
                  </div>
                </div>
              </button>
            </li>))}
          </ul>
        </div>
      </div>
    );
  }

  public onProfileLinkClick (profileLink: string): void{
    window.open(profileLink);
  }

  public componentDidMount(): void {
    this._getUserProperties();
  }

  private _getUserProperties(): void {
    const serviceScope: ServiceScope = ServiceScope.startNewRoot();
    const userProfileServiceKey: ServiceKey<IUserProfileService> = ServiceKey.create<IUserProfileService>("userprofileservicekey", UserProfileService);
    const userProfileServiceInstance: IUserProfileService = serviceScope.createDefaultAndProvide(userProfileServiceKey);
    serviceScope.finish();

    userProfileServiceInstance.getPropertiesForCurrentUSer().then((person: IPerson) => {
      this.setState({ user: person });

      userProfileServiceInstance.getPropertiesForUsers(person.ExtendedManagers).then((mngrs: IPerson[]) => {
        this.setState({ managers: mngrs });
      });

      userProfileServiceInstance.getPropertiesForUsers(person.DirectReports).then((rprts: IPerson[]) => {
        this.setState({ reports: rprts });
      });
    });
  }
}
