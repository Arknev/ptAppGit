import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { IDynamicDataPropertyDefinition, IDynamicDataCallables } from '@microsoft/sp-dynamic-data';
import { Dialog } from '@microsoft/sp-dialog';
import { sp, Web, Item, Items } from '@pnp/sp';
import * as strings from 'PtAppMainExtensionApplicationCustomizerStrings';
import {SiteWideSecMessageHandler} from './components/SecurityMsgHandling';
import {ISiteWideSecMessageHandler, IUserProps} from './components/ISiteWideSecProps';
const LOG_SOURCE: string = 'PtAppMainExtensionApplicationCustomizerStrings';

let DsAppsWeb = new Web('https://staffkyschools.sharepoint.com/sites/dsapps/');
let IsDebugMode: any = '';
let CurrentUserProps = {
  ItemId: '0',
  DispName: '0',
  AccountName: '0',
  EMail: '0',
  DistId: '0',
  DistName: '0',
  ThirdId: '0',
  ThirdName: '0',
  Roles: null,
  RoleIds: '0',
};
let ItemsUserMembership: {
  RoleName: string,
  Title: string,
  Id: string,
  solDSecRolesLU: string,
  solWfActionInProgress: boolean,
}[] = [];

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPtAppExtensionMainApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}
/** A Custom Action which can be run during execution of a Client Side Application */
export default class PtAppExtensionMainApplicationCustomizer extends BaseApplicationCustomizer<IPtAppExtensionMainApplicationCustomizerProperties> implements IDynamicDataCallables {
  private _footerPlaceHolder: PlaceholderContent | undefined;
  /**
   * Return list of dynamic data properties that this dynamic data source returns
   */
  public getPropertyDefinitions(): ReadonlyArray<IDynamicDataPropertyDefinition> {
    return [
      {
        id: 'currUser',
        title: 'Current User'
      },
    ];
  }
  /**
   * Return the current value of the specified dynamic data set
   * @param propertyId ID of the dynamic data set to retrieve the value for
   */
  public getPropertyValue(propertyId: string) {
    if(propertyId === 'currUser'){
      return CurrentUserProps;
    }
    else{
      throw new Error('Bad property id');
    }
  }
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    let QSParams = new URL(window.location.href).searchParams;
    IsDebugMode = QSParams.get('KdeDebugMode');
    let currentPageUrl:string = document.URL;
    let sitePagesLibraryPath = this.context.pageContext.web.serverRelativeUrl + "/SitePages";
    this.context.dynamicDataSourceManager.initializeSource(this);
    return Promise.resolve().then(() => {
      sp.setup({
        spfxContext: this.context
      });
      sp.profiles.myProperties.get()
      .then(ResponseMyProps => {
        CurrentUserProps.DispName = ResponseMyProps.DisplayName ? ResponseMyProps.DisplayName : '0';
        CurrentUserProps.AccountName = ResponseMyProps.AccountName ? ResponseMyProps.AccountName : '0';
        CurrentUserProps.EMail = ResponseMyProps.Email ? ResponseMyProps.Email : '0';
      })
      .then(() => {
        DsAppsWeb.lists.getByTitle("District Security Users").items
          .expand('solDSecDistUsAccount/Id,solDsDistricts/Id,solDSec3rdPartyLU/Id')
          .select(''.concat('',
            'Id,',
            'solDSecDistIdNU,',
            'solDSec3rdPartyGroupIdNU,',
            'solDsDistricts/Title,',
            'solDSec3rdPartyLU/Title,',
            'solDsDistricts/Id,',
            'solDSec3rdPartyLU/Id,',
            'solDSecDistUsAccount/Id,',
            'solDSecDistUsAccount/Name,',
            'solDSecDistUsAccount/LastName,',
            'solDSecDistUsAccount/FirstName,',
            'solDSecDistUsAccount/Title,',
            'solDSecDistUsAccount/EMail,',
            'Title,',
            'solDSecDistUsAccount,',
            'solWfActionInProgress',
            ''))
          .filter('solDSecDistUsAccount/EMail eq \'' + CurrentUserProps.EMail + '\'')
          .top(500)
          .orderBy('solDSecDistUsAccount/LastName,solDSecDistUsAccount/FirstName,solDSecDistUsAccount/Name', true)
          .get()
          .then(ResponseUserItem => {
            ResponseUserItem.map(item => {
              CurrentUserProps.ItemId = item ? item.ID : '0';
              CurrentUserProps.DistId = (item.solDsDistricts) ? item.solDsDistricts.Id : '0';
              CurrentUserProps.ThirdId = (item.solDSec3rdPartyLU) ? item.solDSec3rdPartyLU.Id : '0';
              CurrentUserProps.DistName = (item.solDsDistricts) ? item.solDsDistricts.Title : '0';
              CurrentUserProps.ThirdName = (item.solDSec3rdPartyLU) ? item.solDSec3rdPartyLU.Title : '0';
            });
          })
          .then(() => {
            DsAppsWeb.lists.getByTitle("District Security Membership").items
            .expand('solDsAppName/Id,solDSecRolesLU/Id')
            .select('Id,Title,solDsAppName/Title,solDsAppName/Id,solDSecRolesLU/Title,solDSecRolesLU/Id, solWfActionInProgress')
            .filter('solDSecDistSecUserIdNU eq '.concat("'", CurrentUserProps.ItemId, "'"))
            .top(500)
            .orderBy('solDSecRolesLU/Title', false)
            .get()
            .then(ResponseUserRoles => {
              ResponseUserRoles.map(UserMembershipItem => {
                ItemsUserMembership.push({
                  RoleName: UserMembershipItem.solDSecRolesLU.Title,
                  Title: UserMembershipItem.Title,
                  Id: UserMembershipItem.Id,
                  solDSecRolesLU: UserMembershipItem.solDSecRolesLU.Id,
                  solWfActionInProgress: UserMembershipItem.solWfActionInProgress,
                });
              });
              CurrentUserProps.Roles = ItemsUserMembership;
              let StringifyRoleDetails: string = CurrentUserProps.Roles ? JSON.stringify(CurrentUserProps.Roles) : '0';
              let RoleIdsArray: any = [];
              CurrentUserProps.Roles.forEach(RoleItem => {
                RoleIdsArray.push(RoleItem.solDSecRolesLU);
              });
              CurrentUserProps.Roles = StringifyRoleDetails;
              CurrentUserProps.RoleIds = RoleIdsArray.length >= 0 ? JSON.stringify(RoleIdsArray) : '0';
              /** DEV Display the action button only if the current page is site page. This is likely to change to target specific roles */
              if(sitePagesLibraryPath.toLowerCase() === this.context.pageContext.list.serverRelativeUrl.toLowerCase()){
                this.context.placeholderProvider.changedEvent.add(this, this.RenderPlaceHolders);
              }
              console.log('CurrentUserProps');
              console.log(CurrentUserProps);
              /** Add dynamic data value set function here */
              this.context.dynamicDataSourceManager.notifyPropertyChanged('currUser');
            });
          });
        });
    });
  }
  private RenderPlaceHolders():void{
    if (!this._footerPlaceHolder) {
      this._footerPlaceHolder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, { onDispose: this._onDispose });
      if (!this._footerPlaceHolder) {
          console.error('The expected placeholder (Bottom) was not found.');
          return;
      }
      const element: React.ReactElement<ISiteWideSecMessageHandler> = React.createElement(
        SiteWideSecMessageHandler,
        {
          CurrentUserProps: CurrentUserProps,
          IsDebugMode: IsDebugMode === '1' ? true : false,
        }
      );
      ReactDOM.render(element, this._footerPlaceHolder.domElement);
    }
  }
  private _onDispose(): void {
    console.log('[SitePageMetadataFooterExtension._onDispose] Disposed custom footer placeholders.');
  }
}
