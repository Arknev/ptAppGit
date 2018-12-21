import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneConditionalGroup,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
  IWebPartPropertiesMetadata,
  DynamicDataSharedDepth
} from '@microsoft/sp-webpart-base';

import * as strings from 'PtAppRootWp2WebPartStrings';
import PtAppRootWp1 from './components/PtAppRootWp2';
import { IPtAppRootWp2Props } from './components/IPtAppRootWp2Props';
import { DynamicProperty } from '@microsoft/sp-component-base';

export interface IPtAppRootWp1WebPartProps {
  description: string;
  UserItemId: DynamicProperty<string>;
  UserDispName: DynamicProperty<string>;
  UserAccountName: DynamicProperty<string>;
  UserEMail: DynamicProperty<string>;
  UserDistId: DynamicProperty<string>;
  UserDistName: DynamicProperty<string>;
  UserThirdId: DynamicProperty<string>;
  UserThirdName: DynamicProperty<string>;
  UserRoles: DynamicProperty<string>;
  UserRoleIds: DynamicProperty<string>;
}
export default class PtAppRootWp1WebPart extends BaseClientSideWebPart<IPtAppRootWp1WebPartProps> {
  private _onConfigure = (): void => {
    this.context.propertyPane.open();
  }
  public render(): void {
    const UserItemId: string | undefined = this.properties.UserItemId.tryGetValue();
    const UserDispName: string | undefined = this.properties.UserDispName.tryGetValue();
    const UserAccountName: string | undefined = this.properties.UserAccountName.tryGetValue();
    const UserEMail: string | undefined = this.properties.UserEMail.tryGetValue();
    const UserDistId: string | undefined = this.properties.UserDistId.tryGetValue();
    const UserDistName: string | undefined = this.properties.UserDistName.tryGetValue();
    const UserThirdId: string | undefined = this.properties.UserThirdId.tryGetValue();
    const UserThirdName: string | undefined = this.properties.UserThirdName.tryGetValue();
    const UserRoles: string | undefined = this.properties.UserRoles.tryGetValue();
    const UserRoleIds: string | undefined = this.properties.UserRoleIds.tryGetValue();
    const needsConfiguration: boolean = (!UserItemId && !this.properties.UserItemId.tryGetSource()) || (!UserDispName && !this.properties.UserDispName.tryGetSource());

    const element: React.ReactElement<IPtAppRootWp2Props> = React.createElement(
      PtAppRootWp1,
      {
        needsConfiguration: needsConfiguration,
        description: this.properties.description,
        UserItemId: UserItemId ? UserItemId : 'No Data',
        UserDispName: UserDispName ? UserDispName : 'No Data',
        UserAccountName: UserAccountName ? UserAccountName : 'No Data',
        UserEMail: UserEMail ? UserEMail : 'No Data',
        UserDistId: UserDistId ? UserDistId : 'No Data',
        UserDistName: UserDistName ? UserDistName : 'No Data',
        UserThirdId: UserThirdId ? UserThirdId : 'No Data',
        UserThirdName: UserThirdName ? UserThirdName : 'No Data',
        UserRoles: UserRoles ? UserRoles : 'No Data',
        UserRoleIds: UserRoleIds ? UserRoleIds : 'No Data',
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      'UserItemId': {
        dynamicPropertyType: 'string'
      },
      'UserDispName': {
        dynamicPropertyType: 'string'
      },
      'UserAccountName': {
        dynamicPropertyType: 'string'
      },
      'UserEMail': {
        dynamicPropertyType: 'string'
      },
      'UserDistId': {
        dynamicPropertyType: 'string'
      },
      'UserDistName': {
        dynamicPropertyType: 'string'
      },
      'UserThirdId': {
        dynamicPropertyType: 'string'
      },
      'UserThirdName': {
        dynamicPropertyType: 'string'
      },
      'UserRoles': {
        dynamicPropertyType: 'string'
      },
      'UserRoleIds': {
        dynamicPropertyType: 'string'
      }
    } as any as IWebPartPropertiesMetadata;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              primaryGroup: {
                groupName: strings.BasicGroupName,
                groupFields: [
                  PropertyPaneTextField('UserItemId', {
                    label: strings.ItemIdFieldLabel
                  }),
                  PropertyPaneTextField('UserDispName', {
                    label: strings.DispNameFieldLabel
                  }),
                  PropertyPaneTextField('UserAccountName', {
                    label: strings.AccountNameFieldLabel
                  }),
                  PropertyPaneTextField('UserEMail', {
                    label: strings.EMailFieldLabel
                  }),
                  PropertyPaneTextField('UserDistId', {
                    label: strings.DistIdFieldLabel
                  }),
                  PropertyPaneTextField('UserDistName', {
                    label: strings.DistNameFieldLabel
                  }),
                  PropertyPaneTextField('UserThirdId', {
                    label: strings.ThirdIdFieldLabel
                  }),
                  PropertyPaneTextField('UserThirdName', {
                    label: strings.ThirdNameFieldLabel
                  }),
                  PropertyPaneTextField('UserRoles', {
                    label: strings.RolesFieldLabel
                  }),
                  PropertyPaneTextField('UserRoleIds', {
                    label: strings.RoleIdsFieldLabel
                  }),
                ]
              },
              secondaryGroup: {
                groupName: strings.BasicGroupName,
                groupFields: [
                  PropertyPaneDynamicFieldSet({
                    label: 'UserItemId',
                    fields: [
                      PropertyPaneDynamicField('UserItemId', {
                        label: strings.ItemIdFieldLabel
                      }),
                      PropertyPaneDynamicField('UserDispName', {
                        label: strings.DispNameFieldLabel
                      }),
                      PropertyPaneDynamicField('UserAccountName', {
                        label: strings.AccountNameFieldLabel
                      }),
                      PropertyPaneDynamicField('UserEMail', {
                        label: strings.EMailFieldLabel
                      }),
                      PropertyPaneDynamicField('UserDistId', {
                        label: strings.DistIdFieldLabel
                      }),
                      PropertyPaneDynamicField('UserDistName', {
                        label: strings.DistNameFieldLabel
                      }),
                      PropertyPaneDynamicField('UserThirdId', {
                        label: strings.ThirdIdFieldLabel
                      }),
                      PropertyPaneDynamicField('UserThirdName', {
                        label: strings.ThirdNameFieldLabel
                      }),
                      PropertyPaneDynamicField('UserRoles', {
                        label: strings.RolesFieldLabel
                      }),
                      PropertyPaneDynamicField('UserRoleIds', {
                        label: strings.RoleIdsFieldLabel
                      }),
                        ],
                    sharedConfiguration: {
                      depth: DynamicDataSharedDepth.Property
                    }
                  })
                ]
              },
              showSecondaryGroup: !!this.properties.UserItemId.tryGetSource()
            } as IPropertyPaneConditionalGroup
          ]
        }
      ]
    };
  }
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
}
