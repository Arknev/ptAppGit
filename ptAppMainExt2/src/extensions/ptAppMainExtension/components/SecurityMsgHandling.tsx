import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { MessageBar, MessageBarType, Dialog, DialogType, DialogFooter, DefaultButton, PrimaryButton, Panel, PanelType } from 'office-ui-fabric-react';
import { sp, Web, Item, Items } from '@pnp/sp';
import {ISiteWideSecMessageHandler} from './ISiteWideSecProps';
let DsAppsWeb = new Web('https://staffkyschools.sharepoint.com/sites/dsapps/');

export class SiteWideSecMessageHandler extends React.Component<ISiteWideSecMessageHandler,{
  hideDialog: boolean,
  showPanel: boolean;
  CurrentUserPropsJSON: any;
}>{
  constructor(props) {
    super(props);
    this.state = {
      hideDialog: true,
      showPanel: false,
      CurrentUserPropsJSON: JSON.parse(this.props.CurrentUserProps.Roles),
    };
  }
  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  }
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
  private _setShowPanel = (showPanel: boolean): (() => void) => {
    return (): void => {
      this.setState({ showPanel });
    };
  }
  private RoleItemRender(){
    let RoleJSX: any = null;
    let RoleItemRowJSX: any = [];
    this.state.CurrentUserPropsJSON.forEach(RoleItem => {
      let CurrRoleItemJSX =
        <div>
          <div>{'Role Name: '}{RoleItem.RoleName}{' | Id: '}{RoleItem.solDSecRolesLU}</div>
        </div>
      ;
      RoleItemRowJSX.push(CurrRoleItemJSX);
    });
    RoleJSX =
      <div>
        {RoleItemRowJSX}
      </div>
    ;
    return RoleJSX;
  }
  public render(){
    let UserRoles = this.RoleItemRender();
    let UserSecurityPanelJSX = <span></span>;
    console.log(this.props.IsDebugMode);
    if(this.props.IsDebugMode === true){
      UserSecurityPanelJSX =
        <div>
          <DefaultButton onClick={this._setShowPanel(true)} text="User Security" />
          <Panel
            isOpen={this.state.showPanel}
            onDismiss={this._setShowPanel(false)}
            type={PanelType.medium}
            headerText="Current User Security"
            className = 'ms-font-s'
          >
            <div>
              <div>
                <strong>{'Current User: '}</strong>
                {this.props.CurrentUserProps['DispName']}
              </div>
              <div>
                <strong>{'Email Address: '}</strong>
                {this.props.CurrentUserProps['EMail']}
              </div>
              <div>
                <strong>{'Group Name: '}</strong>
                {
                  this.props.CurrentUserProps['DistName'] ?
                  this.props.CurrentUserProps['DistName'] :
                  this.props.CurrentUserProps['ThirdName']
                }
              </div>
              <div>
                <strong>{'User Roles'}</strong>
                {UserRoles}
              </div>
            </div>
          </Panel>
        </div>
      ;
    }
    // let TestSecMessageJSX =
    //   // <MessageBar
    //   //   messageBarType={MessageBarType.blocked}
    //   //   isMultiline={false}
    //   //   onDismiss={log('test')}
    //   //   dismissButtonAriaLabel="Close"
    //   //   truncated={true}
    //   //   overflowButtonAriaLabel="See more"
    //   // >Testing Error Message</MessageBar>
    //   <div>
    //     <DefaultButton secondaryText="Opens the Sample Dialog" onClick={this._showDialog} text="Open Dialog" />
    //     <Dialog
    //       hidden={this.state.hideDialog}
    //       onDismiss={this._closeDialog}
    //       dialogContentProps={{
    //         type: DialogType.normal,
    //         title: this.props.CurrentUserProps['DispName'],
    //         subText: this.props.CurrentUserProps['EMail'],
    //       }}
    //       modalProps={{
    //         isBlocking: true,
    //         containerClassName: 'ms-dialogMainOverride'
    //       }}
    //     >
    //       <DialogFooter>
    //         <PrimaryButton onClick={this._closeDialog} text="Save" />
    //         <DefaultButton onClick={this._closeDialog} text="Cancel" />
    //       </DialogFooter>
    //     </Dialog>
    //   </div>
    // ;
    return UserSecurityPanelJSX;
  }
}


