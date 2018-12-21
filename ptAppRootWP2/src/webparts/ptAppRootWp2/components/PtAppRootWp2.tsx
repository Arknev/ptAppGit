import * as React from 'react';
import styles from './PtAppRootWp2.module.scss';
import { IPtAppRootWp2Props } from './IPtAppRootWp2Props';
import { escape } from '@microsoft/sp-lodash-subset';
export default class PtAppRootWp1 extends React.Component<IPtAppRootWp2Props, {}> {
  private RoleItemRender(RolesJsonString){
    let RoleJSX: any = null;
    let RoleItemRowJSX: any = [];
    try {
      let RolesJsonObj: object = JSON.parse(RolesJsonString);
      for (let varEachProperty in RolesJsonObj) {
        if (RolesJsonObj.hasOwnProperty(varEachProperty)) {
          let CurrRoleItemJSX =
            <div>
              <div>{'Role Name: '}{RolesJsonObj[varEachProperty].RoleName}{' | Id: '}{RolesJsonObj[varEachProperty].solDSecRolesLU}</div>
            </div>
          ;
          RoleItemRowJSX.push(CurrRoleItemJSX);
        }
      }
    }
    catch (error) {
      console.log('Role Processing Error');
      console.log(error);
      let DummyValue = <span>{'0'}</span>;
      RoleItemRowJSX.push(DummyValue);
    }
    RoleJSX =
      <div>
        {RoleItemRowJSX}
      </div>
    ;
    return RoleJSX;
  }
  public render(): React.ReactElement<IPtAppRootWp2Props> {
    let UserRolesJSX = this.props.UserRoles !== '0' ? this.RoleItemRender(this.props.UserRoles) : '0';
    let SampleDisplay;
    if(this.props.needsConfiguration){
      SampleDisplay =
        <div className={ styles.ptAppRootWp2 }>
          <div className={ styles.container }>
            <div className={ styles.row }>
              <div className={ styles.column }>
                <p className={ styles.description }>{'Needs Config'}</p>
              </div>
            </div>
          </div>
        </div>
      ;
    }
    else{
      SampleDisplay =
        <div className={ styles.ptAppRootWp2 }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <p className={ styles.description }>{'Description'}: {this.props.description}</p>
              <p className={ styles.description }>{'UserIdemId'}: {this.props.UserItemId}</p>
              <p className={ styles.description }>{'UserDispName'}: {this.props.UserDispName}</p>
              <p className={ styles.description }>{'UserAccountName'}: {this.props.UserAccountName}</p>
              <p className={ styles.description }>{'UserEMail'}: {this.props.UserEMail}</p>
              <p className={ styles.description }>{'UserDistName'}: {this.props.UserDistName}</p>
              <p className={ styles.description }>{'UserDistId'}: {this.props.UserDistId}</p>
              <p className={ styles.description }>{'UserThirdName'}: {this.props.UserThirdName}</p>
              <p className={ styles.description }>{'UserThirdId'}: {this.props.UserThirdId}</p>
              <p className={ styles.description }>{'UserRoleIds'}: {this.props.UserRoleIds}</p>
              <p className={ styles.description }>{'UserRoles'}: {UserRolesJSX}</p>
            </div>
          </div>
        </div>
      </div>;
    }
    return SampleDisplay;
  }
}
