/**
 * @file Interfaces for all of DSM2
 */
/**
 * @description Site-Wide security message handler interface.
 * @export
 * @interface ISiteWideSecMessageHandler
 */

export interface ISiteWideSecMessageHandler{
  CurrentUserProps: any;
  IsDebugMode: boolean;
}
export interface IUserProps{
  CurrentUserEmail: string;
  CurrentUserItemId: string;
  CurrentUserRolesJsonStringify: string;
}
