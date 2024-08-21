import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PopUpWebPartStrings';
import PopUp from './components/PopUp';
import { IPopUpProps } from './components/IPopUpProps';

export interface IPopUpWebPartProps {

  buttonText: string;
  popUpText:string;

}

export default class PopUpWebPart extends BaseClientSideWebPart<IPopUpWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';



  /**
   * onTextChange
   */
  private onTextChange = (newText: string) => {
    this.properties.popUpText = newText;
    return newText;
  }

  public render(): void {
    const element: React.ReactElement<IPopUpProps> = React.createElement(
      PopUp,
      {
        buttonText: this.properties.buttonText,
        popUpText:this.properties.popUpText,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        displayMode:this.displayMode,
        onTextChange:this.onTextChange
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
//   protected onBeforeSerialize(): void {
//     super.onBeforeSerialize();
//     // modify the web part's properties here - the modified version will be saved
//     //this.properties.description = this.properties.myRichText
// }

// protected onAfterDeserialize(deserializedObject: any, dataVersion: Version): IPopUpWebPartProps {
//     // handle loaded data object here, modify/convert it if necessary
//     console.log(JSON.stringify(deserializedObject));
//     console.log(JSON.stringify(dataVersion));
//     return super.onAfterDeserialize(deserializedObject, dataVersion);
// }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('buttonText',{
                  label:strings.buttonTextFieldLabel,
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
