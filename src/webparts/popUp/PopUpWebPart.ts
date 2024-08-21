import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PopUpWebPartStrings';
import PopUp from './components/PopUp';
import { IPopUpProps } from './components/IPopUpProps';

export interface IPopUpWebPartProps {

  buttonText: string;
  buttonType:string;
  buttonAlignment:"auto" | "center" | "baseline" | "stretch" | "start" | "end"|undefined;

  popUpText:string;

  backgroundColor:string;


}

export default class PopUpWebPart extends BaseClientSideWebPart<IPopUpWebPartProps> {

  private _isDarkTheme: boolean = false;




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
        displayMode:this.displayMode,
        onTextChange:this.onTextChange,
        backgroundColor:this.properties.backgroundColor,
        buttonType:this.properties.buttonType,
        buttonAlignment:this.properties.buttonAlignment
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
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
                PropertyPaneDropdown('buttonType',{
                  label:'Button Type',
                  options:[{key:'Primary',text:'Primary'},{key:'Default',text:'Default'}]
                }),
                PropertyPaneChoiceGroup('buttonAlignment',{
                  label:"Button Alignment",
                  options:[{key:'start',text:'Start'},{key:'center',text:'Center'},{key:'end',text:'End'}]

                })
              ]
            }
          ]
        }
      ]
    };
  }
}
