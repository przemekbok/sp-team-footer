import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneLink,
  IPropertyPaneDropdownOption,
  IPropertyPaneField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldPeoplePicker, PrincipalType, IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'SpTeamFooterWebPartStrings';
import SpTeamFooter from './components/SpTeamFooter';
import { ISpTeamFooterProps } from './components/ISpTeamFooterProps';

export interface ISpTeamFooterWebPartProps {
  listId: string;
  centerDirector: IPropertyFieldGroupOrPerson[];
}

interface IListInfo {
  id: string;
  title: string;
  entityTypeName: string;
}

export default class SpTeamFooterWebPart extends BaseClientSideWebPart<ISpTeamFooterWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _siteLists: IPropertyPaneDropdownOption[] = [];
  private _listsInfo: { [key: string]: IListInfo } = {};
  private _centerDirectorData: any = null;

  public render(): void {
    const element: React.ReactElement<ISpTeamFooterProps> = React.createElement(
      SpTeamFooter,
      {
        listId: this.properties.listId,
        centerDirector: this.properties.centerDirector ? JSON.stringify(this.properties.centerDirector) : '',
        centerDirectorData: this._centerDirectorData,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        httpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    
    // Load site lists
    await this._loadSiteLists();

    // Load center director data if available
    if (this.properties.centerDirector && this.properties.centerDirector.length > 0) {
      await this._loadCenterDirectorData();
    }

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private async _loadSiteLists(): Promise<void> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Id,Title,EntityTypeName&$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      );
      
      if (response.ok) {
        const data = await response.json();
        this._siteLists = data.value.map((list: any) => ({
          key: list.Id,
          text: list.Title
        }));

        console.log(data.value);

        // Store additional list info separately
        data.value.forEach((list: any) => {
          this._listsInfo[list.Id] = {
            id: list.Id,
            title: list.Title,
            entityTypeName: list.EntityTypeName
          };
        });
      }
    } catch (error) {
      console.error('Error loading site lists:', error);
    }
  }

  private _generateListViewUrl(listId: string): string {
    if (!listId || !this._listsInfo[listId]) return '';
    
    const listInfo = this._listsInfo[listId];
    
    // Try different URL patterns based on list type and availability
    // if (listInfo.entityTypeName) {
    //   // For custom lists, use the EntityTypeName in the URL
    //   return `${this.context.pageContext.web.absoluteUrl}/Lists/${listInfo.entityTypeName}`;
    // } else 
    if (listInfo.title) {
      // Fallback: Use the list title with proper encoding
      const encodedTitle = encodeURIComponent(listInfo.title);
      return `${this.context.pageContext.web.absoluteUrl}/Lists/${encodedTitle}`;
    } else {
      // Final fallback: Use the generic list view with list ID
      return `${this.context.pageContext.web.absoluteUrl}/_layouts/15/listform.aspx?PageType=0&ListId={${listId}}`;
    }
  }

  private async _loadCenterDirectorData(): Promise<void> {
    try {
      if (this.properties.centerDirector && this.properties.centerDirector.length > 0) {
        const userInfo = this.properties.centerDirector[0];
        // const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        //   `${this.context.pageContext.web.absoluteUrl}/_api/web/getuserbyid(${userInfo.id})`,
        //   SPHttpClient.configurations.v1
        // );

        const responseDetailed: SPHttpClientResponse = await this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName='${encodeURIComponent(userInfo.id!)}')`,
          SPHttpClient.configurations.v1
        );

        if (responseDetailed.ok) {
          var xd = await responseDetailed.json();

          console.log(xd);

          this._centerDirectorData = xd;
        }
      }
    } catch (error) {
      console.error('Error loading center director data:', error);
    }
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

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'centerDirector') {
      this.properties.centerDirector = newValue;
      await this._loadCenterDirectorData();
      this.context.propertyPane.refresh();
      this.render();
    } else if (propertyPath === 'listId') {
      this.properties.listId = newValue;
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Generate proper list view URL using list name instead of edit URL
    const listViewUrl = this.properties.listId ? this._generateListViewUrl(this.properties.listId) : '';
    const newListUrl = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/new.aspx?FeatureId={00bfea71-de22-43b2-a848-c05709900100}&ListTemplate=100`;

    const fields: IPropertyPaneField<any>[] = [
      PropertyPaneDropdown('listId', {
        label: 'Select List',
        options: this._siteLists,
        selectedKey: this.properties.listId
      })
    ];

    if (this.properties.listId) {
      fields.push(
        PropertyPaneLink('', {
          text: 'Create New List',
          href: newListUrl,
          target: '_blank'
        }),
        PropertyPaneLink('', {
          text: 'View Selected List',
          href: listViewUrl,
          target: '_blank'
        })
      );
    }

    fields.push(
      PropertyFieldPeoplePicker('centerDirector', {
        label: 'Center Director',
        initialData: this.properties.centerDirector,
        allowDuplicate: false,
        multiSelect: false,
        principalType: [PrincipalType.Users],
        onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
        context: this.context as any,
        properties: this.properties,
        key: 'centerDirectorFieldId'
      })
    );

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: fields
            }
          ]
        }
      ]
    };
  }
}
