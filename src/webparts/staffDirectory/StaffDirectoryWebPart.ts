import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { AadHttpClient, AadHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  CustomCollectionFieldType,
  PropertyFieldCollectionData
} from '@pnp/spfx-property-controls/lib/propertyFields/collectionData';
import * as strings from 'StaffDirectoryWebPartStrings';
import { IDropdownOption } from 'office-ui-fabric-react';
import { PropertyFieldToggleWithCallout } from '@pnp/spfx-property-controls';
import { PropertyPaneGroupSelect } from './controls/GroupSelect/PropertyPaneGroupSelect';
import { update } from '@microsoft/sp-lodash-subset';
import { IStaffDirectoryProps } from './components/IStaffDirectoryProps';
import StaffDirectory from './components/StaffDirectory';
import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';

export interface IStaffDirectoryWebPartProps {
  title: string;
  pageSize: number;
  departments: IDropdownOption[];
  showDepartmentFilter: boolean;
  group: string;
}

export default class StaffDirectoryWebPart extends BaseClientSideWebPart<IStaffDirectoryWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IStaffDirectoryProps> = React.createElement(StaffDirectory, {
      title: this.properties.title,
      group: this.properties.group,
      departments: this.properties.departments,
      showDepartmentFilter: this.properties.showDepartmentFilter,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      pageSize: this.properties.pageSize,
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }

    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        '--bodyText',
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty(
        '--linkHovered',
        semanticColors.linkHovered || null
      );
    }
  }

  private _loadGroups: () => Promise<IDropdownOption[]> = async () => {
    const client: unknown = await this.context.aadHttpClientFactory.getClient(
      'https://graph.microsoft.com'
    );
    const res: AadHttpClientResponse = await (client as AadHttpClient).get(
      `https://graph.microsoft.com/v1.0/groups?$select=id,displayName`,
      AadHttpClient.configurations.v1
    );

    const groups = await res.json();

    return new Promise<IDropdownOption[]>((resolve, reject) => {
      resolve(
        groups.value.map((item: { id: string; displayName: string }) => {
          return {
            key: item.id,
            text: item.displayName
          };
        })
      );
      reject((e?: unknown) => console.error(e));
    });
  }

  private _onGroupChange(path: string, nv: unknown): void {
    update(this.properties, path, (): unknown => {
      return nv;
    });

    this.render();
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
            description: strings.PropertyPaneTitle
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneSlider('pageSize', {
                  label: strings.PageSizeFieldLabel,
                  showValue: true,
                  max: 20,
                  min: 1,
                  step: 1,
                  value: this.properties.pageSize
                }),
                new PropertyPaneGroupSelect('group', {
                  label: strings.GroupSelectFieldLabel,
                  loadOptions: this._loadGroups.bind(this),
                  onPropertyChange: this._onGroupChange.bind(this),
                  selected: this.properties.group,
                  key: 'group',
                  onRender: this.render
                }),
                PropertyFieldToggleWithCallout('showDepartmentFilter', {
                  key: 'showDepartmentFilter',
                  label: 'Show department filter',
                  onText: 'Yes',
                  offText: 'No',
                  checked: this.properties.showDepartmentFilter
                }),
                PropertyFieldCollectionData('departments', {
                  key: 'departments',
                  label: strings.DepartmentListFieldLabel,
                  panelHeader: strings.DepartmentListFieldLabel,
                  manageBtnLabel: 'Manage List',
                  value: this.properties.departments,
                  fields: [
                    {
                      id: 'key',
                      title: 'Department key from Active Directory',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'text',
                      title: 'Display name for department in list',
                      type: CustomCollectionFieldType.string,
                      required: true
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
