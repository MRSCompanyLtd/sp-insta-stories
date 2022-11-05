import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { GraphFI, graphfi, SPFx as grSPFx } from '@pnp/graph';
import "@pnp/graph/users";
import "@pnp/graph/photos";

import { SPFI, spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/sites";

import * as strings from 'CompanyStoriesWebPartStrings';
import CompanyStories from './components/CompanyStories';
import { ICompanyStoriesProps } from './components/ICompanyStoriesProps';
import { IDropdownOption } from 'office-ui-fabric-react';
import { IListInfo } from '@pnp/sp/lists';

export interface ICompanyStoriesWebPartProps {
  title: string;
  timing: string;
  selectedList: string;
}

export default class CompanyStoriesWebPart extends BaseClientSideWebPart<ICompanyStoriesWebPartProps> {

  private _sp: SPFI;
  private _graph: GraphFI;
  private _lists: IDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<ICompanyStoriesProps> = React.createElement(
      CompanyStories,
      {
        title: this.properties.title,
        timing: this.properties.timing,
        sp: this._sp,
        graph: this._graph,
        selectedList: this.properties.selectedList,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    this._sp = spfi().using(SPFx(this.context));
    this._graph = graphfi().using(grSPFx(this.context));

    this._lists = await this._sp.web.lists.filter(`Hidden eq false and BaseTemplate eq 100`)().then((res: IListInfo[]) => {
      const ret: IDropdownOption[] = res.map((item: IListInfo) => {
        return {
          key: item.Id,
          text: item.Title
        }
      });

      return ret;
    });

    return;
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }
    
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  placeholder: strings.TitleFieldPlaceholder
                }),
                PropertyPaneDropdown('selectedList', {
                  label: strings.SelectFieldLabel,
                  options: this._lists
                }),
                PropertyPaneTextField('timing', {
                  label: 'Story Duration (s)',
                  placeholder: strings.TimingFieldPlaceholder,
                  value: '5'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
