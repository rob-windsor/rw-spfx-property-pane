import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloPropertyPaneWebPart.module.scss';
import * as strings from 'HelloPropertyPaneWebPartStrings';

import {
  SPHttpClient
} from '@microsoft/sp-http';

export interface IList {
  Id: string;
  Title: string;
}

export interface IListCollection {
  value: IList[];
}

export interface IHelloPropertyPaneWebPartProps {
  description: string;
  color: string;
  list: string;
}

export default class HelloPropertyPaneWebPart extends BaseClientSideWebPart<IHelloPropertyPaneWebPartProps> {
  private lists: IPropertyPaneDropdownOption[] = null;
  private listsDropdownDisabled: boolean = true;

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloPropertyPane}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description}">${escape(this.properties.description)}</p>
              <p class="${ styles.description}">${escape(this.properties.color)}</p>
              <p class="${ styles.description}">${escape(this.properties.list)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button}">
                <span class="${ styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private validateDescription(value: string): string {
    let result = "";

    if (value == null || value.trim().length === 0) {
      result = "Please enter a description";
    }

    return result;
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {
      let restUrl = this.context.pageContext.web.absoluteUrl +
        "/_api/web/lists?$filter=(Hidden eq false)";

      this.context.spHttpClient.get(restUrl, SPHttpClient.configurations.v1)
        .then((response) => {
          if (response.ok) {
            response.json().then((data: IListCollection) => {
              let result = data.value.map((list) => {
                return {key: list.Id, text: list.Title};
              });

              resolve(result);
            });
          } else {
            response.text().then((errorMessage) => {
              reject(errorMessage);
            });
          }
        });
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this.lists) {
      return;
    }

    this.listsDropdownDisabled = true;

    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        this.render();
      });
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
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                  onGetErrorMessage: this.validateDescription.bind(this)
                }),
                PropertyPaneDropdown("color", {
                  label: "Color",
                  options: [
                    { key: "Red", text: "Red" },
                    { key: "Green", text: "Green" },
                    { key: "Blue", text: "Blue" }
                  ]
                }),
                PropertyPaneDropdown("list", {
                  label: "List",
                  options: this.lists,
                  disabled: this.listsDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
