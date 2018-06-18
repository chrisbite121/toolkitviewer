import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneTextFieldProps,
  IPropertyPaneChoiceGroupOption,
  PropertyPaneChoiceGroup,
  IPropertyPaneChoiceGroupProps
} from '@microsoft/sp-webpart-base';

import * as strings from 'ToolkitViewerWebPartStrings';
import ToolkitViewer from './components/ToolkitViewer';
import { IToolkitViewerProps } from './components/IToolkitViewerProps';

export interface IToolkitViewerWebPartProps {
  library1: string;
  library2: string;
  library3: string;
  library4: string;
  itemLimit: string;
  direction: string;
  orderBy: string;
  queryString: string;
}

export default class ToolkitViewerWebPart extends BaseClientSideWebPart<IToolkitViewerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IToolkitViewerProps> = React.createElement(
      ToolkitViewer,
      {
        library1: this.properties.library1,
        library2: this.properties.library2,
        library3: this.properties.library3,
        library4: this.properties.library4,
        itemLimit: this.properties.itemLimit,
        direction: this.properties.direction,
        orderBy: this.properties.orderBy,
        queryString: this.properties.queryString,
        context: this.context,
        
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    //called when the properties are changed
  }

  protected directionChoices: IPropertyPaneChoiceGroupOption[] = <IPropertyPaneChoiceGroupOption[]>[
    {
      key: 'asc',
      text: 'Ascending'
    },
    {
      key: 'desc',
      text: 'Descending'
    }
  ]

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.page1Description
          },
          groups: [
            {
              groupName: strings.group1Name,
              groupFields: [
                PropertyPaneTextField('library1', <IPropertyPaneTextFieldProps>{
                  label: strings.library1FieldLabel
                }),
                PropertyPaneTextField('library2', <IPropertyPaneTextFieldProps>{
                  label: strings.library2FieldLabel
                }),
                PropertyPaneTextField('library3', <IPropertyPaneTextFieldProps>{
                  label: strings.library3FieldLabel
                }),
                PropertyPaneTextField('library4', <IPropertyPaneTextFieldProps>{
                  label: strings.library4FieldLabel
                })                
              ]
            },
            {
              groupName: strings.group2Name,
              groupFields: [
                PropertyPaneTextField('queryString', <IPropertyPaneTextFieldProps> {
                  label: strings.queryStringLabel
                })
              ]
            },
            {
              groupName: strings.group3Name,
              groupFields: [
                PropertyPaneTextField('itemLimit', <IPropertyPaneTextFieldProps>{
                  label: strings.itemLimitLabel
                }),
                PropertyPaneTextField('orderBy', <IPropertyPaneTextFieldProps>{
                  label: strings.orderByLabel
                }),
                PropertyPaneChoiceGroup('direction', <IPropertyPaneChoiceGroupProps>{
                  label: strings.directionLabel,
                  options: this.directionChoices
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
