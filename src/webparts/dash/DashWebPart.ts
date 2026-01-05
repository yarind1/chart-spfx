import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Environment, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';



//import * as strings from 'DashWebPartStrings';
import Dash from './components/Dash';
import { IDashProps } from './components/IDashProps';
import SharePointSerivce from '../../services/SharePoint/SharePointService';

import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';

export interface IDashWebPartProps {
  listId: string;
  selectedFields: string[];
  chartType: string;
  chartTitle: string;
  color1: string;
  color2: string;
  color3: string;
}

export default class DashWebPart extends BaseClientSideWebPart<IDashWebPartProps> {
  // List options state
  private listOptions: IPropertyPaneDropdownOption[];
  private listOptionsLoading: boolean = false;
  //Field options state
  private fieldOptions: IPropertyPaneDropdownOption[];
  private fieldOptionsLoading: boolean = false;

  public render(): void {
    const element: React.ReactElement<IDashProps> = React.createElement(
      Dash,
      {
        listId: this.properties.listId,
        selectedFields: this.properties.selectedFields,
        chartType: this.properties.chartType,
        chartTitle: this.properties.chartTitle,
        colors: [this.properties.color1, this.properties.color2, this.properties.color3]
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {

    await super.onInit();
    SharePointSerivce.setup(this.context, Environment.type);
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
            description: "Dash Settings"
          },
          groups: [
            {
              groupName: "Chart Data",
              groupFields: [
                PropertyPaneDropdown('listId', {
                  label: 'List',
                  options: this.listOptions,
                  disabled: this.listOptionsLoading,
                }),
                PropertyFieldMultiSelect('selectedFields', {
                  key: 'selectedFields',
                  label: "Slected Fields",
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading,
                  selectedKeys: this.properties.selectedFields
                }),
              ]
            },
            {
              groupName: "Chart Settings",
              groupFields: [
                PropertyPaneDropdown('chartType', {
                  label: 'Chart Type',
                  options: [
                    { key: 'Bar', text: 'Bar' },
                    { key: 'Line', text: 'Line' },
                    { key: 'Pie', text: 'Pie' },
                    { key: 'Doughnut', text: 'Doughnut' },
                  ]
                }),
                PropertyPaneTextField('chartTitle', {
                  label: 'Chart Title'
                }),
              ]
            },
            {
              groupName: "Chart Style",
              groupFields: [
                PropertyFieldColorPicker('color1', {
                  label: 'Color 1',
                  selectedColor: this.properties.color1,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'colorPicker1'
                }),
                PropertyFieldColorPicker('color2', {
                  label: 'Color 2',
                  selectedColor: this.properties.color2,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'colorPicker2'
                }),
                PropertyFieldColorPicker('color3', {
                  label: 'Color 3',
                  selectedColor: this.properties.color3,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  iconName: 'Precipitation',
                  key: 'colorPicker3'
                })

              ]
            },
          ]

        }
      ]
    };
  }
  private async getLists(): Promise<IPropertyPaneDropdownOption[]> {
    const lists = await SharePointSerivce.getLists();
    return lists.value.map(list => ({
      key: list.Id,
      text: list.Title
    }));
  }
  public async getFields(): Promise<IPropertyPaneDropdownOption[]> {
    if (!this.properties.listId) {
      return Promise.resolve([]);
    }
    const fields = await SharePointSerivce.getListFields(this.properties.listId);
    return fields.value.map(field => ({
      key: field.InternalName,
      text: `${field.Title} (${field.TypeAsString})`
    }));
  }
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    try {
      // מחכים שהרשימות יגיעו
      const listOptions = await this.getLists();

      // מעדכנים את המשתנה
      this.listOptions = listOptions;

      // מרעננים את הפאנל
      this.context.propertyPane.refresh();
      const fieldOptions = await this.getFields();
      this.fieldOptions = fieldOptions;
      this.context.propertyPane.refresh();


    } catch (error) {
      console.error('Error loading lists', error);
    }
  }
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'listId' && newValue) {
      this.properties.selectedFields = [];
      this.fieldOptionsLoading = true;
      this.context.propertyPane.refresh();

      try {
        const fieldOptions = await this.getFields();
        this.fieldOptions = fieldOptions;
      } catch (error) {
        console.error('Error loading fields', error);
        this.fieldOptions = [];
      } finally {
        this.fieldOptionsLoading = false;
        this.context.propertyPane.refresh();
      }
    }
  }
}
