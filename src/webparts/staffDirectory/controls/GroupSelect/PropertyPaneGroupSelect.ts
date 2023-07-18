import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-property-pane';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { IPropertyPaneGroupSelectProps } from './IPropertyPaneGroupSelectProps';
import GroupSelect from './GroupSelect';
import { IGroupSelectProps } from './IGroupSelectProps';

export class PropertyPaneGroupSelect
  implements IPropertyPaneField<IPropertyPaneGroupSelectProps>
{
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneGroupSelectProps;
  private elem: HTMLElement;

  constructor(
    targetProperty: string,
    properties: IPropertyPaneGroupSelectProps
  ) {
    
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.label,
      label: properties.label,
      loadOptions: properties.loadOptions,
      onPropertyChange: properties.onPropertyChange,
      selected: properties.selected,
      disabled: properties.disabled,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this)
    };
  }

  public render(): void {
    if (!this.elem) {
      return;
    }

    this.onRender(this.elem);
  }

  private onDispose(element: HTMLElement): void {
    ReactDom.unmountComponentAtNode(element);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element: React.ReactElement<IGroupSelectProps> = React.createElement(
      GroupSelect,
      {
        label: this.properties.label,
        loadOptions: this.properties.loadOptions,
        selected: this.properties.selected,
        onChange: this._onChange.bind(this),
        disabled: false
      }
    );
    ReactDom.render(element, elem);
  }

  private _onChange = (option: IDropdownOption, index?: number): void => {
    if (this.properties?.onPropertyChange) {
      this.properties.onPropertyChange(this.targetProperty, option.key);  
    }
  }
}

