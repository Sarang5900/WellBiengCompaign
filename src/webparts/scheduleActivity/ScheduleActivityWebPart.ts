import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import ScheduleActivity from './components/ScheduleActivity';
import { IScheduleActivityProps } from './components/IScheduleActivityProps';

export interface IScheduleActivityWebPartProps {
  email: string;
  fullName: string;
  description: string;
}

export default class ScheduleActivityWebPart extends BaseClientSideWebPart<IScheduleActivityWebPartProps> {

  private handleActivitySubmit = (activity: any): void => {
    console.log("Activity submitted:", activity);
  };

  public render(): void {
    const element: React.ReactElement<IScheduleActivityProps> = React.createElement(
      ScheduleActivity,
      {
        context :  this.context,
        email: this.properties.email,
        fullName: this.properties.fullName,
        onActivitySubmit: this.handleActivitySubmit
      }
    );

    ReactDom.render(element, this.domElement);
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
