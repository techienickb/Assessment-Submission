import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AssessmentSubmissionMonitorWebPartStrings';
import AssessmentSubmissionMonitor from './components/AssessmentSubmissionMonitor';
import { IAssessmentSubmissionMonitorProps } from './components/IAssessmentSubmissionMonitorProps';

export interface IAssessmentSubmissionMonitorWebPartProps {
}

export default class AssessmentSubmissionMonitorWebPart extends BaseClientSideWebPart <IAssessmentSubmissionMonitorWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAssessmentSubmissionMonitorProps> = React.createElement(
      AssessmentSubmissionMonitor,
      {
        context: this.context
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
      pages: [
        {
          header: {
            description: strings.Title
          },
          groups: [
          ]
        }
      ]
    };
  }
}
