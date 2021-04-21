import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import * as strings from 'AssessmentCollaborativeFolderWebPartStrings';
import AssessmentCollaborativeFolder from './components/AssessmentCollaborativeFolder';
import { IAssessmentCollaborativeFolderProps } from './components/IAssessmentCollaborativeFolderProps';

export interface IAssessmentCollaborativeFolderWebPartProps {
  context: WebPartContext;
}

export default class AssessmentCollaborativeFolderWebPart extends BaseClientSideWebPart <IAssessmentCollaborativeFolderWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAssessmentCollaborativeFolderProps> = React.createElement(
      AssessmentCollaborativeFolder,
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

  protected async onInit(): Promise<void> {
    await super.onInit();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.Description
          },
          groups: [
          ]
        }
      ]
    };
  }
}
