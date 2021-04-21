import * as React from 'react';
import styles from './AssessmentSubmissionMonitor.module.scss';
import { IAssessmentSubmissionMonitorProps } from './IAssessmentSubmissionMonitorProps';
import * as strings from 'AssessmentSubmissionMonitorWebPartStrings';
import { IColumn, IDropdownOption, Dropdown, DetailsList, DetailsListLayoutMode, SelectionMode } from 'office-ui-fabric-react';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { MSGraphClient } from '@microsoft/sp-http';

export interface IAssessmentSubmissionMonitorState {
  data: any[];
  items: IDropdownOption[];
}

export default class AssessmentSubmissionMonitor extends React.Component<IAssessmentSubmissionMonitorProps, IAssessmentSubmissionMonitorState> {
  private _id: string;

  constructor(props: IAssessmentSubmissionMonitorProps) {
    super(props);

    this.state = {
      data: null,
      items: []
    };
  }

  public componentDidMount(): void {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api(`sites/cf.sharepoint.com:${this.props.context.pageContext.web.serverRelativeUrl}`).select('id').get((err, res: MicrosoftGraph.Site) => {
        //get site array from the result value
        //get the id of the current site
        this._id = res.id;

        client.api(`sites/${this._id}/drive/items/root:/Submissions:/children`).get((err2, res2) => {
          let items: MicrosoftGraph.DriveItem[] = res2.value;
          this.setState({...this.state, items: items.map(i => ({ key: i.id, text: i.name }))});
        });
      });
   });
  }

  public onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api(`sites/${this._id}/drive/items/${item.key}/children`).get((err, res) => {
        if (err) console.error(err.message, err);
        let items: MicrosoftGraph.DriveItem[] = res.value;
        this.setState({...this.state, data: items.map(i => i.folder ? ({ "Name": i.name, "Items in Folder": i.folder.childCount }) : null)});
      });
   });
  }

  public render(): React.ReactElement<IAssessmentSubmissionMonitorProps> {
    const { items, data } = this.state;
    return (
      <div className={ styles.assessmentSubmissionMonitor }>
        <div className={ styles.container }>
          <span className={ styles.title }>{strings.Title}</span>
          <Dropdown label={strings.SelectLabel} options={items} placeholder={strings.SelectPlaceholder} onChange={this.onChange} />
          {data && <DetailsList items={data} setKey="set" layoutMode={DetailsListLayoutMode.justified} selectionMode={SelectionMode.none} />}
        </div>
      </div>
    );
  }
}
