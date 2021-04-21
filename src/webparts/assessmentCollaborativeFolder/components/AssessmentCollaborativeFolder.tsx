import * as React from 'react';
import styles from './AssessmentCollaborativeFolder.module.scss';
import { IAssessmentCollaborativeFolderProps } from './IAssessmentCollaborativeFolderProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, ProgressIndicator, Separator, PrimaryButton, MessageBar, MessageBarType, TextField, List, Text, Dropdown, IDropdownOption, ITextField, IDropdown, Button  } from 'office-ui-fabric-react';
import { CSVReader } from 'react-papaparse';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import * as strings from 'AssessmentCollaborativeFolderWebPartStrings';

export enum Stage { Form, Running, Done }

export interface IAssessmentCollaborativeFolderState {
  selectionDetails: IDropdownOption;
  csvdata: any[];
  csvcolumns: IColumn[];
  csvSelected: IDropdownOption;
  csvItems: IDropdownOption[];
  stage: Stage;
  logs: Array<string>;
  errors: Array<string>;
  complete: number;
  removeSelected: IDropdownOption;
  removeFolders: IDropdownOption[];
}

export default class AssessmentCollaborativeFolder extends React.Component<IAssessmentCollaborativeFolderProps, IAssessmentCollaborativeFolderState> {
  private _datacolumns: IColumn[];
  private _data: any[];
  private _upnref = React.createRef<IDropdown>();
  private _aref = React.createRef<ITextField>();
  private _mref = React.createRef<ITextField>();
  private count: number = 0;
  private _id: string;

  constructor(props: IAssessmentCollaborativeFolderProps) {
    super(props);

    this.state = {
      selectionDetails: null,
      csvdata: null,
      csvcolumns: [],
      csvSelected: null,
      csvItems: [],
      stage: Stage.Form,
      logs: [],
      errors: [], 
      complete: 0,
      removeSelected: null,
      removeFolders: []
    };
  }

  public addError = (e: string, o: any):void => {
    console.error(e, o);
    let _log: Array<string> = this.state.errors;
    _log.push(e);
    this.setState({...this.state, errors: _log });
  }

  public addLog = (e: string): void => {
    let _log: Array<string> = this.state.logs;
    _log.push(e);
    this.setState({...this.state, logs: _log });
  }

  public handleOnDrop = (data) => {
    var h = data[0].meta.fields;
    this._data = data.map(r => { return r.data; });
    this._datacolumns = h.map(r => { return { key: r.replace(' ', ''), name: r, fieldName: r, isResizable: true }; });
    this.setState({...this.state, csvcolumns: this._datacolumns, csvdata: this._data, csvItems: h.map(r => ({ key: r.replace(' ', ''), text: r })) });
  }

  public handleOnError = (err, file, inputElem, reason) => {
    console.error(err);
  }

  public handleOnRemoveFile = (data) => {
    this._data = null;
    this.setState({...this.state, csvdata: null });
  }

  public onRun = async (e) => {
    this.count = 0;
    this.setState({ ...this.state, stage: Stage.Running, complete: 0, logs: [], errors: [] });

    //get the graph api client
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      //see if a Collaborate folder already exists
      client.api(`sites/${this._id}/drive/root/children`).filter("name eq 'Collaborate'").get((err2, res2) => {
        if (err2) {
          //any errors log and stop
          this.addError(err2.message, err2);
          return;
        }
        //convert results to a DriveItem array
        let _f1: MicrosoftGraph.DriveItem[] = res2.value;
        //if array length equals 0 create the Collaborate folder
        if (_f1.length == 0) this.makeSubFolder(client, this._id);
        else this.run2(client, _f1);
      });
    });
  }

  public async makeSubFolder(client: MSGraphClient, id: string) {
    await client.api(`sites/${id}/drive/root/children`).post({ name: "Collaborate", folder: { }, '@microsoft.graph.conflictBehavior': "rename" }, (err, res) => {
      if (err) {
        //any errors log and stop
        this.addError(err.message, err);
        return;
      }
      this.addLog("Collaborate Folder Created");
      this.run2(client, [res]);
    });
  }

  public run2 = (client: MSGraphClient, _f1: MicrosoftGraph.DriveItem[]) => {
    //create the named Collaborate folder inside Collaborate
    client.api(`sites/${this._id}/drive/items/${_f1[0].id}/children`).post({ name: this._aref.current.value, folder: { }, '@microsoft.graph.conflictBehavior': "rename" }, (err3, res3: MicrosoftGraph.DriveItem) => {
      if (err3) {
        //any errors log and stop
        this.addError(err3.message, err3);
        return;
      }
      this.addLog(`Folder '${this._aref.current.value}' Created`);
      this._data.forEach(i => {
        this.count++;
        this.makeStuFolder(client, this._id, res3.id, _f1[0].id, i[this.state.csvSelected.text], this._mref.current.value);
        
        this.setState({...this.state, complete: this.state.complete+1});
      });
      this.Done();
    });
  }

  public async makeStuFolder(client: MSGraphClient, id: string, path: string, name: string, upn: string, message: string) {
    client.api(`sites/${id}/drive/items/${path}/invite`).post({ recipients: [{ email: upn }], message: message, requireSignIn: true, sendInvitation: true, roles: [ "write" ] }, (err2, res2) => {
      if (err2) {
        //any errors log and stop
        this.addError(err2.message, err2);
        return;
      }
      this.addLog(`Permission granted and Sharing link sent to ${upn}`);
    });
  }

  public onRunRemove = (e) => {
    this.setState({ ...this.state, stage: Stage.Running, logs: [], errors: [], complete: null  });
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api(`sites/${this._id}/drive/items/${this.state.removeSelected.key}/permissions`).get((err, res) => {
        if (err) {
          //any errors log and stop
          this.addError(err.message, err);
          return;
        }
        let _res2: MicrosoftGraph.Permission[] = res.value;
        this.addLog(`Found ${_res2.length} permissions`);
        _res2.forEach(p => {
          let u: any = p.grantedTo.user;
          if (u.email) {
            client.api(`sites/${this._id}/drive/items/${this.state.removeSelected.key}/permissions/${p.id}`).delete();
            this.addLog(`Removed permission ${p.id} on ${this.state.removeSelected.key}`);
          }
        });
        this.Done();
      });
    });
  }

  public Done = (): void => {
    this.setState({ ...this.state, stage: Stage.Done });
  }

  public onEmailChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ ...this.state, csvSelected: item });
  }

  public onRemoveChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ ...this.state, removeSelected: item });
  }

  public componentDidMount(): void {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client.api(`sites/cf.sharepoint.com:${this.props.context.pageContext.web.serverRelativeUrl}`).select('id').get((err, res: MicrosoftGraph.Site) => {
        //get site array from the result value
        //get the id of the current site
        this._id = res.id;

        client.api(`sites/${this._id}/drive/items/root:/Collaborate:/children`).get((err2, res2) => {
          let items: MicrosoftGraph.DriveItem[] = res2.value;
          this.setState({...this.state, removeFolders: items.map(i => ({ key: i.id, text: i.name }))});
        });
      });
   });
  }


  public render(): React.ReactElement<IAssessmentCollaborativeFolderProps> {
    const { csvItems, csvdata, csvcolumns, stage, csvSelected, logs, errors, complete, removeSelected, removeFolders } = this.state;
    return (
      <div className={ styles.assessmentCollaborativeFolder }>
        <div className={ styles.container }>
          <Text>{strings.Description}</Text>
          {stage == Stage.Done &&  <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>{strings.Done}</MessageBar>}
          {stage == Stage.Running && <ProgressIndicator label={strings.ProgressLabel} description={strings.ProgressDescription} percentComplete={complete}  /> }
          <TextField defaultValue={`Collab ${new Date().getDate()}-${new Date().getMonth() + 1}-${new Date().getFullYear()}`} label={strings.CollabLabel} componentRef={this._aref} />
          <div style={{padding: "5px 0"}}>
            <span>{strings.CSVSelect}</span>
            <CSVReader onDrop={this.handleOnDrop} onError={this.handleOnError} addRemoveButton config={{ header: true, skipEmptyLines: true }} onRemoveFile={this.handleOnRemoveFile}><span>{strings.CSVDrop}</span></CSVReader>
          </div>
          <Dropdown label={strings.EmailLabel} placeholder={strings.EmailPlaceholder} options={csvItems} disabled={!csvdata} componentRef={this._upnref} onChange={this.onEmailChange} />
          <TextField label={strings.InviteLabel} multiline autoAdjustHeight componentRef={this._mref} maxLength={500} />
          <PrimaryButton text={strings.PrimaryButton} onClick={this.onRun} allowDisabledFocus disabled={!csvdata || stage === Stage.Running || !csvSelected} />

          <Separator>{strings.CSVPreview}</Separator>
          {csvdata && <DetailsList items={csvdata} columns={csvcolumns} setKey="set" layoutMode={DetailsListLayoutMode.justified} selectionMode={SelectionMode.none} />}
          <Separator>{strings.PermissionRemoval}</Separator>
          <Dropdown label={strings.RemoveLabel} options={removeFolders} placeholder={strings.RemovePlaceholder} onChange={this.onRemoveChange} />
          <Button text={strings.RemoveButton} onClick={this.onRunRemove} allowDisabledFocus disabled={stage === Stage.Running || !removeSelected } />
          {logs.length > 0 && (<><Separator>{strings.Logs}</Separator><List items={logs} onRenderCell={this._onRenderCell} /></>)}
          {errors.length > 0 && (<><Separator>{strings.Errors}</Separator><List items={errors} onRenderCell={this._onRenderCell} /></>)}
        </div>
      </div>
    );
  }
  private _onRenderCell = (item: any, index: number | undefined): JSX.Element => {
    return (
      <div data-is-focusable={true}>
        <div style={{padding: 2}}>
          {item}
        </div>
      </div>
    );
  }
}
