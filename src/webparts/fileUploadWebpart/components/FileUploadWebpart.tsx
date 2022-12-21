import * as React from 'react';
import styles from './FileUploadWebpart.module.scss';
import { IFileUploadWebpartProps } from './IFileUploadWebpartProps';
import { IFileUploadWebpartState } from './IFileUploadWebpartState';
import { escape } from '@microsoft/sp-lodash-subset';
import { DragDropFiles } from "@pnp/spfx-controls-react/lib/DragDropFiles";
import { FilePicker, IFilePickerResult } from "@pnp/spfx-controls-react/lib/FilePicker";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from '@pnp/sp/presets/all';
import { TextField } from 'office-ui-fabric-react';

export default class FileUploadWebpart extends React.Component<IFileUploadWebpartProps, IFileUploadWebpartState> {
  constructor(props) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      Approvers: [],
      ApproverResponse: '',
      filePickerResult: []
    }
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this.onTextChange = this.onTextChange.bind(this);
  }

  public componentDidMount(): void {
    this.getAllDocuments();
  }

  private async getAllDocuments() {
    await sp.web.getFolderByServerRelativeUrl(`/sites/SPFxCrudDemo/MyDocs`).files.get()
      .then(data => {
        console.log("All Document Data", data);
      })
      .catch(err => {
        console.log("Errors", err);
      })
  }

  private async _getPeoplePickerItems(items: any) {
    console.log('Items:', items);
    let approvers: Array<number> = [];
    items.map(async item => {
      approvers.push(item.id);
    })
    await this.setState({ Approvers: approvers });
    console.log(this.state.Approvers);
  }

  private onTextChange(value: any) {
    this.setState({ ApproverResponse: value });
  }

  private _getDropFiles = (files) => {
    files.map(file => {
      if (file.size <= 10485760) {
        sp.web.getFolderByServerRelativeUrl(`${this.props.siteUrl}/MyDocs`)
          .files.add(file.name, file, true)
          .then((data) => {
            alert("File uploaded sucessfully");
            console.log('uploaded Data', data);
          })
          .catch((error) => {
            alert("Error is uploading");
            console.log('Error in uploaded Data', error);
          });
      }
      else {
        sp.web.getFolderByServerRelativeUrl(`${this.props.siteUrl}/MyDocs`)
          .files.addChunked(file.name, file)
          .then((data) => {
            alert("File uploaded sucessfully");
            console.log('uploaded Data', data);
          })
          .catch((error) => {
            alert("Error is uploading");
            console.log('Error in uploaded Data', error);
          });
      }
    });
  }

  private async saveIntoSharePoint(files: IFilePickerResult[]) {
    let FileMetaData: any = {
      ApproversId: { results: this.state.Approvers },
      ApproverResponse: this.state.ApproverResponse
    }

    files.map(file => {
      file.downloadFileContent()
        .then(async r => {
          await sp.web.getFolderByServerRelativeUrl(`/sites/SPFxCrudDemo/MyDocs`)
            .files.add(file.fileName, r, true)
            .then(f => {
              alert("File uploaded sucessfully");
              f.file.getItem().then(item => {
                item.update(FileMetaData)
                  .then(updatedData => {
                    alert("File metedata updated sucessfully");
                    console.log('Updated MetaData', updatedData);
                  });
              });
            })
            .catch((error) => {
              alert("Error is uploading");
              console.log('Error in uploaded Data', error);
            });
        });
    });
  }

  public render(): React.ReactElement<IFileUploadWebpartProps> {
    const { ApproverResponse } = this.state;
    return (
      <div className={styles.fileUploadWebpart}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <hr />
              <PeoplePicker
                context={this.props.context}
                titleText="Approvers"
                personSelectionLimit={3}
                //groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                required={true}
                disabled={false}
                ensureUser={true}
                onChange={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000} />
              <hr />
              <span className={styles.subTitle}>Select or upload </span>
              <FilePicker
                //label={'Select or upload '}
                buttonClassName={styles.button}
                accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".ico", ".png", ".svg", ".pdf", ".docx", ".doc", ".xls", ".xlsx"]}
                buttonIcon="Upload"
                onSave={(file) => { this.saveIntoSharePoint(file) }}
                onChange={(filePickerResult: IFilePickerResult[]) => this.setState({ filePickerResult })}
                context={this.props.context}
              />
              <hr /><hr />
              <DragDropFiles iconName="Upload"
                labelMessage="My custom upload File"
                dropEffect="move"
                enable={true}
                onDrop={this._getDropFiles}
              >
                Drag and drop here...
              </DragDropFiles>
              <TextField
                label='Approver Resonse'
                value={ApproverResponse}
                onChange={(e, value) => this.onTextChange(value)}
              />
              <br /><hr />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
