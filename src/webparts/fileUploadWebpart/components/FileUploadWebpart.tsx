import * as React from 'react';
import styles from './FileUploadWebpart.module.scss';
import { IFileUploadWebpartProps } from './IFileUploadWebpartProps';
import { IFileUploadWebpartState } from './IFileUploadWebpartState';
import { escape } from '@microsoft/sp-lodash-subset';
import { DragDropFiles } from "@pnp/spfx-controls-react/lib/DragDropFiles";
import { FilePicker, IFilePickerResult } from "@pnp/spfx-controls-react/lib/FilePicker";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from '@pnp/sp/presets/all';
import { TextField, PrimaryButton, Dialog, DialogFooter, DefaultButton, DialogType } from 'office-ui-fabric-react';

let userId: any;
let userName: any;
let loggedInUserEmail: string;
export default class FileUploadWebpart extends React.Component<IFileUploadWebpartProps, IFileUploadWebpartState> {
  constructor(props) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      Approvers: [],
      ApproverResponse: '',
      filePickerResult: [],
      items: [],
      hideDialog: true
    };
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this.onTextChange = this.onTextChange.bind(this);
    userId = this.props.context.pageContext.legacyPageContext.userId;
    userName = this.props.context.pageContext.legacyPageContext.userDisplayName;
    loggedInUserEmail = this.props.context.pageContext.user.email;
  }

  public componentDidMount(): void {
    this.getAllDocumentsByApprovers();
  }

  private async getAllDocumentsByApprovers() {
    let data: Array<any> = [];
    await sp.web.getFolderByServerRelativeUrl(`/sites/SPFxCrudDemo/MyDocs`).files
      .expand("ListItemAllFields/FieldValuesAsText")
      .filter(`substringof('${loggedInUserEmail}',ListItemAllFields/FieldValuesAsText/Approvers)`)
      .get()
      .then(async f => {
        data = f;
        await this.setState({ items: data });
      })
      .catch(err => {
        console.log("Errors", err);
      });
  }

  private async ShowDialog(fileName: any) {
    await this.setState({
      hideDialog: false,
      fileName: fileName
    });
  }

  private async ApproveFile() {
    let FileMetaData: any = {
      ApproverResponse: this.state.ApproverResponse
    };
    await sp.web.getFolderByServerRelativeUrl(`/sites/SPFxCrudDemo/MyDocs`).files
      .expand("ListItemAllFields/FieldValuesAsText").getByName(this.state.fileName).getItem()
      .then(f => {
        f.file.getItem().then(item => {
          console.log("item", item);
          item.update(FileMetaData)
            .then(() => {
              this.setState({ hideDialog: true });
              alert("File metedata updated sucessfully");
            });
        });
      });
  }

  private async _getPeoplePickerItems(items: any) {
    let approvers: Array<number> = [];
    items.map(async item => {
      approvers.push(item.id);
    });
    await this.setState({ Approvers: approvers });
  }

  private onTextChange(value: any) {
    this.setState({ ApproverResponse: value });
  }

  private DownloadFile(ServerRelativeUrl: any, fileName: any) {
    sp.web.getFileByServerRelativePath(ServerRelativeUrl)
      .getBlob()
      .then(blob => {
        const url = URL.createObjectURL(blob);
        const newElement = document.createElement("a");
        newElement.href = url;
        newElement.download = fileName;
        document.body.appendChild(newElement);
        newElement.click();
      });
  }

  // private _getDropFiles = (files) => {
  //   let FileMetaData: any = {
  //     ApproversId: { results: this.state.Approvers },
  //     ApproverResponse: this.state.ApproverResponse
  //   };
  //   files.map(file => {
  //     if (file.size <= 10485760) {
  //       sp.web.getFolderByServerRelativeUrl(`${this.props.siteUrl}/MyDocs`)
  //         .files.add(file.name, file, true)
  //         .then(f => {
  //           alert("File uploaded sucessfully");
  //           f.file.getItem().then(item => {
  //             item.update(FileMetaData)
  //               .then(updatedData => {
  //                 alert("File metedata updated sucessfully");
  //               });
  //           });
  //         })
  //         .catch((error) => {
  //           alert("Error is uploading");
  //           console.log('Error in uploaded Data', error);
  //         });
  //     }
  //     else {
  //       sp.web.getFolderByServerRelativeUrl(`${this.props.siteUrl}/MyDocs`)
  //         .files.addChunked(file.name, file)
  //         .then(f => {
  //           alert("File uploaded sucessfully");
  //           f.file.getItem().then(item => {
  //             item.update(FileMetaData)
  //               .then(updatedData => {
  //                 alert("File metedata updated sucessfully");
  //               });
  //           });
  //         })
  //         .catch((error) => {
  //           alert("Error is uploading");
  //           console.log('Error in uploaded Data', error);
  //         });
  //     }
  //   });
  // }

  private async saveIntoSharePoint(files: IFilePickerResult[]) {
    let FileMetaData: any = {
      ApproversId: { results: this.state.Approvers },
      ApproverResponse: "Pending"
    };

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
                  });
              });
              this.getAllDocumentsByApprovers();
            })
            .catch((error) => {
              alert("Error is uploading");
              console.log('Error in uploaded Data', error);
            });
        });
    });
  }

  public render(): React.ReactElement<IFileUploadWebpartProps> {
    const { ApproverResponse, items, hideDialog } = this.state;
    return (
      <div>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <PeoplePicker
                context={this.props.context}
                titleText="Approvers"
                personSelectionLimit={3}
                showtooltip={true}
                required={true}
                disabled={false}
                ensureUser={true}
                onChange={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
              <FilePicker
                label={'Select or upload '}
                accepts={[".gif", ".jpg", ".jpeg", ".bmp", ".ico", ".png", ".svg", ".pdf", ".docx", ".doc", ".xls", ".xlsx"]}
                buttonIcon="Upload"
                required={true}
                onSave={(file) => { this.saveIntoSharePoint(file); }}
                onChange={(filePickerResult: IFilePickerResult[]) => this.setState({ filePickerResult })}
                context={this.props.context}
              />
              {/* <DragDropFiles iconName="Upload"
                labelMessage="My custom upload File"
                dropEffect="move"
                enable={true}
                onDrop={this._getDropFiles}
              >
                Drag and drop here...
              </DragDropFiles> */}
              <Dialog
                hidden={hideDialog}
                onDismiss={(e) => this.setState({ hideDialog: true })}
                dialogContentProps={{
                  type: DialogType.close,
                  title: 'Enter Approver Remark'
                }}
                modalProps={{
                  isBlocking: true,
                  styles: { main: { maxWidth: '450px' } },
                }}
              >
                <TextField
                  label='Approver Resonse'
                  value={ApproverResponse}
                  onChange={(e, value) => this.onTextChange(value)}
                />
                <DialogFooter>
                  <PrimaryButton onClick={() => this.ApproveFile()}
                    text="Save" />
                  <DefaultButton onClick={(e) => this.setState({ hideDialog: true })}
                    text="Cancel" />
                </DialogFooter>
              </Dialog>
            </div>
          </div>

          <div className={styles.row}>
            <div className={styles.column}>
              <table style={{ border: '1px', borderColor: 'black', borderStyle: 'dotted', width: '100%' }}>
                <thead>
                  <tr>
                    <th>File Name</th>
                    <th>Approver Response</th>
                    {/* <th>Approvers</th> */}
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody style={{ textAlign: "center" }}>
                  {
                    items.length > 0 && items.map((item) => {
                      return <tr>
                        <td>
                          <a onClick={() => this.DownloadFile(item.ServerRelativeUrl, item.Name)}
                            style={{ textDecoration: 'none', wordWrap: "break-word", cursor: "pointer", color: "blue" }}>
                            {item.Name}
                          </a>
                        </td>
                        {/* <td>{item.ListItemAllFields.FieldValuesAsText.Approvers}</td> */}
                        <td>{item.ListItemAllFields.ApproverResponse}</td>
                        <td>
                          <PrimaryButton
                            text='Approve'
                            onClick={() => this.ShowDialog(item.Name)}
                          />
                        </td>
                      </tr>;
                    })
                  }
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div >
    );
  }
}
