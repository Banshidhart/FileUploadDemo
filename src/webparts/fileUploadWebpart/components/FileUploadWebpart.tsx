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
let user: any;
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
      items: []
    };
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this.onTextChange = this.onTextChange.bind(this);
    user = this.props.context.pageContext.user;
    console.log("user", user);
  }

  public componentDidMount(): void {
    this.getAllDocumentsByApprovers();
    this.getAllDocuments();
  }

  private async getAllDocuments() {
    let data: any;
    data = await sp.web.getFolderByServerRelativeUrl(`/sites/SPFxCrudDemo/MyDocs`).files
      .get();
    console.log("Files Data", data);
  }

  private async getAllDocumentsByApprovers() {
    let data: Array<any> = [];
    await sp.web.getFolderByServerRelativeUrl(`/sites/SPFxCrudDemo/MyDocs`).files
      .expand("ListItemAllFields")
      //.filter(`(ListItemAllFields/FieldValuesAsText/Approvers eq '${user.displayName}')`).expand("ListItemAllFields", "Author")
      .get()
      .then(async f => {
        data = f;
        let d = data[0].ListItemAllFields.ApproversId['9']
        await this.setState({ items: data });
      })
      .catch(err => {
        console.log("Errors", err);
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
      ApproverResponse: this.state.ApproverResponse
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
            })
            .catch((error) => {
              alert("Error is uploading");
              console.log('Error in uploaded Data', error);
            });
        });
    });
  }

  public render(): React.ReactElement<IFileUploadWebpartProps> {
    const { ApproverResponse, items } = this.state;
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
              <TextField
                label='Approver Resonse'
                value={ApproverResponse}
                onChange={(e, value) => this.onTextChange(value)}
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
            </div>
          </div>

          <div className={styles.row}>
            <div className={styles.column}>
              <table style={{ border: '1px', borderColor: 'black', borderStyle: 'dotted', width: '100%' }}>
                <thead style={{ backgroundColor: "aqua;" }}>
                  <tr>
                    <th>File Name</th>
                    <th>Approvers</th>
                    <th>Approver Response</th>
                  </tr>
                </thead>
                <tbody style={{ backgroundColor: "cornsilk", textAlign: "center" }}>
                  {
                    items.length > 0 && items.map((item) => {
                      return <tr>
                        <td>
                          <a onClick={() => this.DownloadFile(item.ServerRelativeUrl, item.Name)}
                            style={{ textDecoration: 'none', wordWrap: "break-word", cursor: "pointer", color: "blue" }}>
                            {item.Name}
                          </a>
                        </td>
                        <td>{item.ListItemAllFields.ApproversId.map(i => { return i + ', ' })}</td>
                        <td>{item.ListItemAllFields.ApproverResponse}</td>
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
