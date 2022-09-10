import * as React from 'react';
import styles from './DocumentUploader.module.scss';
import { IDocumentUploaderProps } from './IDocumentUploaderProps';
import { boundMethod } from 'autobind-decorator';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IMeetingForm, IJsonMap, IJsonArray, IItem } from './IDocumentUploaderModel';
import { DefaultButton } from 'office-ui-fabric-react';
import { SPHttpClientResponse } from '@microsoft/sp-http';
import { DataService } from './DataService';
import * as strings from 'DocumentUploaderWebPartStrings';

export default class DocumentUploader extends React.Component<IDocumentUploaderProps, IMeetingForm> {
  private dataService: DataService;
  constructor(props: IDocumentUploaderProps) {
    super(props);
    this.dataService = new DataService(this.props.context);
    this.state = {
      attachements: [],
      peoples: [],
      fileInfo: [],
      status: false
    }
  }
  public render(): React.ReactElement<IDocumentUploaderProps> {
    return (
      <div className={styles.documentUploader} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <PeoplePicker
                context={this.props.context}
                titleText='Employee(s)'
                suggestionsLimit={6}
                personSelectionLimit={15}
                showtooltip={false}
                principalTypes={[PrincipalType.User]}
                selectedItems={this._getPeoplePickerprojectDeputy}
                ensureUser={true}
              />
            </div>

            <div className={styles.column}>
              <input type='file' multiple={true}
                onChange={this.addFile.bind(this)}
                id='projectLogoFile' aria-describedby='inputGroupFileAddon01' />
              <div>
                {
                  this.state.fileInfo.map((ev: any) => {
                    return (
                      <div>
                        <span>{ev.name}</span>
                      </div>
                    )
                  })

                }
              </div>
              <div>
                {this.state.fileInfo.length > 0 ? <b><label onClick={this.clear}>Clear</label></b> : ''}
              </div>
            </div>

          </div>
          <div className={styles.row}>
            <DefaultButton disabled={this.state.status}
              text={this.state.status ? 'Uploading...' : 'Save'} onClick={this.startsaving}></DefaultButton>
          </div>
        </div>
      </div >
    );
  }
  @boundMethod
  private clear() {
    this.setState({ fileInfo: [] });
  }
  @boundMethod
  private addFile(event) {
    let resultFile = event.target.files;
    console.log(resultFile);
    let fileInfos = this.state.fileInfo;
    for (var i = 0; i < resultFile.length; i++) {
      var fileName = resultFile[i].name;
      console.log(fileName);
      var file = resultFile[i];
      var reader = new FileReader();
      reader.onload = (function (file) {
        return function (e) {
          //Push the converted file into array
          fileInfos.push({
            'name': file.name,
            'content': e.target.result
          });
        }
      })(file);
      reader.readAsArrayBuffer(file);
    }
    setTimeout(() => {
      this.setState({ fileInfo: fileInfos });
    }, 1500);
  }
  @boundMethod
  private _getPeoplePickerprojectDeputy(items: IItem[]): void {
    let projectdeptids: string[] = [];
    projectdeptids.push(items[0].loginName);
    console.log('Items:', items[0].loginName);
    projectdeptids = [];
    items.map((ev: IItem, index) => {
      projectdeptids.push(items[index].id);
    });
    this.setState({ peoples: projectdeptids });
  }
  @boundMethod
  private startsaving(): void {
    this.setState({ status: true });
    const listname: string = this.props.description;
    let filecount: number = 0;
    this.state.fileInfo.map((ev: any) => {
      let filename: string = ev.name;
      let file: File = ev.content;

      this.dataService.uploaddocument(listname, filename, file)
        .then((response: SPHttpClientResponse) => {
          console.log(response);
          if (response.status === 200) {
            response.json().then((responseJSON: IItem) => {
              const serverRelativeUrl: string = responseJSON.ServerRelativeUrl as string;
              console.log(serverRelativeUrl);

              this.dataService.getDocId(listname, filename)
                .then((response1: SPHttpClientResponse) => {
                  console.log(response1);
                  response1.json().then((responseJSON1: any) => {
                    console.log(responseJSON1.value[0].ID);

                    const body: string = JSON.stringify({
                      '__metadata': {
                        'type': 'SP.Data.'+listname+'Item'
                      },
                      'Title': '',
                      'PeoplesId': { 'results': this.state.peoples }
                    });
                    this.dataService.updateListItemById(listname, responseJSON1.value[0].ID, body)
                      .then((response2: SPHttpClientResponse) => {
                        console.log(response2.status);
                        filecount = filecount + 1;
                        if (filecount === this.state.fileInfo.length) {
                          this.setState({ status: false });
                          this.setState({ fileInfo: [] });
                        }

                      })
                  })

                })
            });

          }
        })

    })




  }
}
