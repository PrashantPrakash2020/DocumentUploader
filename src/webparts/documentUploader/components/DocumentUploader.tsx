import * as React from 'react';
import styles from './DocumentUploader.module.scss';
import { IDocumentUploaderProps } from './IDocumentUploaderProps';
import { boundMethod } from 'autobind-decorator';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IMeetingForm, IJsonMap, IJsonArray, IItem } from './IDocumentUploaderModel';
import {
  Dropdown,
  IDropdownOption, DropdownMenuItemType, DatePicker
} from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton } from 'office-ui-fabric-react';
import { SPHttpClientResponse } from '@microsoft/sp-http';
import { DataService } from './DataService';
import * as strings from 'DocumentUploaderWebPartStrings';
const typeoption: IDropdownOption[] = [
  { key: 'Select', text: 'Select' }
];
const subtypeoption: IDropdownOption[] = [
  { key: 'Select', text: 'Select' }
];
export default class DocumentUploader extends React.Component<IDocumentUploaderProps, IMeetingForm> {
  private dataService: DataService;
  constructor(props: IDocumentUploaderProps) {
    super(props);
    this.dataService = new DataService(this.props.context);
    this.state = {
      attachements: [],
      peoples: [],
      fileInfo: [],
      status: false,
      type: '',
      subtype: '',
      typech: [],
      subtypech: [],
      date: new Date,
      product: '',
      amount: '',
    }
  }
  public componentWillMount(): void {
    this.Options();
  }
  public render(): React.ReactElement<IDocumentUploaderProps> {
    return (
      <div className={styles.documentUploader} >
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <PeoplePicker
                context={this.props.context}
                titleText='Consignee'
                suggestionsLimit={6}
                personSelectionLimit={15}
                showtooltip={false}
                principalTypes={[PrincipalType.User]}
                selectedItems={this._getPeoplePickerprojectDeputy}
                ensureUser={true}
              />
            </div>
            <div className={styles.column}>
              <TextField label='Product'
                onChanged={e => this.setState({ product: e })}
                value={this.state.product}
              />
            </div>
            <div className={styles.column}>
              <DatePicker
                label='Doc Date'
                onSelectDate={e => this.setState({ date: e })}
                placeholder='Select a date.' value={this.state.date}
              />
            </div>
            <div className={styles.column}>
              <TextField label='Amount' type='number'
                onChanged={e => this.onnumberck(e)}
                value={this.state.amount}
              />
            </div>
            <div className={styles.column}>
              {
                this.rendertype()
              }
            </div>
            <div className={styles.column}>
              {
                this.rendersubtype()
              }
            </div>
            <div className={styles.column}>
            <br/>
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
  private onnumberck(e) {
    const re = /^[0-9\b]+$/;
    if (e === '' || re.test(e)) {
      this.setState({ amount: e })
    }
  }
  @boundMethod
  private Options() {
    this.dataService.getchoice(this.props.description, "DocType")
      .then((response: SPHttpClientResponse) => {
        console.log(response);
        response.json().then((responseJSON: IItem) => {
          let options: string[] = [];
          const alltype: string[] = responseJSON.value[0]['Choices'] as string[];
          alltype.forEach((item: any) => {
            console.log(item);
            typeoption.push({ key: item, text: item });
            options.push(item as string)
          })

          this.setState({
            typech: options
          });

        })
      });

    this.dataService.getchoice(this.props.description, "DocSubtype")
      .then((response: SPHttpClientResponse) => {
        console.log(response);
        response.json().then((responseJSON: IItem) => {
          let options: string[] = [];
          const allsubtype: string[] = responseJSON.value[0]['Choices'] as string[];
          allsubtype.forEach((item: any) => {
            console.log(item);
            subtypeoption.push({ key: item, text: item });
            options.push(item as string)
          })
          this.setState({
            subtypech: options
          });

        })
      });

  }
  private rendertype(): JSX.Element[] {

    const elementArr: JSX.Element[] = [];
    elementArr.push(<Dropdown
      label='Document type'
      onChanged={e => this.setState({ type: e.text })}
      options={typeoption}
    />);
    return elementArr;
  }
  private rendersubtype(): JSX.Element[] {
    const elementArr: JSX.Element[] = [];
    elementArr.push(<Dropdown
      label='Document subtype'
      onChanged={e => this.setState({ subtype: e.text })}
      options={subtypeoption}
    />);
    return elementArr;
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
                        'type': 'SP.Data.' + listname + 'Item'
                      },
                      'Title': '',
                      'PeoplesId': { 'results': this.state.peoples },
                      'Product': this.state.product,
                      'Date': this.state.date,
                      'Amount': this.state.amount,
                      'DocType': this.state.type,
                      'DocSubtype': this.state.subtype

                    });
                    this.dataService.updateListItemById(listname, responseJSON1.value[0].ID, body)
                      .then((response2: SPHttpClientResponse) => {
                        console.log(response2.status);
                        filecount = filecount + 1;
                        if (filecount === this.state.fileInfo.length) {
                          this.setState({ status: false });
                          this.setState({ fileInfo: [] });
                          alert("Document(s) uploaded successfully.");
                          window.location.reload();
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
