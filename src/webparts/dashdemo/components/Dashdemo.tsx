// primary js libraries
import * as React from 'react';
import * as jquery from 'jquery';
import * as lodash from 'lodash';

// Office-Ui Fabric Components
import {
  DocumentCard,
  DocumentCardTitle,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardActions,
  IDocumentCardPreviewProps,
  DocumentCardType
} from 'office-ui-fabric-react/lib/DocumentCard';
import Iframe from 'react-iframe';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { autobind } from '@uifabric/utilities/lib';

// Custom components
import LinkListMarkup from './linkListMarkup';
import { IDashdemoProps } from './IDashdemoProps';

// styling
import styles from '../resources/Dashdemo.module.scss';

export interface linksState {
  links: [
    {
      "Title": string;
      "AuthorId": string;
      "linkURL": string;
      "linkDesc": string;
    }
  ];
  linkSelectedURL: string;
  rowClasses: string;
  embedClasses: string;
  cardType: number;
  folders: [{}];
  folderSelected: {key:'',text:''};
  selectedTitle:string;
  searchValue:string;
}

const siteName = encodeURI('PSC Employee Documents');
let folderList = [];
const sitePath = "People/HR";   // for dev site use "sites/dev"

export default class Dashdemo extends React.Component<IDashdemoProps, linksState> {

  public constructor(props: IDashdemoProps, state: linksState) {
    super(props);
    this.state = {
      links:
      [{
        "Title": '',
        "AuthorId": '',
        "linkURL": '',
        "linkDesc": '',
      }],
      linkSelectedURL: "",
      rowClasses: "ms-Grid-col ms-sm12",
      embedClasses: "",
      cardType: DocumentCardType.normal,
      folders: [{}],
      folderSelected: null,
      selectedTitle: null,
      searchValue: ""
    };
    this.onCardClick = this.onCardClick.bind(this);
    this.folderFilter = this.folderFilter.bind(this);
  }

  // seach functions
  public searchOnChange(searchValue) {
    if (searchValue == '') {
      this.componentDidMount();
    }
  }

  public search(searchValue) {
    console.log("searching...")
    let filteredLinks = this.state.links;
    for (var x = 0; x < filteredLinks.length; x++) {
      if (filteredLinks[x]["Name"].toLowerCase().indexOf(searchValue.toLowerCase()) >= 0) {
        filteredLinks[x]["displayMode"] = true;
      } else if (filteredLinks[x]["Author"].Title.toLowerCase().indexOf(searchValue.toLowerCase()) >= 0) {
        filteredLinks[x]["displayMode"] = true;
      } else {
        filteredLinks[x]["displayMode"] = false;
      }
    }
    this.updateState(filteredLinks, null, null, null, DocumentCardType.normal, null, {key:'',text:''},null,searchValue);
  }

  @autobind
  public folderFilter(item: IDropdownOption) {
    let filteredLinks = this.state.links;
    for (var x = 0; x < filteredLinks.length; x++) {
      if (filteredLinks[x]["folder"].toLowerCase().indexOf(item.key.toString().toLowerCase()) < 0) {
        filteredLinks[x]["displayMode"] = false;
      } else {
        filteredLinks[x]["displayMode"] = true;
      }
    }
    this.updateState(filteredLinks, null, null, null, null, null,item,null,"");
  }

  public updateState(links, linkSelectedURL, rowClasses, embedClasses, cardType, folders, folderSelected, selectedTitle, searchValue) {
    // only update state paramters that are passed in to the function
    if (links == null) { links = this.state.links; }
    if (linkSelectedURL == null) { linkSelectedURL = this.state.linkSelectedURL; }
    if (rowClasses == null) { rowClasses = this.state.rowClasses; }
    if (embedClasses == null) { embedClasses = this.state.embedClasses; }
    if (cardType == null) { cardType = this.state.cardType; }
    if (folders == null) { folders = this.state.folders; }
    if (folderSelected == null) { folderSelected = this.state.folderSelected; }
    if (selectedTitle == null) { selectedTitle = this.state.selectedTitle; }
    if (searchValue == null) { searchValue = this.state.searchValue; }
     // update state
     console.log("update state")
    this.setState({
      links: links,
      linkSelectedURL: linkSelectedURL,
      rowClasses: rowClasses,
      embedClasses: embedClasses,
      cardType: cardType,
      folders: folders,
      folderSelected: folderSelected,
      selectedTitle: selectedTitle,
      searchValue: searchValue
    });
  }

  // card click listener
  public onCardClick(link, e) {
    const linkId = link.ID;
    const fileName = link.Name;
    const fileExt = fileName.substr(fileName.lastIndexOf('.') + 1);
    const fileEmbedList = 'doc~docx~xls~xlsx~ppt~pptx~pdf';
    const selectedTitle = jquery(e.target).closest('.ms-DocumentCard').find('div[class^="documentCardTitle"]').text();
    console.log('oncardclick',selectedTitle)
    if (fileEmbedList.indexOf(fileExt) >= 0) {
      const folder = encodeURI(link.folder);
      let attachmentURL = window.location.origin + "/" + sitePath + "/" + siteName + "/Forms/AllItems.aspx";
      attachmentURL += "?id=/" + sitePath + "/" + siteName + "/" + folder + "/" + fileName;
      attachmentURL += "&parent=/" + sitePath + "/" + siteName + "/" + folder;
      console.log("attachmentURL",attachmentURL)
      this.updateState(null, attachmentURL, 'ms-Grid-col ms-sm4', '', DocumentCardType.compact, null,null,selectedTitle,"");
    } else {
      window.open(link.ServerRelativeUrl, '_blank', 'rel="noopener"').focus();
    }
  }

  public clearSelected() {
    this.updateState(null, '', null, null,  DocumentCardType.normal, null, null,null,"");
  }

  public getFiles(folders, reactHandler) {
    let files = [];
    lodash.map(folders, (folder) => {
      const fileURL = window.location.origin + "/" + sitePath + "/_api/Web/GetFolderByServerRelativeUrl('PSC%20Employee%20Documents/" + folder + "')/Files?$expand=Author";
      jquery.ajax({
        url: fileURL,
        type: "GET",
        dataType: "json",
        headers: { 'Accept': 'application/json; odata=verbose;' },
        success: (fileData) => {
          lodash.map(fileData.d.results, (result) => {
            result['folder'] = folder;
            result['displayMode'] = true;
            files.push(result);
          });
        },
        error: (jqXHR, textStatus, errorThrown) => {
          console.log('jqXHR', jqXHR);
          console.log('text status', textStatus);
          console.log('error', errorThrown);
        },
        complete: (jqXHR, textStatus) => {
          this.updateState(files, '', 'ms-Grid-col ms-sm12', '', DocumentCardType.normal, folders, null,null,"");
        }
      });
    });

  }

  public componentDidMount() {
    console.log("componentDidMount")
    var reactHandler = this;
    const rootUrl = window.location.origin;
    const listName = "DashboardLinks";
    const folderURL = rootUrl + "/" + sitePath + "/_api/Web/GetFolderByServerRelativeUrl('PSC%20Employee%20Documents')/Folders";
    let folders = [];
    jquery.ajax({
      url: folderURL,
      type: "GET",
      dataType: "json",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: (resultData) => {
        lodash.map(resultData.d.results, (result) => {
          if (result['Name'] != 'Forms') {
            folders.push(result['Name']);
          }
        });
      },
      error: (jqXHR, textStatus, errorThrown) => {
        console.log('jqXHR', jqXHR);
        console.log('text status', textStatus);
        console.log('error', errorThrown);
      },
      complete: (jqXHR, status) => {
        this.getFiles(folders, reactHandler);
      }
    });
  }

  public componentDidUpdate() {
    console.log("componentDidUpdate")
    //wait for DOM then scroll to selected document
    if (this.state.selectedTitle != null) {
      console.log("windowrequestanimationframe")
      window.requestAnimationFrame(() => {
        jquery('div[class*="documentCardSelected"]').get(0).scrollIntoView();
      });
    }
  }

  public sortArray(thisArray) {
    thisArray.sort((a, b) => {
      if (a.Name.split(".")[0].trim() < b.Name.split(".")[0].trim()) return -1;
      if (a.Name.split(".")[0].trim() > b.Name.split(".")[0].trim()) return 1;
      return 0;
    });
    return thisArray;
  };

  public render(): React.ReactElement<IDashdemoProps> {
    console.log("render")
    folderList = this.state.folders.map(item => ({ key: item, text: item }));
    let showLinks = this.sortArray(this.state.links);  
    if (showLinks[0].Name == '') {
      return (
        <div>Loading...</div>
      );
    } else if (this.state.linkSelectedURL != '') {
      const docList = 'ms-Grid-col ms-sm3 ' + styles.documentList;
      console.log('this.state',this.state)
      return (
        <div id="mainContainer">
          <div className="ms-Grid">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm12">
                <div className={styles.buttonRight}>
                  <PrimaryButton
                    text='Back'
                    onClick={() => this.clearSelected()}
                  />
                </div>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className={docList}>
                <LinkListMarkup
                  links={this.state.links}
                  rowClasses={this.state.rowClasses}
                  embed={this.state.embedClasses}
                  selected={this.state.linkSelectedURL}
                  handler={this.onCardClick}
                  cardType={this.state.cardType}
                  selectedTitle={this.state.selectedTitle}
                ></LinkListMarkup>
              </div>
              <div className="ms-Grid-col ms-sm9">
                <div>
                  <iframe 
                    src={this.state.linkSelectedURL}
                    height="1000px" 
                    width="100%">
                  </iframe>
                </div>
              </div>
            </div>
          </div>
        </div>
      );
      /*
      <Iframe url={this.state.linkSelectedURL}
        width="100%"
        height="1000px"
        display="initial"
        position="relative"
        allowFullScreen>
      </Iframe>
      */
    } else {
      console.log("else")
      console.log('this.state',this.state)
      const theseIncidents = this.state.links;
      let { folderSelected } = this.state;
      return (
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              <div className={styles.searchFilterContainer}>
                <div className={styles.searchContainer}>
                  <SearchBox
                    labelText='Search file name or author'
                    onChange={(newValue) => this.searchOnChange(newValue)}
                    onSearch={(newValue) => this.search(newValue)}
                    value={this.state.searchValue}
                  />
                </div>
                <div className={styles.searchFilterConnector}>- or -</div> 
                <div className={styles.folderContainer}>
                  <Dropdown
                    label=''
                    id='folderFilterSelect'
                    placeHolder='Filter list on folder'
                    onChanged={this.folderFilter}
                    options={folderList}
                    selectedKey={ folderSelected && folderSelected.key}
                  />
                </div>
              </div>
            </div>
          </div>
          <LinkListMarkup
            links={this.state.links}
            rowClasses={this.state.rowClasses}
            embed={this.state.embedClasses}
            selected={this.state.linkSelectedURL}
            handler={this.onCardClick}
            cardType={this.state.cardType}
            selectedTitle={this.state.selectedTitle}
          ></LinkListMarkup>
        </div>
      );
    }
  }
}
