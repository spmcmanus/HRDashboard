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
import { Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownInternalProps } from 'office-ui-fabric-react/lib/Dropdown';
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
  folderSelected: { key: '', text: '' };
  selectedTitle: string;
  searchValue: string;
  tagSelected: { key: '', text: '' };
  tagList: [{}];
}

const siteName = encodeURI('PSC Employee Documents');
let folderList = [];
let tagList = [];
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
      linkSelectedURL: null,
      rowClasses: "ms-Grid-col ms-sm12",
      embedClasses: "",
      cardType: DocumentCardType.normal,
      folders: [{}],
      folderSelected: null,
      selectedTitle: null,
      searchValue: "",
      tagSelected: null,
      tagList: [{}]
    };
    this.onCardClick = this.onCardClick.bind(this);
    this.folderFilter = this.folderFilter.bind(this);
  }

  // search functions
  public searchOnChange(searchValue) {
    if (searchValue == '') {
      this.componentDidMount();
    }
  }

  public search(searchValue) {
    let filteredLinks = this.state.links;
    for (var x = 0; x < filteredLinks.length; x++) {
      if (filteredLinks[x]["Name"].toLowerCase().indexOf(searchValue.toLowerCase()) >= 0) {
        filteredLinks[x]["displayMode"] = true;
      } else if (filteredLinks[x]["ModifiedBy"].Title.toLowerCase().indexOf(searchValue.toLowerCase()) >= 0) {
        filteredLinks[x]["displayMode"] = true;
      } else {
        filteredLinks[x]["displayMode"] = false;
      }
    }
    this.updateState(filteredLinks, null, null, null, DocumentCardType.normal, null, { key: '', text: '' }, null, searchValue, { key: '', text: '' }, null);
  }

  @autobind
  public folderFilter(item: IDropdownOption) {
    let filteredLinks = this.state.links;
    for (var x = 0; x < filteredLinks.length; x++) {
      if (filteredLinks[x]["folder"]["Name"].toLowerCase().indexOf(item.key.toString().toLowerCase()) < 0) {
        filteredLinks[x]["displayMode"] = false;
      } else {
        filteredLinks[x]["displayMode"] = true;
      }
    }
    this.updateState(filteredLinks, null, null, null, null, null, item, null, "", { key: '', text: '' }, null);
  }

  public folderSelect(folderName:string) {
    let filteredLinks = this.state.links;
    for (var x = 0; x < filteredLinks.length; x++) {
      if (filteredLinks[x]["folder"]["Name"].toLowerCase().indexOf(folderName.toLowerCase()) < 0) {
        filteredLinks[x]["displayMode"] = false;
      } else {
        filteredLinks[x]["displayMode"] = true;
      }
    }
    this.updateState(filteredLinks, '', null, null, null, null, {key:folderName,text:folderName}, null, "", { key: '', text: '' }, null);
  }

  @autobind
  public tagFilter(item: IDropdownOption) {
    let filteredLinks = this.state.links;
    if (item.key != undefined) {
      for (var x = 0; x < filteredLinks.length; x++) {
        let docTags = filteredLinks[x]["ListItemAllFields"].TaxKeyword.results;
        let displayMode = false;
        lodash.map(docTags, (result) => {
          if (result["Label"].toLowerCase().indexOf(item.key.toString().toLowerCase()) >= 0) {
            displayMode = true;
          }
        });
        filteredLinks[x]["displayMode"] = displayMode;
      }
    }
    this.updateState(filteredLinks, null, null, null, null, null, { key: '', text: '' }, null, "", item, null);
  }

  public updateState(links, linkSelectedURL, rowClasses, embedClasses, cardType, folders, folderSelected, selectedTitle, searchValue, tagSelected, tagList) {
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
    if (tagSelected == null) { tagSelected = this.state.tagSelected; }
    if (tagList == null) { tagList = this.state.tagList; }
    // update state
    this.setState({
      links: links,
      linkSelectedURL: linkSelectedURL,
      rowClasses: rowClasses,
      embedClasses: embedClasses,
      cardType: cardType,
      folders: folders,
      folderSelected: folderSelected,
      selectedTitle: selectedTitle,
      searchValue: searchValue,
      tagSelected: tagSelected,
      tagList: tagList
    });
  }

  // card click listener
  public onCardClick(link, e) {
    const linkId = link.ID;
    const fileName = link.Name;
    const fileExt = fileName.substr(fileName.lastIndexOf('.') + 1);
    const fileEmbedList = 'doc~docx~xls~xlsx~ppt~pptx~pdf';
    const selectedTitle = jquery(e.target).closest('.ms-DocumentCard').find('div[class^="documentCardTitle"]').text();
    if (fileEmbedList.indexOf(fileExt) >= 0) {
      const folder = encodeURI(link.folder["Name"]);
      let attachmentURL = window.location.origin + "/" + sitePath + "/" + siteName + "/Forms/AllItems.aspx";
      attachmentURL += "?id=/" + sitePath + "/" + siteName + "/" + folder + "/" + fileName;
      attachmentURL += "&parent=/" + sitePath + "/" + siteName + "/" + folder;
      this.updateState(null, attachmentURL, 'ms-Grid-col ms-sm4', '', DocumentCardType.compact, null, null, selectedTitle, "", null, null);
    } else {
      window.open(link.ServerRelativeUrl, '_blank', 'rel="noopener"').focus();
    }
  }

  public clearSelected() {
    this.updateState(null, '', null, null, DocumentCardType.normal, null, null, null, "", null, null);
  }

  public getFiles(folders, reactHandler) {
    let files = [];
    let tagList = [];
    lodash.map(folders, (folder) => {
      const fileURL = window.location.origin + "/" + sitePath + "/_api/Web/GetFolderByServerRelativeUrl('PSC%20Employee%20Documents/" + folder["Name"] + "')/Files?$expand=ModifiedBy,ListItemAllFields"; //Author";
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
            //update tag list
            lodash.map(result['ListItemAllFields'].TaxKeyword.results, (result) => {
              tagList.push(result['Label']);
            });
          });
        },
        error: (jqXHR, textStatus, errorThrown) => {
          console.log('jqXHR', jqXHR);
          console.log('text status', textStatus);
          console.log('error', errorThrown);
        },
        complete: (jqXHR, textStatus) => {
          tagList = lodash.uniq(tagList).sort();
          let defaultTag = { key: this.props.defaultTag, text: this.props.defaultTag};
          this.updateState(files, '' , 'ms-Grid-col ms-sm12', '', DocumentCardType.normal, folders, null, null, "", defaultTag , tagList);
          this.tagFilter(defaultTag);
        }
      });
    });
  }

  public componentDidMount() {
    var reactHandler = this;
    const rootUrl = window.location.origin;
    const listName = "DashboardLinks";
    const folderURL = rootUrl + "/" + sitePath + "/_api/Web/GetFolderByServerRelativeUrl('PSC%20Employee%20Documents')/Folders?$expand=ListItemAllFields";
    let folders = [];
    jquery.ajax({
      url: folderURL,
      type: "GET",
      dataType: "json",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: (resultData) => {
        lodash.map(resultData.d.results, (result) => {
          if (result['Name'] != 'Forms') {
            folders.push({ Name: result['Name'], OrderBy: result['ListItemAllFields'].OrderBy });
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
    //wait for DOM then scroll to selected document
    if (this.state.selectedTitle != null) {
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
    // Folder List
    folderList = this.state.folders.map(item => ({ key: item["Name"], text: item["Name"], orderBy: item["OrderBy"] }));
    folderList.sort((a,b) => {
      return (a["orderBy"] > b["orderBy"]) ? 1 : ((b["orderBy"] > a["orderBy"]) ? -1 : 0);
    });
    tagList = this.state.tagList.map(item => ({ key: item, text: item }));
    let showLinks = this.sortArray(this.state.links);
    console.log("showLinks",showLinks);
    //if (showLinks[0].Name == '') {
    if (showLinks.length <= 1) {
      return (
        <div className="searchFilterConnector">Loading...</div>
      ); 
    } else if (this.state.linkSelectedURL != '') {
      const docList = 'ms-Grid-col ms-sm3 ' + styles.documentList;
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
    } else {
      const theseIncidents = this.state.links;
      let { folderSelected } = this.state;
      let { tagSelected } = this.state;
      return (
        <div className="ms-Grid">
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12">
              <div className={styles.searchFilterContainer}>
                <div className={styles.searchFilterConnector}>Filter file list on tag: </div>
                <div className={styles.folderContainer}>
                  <Dropdown
                    label=''
                    id='tagFilterSelect'
                    placeHolder='Choose filter...'
                    onChanged={this.tagFilter}
                    options={tagList}
                    selectedKey={tagSelected && tagSelected.key}
                  />
                </div>
                <div className={styles.searchFilterConnector}>Or search file name and author: </div>
                <div className={styles.searchContainer}>
                  <SearchBox
                    labelText='Enter search keyword...'
                    onChange={(newValue) => this.searchOnChange(newValue)}
                    onSearch={(newValue) => this.search(newValue)}
                    value={this.state.searchValue}
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
