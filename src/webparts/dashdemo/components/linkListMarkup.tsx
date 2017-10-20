// React
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import Iframe from 'react-iframe';

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
import { ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Modal } from 'office-ui-fabric-react/lib/Modal';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';

// Custom components and properties
import { IDashdemoProps } from './IDashdemoProps';

// Styling
import styles from '../resources/Dashdemo.module.scss';

// local state
export interface localState {
  rowClasses: string;
  selectedTitle:string;
}

// component class definition
export default class LinkListContainer extends React.Component<any, localState> {

  // constructor
  public constructor(props: IDashdemoProps, state: localState) {
    super(props);
  }

  public getImageSrc(fileExt) {
    let filename = '';
    if (fileExt == 'xlsx' || fileExt == 'xls') {
      filename = 'msexcel.jpg';
    } else if (fileExt == 'docx' || fileExt == 'doc') {
      filename = 'msword.jpg';
    } else if (fileExt == 'pptx' || fileExt == 'ppt') {
      filename = 'mspowerpoint.jpg';
    } else if (fileExt == 'pdf') {
      filename = 'pdf.png';
    } else if (fileExt == 'mp4' || fileExt == 'avi' || fileExt == 'mpeg') {
      filename = 'video.jpg';
    } else {
      filename = null;
    }
    if (filename != null) {
      const fullUrl = window.location.origin + "/sites/HRDashboard/SiteAssets/" + filename;
      //const fullUrl = window.location.origin + "/sites/dev/SiteAssets/" + filename;
      return fullUrl;
    } else {
      return null;
    }
  }
  
  public getIconStyling(fileExt, selected) {
    let fileExtClassNormal = '';
    let fileExtClassCompact = '';

    if (fileExt == 'xlsx' || fileExt == 'xls') {
      fileExtClassNormal = [styles.MSIcons, styles.excelIcon].join(' ');
      fileExtClassCompact = [styles.documentCard, styles.documentCardExcel].join(' ');
    } else if (fileExt == 'docx' || fileExt == 'doc') {
      fileExtClassNormal = [styles.MSIcons, styles.wordIcon].join(' ');
      fileExtClassCompact = [styles.documentCard, styles.documentCardWord].join(' ');
    } else if (fileExt == 'pptx' || fileExt == 'ppt') {
      fileExtClassNormal = [styles.MSIcons, styles.powerpointIcon].join(' ');
      fileExtClassCompact = [styles.documentCard, styles.documentCardPowerpoint].join(' ');
    } else if (fileExt == 'pdf') {
      fileExtClassNormal = [styles.MSIcons, styles.adobeIcon].join(' ');
      fileExtClassCompact = [styles.documentCard, styles.documentCardPdf].join(' ');
    } else {
      fileExtClassNormal = [styles.MSIcons, styles.genericIcon].join(' ');
      fileExtClassCompact = [styles.documentCard, styles.documentCardGeneric].join(' ');
    }
    if (selected == true) {
      fileExtClassCompact = [fileExtClassCompact, styles.documentCardSelected].join(' ');
    }
    return { normal: fileExtClassNormal, compact: fileExtClassCompact };
  }
  // return loading if the incidents state has not yet been set
  public render(): React.ReactElement<IDashdemoProps> {
    const links = this.props.links.slice(0, this.props.showRecentIncidents);
    const handler = this.props.handler;
    const rootUrl = window.location.origin;
    const siteName = "dev";
    const listName = "SiteAssets";
    const fileName = "preview200.jpg";
    const selectedTitle = this.props.selectedTitle;
    let previewURL = rootUrl + "/sites/" + siteName + "/" + listName + "/" + fileName;

    if (!links) {
      return <div>Loading...</div>;
    }
    // return list of incidents
    return (
      <div className="ms-Grid">
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12">
            {links.map((link, key) => {
              
              if (link.displayMode == true) {
                const date = new Date(link.TimeCreated);
                var formatOptions = {
                  day: '2-digit',
                  month: '2-digit',
                  year: 'numeric',
                  hour: '2-digit',
                  minute: '2-digit',
                  hour12: true
                };
                const displayDate = date.toLocaleDateString('en-US', formatOptions);

                if (link.incidentPhotos != null) {
                  previewURL = null;//link.incidentPhotos.Url;
                }
                const thisPreviewProps: IDocumentCardPreviewProps = {
                  previewImages: [
                    {
                      previewImageSrc: previewURL,
                      imageFit: ImageFit.none
                    }
                  ],
                };
                var fileExt = null;
                var fileExtClass = null;
                var fileExtClassSm = null;
                var selected = false;
          

                fileExt = link.Name.split(".")[1].trim();
                var thisTitle = link.Name.split(".")[0].trim();
                if (thisTitle == selectedTitle) {
                  selected = true;
                } else {
                  selected = false;
                }
                var thisFolder = link.folder["Name"];
            
                if (this.props.cardType == 0) {
                  // Normal Document Card Size
                  let tags = link.ListItemAllFields.TaxKeyword.results;
                  let displayTag = [];
                  if (tags.length > 0) {
                    tags.map((tag) => {
                      displayTag.push(tag.Label);
                    });
                  } 
                  return (
                    <DocumentCard
                      className={styles.documentCard}
                      onClick={handler.bind(this, link)}>
                      <div>
                        <div className={styles.documentCardFolder}>
                          {displayTag.join(' \u00B7 ')}
                        </div>
                        <div className={styles.documentCardTitle}>
                          {thisTitle}
                        </div>
                      </div>
                      <div className={styles.documentCardActivity}>
                        <DocumentCardActivity
                          activity={displayDate}
                          people={[{
                            name: link.ModifiedBy.Title,
                            profileImageSrc: this.getImageSrc(fileExt)
                          }]}
                        />
                      </div>
                    </DocumentCard>
                  );
                } else {
                  // Compact Document Card Size          
                  return (
                    <DocumentCard
                      type={DocumentCardType.compact}
                      className={this.getIconStyling(fileExt, selected).compact}
                      onClick={handler.bind(this, link)}>
                      <div className='ms-DocumentCard-details'>
                        <div className={styles.iconContainer}>
                          <div className={[styles.documentCardTitle_small, styles.inline].join(' ')}>{thisTitle}</div>
                        </div>
                        <div className={styles.documentCardActivity}>
                          <DocumentCardActivity
                            activity={displayDate}
                            people={[{
                              name: link.ModifiedBy.Title,
                              profileImageSrc: this.getImageSrc(fileExt)
                            }]}
                          />
                        </div>
                      </div>
                    </DocumentCard>
                  );
                }
              }
            })}
          </div>
        </div>
      </div>
    );
  }
}