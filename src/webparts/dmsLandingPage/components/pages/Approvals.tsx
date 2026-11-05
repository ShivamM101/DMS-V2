/* eslint-disable */
import { useCallback, useEffect, useMemo, useState } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import {
  SearchBox, Panel, PanelType, Toggle, mergeStyleSets,
  DefaultButton,
  FocusTrapZone,
  Layer,
  Popup,
  TextField,
  PrimaryButton
} from '@fluentui/react';
import {
  ArrowLeft20Regular,
  CheckmarkCircle20Regular,
  Person20Regular,
  Calendar20Regular,
  Checkmark20Regular,
  Dismiss20Regular,
  Eye20Regular,
  ArrowDownload20Regular,
  Info20Regular,
  ClipboardTextLtr20Regular,
} from '@fluentui/react-icons';
import { Text } from "@fluentui/react-components";
import StatusBadge from '../../common/component/StatusBadge';
import * as React from "react";
import { fileTypeConfig, getDocumentVersion } from '../../common/commonfunction';
import { format } from 'date-fns';
import { getApprovalData, updateLibrary } from '../../../../Services/GeneralDocument';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ILabel } from '../../../../Intrface/ILabel';
import { getStatusByInternalStatus } from '../../../../Services/StatusSerivce';
import { createHistoryItem } from '../../../../Services/GeneralDocHistoryService';
import { TileSendMail } from '../../../../Services/SendEmail';
import { getConfigActive } from '../../../../Services/ConfigService';
import { getDataByLibraryName } from '../../../../Services/MasTileService';
import PageLoader from '../../common/component/PageLoader';
import { SPHttpClient } from "@microsoft/sp-http";

interface IApprovalsProps {
  context: WebPartContext;
}


const popupStyles = mergeStyleSets({
  overlay: {
    position: "fixed",
    inset: 0,
    backgroundColor: "rgba(0,0,0,0.45)",
    zIndex: 1000,
  },

  content: {
    backgroundColor: "#fff",
    width: "640px",
    maxWidth: "90vw",
    borderRadius: "12px",
    padding: "24px",
    position: "absolute",
    top: "50%",
    left: "50%",
    transform: "translate(-50%, -50%)",
    boxShadow: "0 10px 30px rgba(0,0,0,0.25)",
  },

  title: {
    fontSize: "20px",
    fontWeight: 600,
    marginBottom: "4px",
  },

  description: {
    color: "#605e5c",
    marginBottom: "16px",
  },

  textarea: {
    width: "100%",
    minHeight: "120px",
    marginBottom: "24px",
  },

  fieldGroup: {
    marginBottom: "24px",
  },

  selectWrapper: {
    marginTop: "8px",
  },

  footer: {
    display: "flex",
    justifyContent: "flex-end",
    gap: "12px",
    marginTop: "24px",
  },
});



export type actionType = "APPROVE" | "REJECT";
export default function Approvals({ context }: IApprovalsProps) {
  const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
  const navigate = useNavigate();
  const location = useLocation();
  const [searchQuery, setSearchQuery] = useState('');
  const [metadataDoc, setMetadataDoc] = useState<any | null>(null);
  const [allDocs, setAllDocs] = useState([]);
  const [dynamicControl, setDynamicControl] = useState([]);
  const [configData, setConfigData] = useState<any[]>([]);
  const [isOpen, setIsOpen] = useState<boolean>(false);
  const [comment, setComment] = useState("");
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const [actions, setActions] = useState<actionType>("APPROVE");
  const [isLoading, setIsLoading] = useState(true);

  const fromPath = (location.state as any)?.from || '/';
  const libraryName = (location.state as any)?.libName || "";
  const TileName = (location.state as any)?.tileName || "";

  // const [userName, setUserName] = React.useState("");

  const controls: any[] = dynamicControl; // typed alias, no change to original state
  //const [userNames, setUserNames] = React.useState<{ [key: number]: string }>({});
  const [allUsers, setAllUsers] = React.useState<any[]>([]);

  const currentDocs = useMemo(() => {
    if (!searchQuery.trim()) return allDocs;

    const query = searchQuery.toLowerCase();

    return allDocs.filter(ws =>
      Object.keys(ws).some(key => {
        const value = (ws as any)[key];
        return (
          value !== null &&
          value !== undefined &&
          value
            .toString()
            .toLowerCase()
            .includes(query)
        );
      })
    );
  }, [searchQuery, allDocs]);




  useEffect(() => {
    Promise.all([getFiles(), fetchLibraryDetails()]).finally(() => setIsLoading(false));
  }, []);

  const getFiles = async () => {
    const data = await getApprovalData(context, libraryName, context.pageContext.user.email);
    setAllDocs(data.value || []);
  };

  const fetchLibraryDetails = async () => {
    const dataConfig = await getConfigActive(context.pageContext.web.absoluteUrl, context.spHttpClient);
    const libraryData = await getDataByLibraryName(context.pageContext.web.absoluteUrl, context.spHttpClient, libraryName);
    if (libraryData.value[0]?.DynamicControl) {
      let jsonData = JSON.parse(libraryData.value[0].DynamicControl);
      jsonData = jsonData.filter((ele: any) => ele.IsActiveControl);
      jsonData = jsonData.map((el: any) => {
        if (el.ColumnType === "Person or Group") {
          el.InternalTitleName = `${el.InternalTitleName}Id`;
        }
        return el;
      });
      setDynamicControl(jsonData);
      setConfigData(dataConfig.value);
    }
  };


  const handleApprove = async (doc: any) => {
    setActions("APPROVE");
    setComment("");
    setMetadataDoc(doc);
    setIsDialogOpen(true);
  };

  const Approve = async () => {
    try {
      let dataObj: any = {};
      let InternalStatus = "", ToUser = "";
      if (metadataDoc.InternalStatus === "PendingWithPM" && metadataDoc.PublisherEmail !== null) {
        dataObj.CurrentApprover = metadataDoc.PublisherEmail;
        ToUser = metadataDoc.PublisherEmail;
        InternalStatus = "PendingWithPublisher";
      } else {
        dataObj.CurrentApprover = "";
        let PMEmail = metadataDoc.ProjectmanagerEmail;
        let AuthorEmail = metadataDoc.Author.EMail;
        ToUser = (PMEmail == "" ? AuthorEmail : (PMEmail + ";" + AuthorEmail));
        InternalStatus = "Published";
      }
      dataObj.LatestRemark = comment;
      const status = await getStatusByInternalStatus(context.pageContext.web.absoluteUrl, context.spHttpClient, InternalStatus);

      dataObj.StatusId = status.value[0].ID;
      dataObj.InternalStatus = status.value[0].InternalStatus;
      dataObj.DisplayStatus = status.value[0].StatusName;
      await updateLibrary(context.pageContext.web.absoluteUrl, context.spHttpClient, dataObj, metadataDoc.Id, libraryName);

      const dataHistory = {
        DocumetLID: metadataDoc.Id,
        ActionDate: new Date(),
        Remark: comment,
        Status: status.value[0].StatusName,
        InternalComment: comment,
        LibName: libraryName,
        Action: "Approved"
      };
      await createHistoryItem(context.pageContext.web.absoluteUrl, context.spHttpClient, dataHistory);
      const emailObj: any = {
        To: ToUser,
        FolderPath: metadataDoc.FolderDocumentPath,
        DocName: metadataDoc.ActualName,
        AuthorTitle: metadataDoc.Author.Title,
        TileName: libraryName
      };

      if (InternalStatus == "PendingWithPublisher") {
        emailObj.Sub = DisplayLabel.PublisherEmailSubject + " " + metadataDoc.ReferenceNo;
        emailObj.Status = InternalStatus;
      } else if (InternalStatus == "PendingWithPM") {
        emailObj.Sub = DisplayLabel.PMEmailSubject + " " + metadataDoc.ReferenceNo;
        emailObj.Status = InternalStatus;

      } else {
        emailObj.Sub = DisplayLabel.PublishedEmailSubject + " " + metadataDoc.ReferenceNo;
        emailObj.Msg = DisplayLabel.PublishedEmailMsg;
        emailObj.Status = InternalStatus;
      }
      emailObj.ID = metadataDoc.Id;
      emailObj.libraryName = libraryName;
      await TileSendMail(context, emailObj);
      setComment("");
      setIsDialogOpen(false);
      getFiles();

    } catch (error) {
      console.log("error", error);
    }

  };

  const RejectFile = async () => {

    let InternalStatus = "";
    let dataobj: any = { CurrentApprover: "" };
    InternalStatus = "Rejected";
    dataobj.LatestRemark = comment;
    let ToUser = metadataDoc?.Author.EMail;
    if (metadataDoc?.InternalStatus !== "PendingWithPM") {
      ToUser = (metadataDoc?.ProjectmanagerEmail === "" ? ToUser : (ToUser + ";" + metadataDoc?.ProjectmanagerEmail));
    }

    const status = await getStatusByInternalStatus(context.pageContext.web.absoluteUrl, context.spHttpClient, InternalStatus);

    dataobj.StatusId = status.value[0].ID;
    dataobj.InternalStatus = status.value[0].InternalStatus;
    dataobj.DisplayStatus = status.value[0].StatusName;

    await updateLibrary(context.pageContext.web.absoluteUrl, context.spHttpClient, dataobj, metadataDoc?.Id, libraryName);

    const dataHistory = {
      DocumetLID: metadataDoc?.Id,
      ActionDate: new Date(),
      Remark: comment,
      Status: status.value[0].StatusName,
      InternalComment: comment,
      LibName: libraryName,
      Action: "Rejected"
    };

    await createHistoryItem(context.pageContext.web.absoluteUrl, context.spHttpClient, dataHistory);
    const emailObj = {
      To: ToUser,
      FolderPath: metadataDoc?.FolderDocumentPath,
      DocName: metadataDoc?.ActualName,
      AuthorTitle: metadataDoc?.Author.Title,
      TileName: libraryName,
      Sub: DisplayLabel.RejectEmailSubject + " " + metadataDoc?.ReferenceNo,
      Msg: DisplayLabel.RejectEmailMsg,
      Status: status.value[0].StatusName
    };

    await TileSendMail(context, emailObj);
    setComment("");
    setIsDialogOpen(false);
    getFiles();
  };


  const POPUP_CONFIG: any = {
    APPROVE: {
      title: "Approve Request",
      description: "Are You Sure You Want To Approve This Request?",
      buttonType: "primary",
    },
    REJECT: {
      title: "Reject Request",
      description: "Please Provide A Reason For Rejecting This Request.",
      buttonType: "danger",
    },
  };

  const handleReject = (doc: any) => {
    setMetadataDoc(doc);
    setComment("");
    setActions("REJECT");
    setIsDialogOpen(true);
  };

  // React.useEffect(() => {
  //   const loadUsers = async () => {
  //     if (!metadataDoc) return;

  //     const personFields = controls.filter(
  //       (item: any) => item.ColumnType === "Person or Group"
  //     );

  //     const userMap: { [key: number]: string } = {};

  //     for (const field of personFields) {
  //       const userId =
  //         metadataDoc?.[`${field.InternalTitleName}Id`] ||
  //         metadataDoc?.[field.InternalTitleName];

  //       if (userId && !userMap[userId]) {
  //         try {
  //           const response = await context.spHttpClient.get(
  //             `${context.pageContext.web.absoluteUrl}/_api/web/getUserById(${userId})`,
  //             SPHttpClient.configurations.v1
  //           );

  //           if (response.ok) {
  //             const user = await response.json();
  //             userMap[userId] = user.Title;
  //           }
  //         } catch (error) {
  //           console.error(`Error fetching user ${userId}`, error);
  //         }
  //       }
  //     }

  //     setUserNames(userMap);
  //   };

  //   loadUsers();
  // }, [metadataDoc, controls]);

  React.useEffect(() => {
    const loadUsers = async () => {
      try {
        const response = await context.spHttpClient.get(
          `${context.pageContext.web.absoluteUrl}/_api/web/siteusers?$select=Id,Title,Email`,
          SPHttpClient.configurations.v1
        );

        const data = await response.json();

        console.log("All Users", data.value);

        setAllUsers(data.value || []);
      } catch (error) {
        console.error("Error loading users", error);
      }
    };

    loadUsers();
  }, []);

  const renderDynamicControls = useCallback(() => {
    // return dynamicControl.filter((item: any, index: number) => !item.IsFieldAllowInFile).map((item: any, index: number) => {
    // return dynamicControl.filter((item: any, index: number) => item.IsFieldAllowInFile).map((item: any, index: number) => {
    return controls.map((item: any, index: number) => {

      const filterObj = configData.find((ele) => ele.Id === item.Id);

      if (!filterObj) return null;

      switch (item.ColumnType) {
        case "Dropdown":
        case "Multiple Select":
          return (
            <div className="meta-panel-field">
              <label className="meta-panel-label">{item.Title}</label>
              <div className="meta-panel-plain-value">
                <span>{metadataDoc[item.InternalTitleName]}</span>
                {/* <span className="meta-panel-chevron">&#8964;</span> */}
              </div>
            </div>
          );

        case "Person or Group":
          const userId =
            metadataDoc?.[`${item.InternalTitleName}Id`] ||
            metadataDoc?.[item.InternalTitleName];

          const selectedUser = allUsers.find(
            (user) => Number(user.Id) === Number(userId)
          );


          return (
            <div className="meta-panel-field" >
              <label className="meta-panel-label">{item.Title}</label>
              <div className="meta-panel-plain-value">
                {/* <span>{metadataDoc[item.InternalTitleName]?.Title}</span> */}
                <span>{selectedUser?.Title || ""}</span>
                {/* <span className="meta-panel-chevron">&#8964;</span> */}
              </div>
            </div >
          );

        case "Radio":
          const radioOptions = filterObj.StaticDataObject.split(";").map((ele: string) => ({
            key: ele,
            text: ele,
          }));
          return (
            <div className="meta-panel-field">
              <label className="meta-panel-label">{item.Title}</label>
              <div className="meta-panel-plain-value">
                <span>{metadataDoc[item.InternalTitleName]}</span>
                {/* <span className="meta-panel-chevron">&#8964;</span> */}
              </div>
            </div>
          );
        case "Date and Time":
          return <div className="meta-panel-field" >
            <label className="meta-panel-label">{item.Title}</label>
            <div className="meta-panel-plain-value">
              <span>{format(metadataDoc[item.InternalTitleName], "dd/MM/yyyy")}</span>
              {/* <span className="meta-panel-chevron">&#8964;</span> */}
            </div>
          </div >;

        default:
          return <div className="meta-panel-field" >
            <label className="meta-panel-label">{item.Title}</label>
            <div className="meta-panel-plain-value">
              <span>{metadataDoc[item.InternalTitleName]}</span>
              {/* <span className="meta-panel-chevron">&#8964;</span> */}
            </div>
          </div >;
      }
    });
  }, [dynamicControl, metadataDoc]);


  const openDocument = (el: any) => {
    if (!el?.File?.ServerRelativeUrl) return;
    const ext = el.File.Name?.split(".").pop()?.toLowerCase();
    const officeExts = ["doc", "docx", "xls", "xlsx", "ppt", "pptx"];
    if (ext === "pdf" || !officeExts.includes(ext)) {
      window.open(el.File.ServerRelativeUrl, "_blank");
    } else {
      window.open(
        `${context.pageContext.site.absoluteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${el.File.ServerRelativeUrl}&action=embedview`,
        "_blank"
      );
    }
  };


  const renderDocCard = (doc: any) => {
    const ext = doc.File.Name.split(".").pop();
    const fileConfig = fileTypeConfig[ext] || fileTypeConfig.other;
    const { IconName, className } = fileConfig;

    return (
      <div key={doc.ID} className="approval-card" data-testid={`card-approval-${doc.ID}`}>
        <div className="approval-card-top">
          <div className="approval-card-file-info">
            <div className={`doc-icon-wrap ${className}`}>
              <IconName className="doc-icon-svg" />
            </div>
            <div className="approval-card-name-block">
              <span className="approval-card-name" data-testid={`text-approval-name-${doc.ID}`}>{doc?.ActualName}</span>
              <span className="approval-card-ref" data-testid={`text-apprdotnet dev-certs httpsoval-ref-${doc.ID}`}>{doc?.ReferenceNo}</span>
            </div>
          </div>
          <div className="approval-card-status-wrap">
            <StatusBadge status={doc?.Status.StatusName} />
            <span className="approval-card-version" data-testid={`text-approval-version-${doc.ID}`}>v{getDocumentVersion(doc?.Level, doc?.OData__UIVersionString) || "-"}
            </span>
          </div>
        </div>

        <div className="approval-card-meta">
          <div className="approval-card-meta-item" data-testid={`text-approval-submitter-${doc.ID}`}>
            <Person20Regular className="approval-card-meta-icon" />
            <span>{doc?.Author.Title}</span>
          </div>
          <div className="approval-card-meta-item" data-testid={`text-approval-date-${doc.ID}`}>
            <Calendar20Regular className="approval-card-meta-icon" />
            <span>{format(doc.Created, "MMM dd, yyyy")}</span>
          </div>
          <div className="approval-card-meta-item" data-testid={`text-approval-workspace-${doc.ID}`}>
            <Info20Regular className="approval-card-meta-icon" />
            <span>{doc?.FolderDocumentPath}</span>
          </div>
        </div>

        <div className="approval-card-actions">
          <button
            className="approval-card-metadata-link"
            onClick={() => { setIsOpen(true); setMetadataDoc(doc); }}
            data-testid={`button-metadata-${doc.ID}`}
          >
            <ClipboardTextLtr20Regular className="approval-action-icon" />
            <span>View Metadata</span>
          </button>
          <div className="approval-card-primary-actions">
            <button
              className="approval-action-btn approval-action-reject"
              onClick={() => handleReject(doc)}
              data-testid={`button-reject-${doc.ID}`}
            >
              <Dismiss20Regular className="approval-action-icon" />
              <span>Reject</span>
            </button>
            <button
              className="approval-action-btn approval-action-approve"
              onClick={() => handleApprove(doc)}
              data-testid={`button-approve-${doc.ID}`}
            >
              <Checkmark20Regular className="approval-action-icon" />
              <span>Approve</span>
            </button>
          </div>
        </div>
      </div>
    );
  };

  if (isLoading) {
    return <PageLoader message="Loading approvals..." minHeight="72vh" />;
  }

  return (
    <div className="approval-page" data-testid="page-approvals">
      <div className="approval-topbar">
        <div className="approval-topbar-left">
          <button
            className="approval-back-btn"
            onClick={() => navigate(fromPath)}
            data-testid="link-back"
          >
            <ArrowLeft20Regular />
            <span>Go Back</span>
          </button>
          <div className="approval-topbar-title-group">
            <CheckmarkCircle20Regular className="approval-topbar-icon" />
            <h1 className="approval-topbar-title" data-testid="text-page-title">Document Approvals</h1>
          </div>
        </div>
      </div>

      <div className="approval-body">
        <div className="approval-toolbar">
          <div className="approval-search">
            <SearchBox
              placeholder="Search documents..."
              value={searchQuery}
              onChange={(_, value) => setSearchQuery(value || '')}
              onClear={() => setSearchQuery('')}
              className="dashboard-search-box"
              data-testid="input-search-approvals"
            />
          </div>
        </div>

        <div className="approval-list" data-testid="container-approval-list">
          {currentDocs.length > 0 ? (
            currentDocs.map(doc => renderDocCard(doc))
          ) : (
            <div className="approval-empty" data-testid="text-no-approvals">
              <CheckmarkCircle20Regular className="approval-empty-icon" />
              <span className="approval-empty-title">
                {'No pending approvals'}
              </span>
              <span className="approval-empty-subtitle">
                {searchQuery ? 'Try adjusting your search query.' : 'Documents submitted for approval will appear here.'}
              </span>
            </div>
          )}
        </div>
      </div>

      <Panel
        isOpen={isOpen}
        onDismiss={() => { setIsOpen(false); }}
        type={PanelType.custom}
        customWidth="520px"
        headerText=""
        hasCloseButton={false}
        onRenderHeader={() => {
          if (!metadataDoc) return null;
          return (
            <div className="meta-panel-header" data-testid="container-metadata-panel">
              <div className="meta-panel-header-top">
                <h2 className="meta-panel-title" data-testid="text-metadata-title">View Folder</h2>
                <button
                  className="meta-panel-close"
                  onClick={() => { setMetadataDoc(null); setIsOpen(false); }}
                  data-testid="button-close-metadata"
                >
                  <Dismiss20Regular />
                </button>
              </div>
          
              <div className="meta-panel-quick-actions">
                {/* <button
                  className="meta-panel-quick-btn"
                  onClick={() => {
                    if (metadataDoc?.File?.LinkingUrl === "")
                      window.open(metadataDoc?.File?.ServerRelativeUrl, "_blank");
                    else
                      window.open(metadataDoc?.File?.LinkingUrl, "_blank");
                  }}
                  data-testid="button-view-doc"
                >
                  <Eye20Regular className="meta-panel-quick-icon" />
                  <span>View Document</span>
                </button> */}
                {/* <button
                  className="meta-panel-quick-btn"
                  onClick={() => {
                    window.open(metadataDoc?.File?.ServerRelativeUrl + "?download=1");
                  }}
                  data-testid="button-download-doc"
                >
                  <ArrowDownload20Regular className="meta-panel-quick-icon" />
                  <span>Download</span>
                </button> */}
              </div>

              
            </div>
          );
        }}
        onRenderFooterContent={() => (
          <div className="approval-card-primary-actions" style={{ display: "center" }}>
            <button
              className="approval-action-btn approval-action-reject"
              onClick={() => {
                setActions("REJECT");
                setIsDialogOpen(true);
              }}
              data-testid={`button-reject-${metadataDoc?.ID}`}
            >
              <Dismiss20Regular className="approval-action-icon" />
              <span>Reject</span>
            </button>
            <button
              className="approval-action-btn approval-action-approve"
              onClick={() => {
                setActions("APPROVE");
                setIsDialogOpen(true);
              }}
              data-testid={`button-approve-${metadataDoc?.ID}`}
            >
              <Checkmark20Regular className="approval-action-icon" />
              <span>Approve</span>
            </button>
          </div>
        )}
      // isFooterAtBottom={true}
      >
        {metadataDoc && (
          <div className="meta-panel-body">
            <div className="meta-panel-section">
              <h3 className="meta-panel-section-title">Folder Details</h3>
              <div className="meta-panel-fields">
                <div className="meta-panel-row meta-panel-row-2col">
                  <div className="meta-panel-field">
                    <label className="meta-panel-label">{DisplayLabel.TileName}</label>
                    <span className="meta-panel-plain-value" data-testid="text-meta-tile">{TileName}</span>
                  </div>
                  <div className="meta-panel-field">
                    <label className="meta-panel-label">{DisplayLabel.FolderName}</label>
                    <div className="meta-panel-plain-value" data-testid="text-meta-name">{metadataDoc?.FolderDocumentPath.split("/").pop() || ""}</div>
                  </div>
                </div>
                <div className="meta-panel-row meta-panel-row-2col">
                  <div className="meta-panel-field">
                    <label className="meta-panel-label">{DisplayLabel.ReferenceNo}</label>
                    <span className="meta-panel-plain-value" data-testid="text-meta-tile">{metadataDoc?.ReferenceNo}</span>
                  </div>
                  <div className="meta-panel-field">
                    <label className="meta-panel-label">{DisplayLabel.FolderPath}</label>
                    <div className="meta-panel-plain-value" data-testid="text-meta-name">{metadataDoc?.FolderDocumentPath || ""}</div>
                  </div>
                  
                </div>

                <div className="meta-panel-row meta-panel-row-2col">
                  <div className="meta-panel-field">
                    <label className="meta-panel-label">{DisplayLabel.AuthorName}</label>
                    <div className="meta-panel-plain-value" data-testid="text-meta-name">{metadataDoc?.Author.Title || ""}</div>
                  </div>
                  <div className="meta-panel-field">
                    <label className="meta-panel-label">{DisplayLabel.CreatedDate}</label>
                    {/* <span className="meta-panel-plain-value" data-testid="text-meta-tile">{metadataDoc?.Created}</span> */}
                    <span className="meta-panel-plain-value" data-testid="text-meta-tile">
                      {metadataDoc?.Created
                        ? format(new Date(metadataDoc.Created), "dd/MM/yyyy")
                        : ""}
                    </span>
                  </div>
               
                </div>

                {/* <div className="meta-panel-field">
                  <label className="meta-panel-label">{DisplayLabel.IsSuffixRequired}</label>
                  <Toggle
                    checked={metadataDoc?.IsSuffixRequired}
                    disabled
                    data-testid="toggle-meta-suffix"
                  />
                </div> */}
                {metadataDoc?.IsSuffixRequired && (
                  <>
                    <div className="meta-panel-field">
                      <label className="meta-panel-label">{DisplayLabel.DocumentSuffix}</label>
                      <div className="meta-panel-plain-value" data-testid="text-meta-name">{metadataDoc?.DocumentSuffix || ""}</div>
                    </div>

                    {metadataDoc?.DocumentSuffix === "Other" && (
                      <div className="meta-panel-field">
                        <label className="meta-panel-label">{DisplayLabel.DocumentSuffix}</label>
                        <div className="meta-panel-plain-value" data-testid="text-meta-name">{metadataDoc?.OtherSuffix || ""}</div>
                      </div>
                    )}
                  </>
                )}
              </div>
            </div>

            <div className="meta-panel-section">
              <h3 className="meta-panel-section-title">Classification</h3>
              <div className="meta-panel-row meta-panel-row-2col">
                {renderDynamicControls()}
              </div>
            </div>

               <div className="meta-panel-section">
              <h3 className="meta-panel-section-title">Document</h3>
              <div
                className="meta-panel-doc-summary meta-panel-doc-summary-hover"
                onClick={() => openDocument(metadataDoc)}
                style={{ cursor: "pointer" }}
                data-testid="button-open-document"
              >
                {(() => {
                  const docExt = metadataDoc.File.Name.split(".").pop();
                  const docFileConfig = fileTypeConfig[docExt] || fileTypeConfig.other;
                  const { IconName: DocIcon, className: docIconClass } = docFileConfig;
                  return (
                    <div className="meta-panel-doc-left">
                      <div className={`doc-icon-wrap ${docIconClass}`}>
                        <DocIcon className="doc-icon-svg" />
                      </div>
                      <div className="meta-panel-doc-info">
                        <span className="meta-panel-doc-name" data-testid="text-meta-doc-name">
                          {metadataDoc?.ActualName}
                        </span>
                        <span className="meta-panel-doc-ref">
                          {metadataDoc.referenceNo} &middot; v
                          {getDocumentVersion(metadataDoc?.Level, metadataDoc?.OData__UIVersionString) || "-"}
                        </span>
                      </div>
                    </div>
                  );
                })()}
                <button
                  type="button"
                  className="meta-panel-doc-download-btn"
                  onClick={(e) => {
                    e.stopPropagation();
                    window.open(`${metadataDoc?.File?.ServerRelativeUrl}?download=1`, "_blank");
                  }}
                  title="Download"
                  aria-label="Download"
                  data-testid="button-download-doc"
                >
                  <ArrowDownload20Regular className="meta-panel-doc-download-icon" />
                </button>
              </div>
            </div>

            {/* <div className="meta-panel-section">
              <h3 className="meta-panel-section-title">Workflow</h3>
              <div className="meta-panel-fields">
                <div className="meta-panel-field">
                  <label className="meta-panel-label">Is Approval flow required?</label>
               
                  <Toggle
                    checked={metadataDoc.AllowApprover}
                    disabled
                    data-testid="toggle-meta-approval-flow"
                  />
                </div>
              </div>
            </div> */}
            
          </div>
        )}
        
      </Panel>
      {isDialogOpen && (
        <Layer>
          <Popup
            role="dialog"
            aria-modal="true"
            // onDismiss={() => setIsDialogOpen(false) }
             onDismiss={() => { setIsDialogOpen(false); setComment(""); }}
             className={popupStyles.overlay}
          >
            <FocusTrapZone>
              <div role="document" className={popupStyles.content}>
                <div className={popupStyles.title}>
                  {POPUP_CONFIG[actions]?.title}
                </div>
                <Text className={popupStyles.description}>
                  {POPUP_CONFIG[actions]?.description}
                </Text>


                <TextField
                  multiline
                  rows={4}
                  value={comment}
                  onChange={(_, val) => setComment(val || "")}
                  className={popupStyles.textarea}
                />

                <div className={popupStyles.footer}>
                  <DefaultButton
                    text="Cancel"
                    onClick={() => { setIsDialogOpen(false); setComment(""); }}
                  />
                  <PrimaryButton
                    text={actions === "APPROVE" ? "Approve Request" : "Reject Request"}
                    onClick={() => { actions === "APPROVE" ? Approve() : RejectFile(); }}
                    disabled={!comment.trim()}
                  />
                </div>
              </div>
            </FocusTrapZone>
          </Popup>
        </Layer>
      )}
    </div>
  );
}
