/* eslint-disable */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { DefaultButton, PrimaryButton, PanelType, Panel, DialogType, TextField, TooltipHost, DirectionalHint, Spinner, SpinnerSize, ActionButton, Dropdown, Toggle, IDropdownStyles, IDropdownOption, ITextFieldStyles, IToggleStyles, IToggleStyleProps, IBasePickerStyles } from '@fluentui/react';
import { ArrowUpload20Regular, FolderAdd20Regular, Add20Regular, Home20Regular, ChevronRight12Regular, MoreHorizontalRegular, ChevronRight24Regular, ChevronDown24Regular } from '@fluentui/react-icons';
import Sidebar from "../../common/component/Sidebar";
import { FolderNode } from "../../common/component/FolderTree";
import { buildBreadcrumbPath, buildFolderHierarchy, buildLibraryRootPath, checkButtons, checkExtension, fileTypeConfig, getAllDocuments, getChildFolders, getDocumentVersion, getOpenAppURL, resolveVersionHistoryTarget } from "../../common/commonfunction";
import ReusableDataTable from "../ResuableComponents/ReusableDataTable";
import { spfi, PermissionKind } from "@pnp/sp/presets/all";
import { SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/security";
import { ILabel } from "../../../../Intrface/ILabel";
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import * as FluentIcons from "@fluentui/react-icons";
import { Icon } from '@fluentui/react';
import {
    Menu,
    MenuTrigger,
    MenuPopover,
    MenuList,
    MenuItem,
    Button,
    Badge,
    Input,
    Label,
    Field
} from "@fluentui/react-components";
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { getAllButtons } from "../../../../Services/Buttons";
import { IButtonsProps, IRolePermission } from "../../../../Intrface/IButtonInterface";
import { checkPermissions, commonPostMethod, getApprovalData, getArchiveData, getListData, hasFolderPermission, updateLibrary } from "../../../../Services/GeneralDocument";
import { getHistoryByID } from "../../../../Services/GeneralDocHistoryService";
import { format } from "date-fns";
import { getConfigActive } from "../../../../Services/ConfigService";
import { getDataByLibraryName } from "../../../../Services/MasTileService";
import IFrameDialogPopup from "../../common/component/IFrameDialog";
import AdvancePermission from "../../common/component/AdvancePermission";
import PopupBox, { ConfirmationDialog } from "../../common/component/PopupBox";
import { FolderStructure } from "../../../../Services/FolderStructure";
import { isMember } from "../../../../DAL/Commonfile";
import ProjectEntryForm from "../../common/component/ProjectEntryForm";
import UploadFiles from "../../common/component/UploadFile";
import ApprovalFlow from "../../common/component/ApprovalFlow";
import PageLoader from "../../common/component/PageLoader";
import { getPermittedButtons } from "./ButtonsService";
import { IDmsButton } from "./buttonPermissionHelper";



interface IWorkspaceProps {
    context: WebPartContext;
}
/**
 * Fields returned by SharePoint's listItemAllFields endpoint.  Folder
 * libraries can contain tenant-specific columns, so those values must remain
 * extensible while the tree fields retain their known types.
 */
interface FolderMetadata {
    Id?: number;
    ID?: number;
    Title?: string;
    ProjectmanagerId?: number;
    PublisherId?: number;
    [fieldName: string]: unknown;
}

/** A tree node enriched with all of its backing list item's fields. */
interface SelectedFolder extends FolderNode, FolderMetadata {
    children: FolderNode[];
}

interface Folder extends FolderNode {
    children?: FolderNode[];
}
/* ------------------------------------------------------------------ */
/* Reusable view-only styles - light, non-editable, form-consistent    */
/* ------------------------------------------------------------------ */
// const viewOnlyBackground = "#fafafa";
// const viewOnlyTextColor = "#323130";
// const viewOnlyBorderColor = "#d1d1d1";

// const viewOnlyTextFieldStyles: Partial<ITextFieldStyles> = {
//     root: { width: "100%" },
//     fieldGroup: {
//         backgroundColor: viewOnlyBackground,
//         borderColor: viewOnlyBorderColor,
//         borderRadius: 4,
//         selectors: {
//             ":hover": { borderColor: viewOnlyBorderColor },
//             "&.is-focused::after": { border: `1px solid ${viewOnlyBorderColor}` },
//         },
//     },
//     field: {
//         backgroundColor: viewOnlyBackground,
//         color: viewOnlyTextColor,
//         fontWeight: 400,
//     },
//     subComponentStyles: {
//         label: { root: { color: viewOnlyTextColor, fontWeight: 600 } },
//     },
// };

// const viewOnlyDropdownStyles: Partial<IDropdownStyles> = {
//     root: { width: "100%" },
//     label: { color: viewOnlyTextColor, fontWeight: 600 },
//     title: {
//         backgroundColor: viewOnlyBackground,
//         borderColor: viewOnlyBorderColor,
//         color: viewOnlyTextColor,
//         height: 32,
//         lineHeight: 30,
//         fontSize: 14,
//         borderRadius: 4,
//         selectors: {
//             ":hover": { borderColor: viewOnlyBorderColor },
//         },
//     },
//     caretDownWrapper: { lineHeight: 30 },
//     caretDown: { color: "#605e5c", fontSize: 14 },
// };

// const viewOnlyPeoplePickerStyles: Partial<IBasePickerStyles> = {
//     root: { width: "100%" },
//     text: {
//         backgroundColor: viewOnlyBackground,
//         borderColor: viewOnlyBorderColor,
//         borderRadius: 4,
//         selectors: {
//             ":hover": { borderColor: viewOnlyBorderColor },
//         },
//     },
//     itemsWrapper: { backgroundColor: viewOnlyBackground },
//     input: { backgroundColor: viewOnlyBackground, color: viewOnlyTextColor },
// };

// const viewOnlyToggleStyles = (props: IToggleStyleProps): IToggleStyles => ({
//     root: { marginBottom: 0 },
//     label: { color: viewOnlyTextColor, fontWeight: 600 },
//     container: { marginTop: 4 },
//     pill: {
//         backgroundColor: props.checked ? "#0078d4" : "#f3f2f1",
//         borderColor: viewOnlyBorderColor,
//     },
//     thumb: { backgroundColor: "#ffffff", borderColor: viewOnlyBorderColor },
//     text: { color: viewOnlyTextColor },
// });

const formatDateValue = (value: any): string => {
    if (!value) return "";
    try {
        const date = new Date(value);
        if (isNaN(date.getTime())) return "";
        return format(date, "dd/MM/yyyy");
    } catch {
        return "";
    }
};

const normalizeMultiValues = (value: any): string[] => {
    if (value == null || value === "") return [];
    if (Array.isArray(value)) return value.map((v: any) => String(v));
    const text = String(value);
    if (text.indexOf(";#") > -1) return text.split(";#").filter((v: string) => v !== "");
    return [text];
};

/**
 * Renders a dynamic metadata field of the document View panel as a read-only
 * Fluent UI control (matching the Create/Edit form layout) based on its column
 * type: Text, Choice (Dropdown/Radio/Multiple Select), Date and Time, Person or
 * Group, Yes/No and Multi-line text.
 */
// const renderViewOnlyMetaField = (
//     el: any,
//     filterObj: any,
//     item: any,
//     usersById: Map<number, any>,
//     choiceOptionsMap: { [key: string]: IDropdownOption[] },
//     peoplePickerContext: IPeoplePickerContext
// ): React.ReactElement | null => {
//     if (!filterObj) return null;

//     const columnType = filterObj.ColumnType || el.ColumnType || "Single line of Text";
//     const allFields = item?.ListItemAllFields || {};
//     const hasValue = allFields.hasOwnProperty(el.InternalTitleName);
//     const raw = hasValue ? allFields[el.InternalTitleName] : undefined;
//     const fieldTitle = el.Title || filterObj.Title || "";
//     const fieldKey = el.Id ?? el.InternalTitleName ?? "meta-field";

//     const buildOptions = (currentValues: string[]): IDropdownOption[] => {
//         const fetched = choiceOptionsMap[el.InternalTitleName];
//         if (fetched && fetched.length > 0) return fetched;
//         if (el.IsStaticValue || filterObj.IsStaticValue) {
//             const staticData = el.StaticDataObject || filterObj.StaticDataObject || "";
//             return staticData
//                 .split(";")
//                 .filter((v: string) => v !== "")
//                 .map((v: string) => ({ key: v, text: v }));
//         }
//         if (currentValues.length > 0) {
//             return currentValues.map((v: string) => ({ key: v, text: v }));
//         }
//         return [{ key: "", text: "" }];
//     };

//     let control: React.ReactElement;

//     switch (columnType) {
//         case "Date and Time": {
//             control = (
//                 <TextField
//                     label={fieldTitle}
//                     readOnly
//                     value={formatDateValue(raw)}
//                     styles={viewOnlyTextFieldStyles}
//                 />
//             );
//             break;
//         }
//         case "Person or Group": {
//             const person = raw && typeof raw === "object" ? raw : null;
//             const personId = person?.Id ?? (typeof raw === "number" ? raw : null);
//             const siteUser = personId != null ? usersById.get(Number(personId)) : undefined;
//             const personEmail = siteUser?.Email || person?.EMail || person?.LoginName || "";
//             control = personEmail ? (
//                 <PeoplePicker
//                     titleText={fieldTitle}
//                     context={peoplePickerContext}
//                     personSelectionLimit={20}
//                     showtooltip={false}
//                     showHiddenInUI={false}
//                     principalTypes={[PrincipalType.User]}
//                     defaultSelectedUsers={[personEmail]}
//                     disabled
//                     styles={viewOnlyPeoplePickerStyles}
//                 />
//             ) : (
//                 <TextField
//                     label={fieldTitle}
//                     readOnly
//                     value={person?.Title || siteUser?.Title || ""}
//                     styles={viewOnlyTextFieldStyles}
//                 />
//             );
//             break;
//         }
//         case "Dropdown":
//         case "Radio": {
//             const singleValue = raw == null ? "" : String(raw);
//             control = (
//                 <Dropdown
//                     label={fieldTitle}
//                     options={buildOptions([singleValue])}
//                     selectedKey={singleValue}
//                     disabled
//                     styles={viewOnlyDropdownStyles}
//                 />
//             );
//             break;
//         }
//         case "Multiple Select": {
//             const multiValues = normalizeMultiValues(raw);
//             control = (
//                 <Dropdown
//                     label={fieldTitle}
//                     multiSelect
//                     options={buildOptions(multiValues)}
//                     selectedKeys={multiValues}
//                     disabled
//                     styles={viewOnlyDropdownStyles}
//                 />
//             );
//             break;
//         }
//         case "Yes/No": {
//             const checked = raw === true || raw === 1 || raw === "Yes" || raw === "true" || raw === "1";
//             control = (
//                 <Toggle
//                     label={fieldTitle}
//                     checked={checked}
//                     disabled
//                     styles={viewOnlyToggleStyles}
//                 />
//             );
//             break;
//         }
//         case "Multiple lines of Text": {
//             control = (
//                 <TextField
//                     label={fieldTitle}
//                     multiline
//                     readOnly
//                     value={raw == null ? "" : String(raw)}
//                     styles={viewOnlyTextFieldStyles}
//                 />
//             );
//             break;
//         }
//         default: {
//             control = (
//                 <TextField
//                     label={fieldTitle}
//                     readOnly
//                     value={raw == null ? "" : String(raw)}
//                     styles={viewOnlyTextFieldStyles}
//                 />
//             );
//             break;
//         }
//     }

//     return (
//         <div className="col-md-6" key={fieldKey}>
//             {control}
//         </div>
//     );
// };

const renderViewOnlyMetaField = (
    el: any,
    filterObj: any,
    item: any,
    usersById: Map<number, any>,
    choiceOptionsMap: { [key: string]: IDropdownOption[] },
    peoplePickerContext: IPeoplePickerContext
): React.ReactElement | null => {
    if (!filterObj) return null;

    const columnType = filterObj.ColumnType || el.ColumnType || "Single line of Text";
    const allFields = item?.ListItemAllFields || {};
    const hasValue = allFields.hasOwnProperty(el.InternalTitleName);
    const raw = hasValue ? allFields[el.InternalTitleName] : undefined;
    const fieldTitle = el.Title || filterObj.Title || "";
    const fieldKey = el.Id ?? el.InternalTitleName ?? "meta-field";

    const buildOptions = (currentValues: string[]): IDropdownOption[] => {
        const fetched = choiceOptionsMap[el.InternalTitleName];
        if (fetched && fetched.length > 0) return fetched;
        if (el.IsStaticValue || filterObj.IsStaticValue) {
            const staticData = el.StaticDataObject || filterObj.StaticDataObject || "";
            return staticData
                .split(";")
                .filter((v: string) => v !== "")
                .map((v: string) => ({ key: v, text: v }));
        }
        if (currentValues.length > 0) {
            return currentValues.map((v: string) => ({ key: v, text: v }));
        }
        return [{ key: "", text: "" }];
    };

    let displayValue: string;

    switch (columnType) {
        case "Date and Time": {
            // Guard against empty/null/undefined raw values so they don't
            // fall through to new Date(null) -> 01/01/1970.
            displayValue = raw ? formatDateValue(raw) : "";
            break;
        }
        case "Person or Group": {
            const person = raw && typeof raw === "object" ? raw : null;
            const personId = person?.Id ?? (typeof raw === "number" ? raw : null);
            const siteUser = personId != null ? usersById.get(Number(personId)) : undefined;
            displayValue = person?.Title || siteUser?.Title || "";
            break;
        }
        case "Dropdown":
        case "Radio": {
            const singleValue = raw == null ? "" : String(raw);
            const options = buildOptions([singleValue]);
            const matched = options.find((o) => String(o.key) === singleValue);
            displayValue = matched?.text || singleValue;
            break;
        }
        case "Multiple Select": {
            const multiValues = normalizeMultiValues(raw);
            const options = buildOptions(multiValues);
            displayValue = multiValues
                .map((v) => options.find((o) => String(o.key) === v)?.text || v)
                .join(", ");
            break;
        }
        case "Yes/No": {
            const checked = raw === true || raw === 1 || raw === "Yes" || raw === "true" || raw === "1";
            displayValue = checked ? "Yes" : "No";
            break;
        }
        case "Multiple lines of Text": {
            displayValue = raw == null ? "" : String(raw);
            break;
        }
        default: {
            displayValue = raw == null ? "" : String(raw);
            break;
        }
    }

    return (
        <div className="col-md-6" key={fieldKey}>
            <label className="view-only-label">{fieldTitle}</label>
            <div className="view-only-value">{displayValue}</div>
        </div>
    );
};



const Workspace: React.FunctionComponent<IWorkspaceProps> = ({ context }) => {
    const SiteURL = context.pageContext.web.absoluteUrl;
    const UserID = context.pageContext.legacyPageContext.userId;
    const UserEmailID = context.pageContext.user.email;
    const portalUrl = new URL(context.pageContext.web.absoluteUrl).origin;
    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: SiteURL,
        msGraphClientFactory: context.msGraphClientFactory as any,
        spHttpClient: context.spHttpClient as any,
    };
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const { workspaceId } = useParams<{ workspaceId: string; }>();
    const navigate = useNavigate();
    const [selectedFolder, setSelectedFolder] = useState<SelectedFolder | null>(null);
    const [folders, setFolders] = useState<any>([]);
    const [tileData, setTileData] = useState<any | null>(null);
    const [files, setFiles] = useState<any[]>([]);
    const [buttons, setButtons] = useState<any[]>([]);
    const [itemId, setItemId] = useState<number>(0);
    const [message, setMessage] = useState<string>("");
    const [hideDialog, setHideDialog] = useState<boolean>(false);
    const [actionButton, setActionButton] = useState<React.ReactNode>(null);
    const [panelForm, setPanelForm] = useState<React.ReactNode>(null);
    const [panelTitle, setPanelTitle] = useState("");
    const [isOpenCommonPanel, setIsOpenCommonPanel] = useState(false);
    const [extension, setExtension] = useState("");
    const [fileName, setFileName] = useState("");
    const [fileNameErr, setFileNameErr] = useState("");
    const [panelSize, setPanelSize] = useState(PanelType.medium);
    const [alertMsg, setAlertMsg] = useState("");
    const [isPopupBoxVisible, setIsPopupBoxVisible] = useState<boolean>(false);
    const [comment, setComment] = useState("");
    const [serverRelativeUrl, setServerRelativeUrl] = useState("");
    const [hideDialogCheckOut, setHideDialogCheckOut] = useState<boolean>(false);
    const [isPanelOpen, setIsPanelOpen] = useState(false);
    const [shareURL, setShareURL] = useState("");
    const [iFrameDialogOpened, setIFrameDialogOpened] = useState(false);
    const [isShowCommnPopupBoxVisible, setIsShowCommnPopupBoxVisible] = useState<boolean>(false);
    const [isOpenFolderPanel, setIsOpenFolderPanel] = useState(false);
    const [folderNameErr, setFolderNameErr] = useState("");
    const [folderName, setFolderName] = useState("");
    const invalidCharsRegex = /["*:<>?/\\|]/;
    const [admin, setAdmin] = useState([]);
    const [isValidUser, setIsValidUser] = useState<boolean>(false);
    const [isCreateProjectPopupOpen, setIsCreateProjectPopupOpen] = useState(false);
    const [isOpenUploadPanel, setIsOpenUploadPanel] = useState(false);
    const [fileType, setFileType] = useState<string>("");
    const [formType, setFormType] = useState("EntryForm");
    const [tables, setTables] = useState("");
    const [viewListSetting, setViewListSetting] = useState("");
    const [deletedData, setDeletedData] = useState<any>([]);
    const [approvalData, setApprovalData] = useState<any>([]);
    const [archiveData, setArchiveData] = useState<any>([]);
    const [projectUpdateData, setProjectUpdateData] = useState<any>({});
    const [permittedButtons, setPermittedButtons] = useState<IDmsButton[]>([]);
    const [userPerms, setUserPerms] = useState({
        FullControl: false,
        Contribute: false,
        Edit: false,
        Read: false
    });
    const buttonsCache = useRef<any[] | null>(null);
    const [hasPermission, setHasPermission] = useState<boolean>(false);
    const [isRestrictedView, setIsRestrictedView] = useState(false);
    const [expandedFolders, setExpandedFolders] = useState<string[]>([]);
    const [isVersionsLoading, setIsVersionsLoading] = useState(false);
    const [versionsPanelUrl, setVersionsPanelUrl] = useState("");
    const [isVersionsOverlayVisible, setIsVersionsOverlayVisible] = useState(false);
    const [isWorkspaceLoading, setIsWorkspaceLoading] = useState(true);
    const selectedFolderRef = useRef<SelectedFolder | null>(null);
    const folderSelectionRequestRef = useRef(0);
    const [isFolderMetadataLoading, setIsFolderMetadataLoading] = useState(false);
    const [folderMetadataError, setFolderMetadataError] = useState<string | null>(null);
    const [popupType, setPopupType] = useState<"success" | "warning" | "insert" | "checkin" | "checkout" | "approve" | "reject" | "delete" | "update" | "restore" | "grant" | "remove">("success");

    
    // const canCreateRequest = useMemo(() => {
    //     return isValidUser || tileData?.TileAdminId === UserID;
    // }, [isValidUser, tileData, UserID]);

    const canCreateRequest = useMemo(() => {
        return (
            isValidUser ||
            tileData?.TileAdminId?.includes?.(Number(UserID))
        );
    }, [isValidUser, tileData, UserID]);

    const ShowHideDeleteOption = useMemo(() => {
        return isValidUser || tileData?.TileAdminId === UserID;
    }, [isValidUser, tileData, UserID]);

    //Added New
    const [canShowButtons, setCanShowButtons] = useState(false);

     //Show Child Folder creation button
    const [hasSelectedFolder, setHasSelectedFolder] = useState(false);


    // Workspace page component — runs every time this page mounts
    useEffect(() => {
        const resetScroll = () => {
            // reset window scroll
            window.scrollTo(0, 0);

            // reset SharePoint's actual scroll container(s)
            const selectors = [
                '#spPageCanvasContent',
                '.SPPageChrome',
                '[data-automation-id="contentScrollRegion"]',
            ];
            selectors.forEach(sel => {
                const el = document.querySelector(sel) as HTMLElement | null;
                if (el) el.scrollTop = 0;
            });

            // force sticky/scroll-aware components (Fluent UI) to recompute position
            window.dispatchEvent(new Event('scroll'));
            window.dispatchEvent(new Event('resize'));
        };

        // run once immediately
        resetScroll();

        // run again after layout settles (handles late-mounting sticky elements)
        const rafId = requestAnimationFrame(resetScroll);

        // and once more after a short delay, in case data/content loads async
        // and pushes the sticky bar's offset after our first two attempts
        const timeoutId = setTimeout(resetScroll, 150);

        return () => {
            cancelAnimationFrame(rafId);
            clearTimeout(timeoutId);
        };
    }, []);

    useEffect(() => {
        void fetchTileData();
        void getAdmin();
    }, []);

       // Reset the "user explicitly picked a folder" flag whenever the workspace
    // route changes so the "New Folder" button never carries over from a
    // previously opened workspace/folder.
    useEffect(() => {
        setHasSelectedFolder(false);
    }, [workspaceId]);


    const fetchTileData = async () => {
        const sp = spfi().using(SPFx(context));
        const data = await sp.web.lists.getByTitle("DMS_Mas_Tile").select("*").items.getById(Number(workspaceId))();
        setTileData(data);
    };

    useEffect(() => {
        if (tileData) {
            setIsWorkspaceLoading(true);
            getPermittedButtons(context, tileData.LibraryName).then(setPermittedButtons);
            Promise.all([
                fetchFolder(),
                getDeletedData(),
                getArchiveFile(),
            ]).finally(() => setIsWorkspaceLoading(false));
        }
    }, [tileData]);

    useEffect(() => {
        if (!tileData?.LibraryName) return;
        getPendingApprovalData();
    }, [isOpenUploadPanel, tileData]);

    // useEffect(() => {
    //     if (selectedFolder?.path) {

    //         void fetchButtonsAndPermissions(selectedFolder.path);
    //     }
    // }, [selectedFolder?.path]);

    //Added New
    useEffect(() => {
        if (selectedFolder?.path && tileData?.LibraryName) {
            void fetchButtonsAndPermissions(selectedFolder.path);
        }
    }, [selectedFolder?.path, tileData?.LibraryName]);

    useEffect(() => {
        selectedFolderRef.current = selectedFolder;
    }, [selectedFolder]);

    const findFolderByPath = (nodes: any[], targetPath?: string): any | null => {
        if (!targetPath) return null;

        for (const node of nodes) {
            if (node.path === targetPath) {
                return node;
            }

            if (node.children?.length) {
                const matchedNode = findFolderByPath(node.children, targetPath);
                if (matchedNode) {
                    return matchedNode;
                }
            }
        }

        return null;
    };

    /**
     * Commits a folder selection only after its backing list item has been
     * loaded. The request id prevents a slower earlier click from replacing a
     * more recent selection.
     */
    const selectFolderWithMetadata = useCallback(async (folder: FolderNode): Promise<{ folder: SelectedFolder; requestId: number; } | null> => {
        const requestId = ++folderSelectionRequestRef.current;
        setIsFolderMetadataLoading(true);
        setFolderMetadataError(null);

        try {
            const sp = spfi().using(SPFx(context));
            const fieldsData = await sp.web
                .getFolderByServerRelativePath(folder.path)
                .listItemAllFields() as FolderMetadata;

            // TEMP-DIAG: verify what listItemAllFields returns for the selected folder. Remove after confirming.
            console.log("[DMS-TEMP] Folder Metadata:", fieldsData);

            if (!fieldsData || typeof fieldsData !== "object" || Array.isArray(fieldsData)) {
                throw new Error("SharePoint returned invalid folder metadata.");
            }

            // A newer selection has started while this request was in flight.
            if (requestId !== folderSelectionRequestRef.current) return null;

            const enrichedFolder: SelectedFolder = {
                ...folder,
                ...fieldsData,
                children: folder.children || []
            };

            selectedFolderRef.current = enrichedFolder;
            setSelectedFolder(enrichedFolder);
            return { folder: enrichedFolder, requestId };
        } catch (error) {
            if (requestId === folderSelectionRequestRef.current) {
                console.error("Unable to load selected folder metadata:", error);
                setFolderMetadataError("Unable to load the selected folder's metadata. Please try again.");
            }
            return null;
        } finally {
            if (requestId === folderSelectionRequestRef.current) {
                setIsFolderMetadataLoading(false);
            }
        }
    }, [context]);

    // const collectExpandedFolderIds = (nodes: any[], targetPath?: string, parents: string[] = []): string[] => {
    //     if (!targetPath) return [];

    //     for (const node of nodes) {
    //         const currentBranch = [...parents, String(node.id)];
    //         if (node.path === targetPath) {
    //             return currentBranch;
    //         }

    //         if (node.children?.length) {
    //             const childBranch = collectExpandedFolderIds(node.children, targetPath, currentBranch);
    //             if (childBranch.length) {
    //                 return childBranch;
    //             }
    //         }
    //     }

    //     return [];
    // };



    //Original Code
    // const fetchFolder = async () => {
    //     const sp = spfi().using(SPFx(context));

    //     const allFolders: any[] = [];

    //     const items = await sp.web.lists
    //         .getByTitle(tileData?.LibraryName)
    //         .items
    //         .select("*", "Id", "Title", "FileRef", "FileDirRef", "FSObjType")
    //         .filter("FSObjType eq 1")
    //         .top(5000);

    //     for await (const batch of items) {
    //         allFolders.push(...batch);
    //     }

    //     const rootPath = buildLibraryRootPath(context, tileData?.LibraryName);
    //     const folder = buildFolderHierarchy(allFolders, rootPath);
    //     const folderObj = {
    //         id: 0,
    //         name: tileData?.LibraryName,
    //         path: rootPath,
    //         children: [...folder]
    //     };
    //     const nextFolders = [folderObj];
    //     const preservedFolder = findFolderByPath(nextFolders, selectedFolderRef.current?.path) || folderObj;
    //     setFolders(nextFolders);
    //     expandParentFolders(folderObj);
    //     setSelectedFolder(preservedFolder);
    // };

    const updateFolderNodeState = (targetPath: string, updates: Partial<any>) => {
        setFolders((prevFolders: any[]) => {
            const updateNode = (nodes: any[]): any[] => {
                return nodes.map(node => {
                    if (node.path === targetPath) {
                        return { ...node, ...updates };
                    }
                    if (node.children && node.children.length > 0) {
                        return { ...node, children: updateNode(node.children) };
                    }
                    return node;
                });
            };
            return updateNode(prevFolders);
        });
    };

    const loadChildFolders = async (parentFolder: any): Promise<any[]> => {
        if (!parentFolder || !tileData?.LibraryName) return [];

        if (parentFolder.isLoaded && parentFolder.children) {
            return parentFolder.children;
        }

        updateFolderNodeState(parentFolder.path, { isLoading: true });

        const childFolders = await getChildFolders(context, parentFolder.path);

        const updatedNodeProps = {
            children: childFolders,
            isLoaded: true,
            isLoading: false,
            isLastLevel: childFolders.length === 0
        };

        Object.assign(parentFolder, updatedNodeProps);
        updateFolderNodeState(parentFolder.path, updatedNodeProps);

        return childFolders;
    };

    // Refreshes only the currently selected folder's children.
    // Unlike fetchFolder (which rebuilds the whole tree from root),
    // this preserves the nested folder structure and keeps the user
    // in the same location after creating a folder.
    const refreshCurrentFolder = async () => {
        const folderToRefresh = selectedFolderRef.current || selectedFolder;
        if (!folderToRefresh || !tileData?.LibraryName) return;
        const selectionRequestId = folderSelectionRequestRef.current;

        // Force a re-fetch of the selected folder's children
        const childFolders = await getChildFolders(context, folderToRefresh.path);
        if (selectionRequestId !== folderSelectionRequestRef.current) return;

        const updatedNodeProps = {
            children: childFolders,
            isLoaded: true,
            isLoading: false,
            isLastLevel: childFolders.length === 0
        };

        Object.assign(folderToRefresh, updatedNodeProps);
        updateFolderNodeState(folderToRefresh.path, updatedNodeProps);

        // Keep the selection pointing to the refreshed folder object
        const refreshedFolder: SelectedFolder = {
            ...folderToRefresh,
            ...updatedNodeProps
        };
        selectedFolderRef.current = refreshedFolder;
        setSelectedFolder(refreshedFolder);

        if (childFolders.length === 0) {
            await getDocument(folderToRefresh, selectionRequestId);
        }
    };

    const fetchFolder = async () => {

        if (!tileData?.LibraryName) return;

           // Root is AUTO-selected here on page/tile load, not by an explicit user
        // click, so the "New Folder" button must stay hidden until the user
        // actually opens a folder.
        setHasSelectedFolder(false);

        const rootPath = buildLibraryRootPath(context, tileData?.LibraryName);
        const rootChildFolders = await getChildFolders(context, rootPath);

        const rootFolderObj = {
            id: "0",
            name: tileData?.LibraryName,
            path: rootPath,
            children: rootChildFolders,
            isLoaded: true,
            isLoading: false,
            isLastLevel: rootChildFolders.length === 0,
            FileRef: rootPath,
            FileDirRef: "",
            FSObjType: 1
        };

        const nextFolders = [rootFolderObj];
        const preservedFolder =
            findFolderByPath(nextFolders, selectedFolderRef.current?.path) || rootFolderObj;

        setFolders(nextFolders);
        expandParentFolders(rootFolderObj);
        const selection = await selectFolderWithMetadata(preservedFolder);
        if (selection?.folder.children.length === 0) {
            await getDocument(selection.folder, selection.requestId);
        }
    };

    const getAdmin = async () => {
        const data = await getListData(
            `${SiteURL}/_api/web/lists/getbytitle('DMS_GroupName')/items?`,
            context
        );
        setAdmin(data.value.map((el: any) => el.GroupNameId));
        try {
            const isMembers = await isMember(context, "ProjectAdmin");
            setIsValidUser(
                isMembers?.value?.length > 0
            );
        } catch (error) {
            console.log("User is not a member or access denied:", error);
            setIsValidUser(false);
        }
    };

    const getDeletedData = async () => {
        const deletedData = await getListData(`${SiteURL}/_api/web/lists/getbytitle('${tileData?.LibraryName}')/items?$filter=DeleteFlag eq 'Deleted' and Active eq 0`, context);
        setDeletedData(deletedData.value);
    };

    const handleFolderSelect = async (folder: FolderNode) => {
        const selection = await selectFolderWithMetadata(folder);
        if (!selection) return;

        const { folder: selectedFolderWithMetadata, requestId } = selection;

          // The user explicitly clicked into this folder (tile, list row, sidebar
        // tree node, or breadcrumb segment). Only now may the "New Folder"
        // button become eligible to show.
        setHasSelectedFolder(true);

        setTables(""); // <-- Reset from Archive/Recycle to Folder view
        expandParentFolders(selectedFolderWithMetadata);

        let children = selectedFolderWithMetadata.children || [];
        if (!selectedFolderWithMetadata.isLoaded) {
            children = await loadChildFolders(selectedFolderWithMetadata);
        }

        // Do not let an earlier selection update documents after a later click.
        if (requestId !== folderSelectionRequestRef.current) return;

        if (children.length === 0) {
            await getDocument(selectedFolderWithMetadata, requestId);
        } else {
            setFiles([]);
        }
    };


    const getPendingApprovalData = async () => {
        if (!tileData?.LibraryName) return;
        const pendingApprovalData = await getApprovalData(context, tileData.LibraryName, UserEmailID);
        setApprovalData(pendingApprovalData.value);
    };
    const getArchiveFile = async () => {
        const data = await getArchiveData(context, tileData?.LibraryName);
        setArchiveData(data.value || []);
    };

    /**
     * Resolves the full folder metadata (listItemAllFields) BEFORE opening the
     * ProjectEntryForm so bindFormData() can bind every field. The sidebar tree
     * node alone only carries id/name/path/children, so metadata-bound inputs
     * would otherwise render empty.
     */
    const openFolderEntryForm = async (folder: FolderNode, formType: string) => {
        try {
            let enrichedFolder: any = folder;

            const selected = selectedFolderRef.current;
            if (selected?.path === folder.path) {
                // Already enriched (the currently selected folder) — reuse it.
                enrichedFolder = { ...folder, ...selected };
            } else {
                const sp = spfi().using(SPFx(context));
                const fieldsData = await sp.web
                    .getFolderByServerRelativePath(folder.path)
                    .listItemAllFields();
                enrichedFolder = { ...folder, ...fieldsData };
            }

            setProjectUpdateData(enrichedFolder);
        } catch (error) {
            console.error("Unable to load folder metadata for the entry form:", error);
            // Fall back to the tree node so the form still opens with the folder name.
            setProjectUpdateData(folder);
        } finally {
            setFormType(formType);
            setIsCreateProjectPopupOpen(true);   // open only AFTER data is set
        }
    };

    const handleFolderAction = (action: string, folder: FolderNode) => {
        console.log('Folder action:', action, folder.name);

        // const folderPath = selectedFolder?.path?.replace(context.pageContext.web.serverRelativeUrl, "")?.replace(/^\/+/, "");
        switch (action) {
            case "FView":
                void openFolderEntryForm(folder, "ViewForm");
                break;
            case "FEdit":
                void openFolderEntryForm(folder, "EditForm");
                break;
            case "AdvancePermission":
                setItemId(Number(folder.id)); setIsPanelOpen(true);
                break;
            case "Share":
                setShareURL(`${SiteURL}/_layouts/15/sharedialog.aspx?listId=${tileData?.LibGuidName}&listItemId=${folder.id}&clientId=sharePoint&policyTip=0&folderColor=undefined&ma=0&fullScreenMode=true&itemName=${folder.name}&origin=${portalUrl}`);
                setIFrameDialogOpened(true);
                break;
        }
    };

    const getRestrictedUserData = async (folderPath: string) => {

        let View: any = "viewListItems";
        let Edit: any = "editListItems";

        const canView = await hasFolderPermission(
            context,
            folderPath,
            View
        );

        const canEdit = await hasFolderPermission(
            context,
            folderPath,
            Edit
        );

        // Restricted View Logic
        if (selectedFolderRef.current?.path !== folderPath) return;

        if (canView === true && canEdit === false) {
            setIsRestrictedView(true);
        } else {
            setIsRestrictedView(false);
        }
        console.log("VIEW:", canView, "EDIT:", canEdit, "Restricted:", canView && !canEdit);
    };


    const getPreviewUrl = (filePath: string) => {
        const extension = filePath?.split('.').pop()?.toLowerCase();
        switch (extension) {
            case 'doc':
            case 'docx':
            case 'ppt':
            case 'pptx':
            case 'xls':
            case 'xlsx':
                return <iframe src={`${SiteURL}/_layouts/15/WopiFrame.aspx?sourcedoc=${filePath}&action=embedview`} style={{ width: "100%", height: "80vh" }}></iframe>;

            case 'txt':
                return <iframe src={`${filePath}`} style={{ width: "100%", height: "80vh" }}></iframe>;
            case 'jpg':
            case 'jpeg':
            case 'png':
            case 'gif':

            case 'bmp':
                return <img src={`${filePath}`} alt={DisplayLabel.Preview} />;
            case 'pdf':
                //  return <iframe src={`${filePath}`} style={{ width: "100%", height: "80vh" }}></iframe>;
                // return <iframe 
                //     src={`${filePath}#toolbar=0&navpanes=0&scrollbar=0&view=FitH`} 
                //     style={{ width: "100%", height: "80vh" }} 
                // ></iframe>;

                return (
                    <div
                        // style={{
                        //     position: "relative",
                        //     height: "80vh",
                        //     overflowY: "auto"
                        // }}
                        style={{
                            position: "relative",
                            height: "80vh",
                            userSelect: "none"
                        }}
                        onContextMenu={(e) => e.preventDefault()}

                    >
                        <iframe
                            src={`${filePath}#toolbar=0&navpanes=0&scrollbar=0&view=FitH`}
                            style={{
                                width: "100%",
                                height: "100%",
                                border: "none",
                                //pointerEvents: "none" // disables clicks
                            }}
                            onLoad={(e) => {
                                try {
                                    const iframeDoc =
                                        e.currentTarget.contentDocument ||
                                        e.currentTarget.contentWindow?.document;

                                    iframeDoc?.addEventListener("contextmenu", (event) => {
                                        event.preventDefault();
                                    });
                                } catch (err) {
                                    console.log("Cannot access iframe content");
                                }
                            }}
                        />

                    </div>
                );

        }
    };

    const hideClickableOptionsInVersionsFrame = (iframe: HTMLIFrameElement | null) => {
        if (!iframe) return;

        try {
            const frameDocument = iframe.contentDocument || iframe.contentWindow?.document;
            if (!frameDocument) {
                setIsVersionsOverlayVisible(true);
                return;
            }

            const existingStyle = frameDocument.getElementById("readonly-versions-style");
            if (existingStyle) {
                existingStyle.remove();
            }

            const style = frameDocument.createElement("style");
            style.id = "readonly-versions-style";
            style.textContent = `
                a,
                button,
                input,
                select,
                textarea,
                [role="button"],
                [onclick] {
                    pointer-events: none !important;
                    cursor: default !important;
                }

                a {
                    text-decoration: none !important;
                    color: inherit !important;
                }

                button,
                input[type="button"],
                input[type="submit"],
                input[type="reset"] {
                    opacity: 0.55 !important;
                }
                    
                [id*="Delete" i],
                [class*="delete" i],
                [title*="Delete" i],
                [aria-label*="Delete" i],
                input[value*="Delete" i],
                button[title*="Delete" i],
                a[title*="Delete" i] {
                    display: none !important;
                    visibility: hidden !important;
                }
            `;

            frameDocument.head?.appendChild(style);

            const blockInteraction = (event: Event) => {
                event.preventDefault();
                event.stopPropagation();
                if ("stopImmediatePropagation" in event) {
                    (event as Event & { stopImmediatePropagation: () => void; }).stopImmediatePropagation();
                }
            };

            frameDocument.addEventListener("click", blockInteraction, true);
            frameDocument.addEventListener("dblclick", blockInteraction, true);
            frameDocument.addEventListener("contextmenu", blockInteraction, true);
            frameDocument.addEventListener("submit", blockInteraction, true);
            frameDocument.addEventListener("keydown", (event: KeyboardEvent) => {
                if (event.key === "Enter" || event.key === " ") {
                    blockInteraction(event);
                }
            }, true);
            setIsVersionsOverlayVisible(false);
        } catch (error) {
            console.warn("Unable to update versions iframe actions.", error);
            setIsVersionsOverlayVisible(true);
        } finally {
            setIsVersionsLoading(false);
        }
    };

    const renderVersionsPanel = (url: string) => (
        <div
            // style={{
            //     position: "relative",
            //     minHeight: "80vh",
            //     borderRadius: "16px",
            //     overflow: "hidden",
            //     background: "linear-gradient(180deg, #f8fbff 0%, #eef4fb 100%)",
            //     border: "1px solid #dbe7f3",
            //     boxShadow: "0 10px 30px rgba(15, 108, 189, 0.08)"
            // }}
            style={{
                position: "relative",
                height: "70vh",
                overflowY: "auto",
                overflowX: "hidden",
                borderRadius: "16px",
                background: "#fff",
                border: "1px solid #dbe7f3"
            }}
        >
            {isVersionsLoading && (
                <div
                    style={{
                        position: "absolute",
                        inset: 0,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        background: "rgba(248, 251, 255, 0.96)",
                        zIndex: 2
                    }}
                    onContextMenu={(e) => e.preventDefault()}
                >
                    <div
                        style={{
                            display: "flex",
                            flexDirection: "column",
                            alignItems: "center",
                            gap: "12px",
                            padding: "28px 32px",
                            borderRadius: "18px",
                            background: "#ffffff",
                            boxShadow: "0 14px 40px rgba(15, 108, 189, 0.12)",
                            border: "1px solid #e3edf7"
                        }}
                    >
                        <Spinner size={SpinnerSize.large} label="Loading version history..." />
                        <span style={{ color: "#4b5563", fontSize: "13px" }}>
                            Preparing the SharePoint versions view
                        </span>
                    </div>
                </div>
            )}
            <iframe
                id="frame"
                src={url}
                // style={{
                //     width: "100%",
                //     height: "calc(80vh - 58px)",
                //     border: "none",
                //     backgroundColor: "#fff",
                //     pointerEvents: "none"
                // }}
                style={{
                    width: "100%",
                    height: "1200px", // adjust as needed
                    border: "none",
                    backgroundColor: "#fff"
                }}
                onLoad={(event) => hideClickableOptionsInVersionsFrame(event.currentTarget)}
            >

            </iframe>
            <div
                aria-hidden="true"
                style={{
                    position: "absolute",
                    inset: 0,
                    zIndex: isVersionsLoading ? 1 : 3,
                    background: "transparent",
                    cursor: "default",
                    pointerEvents: "auto"
                }}
                onClick={(event) => event.preventDefault()}
                onDoubleClick={(event) => event.preventDefault()}
                onMouseDown={(event) => event.preventDefault()}
                onMouseUp={(event) => event.preventDefault()}
                onContextMenu={(event) => event.preventDefault()}
            />
        </div>
    );

    const thStyle = {
        border: "1px solid #d1d1d1",
        padding: "10px",
        backgroundColor: "#f5f5f5",
        textAlign: "left" as const,
        fontWeight: 600,
    };

    const tdStyle = {
        border: "1px solid #d1d1d1",
        padding: "10px",
        // textAlign: "left" as const,
        textAlign: "center" as const,
    };

    const handleDocumentAction = async (action: string, item: any) => {
        switch (action) {
            case "OpenInApp":
                getOpenAppURL(item.ServerRelativeUrl, SiteURL);
                break;
            case "Delete":
                setMessage(DisplayLabel.DeleteConfirmMsg);
                setItemId(item.ListItemAllFields.Id);
                setHideDialog(true);
                break;
            case "Versions": {
                setActionButton(null);
                setIsVersionsLoading(true);
                setIsVersionsOverlayVisible(false);
                setPanelSize(PanelType.large);

                // SharePoint's version history page needs the GUID of the library the
                // file really lives in plus the id of the file's list item; both are
                // resolved from SharePoint here (see resolveVersionHistoryTarget) so
                // nested folders, spaces/special characters in the name and every file
                // type (PDF, Word, Excel, PowerPoint, …) open the correct history.
                const versionHistoryTarget = await resolveVersionHistoryTarget(context, item, tileData?.LibraryName);

                if (!versionHistoryTarget.url) {
                    // Invalid document reference: report it instead of showing a page
                    // that would only display "Item does not exist" from SharePoint.
                    setIsVersionsLoading(false);
                    setVersionsPanelUrl("");
                    setAlertMsg(versionHistoryTarget.error || "Version history is not available for this document.");
                    setPopupType("warning");
                    setIsPopupBoxVisible(true);
                    break;
                }

                setPanelTitle(DisplayLabel.Versions);
                setVersionsPanelUrl(versionHistoryTarget.url);
                setPanelForm(null);
                setIsOpenCommonPanel(true);
                break;
            }
            case "Rename":
                setFileNameErr("");
                setItemId(item.ListItemAllFields.Id);
                setPanelTitle(DisplayLabel.Rename);
                const fileDetails = item.ListItemAllFields.ActualName.split(".");
                setExtension(fileDetails[1]);
                setFileName(fileDetails[0]);
                setIsOpenCommonPanel(true);
                break;
            case "Download":
                location.href = `${SiteURL}/_layouts/15/download.aspx?SourceUrl=${item.ServerRelativeUrl}`;
                break;
            case "Preview":
                setActionButton(null);
                setPanelSize(PanelType.smallFluid);
                setPanelTitle(DisplayLabel.Preview);
                const previewData = getPreviewUrl(item.ServerRelativeUrl);
                setPanelForm(previewData);
                setIsOpenCommonPanel(true);
                break;
            case "CheckOut":
                await commonPostMethod(`${SiteURL}/_api/web/GetFileByServerRelativeUrl('${item.ServerRelativeUrl}')/checkout`, context);
                setAlertMsg(DisplayLabel.CheckoutSuccess);
                setPopupType("checkout");
                setIsPopupBoxVisible(true);
                await getDocument(selectedFolderRef.current);
                break;
            case "CheckIn":
                setPanelTitle(DisplayLabel.CheckIn);
                setActionButton(<PrimaryButton text={DisplayLabel.CheckIn} style={{ marginRight: "10px" }} onClick={async () => {
                    await commonPostMethod(`${SiteURL}/_api/web/GetFileByServerRelativeUrl('${item.ServerRelativeUrl}')/checkin(comment='${comment}',checkintype=0)`, context);
                    setAlertMsg(DisplayLabel.CheckInSuccess);
                    setPopupType("checkin");
                    setIsPopupBoxVisible(true);
                    await getDocument(selectedFolderRef.current);
                }} />);
                setIsOpenCommonPanel(true);
                break;
            case "DiscardCheckOut":
                setMessage(DisplayLabel.CheckoutConfirm);
                setServerRelativeUrl(item.ServerRelativeUrl);
                setHideDialogCheckOut(true);
                break;
            case "History":
                setActionButton(null);
                const HistoryData = await getHistoryByID(SiteURL, context.spHttpClient, item.ListItemAllFields.Id, tileData?.LibraryName);
                // const bindData =
                //     HistoryData?.value.length > 0 ? (
                //         HistoryData.value
                //             .sort((a: any, b: any) => {
                //                 return new Date(a.ActionDate).getTime() - new Date(b.ActionDate).getTime();
                //             })
                //             .map((el: any, index: number) => (
                //                 <tr key={index}>
                //                     <td>{index + 1}</td>
                //                     <td>{el.Action}</td>
                //                     <td>{el.Author.Title}</td>
                //                     {/* <td>{el.ActionDate ? format(el.ActionDate, "DD-MM-YYYY hh:mm:ss A") : ""}</td> */}
                //                      <td>
                //                         {el.ActionDate
                //                             ? format(new Date(el.ActionDate), "dd-MM-yyyy hh:mm:ss a")
                //                             : ""}
                //                     </td>
                //                     <td>{el.InternalComment}</td>
                //                 </tr>
                //             ))
                //     ) : (
                //         <tr>
                //             <td colSpan={5}>No Data</td>
                //         </tr>
                //     );
                // setPanelForm(<table className="addoption" style={{ width: '100%', marginTop: '20px', borderCollapse: 'collapse' }}>
                //     <thead>
                //         <tr>
                //             <th>{DisplayLabel?.SrNo}</th>
                //             <th>{DisplayLabel?.Action}</th>
                //             <th>{DisplayLabel?.ActionBy}</th>
                //             <th>{DisplayLabel?.ActionDate}</th>
                //             <th>{DisplayLabel?.Comments}</th>
                //         </tr>
                //     </thead>
                //     <tbody>{bindData}</tbody>
                // </table>);

                const bindData =
                    HistoryData?.value?.length > 0 ? (
                        HistoryData.value
                            .sort((a: any, b: any) => {
                                return (
                                    new Date(a.ActionDate).getTime() -
                                    new Date(b.ActionDate).getTime()
                                );
                            })
                            .map((el: any, index: number) => (
                                <tr key={index}>
                                    <td style={tdStyle}>{index + 1}</td>
                                    <td style={tdStyle}>{el.Action}</td>
                                    <td style={tdStyle}>{el.Author?.Title}</td>
                                    <td style={tdStyle}>
                                        {el.ActionDate
                                            ? format(
                                                new Date(el.ActionDate),
                                                "dd-MM-yyyy hh:mm:ss a"
                                            )
                                            : ""}
                                    </td>
                                    <td style={tdStyle}>{el.InternalComment}</td>
                                </tr>
                            ))
                    ) : (
                        <tr>
                            <td style={tdStyle} colSpan={5}>
                                No Data
                            </td>
                        </tr>
                    );

                setPanelForm(
                    <table
                        style={{
                            width: "100%",
                            marginTop: "20px",
                            borderCollapse: "collapse",
                            border: "1px solid #d1d1d1",
                        }}
                    >
                        <thead>
                            <tr>
                                <th style={thStyle}>{DisplayLabel?.SrNo}</th>
                                <th style={thStyle}>{DisplayLabel?.Action}</th>
                                <th style={thStyle}>{DisplayLabel?.ActionBy}</th>
                                <th style={thStyle}>{DisplayLabel?.ActionDate}</th>
                                <th style={thStyle}>{DisplayLabel?.Comments}</th>
                            </tr>
                        </thead>
                        <tbody>{bindData}</tbody>
                    </table>
                );

                setPanelTitle(DisplayLabel.History);
                setIsOpenCommonPanel(true);


                break;
            case "View":
                setActionButton(null);
                const dataConfig = await getConfigActive(context.pageContext.web.absoluteUrl, context.spHttpClient);
                const libraryData = await getDataByLibraryName(context.pageContext.web.absoluteUrl, context.spHttpClient, tileData.LibraryName);
                const currentSelectedFolder = selectedFolderRef.current || selectedFolder;
                const selectedFolderName =
                    currentSelectedFolder?.name ||
                    currentSelectedFolder?.path?.split("/").filter(Boolean).pop() ||
                    "";
                let jsonData = JSON.parse(libraryData.value[0].DynamicControl);
                jsonData = jsonData.filter((ele: any) => ele.IsActiveControl);

                // Resolve site users once so "Person or Group" fields can bind the
                // PeoplePicker by email (listItemAllFields only returns Id/Title).
                let usersById = new Map<number, any>();
                try {
                    const siteUsersRes = await getListData(`${SiteURL}/_api/web/siteusers?$filter=PrincipalType eq 1`, context);
                    usersById = new Map((siteUsersRes?.value || []).map((u: any) => [Number(u.Id), u]));
                } catch (error) {
                    console.warn("Unable to load site users for PeoplePicker:", error);
                }

                // Pre-fetch options for list-backed choice columns (same logic as the Create/Edit form).
                const choiceOptionsMap: { [key: string]: IDropdownOption[] } = {};
                const listBasedChoiceFields = jsonData.filter(
                    (el: any) =>
                        (el.ColumnType === "Dropdown" || el.ColumnType === "Multiple Select" || el.ColumnType === "Radio") &&
                        !el.IsStaticValue &&
                        el.InternalListName
                );
                await Promise.all(
                    listBasedChoiceFields.map(async (el: any) => {
                        try {
                            const res = await getListData(
                                `${SiteURL}/_api/web/lists/getbytitle('${el.InternalListName}')/items?$top=5000&$filter=Active eq 1&$orderby=${el.DisplayValue} asc`,
                                context
                            );
                            choiceOptionsMap[el.InternalTitleName] = (res?.value || []).map((ele: any) => ({
                                key: String(ele[el.DisplayValue]),
                                text: ele[el.DisplayValue],
                            }));
                        } catch (error) {
                            console.warn(`Unable to load options for "${el.Title}":`, error);
                        }
                    })
                );

                // const htm = (
                //     <>
                //         <div className="row">
                //             <div className="col-md-12">
                //                 <label>{DisplayLabel.Path}: <b>{currentSelectedFolder?.path || ""}</b></label>
                //             </div>
                //         </div>
                //         <div className="grid-2">
                //             <div className="col-md-6" data-testid="text-meta-tile">
                //                 <TextField
                //                     label={DisplayLabel.TileName}
                //                     readOnly
                //                     value={tileData.TileName}
                //                     styles={viewOnlyTextFieldStyles}
                //                 />
                //             </div>
                //             <div className="col-md-6" data-testid="text-meta-name">
                //                 <TextField
                //                     label={DisplayLabel.FolderName}
                //                     readOnly
                //                     value={selectedFolderName}
                //                     styles={viewOnlyTextFieldStyles}
                //                 />
                //             </div>

                //             {item.ListItemAllFields.IsSuffixRequired ? (
                //                 <>
                //                     <div className="col-md-6" data-testid="text-meta-tile">
                //                         <TextField
                //                             label={DisplayLabel.DocumentSuffix}
                //                             readOnly
                //                             value={item.ListItemAllFields.DocumentSuffix || ""}
                //                             styles={viewOnlyTextFieldStyles}
                //                         />
                //                     </div>
                //                     {item.ListItemAllFields.DocumentSuffix === "Other" && (
                //                         <div className="col-md-6" data-testid="text-meta-name">
                //                             <TextField
                //                                 label={DisplayLabel.OtherSuffixName}
                //                                 readOnly
                //                                 value={item.ListItemAllFields.OtherSuffix || ""}
                //                                 styles={viewOnlyTextFieldStyles}
                //                             />
                //                         </div>
                //                     )}
                //                 </>
                //             ) : null}

                //             {jsonData.map((el: any) => {
                //                 const filterObj = dataConfig?.value.find((ele: any) => ele.Id === el.Id);
                //                 return renderViewOnlyMetaField(el, filterObj, item, usersById, choiceOptionsMap, peoplePickerContext);
                //             })}
                //         </div>
                //     </>
                // );

                    const htm = (
                        <>
                            <div className="row">
                                <div className="col-md-12">
                                    <label>{DisplayLabel.Path}: <b>{currentSelectedFolder?.path || ""}</b></label>
                                </div>
                            </div>
                            <div className="grid-2">
                                <div className="col-md-6" data-testid="text-meta-tile">
                                    {/* <label className="view-only-label">{DisplayLabel.TileName}</label> */}
                                    <label className="view-only-label" >{DisplayLabel.TileName}</label>
                                    <div className="view-only-value">{tileData.TileName}</div>
                                </div>
                                <div className="col-md-6" data-testid="text-meta-name">
                                   
                                     <label className="view-only-label" >{DisplayLabel.FolderName}</label>
                                    <div className="view-only-value">{selectedFolderName}</div>
                                </div>

                                {item.ListItemAllFields.IsSuffixRequired ? (
                                    <>
                                        <div className="col-md-6" data-testid="text-meta-tile">
                                            {/* <label className="view-only-label">{DisplayLabel.DocumentSuffix}</label> */}
                                            <label className="view-only-label" >{DisplayLabel.DocumentSuffix}</label>
                                            <div className="view-only-value">{item.ListItemAllFields.DocumentSuffix || ""}</div>
                                        </div>
                                        {item.ListItemAllFields.DocumentSuffix === "Other" && (
                                            <div className="col-md-6" data-testid="text-meta-name">
                                                {/* <label className="view-only-label">{DisplayLabel.OtherSuffixName}</label> */}
                                                      <label className="view-only-label" >{DisplayLabel.OtherSuffixName}</label>
                                                <div className="view-only-value">{item.ListItemAllFields.OtherSuffix || ""}</div>
                                            </div>
                                        )}
                                    </>
                                ) : null}

                                {jsonData.map((el: any) => {
                                    const filterObj = dataConfig?.value.find((ele: any) => ele.Id === el.Id);
                                    return renderViewOnlyMetaField(el, filterObj, item, usersById, choiceOptionsMap, peoplePickerContext);
                                })}
                            </div>
                        </>
                    );

                setPanelForm(htm);
                setPanelTitle(DisplayLabel.View);
                setIsOpenCommonPanel(true);
                break;
            case "AdvancePermission":
                setItemId(item.ListItemAllFields.Id);
                setIsPanelOpen(true);
                break;
            case "Share":
                const URL = `${SiteURL}/_layouts/15/sharedialog.aspx?listId=${tileData.LibGuidName}&listItemId=${item.ListItemAllFields.Id}&clientId=sharePoint&policyTip=0&folderColor=undefined&ma=0&fullScreenMode=true&itemName=${item.ListItemAllFields.ActualName}&origin=${portalUrl}&clientId=sharePoint&ma=1`;
                setShareURL(URL);
                setIFrameDialogOpened(true);
                break;
            case "OpenInBrowser":
                const urls = item.LinkingUri === null ? item.ServerRelativeUrl : item.LinkingUri;
                window.open(urls, '_blank');
                break;
        }
    };

    const folderPathBread = useMemo<FolderNode[]>(() => {
        if (!selectedFolder) return [];
        return buildBreadcrumbPath(selectedFolder, folders);
    }, [selectedFolder]);

    useEffect(() => {
        if (selectedFolder) {
            hasRequiredPermissions(selectedFolder.path);
            void getRestrictedUserData(selectedFolder.path);
        }
    }, [selectedFolder]);

    useEffect(() => {
        getDocument();
    }, [isOpenUploadPanel]);

    //Original Code
    // const getDocument = async (folderNode?: FolderNode | null) => {
    //     const folderToLoad = folderNode || selectedFolderRef.current || selectedFolder;
    //     if (!folderToLoad) return [];
    //     if (folderToLoad.isLastLevel) {
    //         const files = await getAllDocuments(context, folderToLoad.path);
    //         setFiles(files.filter((el: any) => (el.ListItemAllFields.Active && (el.ListItemAllFields.InternalStatus === "Published" || el.ListItemAllFields.AuthorId === UserID))) || []);
    //     } else {
    //         setFiles([]);
    //     }
    // };

    //comment by rupali

    //  const getDocument = async (folderNode?: FolderNode | null) => {
    //     const folderToLoad = folderNode || selectedFolderRef.current || selectedFolder;
    //     if (!folderToLoad) return [];

    //     //const files = await getAllDocuments(context, folderToLoad.path);
    //     const allFiles = await getAllDocuments(context, folderToLoad.path);

    //     const files = allFiles.filter(
    //         (file: any) =>
    //             file.ListItemAllFields?.Active === true &&
    //             file.ListItemAllFields?.DeleteFlag !== "Deleted"
    //     );


    //       console.log("Files:", files);
    //     console.log("Files:", files);

    //     setFiles(files);
    // };

    const getDocument = async (folderNode?: FolderNode | null, selectionRequestId?: number) => {
        const folderToLoad = folderNode || selectedFolderRef.current || selectedFolder;
        if (!folderToLoad) return [];

        const allFiles = await getAllDocuments(context, folderToLoad.path);

        const files = allFiles.filter((file: any) => {
            const item = file.ListItemAllFields;

            // Existing filters
            if (
                item?.Active !== true ||
                item?.DeleteFlag === "Deleted"
            ) {
                return false;
            }

            const status = item?.InternalStatus;
            const authorId = item?.AuthorId;

            // Pending documents visible only to author
            if (
                ["PendingWithPublisher", "PendingWithPM", "Rejected"].includes(status)
            ) {
                return authorId === UserID;
            }

            // Published and other statuses visible to everyone
            return true;
        });

        if (selectionRequestId !== undefined && selectionRequestId !== folderSelectionRequestRef.current) {
            return [];
        }

        setFiles(files);
    };

    const fetchButtonsAndPermissions = useCallback(async (targetPath: string) => {
        try {
            //added new
            // setButtons([]);
            //   setUserPerms({
            //         FullControl: false,
            //         Edit: false,
            //         Contribute: false,
            //         Read: false
            //     });

            const sp = spfi().using(SPFx(context));

            // Determine if the target path is the root of the library
            const rootPath = buildLibraryRootPath(context, tileData?.LibraryName);
            const isRoot = targetPath.toLowerCase() === rootPath.toLowerCase();

            // Parallel fetch of buttons and current effective permissions
            const [btnsRes, perms] = await Promise.all([
                buttonsCache.current
                    ? Promise.resolve(buttonsCache.current)
                    : sp.web.lists.getByTitle("DMS_Buttons").items
                        .filter("Active eq 1")
                        .select("Title", "InternalName", "ButtonType", "ButtonDisplayName", "Icons", "Sequence", "FullControl", "Contribute", "EditPermission", "ReadPermission")
                        .orderBy("Sequence", true)(),
                isRoot
                    ? sp.web.lists.getByTitle(tileData?.LibraryName).getCurrentUserEffectivePermissions()
                    : sp.web.getFolderByServerRelativePath(targetPath).getItem().then(item => item.getCurrentUserEffectivePermissions())
            ]);

            if (!buttonsCache.current) buttonsCache.current = btnsRes;

            const isFullControl = sp.web.hasPermissions(perms, PermissionKind.ManagePermissions);
            const isEdit = sp.web.hasPermissions(perms, PermissionKind.EditListItems);
            const isContribute = sp.web.hasPermissions(perms, PermissionKind.AddListItems);
            const isRead = sp.web.hasPermissions(perms, PermissionKind.ViewListItems);

            setCanShowButtons(
                isFullControl || isEdit || isContribute
            );

            // Strictly Hierarchical Resolution: Set only the highest permission tier to true
            setUserPerms({
                FullControl: isFullControl,
                Edit: isEdit && !isFullControl,
                Contribute: isContribute && !isEdit && !isFullControl,
                Read: isRead && !isContribute && !isEdit && !isFullControl
            });

            setButtons(btnsRes.map((btn: any) => ({ ...btn, key: btn.InternalName })));
        } catch (err) {
            console.error('Error in fetchButtonsAndPermissions:', err);
        }
    }, [context, tileData]);

    const visibleButtons = useMemo(() => {
        return buttons.filter(btn => {
            if (userPerms.FullControl) return true; // Full Control sees all
            if (userPerms.Edit && (btn.EditPermission || btn.Contribute || btn.ReadPermission)) return true;
            if (userPerms.Contribute && (btn.Contribute || btn.ReadPermission)) return true;
            if (userPerms.Read && btn.ReadPermission) return true;
            return false;
        }).sort((a, b) => (a.Sequence || 0) - (b.Sequence || 0));
    }, [buttons, userPerms]);

    const createMenuProps = (item: any) => {
        return visibleButtons.filter((btn) => btn.ButtonType === "Document")
            .filter((btn) => {
                switch (btn.key) {
                    // case "Delete":
                    //    return !tileData?.IsArchiveRequired;

                    case "Delete":

                        //multiple Tile Admin
                        //   return (
                        //     !tileData?.IsArchiveRequired &&
                        //     (
                        //         tileData?.AuthorId === UserID ||
                        //         tileData?.TileAdminId?.some(
                        //             (admin: any) => admin.Id === UserID
                        //         ) ||
                        //         isValidUser
                        //     )
                        // );

                        return (
                            !tileData?.IsArchiveRequired &&
                            (
                                item?.data?.ListItemAllFields?.AuthorId === UserID ||
                                // tileData?.TileAdminId === UserID ||
                                tileData?.TileAdminId?.includes?.(Number(UserID)) ||
                                isValidUser
                            )
                        );

                    case "OpenInApp":
                        const isCheck = checkExtension(item.data.Name);
                        return isCheck;
                    case "CheckIn":
                        return item.data.CheckOutType === 0 && item.data.CheckedOutByUser?.Id === UserID;
                    case "DiscardCheckOut":
                        return item.data.CheckOutType === 0 && item.data.CheckedOutByUser?.Id === UserID;
                    case "CheckOut":
                        return item.data.CheckOutType === 2;
                    case "Preview":
                        return !checkExtension(item.data.Name);
                    default:
                        return checkButtons(btn.key);
                }
            }).map((btn: any) => ({
                key: btn.key,
                text: btn.ButtonDisplayName,
                Icons: btn?.Icons
            }));
    };

    const handleButtonClick = (internalName: string) => {
        switch (internalName) {
            case "NewRequest":
                projectCreation();
                break;
            case "NewFolder":
                setIsOpenFolderPanel(true);
                setFolderName("");
                setFolderNameErr("");
                break;
            case "Upload":
                setFileType("upload");
                setIsOpenUploadPanel(true);
                break;
            case "AdvancedSearch":
                navigate('/Search', { state: { from: `/workspace/${workspaceId}`, libName: tileData?.LibraryName } });
                break;
            default:
                console.log("Action triggered:", internalName);
        }
    };

    const getStatusStyles = (status: any) => {
        switch (status) {
            case "Pending With Approver":
                return { backgroundColor: "#f1faff", color: "#009ef7" };
            case "Published":
                return { backgroundColor: "#e8fff3", color: "#50cd89" };
            case "Pending With Publisher":
                return { backgroundColor: "#fff8dd", color: "#ffc700" };
            case "Rejected":
                return { backgroundColor: "#fff5f8", color: "#ed1c24" };
        }
    };

    const renderDocName = (item: any) => {
        const ext = item.Name.split(".").pop();
        const config = fileTypeConfig[ext] || fileTypeConfig.other;
        const { IconName, className } = config;
        const checkedOutUser = item?.CheckedOutByUser;
        const isCheckedOut = item?.CheckOutType === 0;
        const isCheckedOutByCurrentUser = checkedOutUser?.Id === UserID;
        return (
            <div className="doc-name-cell" data-testid={`link-document-${item.id}`}>
                <div className={`doc-icon-wrap ${className}`}>
                    <IconName className="doc-icon-svg" />
                </div>
                <span
                    className="table-cell-link"
                    onClick={() => {

                        if (item.LinkingUrl === "") {
                            if (isRestrictedView === true) {
                                const filePath = item.ServerRelativeUrl;
                                const folderPath = filePath.substring(0, filePath.lastIndexOf("/"));

                                const previewUrl = `${SiteURL}/${tileData?.LibraryName}/Forms/AllItems.aspx?id=${encodeURIComponent(filePath)
                                    }&parent=${encodeURIComponent(folderPath)}`;

                                window.open(previewUrl, "_blank");
                                return;
                            } else {
                                window.open(item.ServerRelativeUrl, "_blank");
                            }
                        }
                        else
                            window.open(item.LinkingUrl, "_blank");
                    }}
                >
                    {item?.ListItemAllFields?.ActualName}
                </span>
                {isCheckedOut && (
                    <TooltipHost
                        content={`${checkedOutUser?.Title} ${DisplayLabel.CheckedOutThisItem}`}
                        directionalHint={DirectionalHint.rightCenter} // Positioning
                        styles={{
                            root: { display: 'inline-block', maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }
                        }}
                    >
                        <Icon
                            iconName={isCheckedOutByCurrentUser ? "CheckedOutByYou12" : "CheckedOutByOther12"}
                            style={{ marginLeft: "5px", marginTop: '5px', color: isCheckedOutByCurrentUser ? "#a4262c" : "#605e5c", cursor: "pointer" }}
                        />

                    </TooltipHost>
                )}
            </div>
        );
    };

    const dismissCommanPanel = () => {
        setIsOpenCommonPanel(false);
        setActionButton(null);
        setPanelForm(null);
        setPanelSize(PanelType.medium);
        setVersionsPanelUrl("");
        setIsVersionsLoading(false);
    };
    const onDismiss: any = useCallback(() => { setIsPanelOpen(false); }, []);
    const closeDialog = useCallback(() => setHideDialog(false), []);
    const closeDialogCheckOut = useCallback(() => setHideDialogCheckOut(false), []);
    const hidePopup = useCallback(() => { setIsPopupBoxVisible(false); }, [isPopupBoxVisible]);
    const hideCommonPopup = useCallback(() => { setIsShowCommnPopupBoxVisible(false); }, []);
    const dismissFolderPanel = () => { setIsOpenFolderPanel(false); };
    const dissmissProjectCreationPanel = useCallback((value: boolean) => { setIsCreateProjectPopupOpen(value); refreshCurrentFolder(); }, [isCreateProjectPopupOpen]);

    const dissmissSharePopup = useCallback((value: boolean) => { setIFrameDialogOpened(value); }, []);
    const dismissUploadPanel = useCallback(() => { setIsOpenUploadPanel(false); }, []);

    const handleConfirm = useCallback(
        async (value: boolean) => {
            if (value) {
                setHideDialog(false);
                setIsPanelOpen(true);
                deleteDoc();
            }
        },
        [itemId]
    );
    const handleConfirmCheckOut = useCallback(async (value: boolean) => {
        if (value) {
            await commonPostMethod(`${SiteURL}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')/undocheckout()`, context);
            setAlertMsg(DisplayLabel.DiscardedCheckOut);
            setIsPopupBoxVisible(true);
            await getDocument(selectedFolderRef.current);
        }
    }, [serverRelativeUrl]);

    const deleteDoc = async () => {
        const obj = {
            Active: false,
            DeleteFlag: "Deleted",
        };
        await updateLibrary(SiteURL, context.spHttpClient, obj, itemId, tileData.LibraryName);
        setAlertMsg(DisplayLabel.DeletedMsg);
        setIsPopupBoxVisible(true);
        getDocument();
    };

    useEffect(() => {
        setPanelForm(<div className="grid">
            <div className="grid-item-large">
                <TextField value={fileName} required onChange={(_, val) => {
                    setFileName(val || "");
                }} />

                <Label style={{ color: "red" }}>{fileNameErr}</Label>
            </div>
            <div className="grid-item-small"><TextField readOnly value={extension} /></div>
        </div>);
        setActionButton(<PrimaryButton text={DisplayLabel.Rename} style={{ marginRight: "10px" }} onClick={() => renameTheFile(itemId)} />);
    }, [fileName, extension, fileNameErr]);


    useEffect(() => {
        setPanelForm(<>
            <div className="col-md-10">
                <TextField value={comment} label={DisplayLabel?.Comments} onChange={(_, val) => setComment(val || "")} />
            </div>
        </>);

    }, [comment]);

    const renameTheFile = (id: number) => {
        if (fileName === "") {
            setFileNameErr(DisplayLabel.ThisFieldisRequired);
        }
        else {
            const obj = {
                ActualName: `${fileName}.${extension}`
            };
            updateLibrary(SiteURL, context.spHttpClient, obj, id, tileData.LibraryName).then((response) => {

                if (response) {
                    dismissFolderPanel();
                    setAlertMsg(DisplayLabel.SubmitMsg);
                    setIsPopupBoxVisible(true);
                    getDocument();
                }
                else {
                    dismissFolderPanel();
                    setAlertMsg(DisplayLabel.RenameAlertMsg);
                    setIsShowCommnPopupBoxVisible(true);
                }

            });
        }
    };

    const createFolder = (): void => {
        setFolderNameErr("");

        if (folderName === "") {
            setFolderNameErr(DisplayLabel.ThisFieldisRequired);
            return;
        }
        if (invalidCharsRegex.test(folderName)) {
            setFolderNameErr(DisplayLabel.FolderSpecialCharacterValidation);
            return;
        }
        if (!selectedFolder) return;

        // const isDuplicate = selectedFolder.children.filter((el: any) => el.Name === folderName);

        // if (isDuplicate.length > 0) {
        //     setFolderNameErr(DisplayLabel.FolderAlreadyExist);
        //     return;
        // }

        const isDuplicate = selectedFolder.children
            .flat(Infinity)
            .some(
                (el: any) =>
                    el?.name?.trim().toLowerCase() === folderName.trim().toLowerCase()
            );

        if (isDuplicate) {
            setFolderNameErr(DisplayLabel.FolderAlreadyExist);
            return;
        }

        // const users = [selectedFolder?.ProjectmanagerId, selectedFolder?.PublisherId, ...admin, tileData?.TileAdminId];

        const users = [
            { id: selectedFolder?.ProjectmanagerId, type: 'FolderAccess' },
            { id: selectedFolder?.PublisherId, type: 'FolderAccess' },
            ...admin.map((id: any) => ({ id, type: 'Admin' })),
            // ...(tileData?.TileAdminId
            //     ? [{ id: tileData.TileAdminId, type: 'TileAdmin' }]
            //     : []),
            ...(tileData?.TileAdminId?.length
                ? tileData.TileAdminId.map((id: number) => ({
                    id,
                    type: 'TileAdmin',
                }))
             : []),
        ];
        const siteRelative = context.pageContext.web.serverRelativeUrl;


        const urlAfterSite = selectedFolder.path.replace(siteRelative, "").replace(/^\/+/, "");

        const IsRootFolder = false;
        //folderPathBread        
        FolderStructure(context, `${urlAfterSite}/${folderName}`, users, tileData.LibraryName, tileData.AllowChildInheritance, IsRootFolder).then(async (response) => {
            const sp = spfi().using(SPFx(context));
            const folderMetadata = await sp.web.getFolderByServerRelativePath(selectedFolder?.path).listItemAllFields();
            const folderData = JSON.parse(JSON.stringify(folderMetadata, (key, value) => (value === null || (Array.isArray(value) && value.length === 0)) ? undefined : value));
            let obj: any = {
                ...folderData
            };

            updateLibrary(SiteURL, context.spHttpClient, obj, response, tileData.LibraryName).then((response) => {
                dismissFolderPanel();
                setAlertMsg(DisplayLabel.SubmitMsg);
                setIsPopupBoxVisible(true);
                // Refresh only the currently selected folder's children
                // instead of rebuilding the whole tree from root.
                // This preserves the nested folder structure and keeps
                // the user in the same location after creating a folder.
                refreshCurrentFolder();
            });
        });
    };



    const columns = React.useMemo(() => {
        return [
            {
                headerName: DisplayLabel.SrNo || "Sr.No",
                filter: false,
                resizable: false,
                maxWidth: 80,
                valueGetter: (params: any) => params.node.rowIndex + 1
            },
            {
                headerName: DisplayLabel.FileName || "File Name",
                filter: true,
                sortable: true,
                field: "Name",
                maxWidth: 400,
                minWidth: 400,
                cellRenderer: (item: any) => renderDocName(item.data)
            },
            {
                headerName: DisplayLabel.ReferenceNo || "Reference No",
                filter: true,
                sortable: true,
                field: "ListItemAllFields.ReferenceNo",
                maxWidth: 160,
                minWidth: 120,
            },
            {
                headerName: DisplayLabel.Versions || "Versions",
                filter: true,
                sortable: true,
                field: "ListItemAllFields.Level",
                maxWidth: 80,
                cellRenderer: (item: any) =>
                    <span className="table-cell-text table-cell-version" data-testid={`text-version-${item.id}`}>
                        v{getDocumentVersion(
                            item?.data?.ListItemAllFields?.Level,
                            item?.data?.ListItemAllFields?.OData__UIVersionString
                        ) || "-"}
                    </span>
            },
            {
                headerName: DisplayLabel.Status || "Status",
                filter: true,
                sortable: true,
                field: "ListItemAllFields.DisplayStatus",
                cellRenderer: (item: any) => {
                    const style = getStatusStyles(item.data.ListItemAllFields.DisplayStatus);
                    return <div>
                        <Badge
                            style={{ ...style }}
                        >
                            {item.data.ListItemAllFields.DisplayStatus}
                        </Badge>
                    </div>;
                }
            },
            {
                headerName: DisplayLabel.Action || "Action",
                filter: true,
                sortable: true,
                minWidth: 100,
                maxWidth: 120,
                cellRenderer: (item: any) => {
                    const menuProps = createMenuProps(item);
                    return <Menu>
                        <MenuTrigger disableButtonEnhancement>
                            <Button
                                appearance="subtle"
                                icon={<MoreHorizontalRegular className="table-action-btn" />}
                            />
                        </MenuTrigger>

                        <MenuPopover
                            style={{
                                boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
                                padding: "15px"
                            }}
                        >
                            <MenuList>
                                {menuProps.map((e) => {
                                    // const IconComponent = FluentIcons[e.Icons as keyof typeof FluentIcons] as React.FC ?? <ChevronRight24Regular />;
                                    const IconComponent = (
                                        FluentIcons[e.Icons as keyof typeof FluentIcons] ??
                                        ChevronRight24Regular
                                    ) as React.ComponentType<React.SVGProps<SVGSVGElement>>;
                                    return <MenuItem
                                        key={e.key}
                                        icon={<IconComponent className="table-action-btn" />}
                                        onClick={() => handleDocumentAction(e.key, item?.data)}
                                    >
                                        {e.text}
                                    </MenuItem>;
                                })}
                            </MenuList>
                        </MenuPopover>
                    </Menu>;
                }
            }
        ];
    }, [buttons]);


    const renderRightFolder = (nodes: Folder[]) => {
        return (
            <div className="folder-grid">
                {nodes.map((node: any) => (
                    <div
                        key={node?.id}
                        className="folder-card"
                        onClick={() => handleFolderSelect(node)}
                    >
                        <FluentIcons.Folder20Filled className="folder-icon" />
                        <span className="folder-name">{node?.name}</span>
                    </div>
                ))}
            </div>
        );
    };

    const expandParentFolders = (folder: any) => {
        setExpandedFolders(prev => {
            if (prev.includes(folder?.id)) {
                return prev.filter(id => id !== folder?.id);
            } else {
                return [...prev, folder?.id];
            }
        });
    };


    // const expandParentFolders = (folder: any) => {
    //     setExpandedFolders(prev => {
    //         if (prev.includes(folder.id)) {
    //             return prev;
    //         }
    //         return [...prev, folder.id];
    //     });
    // };

    const foldersColumn = React.useMemo(() => {
        return [
            {
                headerName: DisplayLabel.SrNo || "Sr.No",
                filter: false,
                resizable: false,
                maxWidth: 400,
                valueGetter: (params: any) => params.node.rowIndex + 1
            },
            {
                headerName: DisplayLabel.FileName || "Folder Name",
                filter: true,
                sortable: true,
                // field: "Name",
                field: "name", // FIXED
                maxWidth: 300,
                minWidth: 400,
                cellRenderer: (item: any) => <a href="javascript:void()" onClick={() => handleFolderSelect(item?.data)} style={{ color: "rgb(0, 158, 247)" }}>{item?.data?.name}</a>
            },
            {
                headerName: DisplayLabel.LastModified,
                //filter: false,
                resizable: false,
                filter: "agDateColumnFilter",
                maxWidth: 80,
                valueGetter: (params: any) => params?.data?.Modified ? format(params.data.Modified, "dd-MM-yyyy hh:mm a") : ""
            },
            // {
            //     headerName: DisplayLabel.LastModifiedBy || "Modified By",
            //     filter: true,
            //     sortable: true,
            //     field: "Editor.Title",
            //     maxWidth: 180
            // },
            {
                headerName: DisplayLabel.CreatedDate || "Created Date",
                resizable: false,
                filter: "agDateColumnFilter",
                maxWidth: 180,
                valueGetter: (params: any) => params?.data?.Created ? format(params.data.Created, "dd-MM-yyyy hh:mm a") : ""
            },
            {
                headerName: DisplayLabel.CreatedBy || "Created By",
                filter: true,
                sortable: true,
                maxWidth: 180,
                valueGetter: (params: any) => params?.data?.Author?.Title || (typeof params?.data?.Author === "string" ? params.data.Author : "") || ""
            },
        ];
    }, []);

    const getItemStyle = (type: string) => ({
        display: "flex",
        alignItems: "center",
        gap: "10px",
        padding: "8px 12px",
        borderRadius: "6px",
        cursor: "pointer",
        backgroundColor: viewListSetting === type ? "#EAF3FC" : "transparent",
        color: viewListSetting === type ? "#0F6CBD" : "#323130",
        fontWeight: viewListSetting === type ? 600 : 400,
        transition: "all 0.2s ease"
    });

    const projectCreation = useCallback(() => { setIsCreateProjectPopupOpen(true); setFormType("EntryForm"); setProjectUpdateData({}); }, []);
    const hasRequiredPermissions = (folderPath: string) => {
        checkPermissions(context, folderPath).then((permission: boolean) => {
            if (selectedFolderRef.current?.path === folderPath) {
                setHasPermission(permission);
            }
        });
    };
    const bindTable = () => {

        if (tables === "Approver") {
            return <ApprovalFlow context={context} libraryName={tileData?.LibraryName} userEmail={UserEmailID} action="Approver" />;
        }
        else if (tables === "Recycle") {
            return <ApprovalFlow context={context} libraryName={tileData?.LibraryName} userEmail={UserEmailID} action="Recycle" />;
        }
        else if (tables === "Archive") {
            return <ApprovalFlow context={context} libraryName={tileData?.LibraryName} userEmail={UserEmailID} action="Archive" />;
        }
        else {
            return (!selectedFolder?.children || selectedFolder?.children.length === 0) ?
                <ReusableDataTable rowData={files} columnDefs={columns} />
                :
                <div>
                    {viewListSetting === "List View" ? (
                        <ReusableDataTable rowData={selectedFolder?.children} columnDefs={foldersColumn} />
                    ) : (
                        <div >
                            {renderRightFolder(selectedFolder?.children)}
                        </div>
                    )}
                </div>;

        }

    };


    // const handleShareDialogCancel = useCallback(() => {
    //     setIFrameDialogOpened(false);
    // }, []);

    if (isWorkspaceLoading || !tileData || folders.length === 0) {
        return <div className="workspace-page"><PageLoader message="Loading workspace..." minHeight="72vh" /></div>;
    }

    const isRootFolder =
        selectedFolder?.path === buildLibraryRootPath(context, tileData?.LibraryName);

    // const showNewFolderButton =
    //     files.length === 0 &&
    //     (
    //         !isRootFolder ||                    // Any nested folder
    //         selectedFolder?.children?.length > 0 // Root with existing children
    //     );
    	
  const showNewFolderButton =
    !!selectedFolder &&
    folderPathBread.length > 0 &&
    files.length === 0 &&
    !isRootFolder;


    return (
        <div className="workspace-page" data-testid="page-workspace-explorer">
            <div className="workspace-topbar">
                <div className="workspace-topbar-breadcrumb" data-testid="nav-top-breadcrumb">
                    <span
                        className="workspace-topbar-link"
                        onClick={() => navigate('/')}
                        data-testid="link-dashboard"
                    >
                        <Home20Regular className="workspace-topbar-home-icon" />
                        <span>Dashboard</span>
                    </span>
                    <ChevronRight12Regular className="workspace-topbar-separator-icon" />
                    <span className="workspace-topbar-current" data-testid="text-workspace-name">
                        {tileData?.TileName}
                    </span>
                </div>
                <div className="workspace-topbar-actions">
                    {canCreateRequest && (
                        <DefaultButton
                            className="workspace-new-request-btn"
                            onClick={projectCreation}
                            data-testid="button-new-request"
                        >
                            <Add20Regular className="workspace-btn-icon" />
                            <span>New Request</span>
                        </DefaultButton>
                    )}
                </div>
            </div>

            <div className="workspace-body">
                <Sidebar
                    folders={folders}
                    selectedFolderId={selectedFolder ? selectedFolder.id : "0"}
                    onFolderSelect={handleFolderSelect}
                    onFolderAction={handleFolderAction}
                    recycleBinCount={deletedData.length}
                    approvalCount={approvalData.length}
                    onRecycleBinClick={() => setTables("Recycle")}
                    onApprovalClick={() => navigate('/approvals', { state: { from: `/workspace/${workspaceId}`, libName: tileData?.LibraryName, tileName: tileData?.TileName } })}
                    onAdvancedSearchClick={() => navigate('/Search', { state: { from: `/workspace/${workspaceId}`, libName: tileData?.LibraryName } })}
                    onArchiveClick={() => { setTables("Archive"); }}
                    LibDetails={tileData}
                    archiveCount={archiveData.length}
                    buttons={visibleButtons.filter((btn) => btn.ButtonType === "Folder")}
                    permittedButtons={permittedButtons}
                    expandedFolders={expandedFolders}
                />

                <div className="workspace-content">
                    {isFolderMetadataLoading && (
                        <div role="status" aria-live="polite" className="workspace-folder-metadata-loading">
                            <Spinner size={SpinnerSize.small} label="Loading folder details..." />
                        </div>
                    )}
                    {folderMetadataError && (
                        <div role="alert" className="workspace-folder-metadata-error">
                            {folderMetadataError}
                        </div>
                    )}
                    {selectedFolder && folderPathBread.length > 0 && (
                        <div className="workspace-content-header">
                            <div className="workspace-folder-breadcrumb" data-testid="nav-folder-breadcrumb">
                                {folderPathBread.map((node, i) => (
                                    <span key={node.id} className="workspace-folder-breadcrumb-segment">
                                        {i > 0 && <ChevronRight12Regular className="workspace-folder-breadcrumb-chevron" />}
                                        <span
                                            className={`workspace-folder-breadcrumb-item ${i === folderPathBread.length - 1 ? 'workspace-folder-breadcrumb-current' : ''}`}
                                            onClick={() => {
                                                if (i < folderPathBread.length - 1) handleFolderSelect(node);
                                            }}
                                            data-testid={`breadcrumb-folder-${node.id}`}
                                        >
                                            {node.name}
                                        </span>
                                    </span>
                                ))}
                            </div>
                            <div className="workspace-content-actions">
                                {canShowButtons && (
                                    <>
                                        {tables === "" ? <>
                                            {(selectedFolder?.children.length === 0 && selectedFolder?.name !== tileData?.LibraryName) ?
                                                <Menu>
                                                    <MenuTrigger disableButtonEnhancement>
                                                        <Button
                                                            appearance="subtle"
                                                            iconPosition="after"
                                                            icon={<ChevronDown24Regular className="table-action-btn" />}
                                                            className="workspace-upload-btn"
                                                        ><span>Create or Upload</span></Button>
                                                    </MenuTrigger>

                                                    <MenuPopover
                                                        style={{
                                                            boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
                                                            padding: "15px"
                                                        }}
                                                    >
                                                        <MenuList>
                                                            <MenuItem
                                                                key="folder"
                                                                icon={<ArrowUpload20Regular style={{ color: "#0078D4" }} />}
                                                                onClick={() => {
                                                                    setFileType("upload");
                                                                    setIsOpenUploadPanel(true);
                                                                }}
                                                            >
                                                                Files Upload
                                                            </MenuItem>
                                                            <MenuItem
                                                                key="word"
                                                                icon={<Icon iconName="WordDocument" style={{ color: "#2B579A", fontSize: 20 }} />}
                                                                onClick={() => {
                                                                    setFileType("docx");
                                                                    setIsOpenUploadPanel(true);
                                                                }}
                                                            >
                                                                Word Document
                                                            </MenuItem>
                                                            <MenuItem
                                                                key="excel"
                                                                icon={<Icon iconName="ExcelDocument" style={{ color: "#217346", fontSize: 20 }} />}

                                                                onClick={() => {
                                                                    setFileType("xlsx");
                                                                    setIsOpenUploadPanel(true);
                                                                }}
                                                            >
                                                                Excel Document
                                                            </MenuItem>
                                                        </MenuList>
                                                    </MenuPopover>
                                                </Menu>
                                                // <DefaultButton text="Create or Upload" menuProps={uploadMenuProps} styles={{ root: { marginRight: 8 } }} />



                                                : <></>}
                                            {/* {files.length === 0 ?
                                            <PrimaryButton
                                                onClick={() => { setIsOpenFolderPanel(true); setFolderName(""); setFolderNameErr(""); }}
                                                className="workspace-new-folder-btn"
                                                data-testid="button-new-folder"
                                            >
                                                <FolderAdd20Regular className="workspace-btn-icon" />
                                                <span>{DisplayLabel.NewFolder} </span>
                                            </PrimaryButton> : <></>} */}

                                            {showNewFolderButton && (
                                                <PrimaryButton
                                                    onClick={() => {
                                                        setIsOpenFolderPanel(true);
                                                        setFolderName("");
                                                        setFolderNameErr("");
                                                    }}
                                                    className="workspace-new-folder-btn"
                                                    data-testid="button-new-folder"
                                                >
                                                    <FolderAdd20Regular className="workspace-btn-icon" />
                                                    <span>{DisplayLabel.NewFolder}</span>
                                                </PrimaryButton>
                                            )}
                                        </> : <> </>
                                        }


                                    </>
                                )}



                                {selectedFolder?.children.length !== 0 && (
                                    <Menu>
                                        <MenuTrigger disableButtonEnhancement>
                                            <Button
                                                appearance="transparent"
                                                iconPosition="after"
                                                icon={<FluentIcons.Board24Regular />}
                                            />
                                        </MenuTrigger>

                                        <MenuPopover
                                            style={{
                                                padding: "8px",
                                                borderRadius: "8px",
                                                boxShadow: "0 8px 24px rgba(0,0,0,0.15)",
                                                minWidth: "140px"
                                            }}
                                        >
                                            <div
                                                style={getItemStyle('List View')}
                                                onClick={() => setViewListSetting('List View')}
                                                onMouseEnter={(e) => {
                                                    if (viewListSetting !== 'List View')
                                                        e.currentTarget.style.backgroundColor = "#F3F2F1";
                                                }}
                                                onMouseLeave={(e) => {
                                                    if (viewListSetting !== 'List View')
                                                        e.currentTarget.style.backgroundColor = "transparent";
                                                }}
                                            >
                                                <FluentIcons.List20Regular
                                                    style={{
                                                        color: viewListSetting === 'List View' ? "#0F6CBD" : "#605E5C"
                                                    }}
                                                />
                                                List
                                            </div>
                                            <div
                                                style={getItemStyle('Tiles View')}
                                                onClick={() => setViewListSetting('Tiles View')}
                                                onMouseEnter={(e) => {
                                                    if (viewListSetting !== 'Tiles View')
                                                        e.currentTarget.style.backgroundColor = "#F3F2F1";
                                                }}
                                                onMouseLeave={(e) => {
                                                    if (viewListSetting !== 'Tiles View')
                                                        e.currentTarget.style.backgroundColor = "transparent";
                                                }}
                                            >
                                                <FluentIcons.Grid20Regular
                                                    style={{
                                                        color: viewListSetting === 'Tiles View' ? "#0F6CBD" : "#605E5C"
                                                    }}
                                                />
                                                Tiles
                                            </div>
                                        </MenuPopover>
                                    </Menu>
                                )}

                            </div>
                        </div>
                    )}

                    {selectedFolder ? (
                        <>
                            {bindTable()}
                        </>
                        // <></>
                    ) : (
                        <div className="empty-state">
                            <div className="empty-state-icon">
                                <span className="empty-state-emoji">📁</span>
                            </div>
                            <h2 className="empty-state-title" data-testid="text-empty-title">Select a Folder</h2>
                            <p className="empty-state-description" data-testid="text-empty-desc">
                                Choose a folder from the sidebar to view its contents.
                                Documents are displayed only at the final folder level.
                            </p>
                        </div>
                    )}
                </div>
            </div>
            <Panel
                headerText={panelTitle}
                isOpen={isOpenCommonPanel}
                onDismiss={dismissCommanPanel}
                closeButtonAriaLabel="Close"
                type={panelSize}
                onRenderFooterContent={() => <>{actionButton}<DefaultButton onClick={dismissCommanPanel} >Cancel</DefaultButton></>}
                isFooterAtBottom={true}
            >
                <div
                    //style={{ marginTop: "10px" }}
                    style={{
                        overflowY: "auto",
                        maxHeight: "80vh"
                    }}
                >
                    {/* <div className="grid">
                        <div className="row"> */}
                    {versionsPanelUrl ? renderVersionsPanel(versionsPanelUrl) : panelForm}
                    {/* </div>
                    </div> */}
                </div>
            </Panel>
            <Panel
                headerText={DisplayLabel.AddNewFolder}
                isOpen={isOpenFolderPanel}
                onDismiss={dismissFolderPanel}
                closeButtonAriaLabel="Close"
                type={PanelType.medium}
                onRenderFooterContent={() => (<>
                    <PrimaryButton onClick={createFolder} styles={{ root: { marginRight: 8 } }}>{DisplayLabel.Submit}</PrimaryButton>
                    <DefaultButton onClick={dismissFolderPanel}>{DisplayLabel.Cancel}</DefaultButton>
                </>)}
                isFooterAtBottom={true}
            >
                <Field>
                    <label>{DisplayLabel.Path}: <b>{
                        selectedFolder?.path
                            ?.replace(context.pageContext.web.serverRelativeUrl, "")
                            ?.replace(/^\/+/, "")
                    }</b></label>
                </Field>

                <Field >
                    <TextField
                        label={DisplayLabel.FolderName}
                        required value={folderName} onChange={(_, val) => {

                            setFolderName(val as string);

                            if (invalidCharsRegex.test(val as string)) {
                                setFolderNameErr(
                                    "Please enter a name that doesn't include any of these characters: \" * : < > ? / \\ |"
                                );
                            } else {
                                setFolderNameErr("");
                            }
                        }}
                        errorMessage={folderNameErr}
                    />
                    {/* <span style={{ color: "red" }}>{folderNameErr}</span> */}
                </Field>
            </Panel>

            {/* <IFrameDialog
                url={shareURL}
                width="800px !important"
                height="600px"
                hidden={!iFrameDialogOpened}
                onDismiss={() => setIFrameDialogOpened(false)}
                iframeOnLoad={(iframe) => console.log('Iframe loaded:', iframe)}
                modalProps={{
                    isBlocking: true,

                }}
                dialogContentProps={{
                    type: DialogType.close,
                    showCloseButton: true
                }}
            /> */}
            {iFrameDialogOpened && (
                <IFrameDialogPopup url={shareURL} isOpen={iFrameDialogOpened} dismissPanel={dissmissSharePopup} />
            )}



            <AdvancePermission isOpen={isPanelOpen} context={context} folderId={itemId} LibraryName={tileData?.LibraryName} dismissPanel={onDismiss} />
            {/* {tileData && <ProjectEntryForm isOpen={isCreateProjectPopupOpen} dismissPanel={dissmissProjectCreationPanel} context={context} LibraryDetails={tileData} admin={admin} FormType={formType} folderObject={projectUpdateData} folderPath={selectedFolder?.path} ChildFolderRoleInheritance={tileData?.AllowChildInheritance} />} */}

            {tileData &&
                <ProjectEntryForm
                    isOpen={isCreateProjectPopupOpen}
                    dismissPanel={dissmissProjectCreationPanel}
                    context={context}
                    LibraryDetails={tileData}
                    admin={admin}
                    FormType={formType}
                    folderObject={projectUpdateData}
                    folderPath={selectedFolder?.path || ""}
                    ChildFolderRoleInheritance={tileData?.AllowChildInheritance}
                    onFolderCreated={refreshCurrentFolder}   // NEW
                />

            }
            <UploadFiles context={context} isOpenUploadPanel={isOpenUploadPanel} folderName={selectedFolder?.name || ""} folderPath={selectedFolder?.path?.replace(context.pageContext.web.serverRelativeUrl, "")?.replace(/^\/+/, "") || ""} dismissUploadPanel={dismissUploadPanel} libName={tileData?.LibraryName} files={files} folderObject={selectedFolder} LibraryDetails={tileData} filetype={fileType} FileData={files} />

            <ConfirmationDialog hideDialog={hideDialog} closeDialog={closeDialog} handleConfirm={handleConfirm} msg={message} />
            <ConfirmationDialog hideDialog={hideDialogCheckOut} closeDialog={closeDialogCheckOut} handleConfirm={handleConfirmCheckOut} msg={message} />
            <PopupBox isPopupBoxVisible={isPopupBoxVisible} hidePopup={hidePopup} msg={alertMsg} type={popupType} />
            <PopupBox isPopupBoxVisible={isShowCommnPopupBoxVisible} hidePopup={hideCommonPopup} msg={alertMsg} type="warning" />
        </div>
    );
};

export default React.memo(Workspace);
