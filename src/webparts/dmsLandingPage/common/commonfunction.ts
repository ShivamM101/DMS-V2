import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";

import { DocumentPdf20Regular, Document20Regular, DocumentTable20Regular, Image20Regular, Cube20Regular, DocumentText20Regular } from '@fluentui/react-icons';

export function buildBreadcrumbPath(folder: any, allFolders: any[]): any[] {
    const path: any[] = [];

    function findPath(nodes: any[], target: string): boolean {
        for (const node of nodes) {
            if (node.id === target) {
                path.push(node);
                return true;
            }
            if (node.children) {
                if (findPath(node.children, target)) {
                    path.unshift(node);
                    return true;
                }
            }
        }
        return false;
    }

    findPath(allFolders, folder.id);
    return path;
}
export const buildFolderHierarchy = (
    folders: any[],
    libraryRoot: string
): any[] => {

    const map = new Map<string, any>();
    const tree: any[] = [];
    folders.forEach(folder => {
        const name = folder.FileRef.split("/").pop() || "";

        map.set(folder.FileRef, {
            // id: folder.Id,
            id: folder.ID,
            name,
            path: folder.FileRef,
            children: [],
            isLastLevel: false,
            ...folder,
        });
    });

    folders.forEach(folder => {
        const node = map.get(folder.FileRef)!;
        if (folder.FileDirRef === libraryRoot) {
            tree.push(node);
        } else {
            const parent = map.get(folder.FileDirRef);
            if (parent) {
                parent.children.push(node);
            }
        }
    });

    const sortChildren = (nodes: any[]) => {
        nodes.sort((a: any, b: any) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
        nodes.forEach(node => {
            if (node.children.length > 0) {
                sortChildren(node.children);
            }
        });
    };

    const markLastLevel = (nodes: any[]) => {
        nodes.forEach(node => {
            if (node.children.length === 0) {
                node.isLastLevel = true;
            } else {
                markLastLevel(node.children);
            }
        });
    };

    markLastLevel(tree);
    sortChildren(tree);
    return tree;
};

export const buildLibraryRootPath = (context: WebPartContext, libName: string) => {
    const webRelativeUrl = context.pageContext.web.serverRelativeUrl;

    return webRelativeUrl === "/"
        ? `/${libName}`
        : `${webRelativeUrl}/${libName}`;
};

export const getAllDocuments = async (
    context: WebPartContext,
    folderPath: string
) => {
    const sp = spfi().using(SPFx(context));

    const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files
        .select("*,ListItemAllFields/*,CheckedOutByUser")
        .expand("ListItemAllFields,CheckedOutByUser")();

    return files;
};

export const getChildFolders = async (
    context: WebPartContext,
    folderPath: string
) => {
    const sp = spfi().using(SPFx(context));

    try {
        const subfolders = await sp.web
            .getFolderByServerRelativePath(folderPath)
            .folders
            .select("Name", "ServerRelativeUrl", "ItemCount", "ListItemAllFields/*")
            .expand("ListItemAllFields")();

        const nodes = subfolders
            .filter((f: any) => f.Name && f.Name.toLowerCase() !== "forms" && !f.Name.startsWith("."))
            .map((f: any) => ({
                id: f.ListItemAllFields?.ID ? String(f.ListItemAllFields.ID) : f.ServerRelativeUrl,
                name: f.Name,
                path: f.ServerRelativeUrl,
                children: [],
                isLoaded: false,
                isLoading: false,
                isLastLevel: false,
                FileRef: f.ServerRelativeUrl,
                FileDirRef: folderPath,
                FSObjType: 1,
                ItemCount: f.ItemCount,
                Modified: f.ListItemAllFields?.Modified,
                Created: f.ListItemAllFields?.Created,
                Editor: f.ListItemAllFields?.Editor,
                Author: f.ListItemAllFields?.Author
            }))
            .sort((a: any, b: any) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

        // SharePoint's ListItemAllFields projection on the /Folders endpoint does
        // not always include the display names of the person fields (Author/Editor).
        // Enrich them from the library's list items using the same
        // Author/Title + expand("Author,Editor") pattern used across the app.
        // This step is optional: if it fails, the folder list still works exactly
        // as before and the grid simply keeps the (possibly blank) person columns.
        try {
            const listId = (subfolders[0] as any)?.ListItemAllFields?.ListId;
            // Fallback: when ListItemAllFields does not expose ListId, derive the
            // library name from the folder path (libraries live at the web root).
            const webRelativeUrl = context.pageContext.web.serverRelativeUrl;
            const libraryName = folderPath.replace(webRelativeUrl, "").split("/").filter(Boolean)[0];

            let folderItems: any[] = [];
            if (listId && nodes.length > 0) {
                folderItems = await sp.web.lists.getById(listId).items
                    .select("Id", "FileRef", "Author/Title", "Editor/Title")
                    .expand("Author,Editor")
                    .filter(`FSObjType eq 1 and FileDirRef eq '${folderPath}'`)
                    .top(5000)();
            } else if (libraryName && nodes.length > 0) {
                folderItems = await sp.web.lists.getByTitle(libraryName).items
                    .select("Id", "FileRef", "Author/Title", "Editor/Title")
                    .expand("Author,Editor")
                    .filter(`FSObjType eq 1 and FileDirRef eq '${folderPath}'`)
                    .top(5000)();
            }
            const itemsByPath = new Map<string, any>(folderItems.map((it: any) => [it.FileRef, it]));
            nodes.forEach((node: any) => {
                const item = itemsByPath.get(node.path);
                if (item) {
                    if (item.Author?.Title) node.Author = { Id: item.Author.Id, Title: item.Author.Title };
                    if (item.Editor?.Title) node.Editor = { Id: item.Editor.Id, Title: item.Editor.Title };
                }
            });

            // TEMP-DIAG: verify person metadata on the grid rows. Remove after confirming.
            if (nodes.length > 0) {
                console.log("[DMS-TEMP] child folders enriched:", {
                    folderPath,
                    count: nodes.length,
                    sample: {
                        name: nodes[0].name,
                        Created: nodes[0].Created,
                        Modified: nodes[0].Modified,
                        Author: nodes[0].Author,
                        Editor: nodes[0].Editor
                    }
                });
            }
        } catch (enrichError) {
            console.warn("Unable to enrich folder author/editor names for path:", folderPath, enrichError);
        }

        return nodes;
    } catch (error) {
        console.error("Error fetching child folders for path:", folderPath, error);
        return [];
    }
};


export const fileTypeConfig: Record<string, { IconName: typeof DocumentPdf20Regular; className: string; label: string; }> = {
    pdf: { IconName: DocumentPdf20Regular, className: 'doc-icon-pdf', label: 'PDF' },
    docx: { IconName: DocumentText20Regular, className: 'doc-icon-word', label: 'Word' },
    xlsx: { IconName: DocumentTable20Regular, className: 'doc-icon-excel', label: 'Excel' },
    png: { IconName: Image20Regular, className: 'doc-icon-image', label: 'Image' },
    jpg: { IconName: Image20Regular, className: 'doc-icon-image', label: 'Image' },
    dwg: { IconName: Cube20Regular, className: 'doc-icon-cad', label: 'AutoCAD' },
    other: { IconName: Document20Regular, className: 'doc-icon-other', label: 'File' },
};

/**
 * Version number of one document in a DMS library ('1.0', '2.0', …) – the value
 * the listings print behind their 'v'.
 *
 * 'Level' is the DMS version column: the SOP publishing service writes the SOP's
 * own version number into it ('Versions' = 'v2.0' → '2.0'), and a document
 * uploaded straight into the DMS gets its own level there as well. It therefore
 * holds the version of the DOCUMENT for EVERY file type – Excel and Word exactly
 * like PDF – and it is the value the SOP listing (SOP Library / BRCQS) shows.
 *
 * The built-in SharePoint file version ('OData__UIVersionString') is NOT that
 * version: it only counts how often that physical file was overwritten, so a
 * freshly uploaded revision of a SOP starts at '1.0' again even when it is
 * version 2.0 of the SOP. It is used only as a fallback for legacy rows that
 * carry no 'Level', so those rows keep the value they showed before.
 *
 * Returns '' when neither column holds a value – the caller decides how to
 * render the empty cell.
 */
export const getDocumentVersion = (
    level?: string | number,
    fileVersion?: string | number
): string => {
    const fromLevel = String(level ?? "").trim();
    if (fromLevel) return fromLevel.replace(/^v/i, "").trim();
    return String(fileVersion ?? "").trim().replace(/^v/i, "").trim();
};

export const checkExtension = (fileName: string): boolean => {
    if (!fileName) return false;
    const extension = fileName.split(".").pop()?.toLowerCase();
    const allowedExtensions = ["pdf", "txt", "jpg", "jpeg", "png", "gif", "bmp"];
    return !allowedExtensions.includes(extension || "");
};

export const checkButtons = (input: string): boolean => {
    if (!input) return false;
    const buttonTypes = ["OpenInApp", "CheckIn", "DiscardCheckOut", "CheckOut", "Preview"];
    return !buttonTypes.includes(input);
};

export const getOpenAppURL = (filePath: string, SiteURL: string) => {
    const portalUrl = new URL(SiteURL).origin;
    if (!filePath) return;
    const extension = filePath.split('.').pop()?.toLowerCase();
    if (!extension) return;

    let appUrl: string | null = null;
    switch (extension) {
        case 'xls':
        case 'xlsx':
            appUrl = `ms-excel:ofe|u|${portalUrl}${filePath}`;
            break;
        case 'doc':
        case 'docx':
            appUrl = `ms-word:ofe|u|${portalUrl}${filePath}`;
            break;
        case 'ppt':
        case 'pptx':
            appUrl = `ms-powerpoint:ofe|u|${portalUrl}${filePath}`;
            break;
    }

    if (appUrl) {
        window.open(appUrl, '_blank');
    }
};

/* -------------------------------------------------------------------------- *
 * SharePoint version history (Versions.aspx)
 *
 * The built-in version history page addresses ONE document through the id of
 * the library the file belongs to plus the id of that file's list item:
 *
 *     /_layouts/15/Versions.aspx?list={library id}&ID={list item id}
 *
 * Those two ids are the only reference SharePoint needs, and they are
 * unambiguous for every file type (PDF, Word, Excel, PowerPoint, …), for every
 * folder depth and for names containing spaces or special characters.
 *
 * The previous implementation built the link as
 *
 *     Versions.aspx?list=<library title>&FileName=<server relative url>&IsDlg=<item id>
 *
 * which is why only some documents worked: 'IsDlg' is the dialog flag of the
 * layouts page and not a place for a document id, so no 'ID' was ever sent and
 * SharePoint had to resolve the item from the library *title* plus the file name
 * given as a raw string. Whenever that string lookup did not match (nested
 * folders, encoded or special characters in the path, library title differing
 * from its URL) the page failed with
 * "Item does not exist. It may have been deleted by another user."
 *
 * Both ids are resolved from SharePoint at run time – no library id, item id or
 * path is hardcoded – and the caller gets a clear message instead of a broken
 * page when a document cannot be resolved.
 * -------------------------------------------------------------------------- */
export interface IVersionHistoryTarget {
    /** Version history page of the document – set only when the reference is valid. */
    url?: string;
    /** Message for the user – set instead of 'url' when the document cannot be opened. */
    error?: string;
}

/** Rows of the document listing: only the properties the version history needs. */
export interface IVersionHistoryFile {
    ServerRelativeUrl?: string;
    ListItemAllFields?: {
        Id?: number;
        ID?: number;
        FileRef?: string;
    };
}

/** Library ids resolved during this page load, keyed per web + library URL. */
const libraryIdCache = new Map<string, string>();

const SHAREPOINT_GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Server-relative path as SharePoint's REST endpoints return it
 * ('/sites/DMS/BRCQS/Rev%201.0/File%20v1.0.xlsx', spaces escaped as %20) turned
 * back into the literal path. PnP path helpers encode the value they are given,
 * so they need this decoded form – handing them the encoded one escapes the '%'
 * a second time and the file is not found.
 */
export const decodeServerRelativePath = (path: string): string => {
    const value = String(path ?? "").trim();
    if (!value) return "";
    try {
        return decodeURIComponent(value);
    } catch {
        // A literal '%' that is not an escape sequence: keep the path unchanged.
        return value;
    }
};

/** List/library GUID exactly as SharePoint returns it, or '' when the value is not a GUID. */
const normalizeSharePointGuid = (value?: string): string => {
    const guid = String(value ?? "").trim().replace(/^\{/, "").replace(/\}$/, "");
    return SHAREPOINT_GUID_PATTERN.test(guid) ? guid : "";
};

/** List item id as a positive integer, or 0 when the value is not a usable id. */
const toListItemId = (value?: number | string): number => {
    const itemId = Number(value);
    return Number.isInteger(itemId) && itemId > 0 ? itemId : 0;
};

/**
 * Version history URL of one document inside one library.
 *
 * 'siteUrl' is the web the library lives in, 'libraryId' the GUID of that
 * library and 'itemId' the id of the document's list item. Callers get these
 * from SharePoint (see resolveVersionHistoryTarget); building the URL from a
 * library or file NAME is what used to break.
 */
export const buildVersionHistoryUrl = (
    siteUrl: string,
    libraryId: string,
    itemId: number | string
): string => {
    const baseUrl = `${String(siteUrl ?? "").replace(/\/+$/, "")}/_layouts/15/Versions.aspx`;
    const listId = encodeURIComponent(normalizeSharePointGuid(libraryId) || String(libraryId ?? "").trim());
    const listItemId = encodeURIComponent(String(itemId ?? "").trim());

    return `${baseUrl}?list=${listId}&ID=${listItemId}&IsDlg=1`;
};

/**
 * Server-relative URL of the library a file lives in. Every document library of
 * the DMS sits directly below the web ('/sites/DMS/<library>/…'), so the file's
 * own path is the source of truth; the library of the workspace is only used
 * when the file path cannot be read.
 */
const getLibraryPathForFile = (webRelativeUrl: string, filePath: string, libraryName: string): string => {
    const configuredPath = `${webRelativeUrl}/${libraryName}`.replace(/\/+$/, "");
    if (libraryName && filePath.toLowerCase().startsWith(`${configuredPath.toLowerCase()}/`)) {
        return configuredPath;
    }

    const relativePath = (webRelativeUrl && filePath.toLowerCase().startsWith(`${webRelativeUrl.toLowerCase()}/`))
        ? filePath.substring(webRelativeUrl.length + 1)
        : filePath.replace(/^\/+/, "");
    const [librarySegment] = relativePath.split("/").filter(Boolean);

    return librarySegment ? `${webRelativeUrl}/${librarySegment}` : configuredPath;
};

/**
 * GUID of a document library, resolved from SharePoint by its server-relative
 * URL (title as fallback) and cached for the rest of the page load.
 * Returns '' when the library cannot be found.
 */
const getLibraryId = async (context: WebPartContext, libraryPath: string, libraryName: string): Promise<string> => {
    const cacheKey = `${context.pageContext.web.absoluteUrl.toLowerCase()}|${libraryPath.toLowerCase()}`;
    const cached = libraryIdCache.get(cacheKey);
    if (cached) return cached;

    const sp = spfi().using(SPFx(context));
    const lookups: { source: string; run: () => Promise<{ Id?: string }> }[] = [];

    if (libraryPath) {
        lookups.push({ source: libraryPath, run: () => sp.web.getList(libraryPath).select("Id")() });
    }
    if (libraryName && libraryName.toLowerCase() !== libraryPath.toLowerCase()) {
        lookups.push({ source: libraryName, run: () => sp.web.lists.getByTitle(libraryName).select("Id")() });
    }

    for (const lookup of lookups) {
        try {
            const libraryId = normalizeSharePointGuid((await lookup.run())?.Id);
            if (libraryId) {
                libraryIdCache.set(cacheKey, libraryId);
                return libraryId;
            }
        } catch (error) {
            console.warn(`Version history: the library '${lookup.source}' could not be resolved.`, error);
        }
    }

    return "";
};

/**
 * Resolves everything the SharePoint version history page needs for one file of
 * the document listing and returns either its URL or a message for the user.
 *
 * The library comes from the file's own path, the item id from the listing. When
 * the listing did not carry a usable item id SharePoint is asked for it – and if
 * that fails too the document really is gone, so no version dialog is shown at
 * all and the user gets a meaningful message instead.
 */
export const resolveVersionHistoryTarget = async (
    context: WebPartContext,
    file: IVersionHistoryFile,
    libraryName: string
): Promise<IVersionHistoryTarget> => {
    const webRelativeUrl = String(context.pageContext.web.serverRelativeUrl ?? "").replace(/\/+$/, "");
    const library = String(libraryName ?? "").trim();
    const fileServerRelativeUrl = decodeServerRelativePath(file?.ServerRelativeUrl || file?.ListItemAllFields?.FileRef || "");

    if (!fileServerRelativeUrl) {
        return {
            error: "Version history could not be opened because this document has no SharePoint file reference. Refresh the folder and try again."
        };
    }

    const libraryPath = getLibraryPathForFile(webRelativeUrl, fileServerRelativeUrl, library);
    const libraryId = await getLibraryId(context, libraryPath, library);
    if (!libraryId) {
        return {
            error: `Version history could not be opened because the library '${libraryPath.split("/").pop()}' could not be found in SharePoint.`
        };
    }

    let itemId = toListItemId(file?.ListItemAllFields?.Id ?? file?.ListItemAllFields?.ID);
    if (!itemId) {
        try {
            const sp = spfi().using(SPFx(context));
            const fileInfo = await sp.web
                .getFileByServerRelativePath(fileServerRelativeUrl)
                .select("ServerRelativeUrl", "ListItemAllFields/Id")
                .expand("ListItemAllFields")() as { ListItemAllFields?: { Id?: number } };
            itemId = toListItemId(fileInfo?.ListItemAllFields?.Id);
        } catch (error) {
            console.warn("Version history: the list item of the document could not be read.", error);
        }
    }

    if (!itemId) {
        return {
            error: "Version history could not be opened because this document no longer exists in SharePoint. It may have been moved, renamed or deleted by another user."
        };
    }

    return { url: buildVersionHistoryUrl(context.pageContext.web.absoluteUrl, libraryId, itemId) };
};

