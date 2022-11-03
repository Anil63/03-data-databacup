import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ServiceBase } from "./ServiceBase";
export declare class SharePointServices extends ServiceBase {
    constructor(context: WebPartContext);
    private getUpdateHeader;
    private getDeleteHeader;
    getListData(listName: string, filter?: string): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    getListItemById(listName: any, itemId: any, selectQuery: string): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    uploadFile(folderPath: string, fileName: string, fileContent: any): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    uploadAttachment(listName: string, itemId: string, fileBuffer: ArrayBuffer, fileName: string): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    ensureFolder(folderURL: string): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    sendEmail(body: any): Promise<any>;
    createOrUpdateListItem(listName: string, body: any, itemID?: number): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    deleteListItem(listName: string, itemId: string | number): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    deleteAttachmentItem(listName: string, itemId: string, fullName: string): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    getLoggedInUserGroups(): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    getLoggedInUserInfo(): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    getFields(listName: string, filter?: string): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    bulkDeleteItems(listName: any, itemIds: number[]): Promise<unknown>;
    createFolderStructure(listName: string, itemId: string): Promise<unknown>;
    private createFolder;
    private checkFolderExistOrNot;
    getItemVersions(listName: string, itemId: string, filter: string): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
    publishFile(fileServerRelativeUrl: string): Promise<import("@microsoft/sp-http").SPHttpClientResponse>;
}
//# sourceMappingURL=SharePointServices.d.ts.map