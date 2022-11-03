/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable promise/param-names */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable eqeqeq */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@microsoft/sp-http';
import { ServiceBase } from "./ServiceBase";
// import { CommonServices } from "./CommonServices";
// import { ICustomerContact, IProposalAlternate } from "../CommonProps/CommonProps";



export class SharePointServices extends ServiceBase {
    constructor(context: WebPartContext) {
        super(context);
    }

    private getUpdateHeader() {
        return {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
        };
    }

    private getDeleteHeader() {
        return {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
        };
    }

    public getListData(listName: string, filter: string = null) {
        return this.context.spHttpClient.get(
            this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/Items` + (filter ? "?" + filter : ""),
            SPHttpClient.configurations.v1);
    }
    public getListItemById(listName: any, itemId: any, selectQuery: string) {
        return this.context.spHttpClient.get(
            this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${listName}')/Items(${itemId})` + (selectQuery ? "?" + selectQuery : ""),
            SPHttpClient.configurations.v1);
    }

    public uploadFile(folderPath: string, fileName: string, fileContent: any) {
        const url = this.context.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + folderPath + "')/Files/Add(url='" + fileName + "', overwrite=" + true + ")?$Expand=ListItemAllFields";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: fileContent });
    }

    public uploadAttachment(listName: string,
        itemId: string,
        fileBuffer: ArrayBuffer,
        fileName: string) {

        const url: string = this.context.pageContext.web.absoluteUrl +
            '/_api/web/lists/getbytitle(\'' +
            listName + '\')/items(' + itemId +
            ')/AttachmentFiles/add(FileName=\'' + encodeURIComponent(fileName) + '\')';
        return this.context.spHttpClient.post(
            url,
            SPHttpClient.configurations.v1,
            { body: fileBuffer });
    }

    public ensureFolder(folderURL: string) {
        const data = JSON.stringify({ 'ServerRelativeUrl': folderURL });
        const url = this.context.pageContext.web.absoluteUrl + "/_api/web/folders";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: data });
    }

    public sendEmail(body: any): Promise<any> {
        const url = this.context.pageContext.web.absoluteUrl + "/_api/SP.Utilities.Utility.SendEmail";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: JSON.stringify(body) });
    }

    // public createListItem(listName, body) {
    //     var url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
    //     return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: JSON.stringify(body) });
    // }

    public createOrUpdateListItem(listName: string, body: any, itemID = 0) {
        const request: any = {};
        const obj = JSON.parse(JSON.stringify(body))
        delete obj.Id;
        request.body = JSON.stringify(obj);
        let url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
        if (itemID) {
            url += "(" + itemID + ")";
            request.headers = this.getUpdateHeader();
        }
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, request);
    }

    public deleteListItem(listName: string, itemId: string | number) {
        const url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items(" + itemId + ")";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { headers: this.getDeleteHeader() });
    }

    public deleteAttachmentItem(listName: string, itemId: string, fullName: string) {
        const url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('" + listName + "')/GetItemById(" + itemId + ")/attachmentfiles('" + fullName + "')";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { headers: this.getDeleteHeader() })

    }

    public getLoggedInUserGroups() {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/currentuser/groups", SPHttpClient.configurations.v1);
    }

    public getLoggedInUserInfo() {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/currentuser", SPHttpClient.configurations.v1);
    }

    public getFields(listName: string, filter: string = null) {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/fields" + (filter ? "?" + filter : ""), SPHttpClient.configurations.v1)
    }
    // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
    public bulkDeleteItems(listName: any, itemIds: number[]) {
        let completed = 0;
        return new Promise((reolve, reject) => {
            itemIds.forEach(e => {
                this.deleteListItem(listName, e).then((_res: any) => {
                    completed++;
                    if (completed == itemIds.length) {
                        reolve(true);
                    }
                });
            });
        });
    }
    // public bulkSaveAlternates(listName: any, alternates: IProposalAlternate[], proposalId: any, commonService: CommonServices) {
    //     let completed = 0;
    //     return new Promise((resolve, reject) => {
    //         if (alternates.length > 0) {
    //             alternates.forEach(e => {
    //                 delete e.tempId;
    //                 e.ProposalId = proposalId;
    //                 e.Add = commonService.getNumberFromString(e.Add);
    //                 e.Deduct = commonService.getNumberFromString(e.Deduct);
    //                 this.createOrUpdateListItem(listName, e, e.Id).then((Res: any) => {
    //                     completed++;
    //                     if (completed == alternates.length) {
    //                         resolve(true);
    //                     }
    //                 })
    //             });
    //         } else {
    //             resolve(true);
    //         }
    //     })
    // }
    // public bulkSaveCustomerContacts(listName: any, contacts: ICustomerContact[], customerId: { toString: () => any; }) {
    //     let completed = 0;
    //     return new Promise((resolve, reject) => {
    //         if (contacts.length > 0) {
    //             contacts.forEach(e => {
    //                 e.CustomerId = customerId.toString();
    //                 delete e.tempId
    //                 this.createOrUpdateListItem(listName, e, e.Id).then((Res: any) => {
    //                     completed++;
    //                     if (completed == contacts.length) {
    //                         resolve(true);
    //                     }
    //                 })
    //             });
    //         } else {
    //             resolve(true);
    //         }
    //     })
    // }
    public createFolderStructure(listName: string, itemId: string) {
        const roolFolderPath = this.context.pageContext.web.serverRelativeUrl + "/" + listName;
        return new Promise((resolve, _reject) => {
            this.createFolder(roolFolderPath, itemId).then(_res => {
                resolve(true);
            });
        });
    }
    private createFolder(docLibName: string, newFolderName: string) {
        const url = this.context.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1='" + docLibName + "'&@a2='" + newFolderName + "'";
        return new Promise((resolve, _rejectt) => {
            this.checkFolderExistOrNot(docLibName + "/" + newFolderName).then(res => {
                if (!res) {
                    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: null }).then((_Res: any) => {
                        resolve(true);
                    })
                } else {
                    resolve(true);
                }
            });
        });
    }
    private checkFolderExistOrNot(folderPath: string) {
        const url = this.context.pageContext.web.absoluteUrl + "/_api/web/getfolderbyserverrelativeurl('" + folderPath + "')/Exists";
        return new Promise((resolve, _reject) => {
            this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((res: { ok: any; json: () => Promise<any>; }) => {
                if (res.ok) {
                    res.json().then((resJSON: any) => {
                        resolve(resJSON.value)
                    })
                }
            });
        });
    }
    public getItemVersions(listName: string, itemId: string, filter: string) {
        // $Filter=substringof(%27.0%27,VersionLabel)
        const url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listName + "')/items(" + itemId + ")/versions" + (filter ? "?" + filter : "")
        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    }
    public publishFile(fileServerRelativeUrl: string) {
        const url = this.context.pageContext.web.absoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + this.context.pageContext.site.serverRelativeUrl + "/" + fileServerRelativeUrl + "')/Publish()";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
            body: null

        })
    }
}