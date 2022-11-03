var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { SPHttpClient } from '@microsoft/sp-http';
import { ServiceBase } from "./ServiceBase";
// import { CommonServices } from "./CommonServices";
// import { ICustomerContact, IProposalAlternate } from "../CommonProps/CommonProps";
var SharePointServices = /** @class */ (function (_super) {
    __extends(SharePointServices, _super);
    function SharePointServices(context) {
        return _super.call(this, context) || this;
    }
    // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
    SharePointServices.prototype.getUpdateHeader = function () {
        return {
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
        };
    };
    SharePointServices.prototype.getDeleteHeader = function () {
        return {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
        };
    };
    SharePointServices.prototype.getListData = function (listName, filter) {
        if (filter === void 0) { filter = null; }
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('".concat(listName, "')/Items") + (filter ? "?" + filter : ""), SPHttpClient.configurations.v1);
    };
    SharePointServices.prototype.getListItemById = function (listName, itemId, selectQuery) {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('".concat(listName, "')/Items(").concat(itemId, ")") + (selectQuery ? "?" + selectQuery : ""), SPHttpClient.configurations.v1);
    };
    SharePointServices.prototype.uploadFile = function (folderPath, fileName, fileContent) {
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativeUrl('" + folderPath + "')/Files/Add(url='" + fileName + "', overwrite=" + true + ")?$Expand=ListItemAllFields";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: fileContent });
    };
    SharePointServices.prototype.uploadAttachment = function (listName, itemId, fileBuffer, fileName) {
        var url = this.context.pageContext.web.absoluteUrl +
            '/_api/web/lists/getbytitle(\'' +
            listName + '\')/items(' + itemId +
            ')/AttachmentFiles/add(FileName=\'' + encodeURIComponent(fileName) + '\')';
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: fileBuffer });
    };
    SharePointServices.prototype.ensureFolder = function (folderURL) {
        var data = JSON.stringify({ 'ServerRelativeUrl': folderURL });
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/folders";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: data });
    };
    SharePointServices.prototype.sendEmail = function (body) {
        var url = this.context.pageContext.web.absoluteUrl + "/_api/SP.Utilities.Utility.SendEmail";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: JSON.stringify(body) });
    };
    // public createListItem(listName, body) {
    //     var url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
    //     return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: JSON.stringify(body) });
    // }
    SharePointServices.prototype.createOrUpdateListItem = function (listName, body, itemID) {
        if (itemID === void 0) { itemID = 0; }
        var request = {};
        var obj = JSON.parse(JSON.stringify(body));
        delete obj.Id;
        request.body = JSON.stringify(obj);
        // eslint-disable-next-line no-var
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items";
        if (itemID) {
            url += "(" + itemID + ")";
            request.headers = this.getUpdateHeader();
        }
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, request);
    };
    SharePointServices.prototype.deleteListItem = function (listName, itemId) {
        // eslint-disable-next-line no-var
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/items(" + itemId + ")";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { headers: this.getDeleteHeader() });
    };
    // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
    SharePointServices.prototype.deleteAttachmentItem = function (listName, itemId, fullName) {
        // eslint-disable-next-line no-var
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('" + listName + "')/GetItemById(" + itemId + ")/attachmentfiles('" + fullName + "')";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { headers: this.getDeleteHeader() });
    };
    SharePointServices.prototype.getLoggedInUserGroups = function () {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/currentuser/groups", SPHttpClient.configurations.v1);
    };
    SharePointServices.prototype.getLoggedInUserInfo = function () {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/currentuser", SPHttpClient.configurations.v1);
    };
    SharePointServices.prototype.getFields = function (listName, filter) {
        if (filter === void 0) { filter = null; }
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/fields" + (filter ? "?" + filter : ""), SPHttpClient.configurations.v1);
    };
    SharePointServices.prototype.bulkDeleteItems = function (listName, itemIds) {
        var _this = this;
        var completed = 0;
        return new Promise(function (reolve) {
            itemIds.forEach(function (e) {
                _this.deleteListItem(listName, e).then(function () {
                    completed++;
                    if (completed == itemIds.length) {
                        reolve(true);
                    }
                });
            });
        });
    };
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
    SharePointServices.prototype.createFolderStructure = function (listName, itemId) {
        var _this = this;
        var roolFolderPath = this.context.pageContext.web.serverRelativeUrl + "/" + listName;
        return new Promise(function (resolve, reject) {
            _this.createFolder(roolFolderPath, itemId).then(function (res) {
                resolve(true);
            });
        });
    };
    SharePointServices.prototype.createFolder = function (docLibName, newFolderName) {
        var _this = this;
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1='" + docLibName + "'&@a2='" + newFolderName + "'";
        return new Promise(function (resolve, rejectt) {
            _this.checkFolderExistOrNot(docLibName + "/" + newFolderName).then(function (res) {
                if (!res) {
                    _this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, { body: null }).then(function (Res) {
                        resolve(true);
                    });
                }
                else {
                    resolve(true);
                }
            });
        });
    };
    SharePointServices.prototype.checkFolderExistOrNot = function (folderPath) {
        var _this = this;
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/getfolderbyserverrelativeurl('" + folderPath + "')/Exists";
        return new Promise(function (resolve, reject) {
            _this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then(function (res) {
                if (res.ok) {
                    res.json().then(function (resJSON) {
                        resolve(resJSON.value);
                    });
                }
            });
        });
    };
    SharePointServices.prototype.getItemVersions = function (listName, itemId, filter) {
        // $Filter=substringof(%27.0%27,VersionLabel)
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + listName + "')/items(" + itemId + ")/versions" + (filter ? "?" + filter : "");
        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    };
    SharePointServices.prototype.publishFile = function (fileServerRelativeUrl) {
        var url = this.context.pageContext.web.absoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + this.context.pageContext.site.serverRelativeUrl + "/" + fileServerRelativeUrl + "')/Publish()";
        return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
            body: null
        });
    };
    return SharePointServices;
}(ServiceBase));
export { SharePointServices };
//# sourceMappingURL=SharePointServices.js.map