import { WebPartContext } from "@microsoft/sp-webpart-base";

export class ServiceBase {
    protected context: WebPartContext;

    constructor(context: WebPartContext) {
        this.context = context;
    }
}