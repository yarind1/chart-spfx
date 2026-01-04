import { EnvironmentType } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { IListCollection } from "./IList"; 
import { IListFieldCollection } from "./IListField";
import {IListItemCollection } from "./IListItem";

export class SharePointSerivceManager {
    public context: WebPartContext;
    public environmentType: EnvironmentType;

    public setup(context: WebPartContext, environmentType: EnvironmentType): void {
        this.context = context;
        this.environmentType = environmentType;
    }   

public async get(relativeEndpointUrl: string): Promise<any> {
    try {
        const response = await this.context.spHttpClient.get(
            `${this.context.pageContext.web.absoluteUrl}/_api/${relativeEndpointUrl}`,
            SPHttpClient.configurations.v1
        );

        
        if (!response.ok) {
            const errorResponse = await response.json();
            throw new Error(errorResponse.error?.message || response.statusText);
        }

        return await response.json();
    } catch (error) {
        return Promise.reject(error);
    }
}
    public getLists(showHiddenLists: boolean =false): Promise<IListCollection> {

        return this.get(`lists${showHiddenLists ? '?$filter=Hidden eq false' : ''}`);
    }
    
    public getListItems(listId:string, selectedFields?:string[]):Promise<IListItemCollection>{
        return this.get(`lists/getbyid('${listId}')/items${selectedFields ? `?$select=${selectedFields.join(',')}` : ''}`);
    }
       
    public getListFields(listId:string, ShowHiddenFields:boolean=false):Promise<IListFieldCollection>{
        return this.get(`lists/getbyid('${listId}')/fields${ShowHiddenFields ? '' : '?$filter=Hidden eq false'}`);
    }
}


const SharePointService = new SharePointSerivceManager();
export default SharePointService;

