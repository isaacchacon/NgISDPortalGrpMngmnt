import {Injectable} from '@angular/core';
import {Headers, Http} from '@angular/http';
import 'rxjs/add/operator/toPromise';
declare var $:any;

import {SharepointListItem} from './sharepoint-list-item';
import {SharepointListItemConstructor} from './sharepoint-list-item-constructor';


@Injectable()
export class SharepointListsWebService{

	constructor(private http: Http) { }
	private serviceUrl = '_vti_bin/Lists.asmx';
	private xmlPayloadWrapperStart = `<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
  <soap12:Body>`;
	private xmlPayloadWrapperEnd = '</soap12:Body></soap12:Envelope>';
	private getListItemsPayload = `
    <GetListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/">
      <listName>listNamePayLoad</listName>
      <viewName>viewNamePayLoad</viewName>
      <query>queryPayload</query>
      <viewFields>viewFieldsPayload</viewFields>
      <rowLimit>rowLimitPayload</rowLimit>
      <QueryOptions>queryOptionsPayload</QueryOptions>
    </GetListItems>`;
	
	private updateListItemsPayload = `<UpdateListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>listNamePayLoad</listName><updates>updatesPayLoad</updates></UpdateListItems>`;

	private headers = new Headers({
								 'Content-Type': 'application/soap+xml; charset=utf-8',
							 });

	///If Filter query is not empty, we will use it to filter the CAML QUery.						 
    getListItems(ctor:SharepointListItemConstructor, filterQuery:[string,string]):Promise<SharepointListItem[]>{
		let dummyInstance = new ctor();
		let currentPayload = this.xmlPayloadWrapperStart+this.getListItemsPayload+this.xmlPayloadWrapperEnd;
		currentPayload = currentPayload.replace("listNamePayLoad",dummyInstance.getListName());
		currentPayload = currentPayload.replace("viewNamePayLoad", "");
		if(filterQuery){
			let camlQuery = "<Query><Where><Eq><FieldRef Name='FieldName' /><Value Type='Text'>ValueName</Value></Eq></Where></Query>"
			camlQuery = camlQuery.replace("FieldName", filterQuery[0]);
			camlQuery = camlQuery.replace("ValueName", filterQuery[1]);
			currentPayload = currentPayload.replace("queryPayload", camlQuery );
		}else{
			currentPayload = currentPayload.replace("queryPayload", "");
		}
		currentPayload = currentPayload.replace("viewFieldsPayload", "");
		currentPayload = currentPayload.replace("rowLimitPayload", "");
		currentPayload = currentPayload.replace("queryOptionsPayload", "");		
		return this.http.post(dummyInstance.getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,})
		.toPromise()
		.then(function(res)
			{
				let filledResponse:SharepointListItem[] = [];
				$(res.text()).find("z\\:row, row").each(function( index:any ) {
					filledResponse.push(new ctor($(this)[0].attributes)); 
				});
				return filledResponse;
		})
		.catch(this.handleError);
	}
	
	
	
	
	updateListItem(itemToUpdate:SharepointListItem,newValue:string):Promise<void>{
		let currentPayload = this.xmlPayloadWrapperStart+this.updateListItemsPayload+this.xmlPayloadWrapperEnd;
		currentPayload = currentPayload.replace("listNamePayLoad", itemToUpdate.getListName());
		currentPayload = currentPayload.replace("updatesPayLoad", `<Batch OnError="Continue"><Method ID="1" Cmd="Update"><Field Name="ID">`+itemToUpdate.itemId+`</Field><Field Name="`+itemToUpdate.getFieldToUpdate()+`">`+newValue+`</Field></Method></Batch>`);
	return this.http.post(itemToUpdate.getSiteUrl()+this.serviceUrl, currentPayload,{headers:this.headers,})
		.toPromise()
		.then(()=> null)
	.catch(this.handleError);
	}
	
	 private handleError(error: any): Promise<any> {
	
    console.error('An error occurred', error); // for demo purposes only
    return Promise.reject(error.message || error);
  }
}


/*private headersUpdateListItems = new Headers({
								 'Content-Type': 'application/soap+xml; charset=utf-8',
								 'SOAPAction': 'http://schemas.microsoft.com/sharepoint/soap/UpdateListItems',
							 });*/