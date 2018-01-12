import {SharepointListsWebService} from 'ng-tax-share-point-web-services-module';
import {SharepointUserGroupWebService} from 'ng-tax-share-point-web-services-module';
import {TaxSpUser} from 'ng-tax-share-point-web-services-module';
import {GroupEntry} from './group-entry';
import {SharepointListItem} from 'ng-tax-share-point-web-services-module';
import {UserInfoListEntry} from 'ng-tax-share-point-web-services-module';
///This is meant to be a business class that handles all the integration between sharepoint-list-item
/// and sharepoint groups.

export class TaxGroupManagement{

	constructor(private sharepointListsWebService: SharepointListsWebService,
				private sharepointUserGroupWebService: SharepointUserGroupWebService){}

				
	//Will search for an item that contains the provided term in either the display name ,the 
	// email or the Account Number. It is case insensitive search!!! What else could you ask for!! ? ?
	searchForPeople(term:string): Promise<SharepointListItem[]>{
		if(term){
			let trimmedTerm = term.trim();
			if(trimmedTerm.length>1){			
				if(trimmedTerm){
					return this.sharepointListsWebService.getListItems(UserInfoListEntry, null ,
						"<Query><OrderBy><FieldRef Name ='Title'/></OrderBy><Where><And><Eq><FieldRef Name='ContentType' />"+
						"<Value Type='Text'>Person</Value></Eq><Or><Or><Contains><FieldRef Name='EMail' /><Value Type='Text'>"+
						trimmedTerm+"</Value></Contains><Contains><FieldRef Name='Title' /><Value Type='Text'>"+
						trimmedTerm+"</Value></Contains></Or><Contains><FieldRef Name='Name' /><Value Type='Text'>"+
						trimmedTerm+"</Value></Contains></Or></And></Where></Query>", null);
				}
			}
		}
		
		//if any of the condition were false, return empty array.
		return Promise.resolve([]);
	}
				
				
				
	//tHIS ONE LOOKS UP USERS IN THE USER INFO TABLE INSTEAD OF THE USERGROUP.ASMX SERVICE
	getLoginsFromUserProfile(emails:string[]):Promise<TaxSpUser[]>{
		let promises: Promise<SharepointListItem[]>[] = [];
		for(let email of emails){
				promises.push(this.sharepointListsWebService.getListItems(UserInfoListEntry,["EMail", email], null, null));
		}
		return Promise.all(promises).then(
			function(internalRes){
				let users:TaxSpUser[]=[];
				let tempCast:UserInfoListEntry[];
				for(let i of internalRes){
					if(i){
						tempCast = <UserInfoListEntry[]> i;
						if(tempCast.length>0){
							users.push({id:0,login:tempCast[0]["name"],email:"",displayName:""});
						}
					}
				}
				return Promise.resolve(users);
			}
		)
	}
	
	
	
	///Business method that will try to remove the specified emails.
	///IT first removes them from the list Item.
	///Then , if spGroupName is present, it tries to remove the members from the group.
	///  For that second action, it looks up  the login id's, deending on:
	///  - if it's test, it'll look into the user information list.
	///  - if it's not test, it'll look into the getUserLoginFromEmail from the usergroup web service.
	removeEmails(group:GroupEntry, emailsToErase:string[],isTest:boolean):Promise<string[]>{
		let promises: Promise<any>[] = [];
		let promiseOfSpTaxUser:Promise<TaxSpUser[]>;
		let finalArray:string[]= group.arrayOfEmails.slice();
		let localGrpName:string;
		let localSiteUrl:string;
		for(let x = 0 ;x<emailsToErase.length;x++){
			finalArray.splice(finalArray.indexOf(emailsToErase[x]),1);
		}
		promises.push(this.sharepointListsWebService.updateListItem(group,finalArray.toString()));
		///If need to update the associated SharePoint Group...
		if(group.spGroupName){
			localGrpName = group.spGroupName;
			localSiteUrl = group.getSiteUrl();
			if(isTest){
				promiseOfSpTaxUser = this.getLoginsFromUserProfile(emailsToErase);
			}else{
				promiseOfSpTaxUser = this.sharepointUserGroupWebService.getUserLoginFromEmail(emailsToErase,localSiteUrl);
			}
			promises.push(promiseOfSpTaxUser.
			then((spTaxUsers)=>this.sharepointUserGroupWebService
			.removeUserCollectionFromGroup(spTaxUsers,localGrpName,localSiteUrl)));
		}
		return Promise.all(promises).then(()=> finalArray);	
	}
	
	///Business method that will try to add the specified emails.
	///IT first adds them from the list Item.
	///Then , if spGroupName is present, it tries to add the members from the group.
	///  For that second action, it looks up  the login id's, depending on:
	///  - if it's test, it'll look into the user information list.
	///  - if it's not test, it'll look into the getUserLoginFromEmail from the usergroup web service.
	addEmails(group:GroupEntry, emailsToAdd:string[],isTest:boolean):Promise<string[]>{
		let promises: Promise<any>[] = [];
		let promiseOfSpTaxUser:Promise<TaxSpUser[]>;
		let finalArray:string[];
		let localGrpName:string;
		let localSiteUrl:string;
		finalArray=group.arrayOfEmails.concat(emailsToAdd);
		finalArray.sort(GroupEntry.sortInsensitive);
		promises.push(this.sharepointListsWebService.updateListItem(group,finalArray.toString()));
		///If need to update the associated SharePoint Group...
		if(group.spGroupName){
			localGrpName = group.spGroupName;
			localSiteUrl = group.getSiteUrl();
			if(isTest){
				promiseOfSpTaxUser = this.getLoginsFromUserProfile(emailsToAdd);
			}else{
				promiseOfSpTaxUser = this.sharepointUserGroupWebService.getUserLoginFromEmail(emailsToAdd,localSiteUrl);
			}
			promises.push(promiseOfSpTaxUser.
			then((spTaxUsers)=>this.sharepointUserGroupWebService
			.addUserCollectionToGroup(spTaxUsers,localGrpName,localSiteUrl)));
		}
		return Promise.all(promises).then(()=> finalArray);	
	}

}