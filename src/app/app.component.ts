import { Component } from '@angular/core';
import { OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';

import {SharepointListsWebService} from './sharepoint-lists-web.service';
import {SharepointUserGroupWebService} from './sharepoint-usergroup-web.service';
import {TaxSpGroup} from './tax-sp-group';
import {GroupEntry} from './group-entry';
import {TaxGroupManagement} from './tax-group-management';
import {TaxSpUser} from './tax-sp-user';
declare var window:any;

@Component({
  selector: 'app-root',
  styles:[`.taxActive {background-color: #D9EDF7 !important;}
			.taxTableRow {cursor:pointer;}
			.taxSearchFound {background-color:yellow;}
			.container h1{text-align:center;}
			.table-striped > tbody > tr:nth-of-type(2n+1) {background-color: #f2f2f2;}
	`],
  templateUrl: './app.component.html',
	providers: [SharepointListsWebService,SharepointUserGroupWebService]
})
export class AppComponent implements OnInit{
  
	constructor(private sharepointListsWebService: SharepointListsWebService, private sharepointUserGroupWebService: SharepointUserGroupWebService){
		this.taxGroupManagement = new TaxGroupManagement(this.sharepointListsWebService, this.sharepointUserGroupWebService);
	}
	
	taxGroupManagement: TaxGroupManagement;
	groups: GroupEntry[];
	selectedGroup:GroupEntry ;
	emailsToErase:string[]= [];
	emailsToAdd:string;
	errorAdding:string;
	errorRemoving:string;
	successMessage:string;
	originalGroups:GroupEntry[]=null;
	searchTerm:string;
	
	
	 ngOnInit(): void {
	 this.sharepointListsWebService.getListItems(GroupEntry, null).then
	 (result => this.groups = <GroupEntry[]>result);
	}
	
	
	add(){
		var reg:RegExp = /^(([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)(\s*;\s*|\s*$))+$/g
		this.errorAdding = '';
		this.successMessage="";
		let finalArray:string[]=[];
		let filteredArray:string[]=[];
		let duplicatedArray:string[]=[];
		if(this.emailsToAdd){
			this.emailsToAdd = this.emailsToAdd.trim();
			if(this.emailsToAdd.length>0){
				if(reg.test(this.emailsToAdd)){
					finalArray = GroupEntry.stringToArray(this.emailsToAdd);
					for(let x = 0;x<finalArray.length;x++){
						if(this.selectedGroup.arrayOfEmails.indexOf(finalArray[x])==-1){
							filteredArray.push(finalArray[x]);
						}
						else{
							duplicatedArray.push(finalArray[x]);
						}
					}
					if(filteredArray.length >0){				
						this.taxGroupManagement.addEmails(this.selectedGroup, filteredArray, true).then((resultArray)=>{
								this.successMessage = resultArray.length - this.selectedGroup.arrayOfEmails.length + " emails were added!!!";
								if(duplicatedArray.length > 0){
									this.successMessage+="\n The following emails were not added because were already part of the group: "+duplicatedArray.toString();
								}
								this.selectedGroup.arrayOfEmails = resultArray;
								this.selectedGroup.emails = resultArray.toString();
								this.emailsToAdd = '';
						}).catch((error:any)=> this.errorAdding = "An error occurred while adding: "+(error.message || error));
					}
					else
					this.errorAdding='The specified email(s) already exist in this group';
				}
				else
				this.errorAdding='Invalid entry. Please enter an email or several emails delimited by a semi-colon.';
			}
			else
				this.errorAdding='Please enter an email or several emails delimited by a semi-colon';
		}
		else
		this.errorAdding='Please enter an email or several emails delimited by a semi-colon';		
	}
	
	removeSelected(){
		this.errorRemoving = '';
		this.successMessage ="";
		this.taxGroupManagement.removeEmails(this.selectedGroup, this.emailsToErase, true).
		then((newArrayOfEmails)=>{
			this.successMessage = "The following email(s) were successfully removed: "+this.emailsToErase.toString(); 
			this.selectedGroup.arrayOfEmails = newArrayOfEmails;
			this.selectedGroup.emails = newArrayOfEmails.toString();
			this.emailsToErase = [];
		})
		.catch((error:any)=> this.errorAdding = "An error occurred while removing: "+(error.message || error));
	}
	
	onChange2(emailEntry2:string){
		if(this.emailsToErase.indexOf(emailEntry2)>=0){
			this.emailsToErase.splice(this.emailsToErase.indexOf(emailEntry2), 1);
		}
		else{
			this.emailsToErase.push(emailEntry2);
		}		
		this.errorAdding = '';
		this.successMessage ='';
		this.errorRemoving = '';
	}
	
	search(){	
		if(this.searchTerm){
			if(this.originalGroups==null){
				this.originalGroups = this.groups;
			}
			this.groups = this.groups.filter(this.filterGroupEntry, this);
		}
		else if (this.originalGroups){
			this.groups = this.originalGroups;
			this.originalGroups = null;
		}
		
	}
	
	filterGroupEntry(entry:GroupEntry):boolean{
		return entry.emails.toLowerCase().indexOf(this.searchTerm.toLowerCase())>=0;
	}
	
	filterWorker(workerEmail:string):boolean{
		return workerEmail.toLowerCase().trim().indexOf(this.searchTerm.toLowerCase())>=0;
	}
	
	onSelect(group:GroupEntry):void{
		this.selectedGroup = group;
		this.emailsToErase = [];
		this.errorAdding = '';
		this.successMessage ='';
		this.errorRemoving = '';
	}
	
	exit():void{
		window.location.href='/';
	}
}
