import { Component } from '@angular/core';
import { OnInit } from '@angular/core';
import { ViewChild } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { Observable } from 'rxjs/Observable';
import { Subject } from 'rxjs/Subject';


import {SharepointListsWebService} from 'ng-tax-share-point-web-services-module';
import {SharepointUserGroupWebService} from 'ng-tax-share-point-web-services-module';
import {TaxSpUser} from 'ng-tax-share-point-web-services-module';
import {UserInfoListEntry} from 'ng-tax-share-point-web-services-module';

import {GroupEntry} from './group-entry';
import {TaxGroupManagement} from './tax-group-management';

import { MatAutocompleteTrigger } from '@angular/material';

import 'rxjs/add/operator/debounceTime';
import 'rxjs/add/operator/distinctUntilChanged';
import 'rxjs/add/operator/switchMap';

declare var window:any;

@Component({
  selector: 'app-root',
  styles:[`.taxActive {background-color: #D9EDF7 !important;}
			.taxTableRow {cursor:pointer;}
			.taxSearchFound {background-color:yellow;}
			.container h1{text-align:center;}
			.table-striped > tbody > tr:nth-of-type(2n+1) {background-color: #f2f2f2;}
			.taxUnresolvedPicker{
		font-style:italic;
		color: #a94442;
	}
	
	.taxResolvedPicker{	
		text-decoration:underline;
	}
	
	.mat-option{
		line-height:inherit;
		border-bottom-color:#DDD;
		border-bottom-style:solid;
		border-bottom-width:1px;
	}
	
	.taxPeoplePickerContainer{
		display:inline-block
		width:100%;
	}
	
	.taxPeoplePickerResults{
		width:100%;
	}
	
	.taxPeoplePickerText{
			line-height:normal;
			padding-top:0px;
			padding-bottom:0px;
			height:25px;
	}
	.taxMainResultPicker{
		font-size:14px;
		font-family: "Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif;
		font-weight:400;
		color:#333;
		padding-top:4px;
	}
	.taxSecResultPicker{
		font-size:12px;
		color:#666;
		font-family: "Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif;
		font-weight:400;
		margin-top:2px;	
	}
	
	input.ng-invalid  {
			border-left: 5px solid #a94442; /* red */
		}
		
	.mattooltiptax{
		font-size:14px;
		font-family: "Segoe UI Regular WestEuropean","Segoe UI",Tahoma,Arial,sans-serif;
		font-weight:600;
		color:#eee;
		padding-top:4px;
	}
	input[type=text]::-ms-clear {
		display: none;
	}
	`],
  templateUrl: './app.component.html',
	providers: [SharepointListsWebService,SharepointUserGroupWebService]
})
export class AppComponent implements OnInit{
  
	
	
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
	processingAdding=false;
	processingDeleting = false;
	
//begin of picker related props
	//access to the input that triggers the autocomplete.
	@ViewChild('term', { read: MatAutocompleteTrigger }) 
	autoCompleteInput: MatAutocompleteTrigger;
	items: Observable<UserInfoListEntry[]>;
	currentItems : UserInfoListEntry[] = [];
	private searchTermStream = new Subject<string>();
	//empTitle:string = "";
	isResolved:boolean = false;
	entryNotValid:boolean= false;
	selectedEmp: UserInfoListEntry = null;
	numberOfActiveRequests:number=0;
	userHitEnter:boolean = false;
	flag=false;
	pickerTooltip = "Please start typing a name";
	hideNoResultsFound = true;  /// this needs be changed for errorAdding.
	//end o picker related props.
	
	
	
	///pickerrelated methods begin
	private cleanPicker(){
		this.pickerTooltip = "Please start typing a name";
		this.isResolved = false;
		this.userHitEnter = false;
		this.currentItems = [];
		this.numberOfActiveRequests = 0;
		this.selectedEmp = null;
		this.entryNotValid = false;
	}
	
	private resolvePicker(employee: UserInfoListEntry){
		this.pickerTooltip = employee['email'] +' - ' +employee['name'];
		this.isResolved =  true;
		this.selectedEmp = employee;
		this.entryNotValid = false;
		//this.group.get('insideTextbox').patchValue(employee.ID, {emitEvent:false}); not needed? 
	}
	
	private toTitleCase(term: string) {
		return term.replace(/\w\S*/g, (term) => { return term.charAt(0).toUpperCase() + term.substr(1).toLowerCase(); });
	}
	
	detectKeyDown(event:KeyboardEvent){
	//if(this.isResolved&& this.selectedEmp && this.empTitle == this.selectedEmp.title){
		//if(event && event.keyCode ==8&& this.selectedEmp && this.empTitle == this.selectedEmp.title){
		//if(event && event.keyCode ==8&& this.isResolved){
		
		let safeKeys :number[]= [9,13,35,36,37,38,39,40] ;
		if(event && (this.isResolved || this.entryNotValid)&& safeKeys.filter(x=> x==event.keyCode).length==0){
			//so that i can put here an if isresolved then cleanup the picker.
			this.cleanPicker();
			this.emailsToAdd = '';
			this.flag = true;	// question this line of code? ? ?? ??
			// cleaning the search results.
			this.search("", null);
		}
	}
	
	///Sets employee via the autocomplete.
	/// Either by hitting enter (keyboard) or by clicking on an option.
	setEmployee(emp: UserInfoListEntry, event:any) {
		//this.selectedEmp = emp;
		//commenting the below line for reactive form.
		//this.empTitle= '';
		this.search("", null);
		this.resolvePicker(emp);
		//this.isResolved = true;
	}
	
	
	blurEvent(term:string){
		if(!this.isResolved &&this.emailsToAdd){
			this.userHitEnter= true;
			//we are tricking the distinctUntilChanged with the addition of a space.
			//need better code.
			this.search(term+" ",  null);
		}
	}
	
	search(term: string, event:KeyboardEvent) {	
		if(this.entryNotValid ||(event && event.keyCode > 34 && event.keyCode<41)){
			//disregard arrow keys: 37, 38, 39, 40.
			//disregard end and home: 35, 36.
			return ; 
		}
	
		if(this.flag){
			//prevent coming twice to the same method.
			this.flag = false;
			return;
		}
		if(this.userHitEnter && this.currentItems){
			let filteredResults:UserInfoListEntry[];
			filteredResults = this.currentItems.filter(x =>{ 
			var y = <any> x;
			var z = this.emailsToAdd.toUpperCase().trim();
			return y.title.toUpperCase() ==z || y.email.toUpperCase() == z
			|| y.name.toUpperCase() == z || y.name.toUpperCase() == ("ID\\"+z)
			});
			if(filteredResults && filteredResults.length ==1){
				//this.resolvePicker(filteredResults[0]);
				this.userHitEnter = false;
				this.flag = true;
				this.emailsToAdd = filteredResults[0].title;
				this.resolvePicker(filteredResults[0]);
				this.autoCompleteInput.closePanel();
				this.searchTermStream.next("");			
				return;
			}
		}
		this.hideNoResultsFound = true;//it was on the on key down at the beginning.
		this.searchTermStream.next(this.toTitleCase(term));
	}
	
	
	constructor(private sharepointListsWebService: SharepointListsWebService, private sharepointUserGroupWebService: SharepointUserGroupWebService){
		
		this.taxGroupManagement = new TaxGroupManagement(this.sharepointListsWebService, this.sharepointUserGroupWebService);
	 this.items = <Observable<UserInfoListEntry[]>>this.searchTermStream
		  .debounceTime(300)
		  .distinctUntilChanged()
		  .switchMap((term: string) => {
			this.numberOfActiveRequests+=1;
			return this.taxGroupManagement.searchForPeople(term).then(x=>{
				if(x){
					if(this.hideNoResultsFound&& term.length > 1){
						this.hideNoResultsFound = false;
					}
					let tempResults:UserInfoListEntry[]=<UserInfoListEntry[]> x;
					let filteredResults:UserInfoListEntry[];
					this.currentItems = tempResults;
					if(this.userHitEnter){
						filteredResults = tempResults.filter(x => {
							var y = <any> x;
							var z = this.emailsToAdd.toUpperCase().trim();
							return y.title.toUpperCase() ==z || y.email.toUpperCase() == z
							|| y.name.toUpperCase() == z || y.name.toUpperCase() == ("ID\\"+z)
						});
						this.userHitEnter = false;
						if(filteredResults && filteredResults.length ==1){
							//this.resolvePicker(filteredResults[0]);
							this.autoCompleteInput.closePanel();
							this.emailsToAdd = filteredResults[0].title;
							this.resolvePicker(filteredResults[0]);
							tempResults = [];
						}else if(!this.isResolved){
							this.entryNotValid = true;
						}
					}
					this.numberOfActiveRequests-=1;
					
					return tempResults;
				}
				else if(this.userHitEnter&& !this.isResolved){
					this.entryNotValid = true;
				}
				this.numberOfActiveRequests-=1;
				
				return x;
				});
		  });
	 }
	
	///picker related methods end.
	
	
	
	
	 ngOnInit(): void {
	 this.sharepointListsWebService.getListItems(GroupEntry, null, null, null).then
	 (result => this.groups = <GroupEntry[]>result);
	}
	
	
	add(){
		this.cleanErrorsAndSuccess();
		if(this.selectedGroup.arrayOfEmails.map((z)=>z.toLowerCase()).indexOf(this.selectedEmp["email"].toLowerCase())==-1){
			this.processingAdding = true;
			
			this.taxGroupManagement.addEmails(this.selectedGroup, [this.selectedEmp["email"]],false).then((resultArray)=>{
					this.successMessage =this.selectedEmp["email"]+ " was added!!!";
					this.selectedGroup.arrayOfEmails = resultArray;
					this.selectedGroup.emails = resultArray.toString();
					this.emailsToAdd = '';
					this.cleanPicker();
					this.processingAdding = false;
			}).catch((error:any)=> {
				this.errorAdding = "An error occurred while adding: "+(error.message || error);
				this.processingAdding = false;
			});
		}
		else{
			this.errorAdding='The specified person already exist in this group';
		}
	}
	
	removeSelected(){
		this.cleanErrorsAndSuccess();
		this.processingDeleting = true;
		this.taxGroupManagement.removeEmails(this.selectedGroup, this.emailsToErase, true).
		then((newArrayOfEmails)=>{
			this.successMessage = "The following email(s) were successfully removed: "+this.emailsToErase.toString(); 
			this.selectedGroup.arrayOfEmails = newArrayOfEmails;
			this.selectedGroup.emails = newArrayOfEmails.toString();
			this.emailsToErase = [];
			this.processingDeleting = false;
		})
	.catch((error:any)=> {
		this.errorAdding = "An error occurred while removing: "+(error.message || error);
		this.processingDeleting = false;
	});
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
	
	search2(){	
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
		this.cleanErrorsAndSuccess();
		this.emailsToAdd = "";
		this.cleanPicker();
	}
	
	cleanErrorsAndSuccess(){
		this.errorAdding = '';
		this.successMessage ='';
		this.errorRemoving = '';
	}
	
	exit():void{
		window.location.href='/';
	}
}
