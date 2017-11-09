import {SharepointListItem} from './sharepoint-list-item';
declare var $:any;


export class GroupEntry extends SharepointListItem{
	
	constructor(rawResponse?:any){
		super(rawResponse);
		if(rawResponse){
			
			
			this.title= this.findString('ows_requesttype',rawResponse);
			this.description = this.findString('ows_group_x0020_description',rawResponse);
			this.emails = this.findString('ows_emails',rawResponse);
			this.arrayOfEmails=GroupEntry.stringToArray(this.emails);
			this.spGroupName = this.findString('ows_sharepointgroupname', rawResponse);
		}
	}
	
	
	
	getSiteUrl():string{
		return '/forms/isdportal/';
	}
	getListName():string{
		return 'Workers';
	}
	
	getFieldToUpdate():string{
		return 'Emails';
	}
	
	static stringToArray(stringOfEmails:string):string[]{
		let tempArray:string[]=[];
		let tempArray2:string[];
		let tempArray3:string[];
		if(stringOfEmails){
			stringOfEmails = stringOfEmails.trim();
			if(stringOfEmails.length>0){
				tempArray3 = stringOfEmails.split(';');
				tempArray3.sort(GroupEntry.sortInsensitive);
				tempArray3.forEach(function(item:string, index:number){
					item = item.replace(/(\r\n|\n|\r)/gm,"").trim();
					if(item.indexOf(",")>=0){
						tempArray2 = item.split(',');
						tempArray2.forEach(function(item2:string, index2:number){
							item2= item2.replace(/(\r\n|\n|\r)/gm,"").trim();
						});
						tempArray=tempArray.concat(tempArray2);
					}
					else
					{
						tempArray.push(item);
					}
				});
				tempArray.sort(GroupEntry.sortInsensitive);
				tempArray = tempArray.filter(a=> a.length>0);
			}
		}
		return tempArray;
	}
	
	static sortInsensitive(a:string, b:string) {
		return a.toLowerCase().localeCompare(b.toLowerCase());
	}
	
	

	emails:string;
	description:string;
	arrayOfEmails:string[];
	spGroupName:string;
	
	private _emailsString:string;
	get emailsString():string{
	return this._emailsString;
	}
	set emailsString(newString:string){
		this._emailsString = newString;
	}	
}
