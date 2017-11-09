import {SharepointListItem} from './sharepoint-list-item';
declare var $:any;

export class UserInfoListEntry extends SharepointListItem{

constructor(rawResponse?:any){
		super(rawResponse);
		if(rawResponse){
			this.title= this.findString('ows_title',rawResponse);
			this.loginName= this.findString('ows_name',rawResponse);
		}
	}

	loginName:string;
	
	///user profile service always run on root web .
	getSiteUrl():string{
		return '/';
	}
	getListName():string{
		return 'User Information List';
	}
	getFieldToUpdate():string{
		return 'Not implemented';
	}

}