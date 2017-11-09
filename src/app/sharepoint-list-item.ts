export abstract class SharepointListItem{
itemId:number;
title?:string;	

constructor(rawResponse?:any){
	if(rawResponse){
		this.itemId = +this.findString('ows_id',rawResponse);
		}
}
	
	protected findString(searchTerm:string, myArray:any):string{
		for(let x:number = 0;x<myArray.length;x++)
			{
				if(myArray[x].name ==searchTerm)
				{
					return myArray[x].value;
				}
			}
			return '';
	}
	
	abstract getSiteUrl():string;
	abstract getListName():string;
	abstract getFieldToUpdate():string;

}