import {TaxSpGroup} from './tax-sp-group';
export const MOCKGROUPS: TaxSpGroup[] =[
	{requestType:'Data Or Query',description:'Worker For the data Or Query Request'},
	{requestType:'Data Or Query Dispatcher Notification',description:`THis group is for the people that get first the Data or Query or Report Notification. then they will reassign to someone else to work on it.

This list has to be a subset of the Data Or Query or Report entry.`},
	{requestType:'Data, Query or Report',description:'Worker group for the new and enhanced version of Data or Query or Report.'},
	{requestType:'Debug Users',description:'Group of users that are in testing mode and will test/ uat the functionality before it is released'},
	{requestType:'Equipment',description:''},
	{requestType:'Equipment Approver Group',description:'administrator approval for equipment request, software request, move equipment request and return equipment requests and local printer request.'}
];