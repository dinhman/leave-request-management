import { Region } from './ILeaveRequest';

export interface ILeaveBalance {
  Id: number;
  UserEmail: string;
  EmployeeName: string;
  AnnualLeaveTotal: number;
  AnnualLeaveUsed: number;
  AnnualLeaveRemaining: number;
  Region: Region;
}
