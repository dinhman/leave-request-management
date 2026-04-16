export type LeaveStatus = 'Awaiting approval' | 'Approved' | 'Rejected' | 'Draft';
export type Region = 'VN' | 'ID_SG';

export interface ILeaveRequest {
  Id: number;
  RequesterName: string;
  RequesterEmail: string;
  StartDate: string;
  EndDate: string;
  SelectedDates?: string;
  Reason: string;
  Status: LeaveStatus;
  
  // 2nd Level Approval Fields
  Approver1Name?: string;
  Approver1Email?: string;
  Approver1Comment?: string;
  
  Approver2Name?: string;
  Approver2Email?: string;
  Approver2Comment?: string;
  
  Region: Region;
  NumberOfDays: number;
  Attachments?: Array<{
    FileName: string;
    ServerRelativeUrl: string;
  }>;
}
