import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { ILeaveRequest, LeaveStatus, Region } from '../models/ILeaveRequest';
import { getSP } from '../pnpjsConfig';

export interface ISPService {
  getLeaveRequests(region: Region, userEmail: string, isAdmin: boolean): Promise<ILeaveRequest[]>;
  checkAdminRole(userEmail: string): Promise<{ isAdmin: boolean; adminRegions: ('VN' | 'ID_SG' | 'Global')[] }>;
}

interface ILeaveRequestItem {
  Id: number;
  Title: string;
  StartDate: string;
  EndDate: string;
  SelectedDates: string;
  LeaveStatus: string;
  Reason: string;
  RequesterEmail: { Title: string; EMail: string };
  Approver1: { Title: string };
  Approver2: { Title: string };
  Approver1Comment: string;
  Approver2Comment: string;
  Region: string;
  NumberOfDays: number;
  AttachmentFiles: Array<{ FileName: string; ServerRelativeUrl: string }>;
}

export class SPService implements ISPService {
  private readonly _siteUrl = "https://longangroup.sharepoint.com/sites/CNGTYTNHHMUABNNQUCTVITNAM";

  public async checkAdminRole(userEmail: string): Promise<{ isAdmin: boolean; adminRegions: ('VN' | 'ID_SG' | 'Global')[] }> {
    const sp = getSP();
    if (!sp) throw new Error("PnPjs SP instance is not initialized.");

    try {
      const siteWeb = Web([sp.web, this._siteUrl]);
      const listUrl = "/sites/CNGTYTNHHMUABNNQUCTVITNAM/Lists/LeaveAdminRoles";
      
      const adminItems = await siteWeb.getList(listUrl).items
        .filter(`AdminUser/EMail eq '${userEmail}'`)
        .select("AdminRegion", "AdminUser/EMail", "AdminUser/Id", "AdminUser/Title")
        .expand("AdminUser")();

      if (adminItems && adminItems.length > 0) {
        // Collect all regions the user is admin for (they might have multiple entries or multi-select)
        const regions = adminItems.map(item => item.AdminRegion as 'VN' | 'ID_SG' | 'Global');
        return {
          isAdmin: true,
          adminRegions: regions
        };
      }

      return { isAdmin: false, adminRegions: [] };
    } catch (error) {
      console.warn("Could not fetch LeaveAdminRoles check (list might not exist yet):", error);
      return { isAdmin: false, adminRegions: [] };
    }
  }

  public async getLeaveRequests(region: Region, userEmail: string, isAdmin: boolean): Promise<ILeaveRequest[]> {
    const sp = getSP();
    
    if (!sp) {
      throw new Error("PnPjs SP instance is not initialized.");
    }

    try {
      const siteWeb = Web([sp.web, this._siteUrl]);
      const consolidatedListUrl = "/sites/CNGTYTNHHMUABNNQUCTVITNAM/Lists/LeaveRequestList";

      // Base filter by region
      let filter = `Region eq '${region}'`;

      // Security: If not admin, strictly fetch only their own records
      if (!isAdmin) {
        filter += ` and RequesterEmail/EMail eq '${userEmail}'`;
      }

      const items: ILeaveRequestItem[] = await siteWeb.getList(consolidatedListUrl).items
        .filter(filter)
        .select(
          "Id", 
          "Title", 
          "StartDate", 
          "EndDate", 
          "SelectedDates",
          "LeaveStatus", 
          "Reason", 
          "RequesterEmail/Title",
          "RequesterEmail/EMail",
          "Approver1/Title", 
          "Approver2/Title",
          "Approver1Comment",
          "Approver2Comment",
          "Region",
          "NumberOfDays",
          "AttachmentFiles/FileName",
          "AttachmentFiles/ServerRelativeUrl"
        )
        .expand("Approver1", "Approver2", "RequesterEmail", "AttachmentFiles")
        .orderBy("Id", false)();

      return items.map(item => ({
        Id: item.Id,
        RequesterName: item.RequesterEmail?.Title || item.Title || "N/A", 
        RequesterEmail: item.RequesterEmail?.EMail || "", 
        StartDate: item.StartDate || "",
        EndDate: item.EndDate || "",
        SelectedDates: item.SelectedDates || undefined,
        Reason: item.Reason || "",
        Status: (item.LeaveStatus as LeaveStatus) || "Awaiting approval",
        Approver1Name: item.Approver1?.Title || "N/A",
        Approver1Comment: item.Approver1Comment || "",
        Approver2Name: item.Approver2?.Title || "N/A",
        Approver2Comment: item.Approver2Comment || "",
        Region: item.Region as Region,
        NumberOfDays: item.NumberOfDays || 0,
        Attachments: item.AttachmentFiles || []
      }));
    } catch (error) {
      console.error(`Error fetching items from consolidated LeaveRequestList:`, error);
      throw error;
    }
  }
}
