import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { ILeaveRequest, LeaveStatus, Region } from '../models/ILeaveRequest';
import { ILeaveBalance } from '../models/ILeaveBalance';
import { getSP } from '../pnpjsConfig';

export interface ISPService {
  getLeaveRequests(region: Region, userEmail: string, isAdmin: boolean): Promise<ILeaveRequest[]>;
  checkAdminRole(userEmail: string): Promise<{ isAdmin: boolean; adminRegions: ('VN' | 'ID_SG' | 'Global')[] }>;
  getUserBalance(userEmail: string, region: Region): Promise<ILeaveBalance | undefined>;
  getAllBalances(region: Region): Promise<ILeaveBalance[]>;
  updateUserBalance(id: number, balanceData: Partial<ILeaveBalance>): Promise<void>;
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

  public async getUserBalance(userEmail: string, region: Region): Promise<ILeaveBalance | undefined> {
    const sp = getSP();
    if (!sp) throw new Error("PnPjs SP instance is not initialized.");

    try {
      const siteWeb = Web([sp.web, this._siteUrl]);
      const listUrl = "/sites/CNGTYTNHHMUABNNQUCTVITNAM/Lists/LeaveBalance";

      const items = await siteWeb.getList(listUrl).items
        .filter(`Title eq '${userEmail}' and Region eq '${region}'`)
        .select("Id", "Title", "EmployeeName", "AnnualLeaveTotal", "AnnualLeaveUsed", "Region")();

      console.log(`Checking balance for ${userEmail} in ${region}... Found:`, items.length);

      if (items && items.length > 0) {
        const item = items[0];
        const total = item.AnnualLeaveTotal || 0;
        const used = item.AnnualLeaveUsed || 0;
        return {
          Id: item.Id,
          UserEmail: item.Title,
          EmployeeName: item.EmployeeName || "N/A",
          AnnualLeaveTotal: total,
          AnnualLeaveUsed: used,
          AnnualLeaveRemaining: total - used,
          Region: item.Region as Region
        };
      }
      return undefined;
    } catch (error) {
      console.error("Critical Error fetching user balance from LeaveBalance list:", error);
      // Giúp user biết lỗi do thiếu cột hay thiếu list
      if (error.message && error.message.indexOf('column') > -1) {
        console.error("Gợi ý: Có vẻ bạn đặt tên cột trong SharePoint chưa khớp với code (AnnualLeaveTotal, AnnualLeaveUsed, EmployeeName, Region).");
      }
      return undefined;
    }
  }

  public async getAllBalances(region: Region): Promise<ILeaveBalance[]> {
    const sp = getSP();
    if (!sp) throw new Error("PnPjs SP instance is not initialized.");

    try {
      const siteWeb = Web([sp.web, this._siteUrl]);
      const listUrl = "/sites/CNGTYTNHHMUABNNQUCTVITNAM/Lists/LeaveBalance";

      const items = await siteWeb.getList(listUrl).items
        .filter(`Region eq '${region}'`)
        .select("Id", "Title", "EmployeeName", "AnnualLeaveTotal", "AnnualLeaveUsed", "Region")();

      return items.map(item => {
        const total = item.AnnualLeaveTotal || 0;
        const used = item.AnnualLeaveUsed || 0;
        return {
          Id: item.Id,
          UserEmail: item.Title,
          EmployeeName: item.EmployeeName || "N/A",
          AnnualLeaveTotal: total,
          AnnualLeaveUsed: used,
          AnnualLeaveRemaining: total - used,
          Region: item.Region as Region
        };
      });
    } catch (error) {
      console.error("Error fetching all balances:", error);
      throw error;
    }
  }

  public async updateUserBalance(id: number, balanceData: Partial<ILeaveBalance>): Promise<void> {
    const sp = getSP();
    if (!sp) throw new Error("PnPjs SP instance is not initialized.");

    try {
      const siteWeb = Web([sp.web, this._siteUrl]);
      const listUrl = "/sites/CNGTYTNHHMUABNNQUCTVITNAM/Lists/LeaveBalance";

      const updatePayload: any = {};
      if (balanceData.AnnualLeaveTotal !== undefined) updatePayload.AnnualLeaveTotal = balanceData.AnnualLeaveTotal;
      if (balanceData.AnnualLeaveUsed !== undefined) updatePayload.AnnualLeaveUsed = balanceData.AnnualLeaveUsed;

      await siteWeb.getList(listUrl).items.getById(id).update(updatePayload);
    } catch (error) {
      console.error("Error updating user balance:", error);
      throw error;
    }
  }
}
