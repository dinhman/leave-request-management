import * as React from 'react';
import { 
  Stack, 
  IStackTokens, 
  Spinner, 
  SpinnerSize, 
  MessageBar, 
  MessageBarType,
  getTheme,
  mergeStyles,
  Text,
  Icon,
  ActionButton,
  AnimationClassNames,
  SearchBox,
  Dropdown,
  IDropdownOption,
  Pivot,
  PivotItem,
  DefaultButton
} from '@fluentui/react';
import { LeaveRequestList } from './subcomponents/LeaveRequestList';
import { LeaveRequestDetail } from './subcomponents/LeaveRequestDetail';
import { RegionSelector } from './subcomponents/RegionSelector';
import { LeaveReports } from './subcomponents/LeaveReports';
import { LeaveRequestForm } from './subcomponents/LeaveRequestForm';
import { ILeaveRequest, Region } from '../../../models/ILeaveRequest';
import { ISPService, SPService } from '../../../services/SPService';
import { getSP } from '../../../pnpjsConfig';

import { ILeaveRequestDashboardProps } from './ILeaveRequestDashboardProps';

const theme = getTheme();
const stackTokens: IStackTokens = { childrenGap: 25 };

const dashboardContainerClass = mergeStyles(
  {
    backgroundColor: '#f3f2f1', 
    minHeight: '600px',
    display: 'flex',
    flexDirection: 'column',
  },
  AnimationClassNames.fadeIn400
);

const contentCardClass = mergeStyles({
  margin: '0 0 20px 0',
  padding: '25px',
  backgroundColor: theme.palette.white,
  borderRadius: '0 0 12px 12px',
  boxShadow: '0 4px 12px rgba(0, 0, 0, 0.05)',
  flexGrow: 1,
});

const kpiCardClass = mergeStyles({
  padding: '12px 18px',
  backgroundColor: theme.palette.white,
  borderRadius: '10px',
  boxShadow: '0 2px 6px rgba(0, 0, 0, 0.04)',
  flex: '1 1 150px',
  borderLeft: '4px solid transparent',
  display: 'flex',
  flexDirection: 'column',
  gap: '4px',
});

const filterContainerClass = mergeStyles({
  padding: '15px 25px',
  backgroundColor: theme.palette.white,
  borderBottom: `1px solid ${theme.palette.neutralLighter}`,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: '20px',
  flexWrap: 'wrap',
});

const headerContainerClass = mergeStyles({
  height: '70px',
  padding: '0 30px',
  backgroundColor: theme.palette.white,
  borderBottom: `1px solid ${theme.palette.neutralLighter}`,
  boxShadow: '0 2px 4px rgba(0, 0, 0, 0.05)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  position: 'sticky',
  top: 0,
  zIndex: 100,
});

const logoSectionClass = mergeStyles({
  display: 'flex',
  alignItems: 'center',
  gap: '15px',
});

const regionBadgeClass = mergeStyles({
  padding: '4px 12px',
  borderRadius: '16px',
  backgroundColor: theme.palette.themeLighterAlt,
  color: theme.palette.themePrimary,
  fontWeight: 600,
  fontSize: '12px',
  textTransform: 'uppercase',
  border: `1px solid ${theme.palette.themeLighter}`,
});

const FLOW_CONFIG = {
  VN: "https://default54b7b7f48d8f40b0918ecd5dc260d2.2b.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/cfb98a65a5b644d197c1bae02964edb0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=M3AyH0VhnLeoTKEQu0gKftCakvvrKq8gKYIbz_Sb8cI",
  ID_SG: "https://default54b7b7f48d8f40b0918ecd5dc260d2.2b.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/3117c3e07e3c4f93a525a6c6e855e929/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=uJm66-sq3pAuG0ygcY7Dr5C4ovBFvKZ5IKOp1E3e9kA"
};

const KpiCard: React.FC<{ title: string; count: number; icon: string; color: string }> = ({ title, count, icon, color }) => (
  <div className={kpiCardClass} style={{ borderLeftColor: color }}>
    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
      <Text variant="small" style={{ color: theme.palette.neutralSecondary, fontWeight: 600, textTransform: 'uppercase', fontSize: '10px' }}>
        {title}
      </Text>
      <Icon iconName={icon} style={{ color: color, fontSize: '14px' }} />
    </Stack>
    <Text variant="xxLarge" style={{ fontWeight: 700, color: theme.palette.neutralPrimary }}>
      {count}
    </Text>
  </div>
);

const AppHeader: React.FC<{ 
  currentRegion?: Region; 
  onResetRegion: () => void;
}> = (props) => {
  const { currentRegion, onResetRegion } = props;
  const logoUrl = "https://static.mycareersfuture.gov.sg/images/company/logos/057e337e2f79861429ef03e4afbb4971/longan-group.png";

  return (
    <div className={headerContainerClass}>
      <div className={logoSectionClass}>
        <div style={{ padding: '4px', borderRadius: '8px' }}>
           <img src={logoUrl} alt="Longan Group Logo" style={{ height: '45px', objectFit: 'contain' }} />
        </div>
        <Stack>
          <Text variant="large" style={{ fontWeight: 700, color: theme.palette.neutralPrimary }}>
            Longan Group
          </Text>
          <Text variant="small" style={{ color: theme.palette.neutralSecondary, marginTop: '-4px' }}>
            Leave Management System
          </Text>
        </Stack>
      </div>

      {currentRegion && (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }}>
          <div className={regionBadgeClass}>
            <Icon iconName="World" style={{ marginRight: '6px', verticalAlign: 'middle' }} />
            {currentRegion === 'VN' ? 'Vietnam' : 'ID & SG'}
          </div>
          <ActionButton 
            iconProps={{ iconName: 'Sync' }} 
            onClick={onResetRegion}
            styles={{ root: { color: theme.palette.neutralSecondary } }}
          >
            Switch Region
          </ActionButton>
        </Stack>
      )}
    </div>
  );
};

// ADMIN_EMAILS hardcode has been removed in favor of dynamic SharePoint List check.

const LeaveRequestDashboardFunctional: React.FC<ILeaveRequestDashboardProps> = (_props) => {
  const [items, setItems] = React.useState<ILeaveRequest[]>([]);
  const [filteredItems, setFilteredItems] = React.useState<ILeaveRequest[]>([]);
  const [selectedItem, setSelectedItem] = React.useState<ILeaveRequest | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [isDetailOpen, setIsDetailOpen] = React.useState<boolean>(false);
  const [selectedRegion, setSelectedRegion] = React.useState<Region | undefined>(undefined);
  
  // Filtering states
  const [searchText, setSearchText] = React.useState<string>("");
  const [statusFilter, setStatusFilter] = React.useState<string>("All");
  const [currentPage, setCurrentPage] = React.useState<number>(1);
  const pageSize = 10;
  
  // Admin Authorization State
  const [isAdmin, setIsAdmin] = React.useState<boolean>(false);
  const [adminRegions, setAdminRegions] = React.useState<('VN' | 'ID_SG' | 'Global')[]>([]);

  const currentUserEmail = _props.context.pageContext.user.email.toLowerCase();

  // Initialize PnPjs with context
  getSP(_props.context);
  const spService: ISPService = new SPService();

  // Check Admin Role on Mount
  React.useEffect(() => {
    spService.checkAdminRole(currentUserEmail).then((roleConfig) => {
      setIsAdmin(roleConfig.isAdmin);
      setAdminRegions(roleConfig.adminRegions);
    }).catch(console.error);
  }, [currentUserEmail]);

  // Determine if the current user has admin rights for the currently selected region
  const hasAccessToAdminDashboard = React.useMemo(() => {
    if (!isAdmin || !selectedRegion) return false;
    return adminRegions.includes('Global') || adminRegions.includes(selectedRegion);
  }, [isAdmin, adminRegions, selectedRegion]);

  const fetchData = async (region: Region): Promise<void> => {
    try {
      setLoading(true);
      const data = await spService.getLeaveRequests(region, currentUserEmail, isAdmin);
      setItems(data || []);
      setFilteredItems(data || []);
      setError(undefined);
    } catch (err) {
      console.error('Error fetching data:', err);
      setError(`Failed to load leave requests from the ${region} site.`);
    } finally {
      setLoading(false);
    }
  };

  // Filter effect
  React.useEffect(() => {
    let filtered = items;
    
    if (searchText) {
      const lowerSearch = searchText.toLowerCase();
      filtered = filtered.filter(i => 
        i.RequesterName.toLowerCase().indexOf(lowerSearch) > -1 || 
        (i.RequesterEmail && i.RequesterEmail.toLowerCase().indexOf(lowerSearch) > -1) ||
        (i.Reason && i.Reason.toLowerCase().indexOf(lowerSearch) > -1) ||
        (i.Approver1Name && i.Approver1Name.toLowerCase().indexOf(lowerSearch) > -1)
      );
    }
    
    if (statusFilter !== "All") {
      filtered = filtered.filter(i => i.Status === statusFilter);
    }
    
    setFilteredItems(filtered);
    setCurrentPage(1); // Reset to first page on filter change
  }, [items, searchText, statusFilter]);

  const onSelectRegion = (region: Region): void => {
    setSelectedRegion(region);
    fetchData(region).catch(() => {});
  };

  const onSelectItem = (item: ILeaveRequest): void => {
    setSelectedItem(item);
    setIsDetailOpen(true);
  };

  const onDismissDetail = (): void => {
    setIsDetailOpen(false);
  };

  const onDismissedDetail = (): void => {
    setSelectedItem(undefined);
  };

  const resetRegion = (): void => {
    setSelectedRegion(undefined);
    setItems([]);
    setFilteredItems([]);
  };

  const getKpiData = (): { total: number; pending: number; approved: number; rejected: number } => {
    return {
      total: items.length,
      pending: items.filter(i => i.Status === 'Awaiting approval').length,
      approved: items.filter(i => i.Status === 'Approved').length,
      rejected: items.filter(i => i.Status === 'Rejected').length,
    };
  };

  const statusOptions: IDropdownOption[] = [
    { key: 'All', text: 'All Statuses' },
    { key: 'Awaiting approval', text: 'Awaiting approval' },
    { key: 'Approved', text: 'Approved' },
    { key: 'Rejected', text: 'Rejected' },
    { key: 'Draft', text: 'Draft' },
  ];

  if (!selectedRegion) {
    return (
      <div className={dashboardContainerClass} style={{ backgroundColor: theme.palette.white }}>
        <AppHeader onResetRegion={resetRegion} />
        <div style={{ marginTop: '40px' }}>
          <RegionSelector onSelectRegion={onSelectRegion} />
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div className={dashboardContainerClass}>
        <AppHeader currentRegion={selectedRegion} onResetRegion={resetRegion} />
        <Stack verticalAlign="center" horizontalAlign="center" style={{ flexGrow: 1 }}>
          <Spinner size={SpinnerSize.large} label={`Syncing with ${selectedRegion} SharePoint...`} />
        </Stack>
      </div>
    );
  }

  const kpis = getKpiData();

  return (
    <div className={dashboardContainerClass} style={{ width: '100%' }}>
      <AppHeader currentRegion={selectedRegion} onResetRegion={resetRegion} />
      
      <Pivot 
        aria-label="Dashboard Views" 
        style={{ padding: '0 10px', marginTop: '10px' }}
        styles={{ 
          root: { 
            backgroundColor: theme.palette.white, 
            borderBottom: `1px solid ${theme.palette.neutralLighter}` 
          },
          linkIsSelected: {
            selectors: {
              '::before': {
                height: '3px',
                backgroundColor: theme.palette.themePrimary
              }
            }
          }
        }}
      >
        <PivotItem headerText="Submit Request" itemIcon="Send">
          <div style={{ marginTop: '25px', paddingBottom: '30px' }}>
            <LeaveRequestForm 
              userDisplayName={_props.userDisplayName}
              userEmail={_props.context.pageContext.user.email}
              automateUrl={selectedRegion === 'VN' ? FLOW_CONFIG.VN : FLOW_CONFIG.ID_SG}
              httpClient={_props.context.httpClient}
              region={selectedRegion}
              onSuccess={() => fetchData(selectedRegion)}
            />
          </div>
        </PivotItem>

        <PivotItem headerText="My Requests" itemIcon="History">
          <div style={{ marginTop: '20px' }}>
            <div className={contentCardClass} style={{ margin: '0 10px 20px 10px' }}>
              <Stack tokens={stackTokens}>
                <LeaveRequestList 
                  items={items.filter(i => i.RequesterEmail.toLowerCase() === _props.context.pageContext.user.email.toLowerCase())} 
                  onSelectItem={onSelectItem} 
                />
              </Stack>
            </div>
          </div>
        </PivotItem>

        {hasAccessToAdminDashboard && (
          <PivotItem headerText="Requests Dashboard" itemIcon="AppIconDefaultList">
            <div style={{ marginTop: '15px' }}>
              {/* KPI Section */}
              <Stack horizontal tokens={{ childrenGap: 12 }} style={{ padding: '0px 10px 15px 10px' }}>
                <KpiCard title="Total Requests" count={kpis.total} icon="AllApps" color={theme.palette.themePrimary} />
                <KpiCard title="Awaiting approval" count={kpis.pending} icon="Clock" color={theme.palette.yellow} />
                <KpiCard title="Approved" count={kpis.approved} icon="Completed" color={theme.palette.green} />
                <KpiCard title="Rejected" count={kpis.rejected} icon="ErrorBadge" color={theme.palette.red} />
              </Stack>

              {/* Filter Section */}
              <div className={filterContainerClass} style={{ margin: '0 10px' }}>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }} style={{ flexGrow: 1 }}>
                  <SearchBox 
                    placeholder="Search requester, reason, or approver..." 
                    onSearch={setSearchText}
                    onChange={(_, newValue) => setSearchText(newValue || "")}
                    styles={{ root: { width: 400 } }}
                  />
                  <Dropdown 
                    placeholder="Filter by Status"
                    options={statusOptions}
                    selectedKey={statusFilter}
                    onChange={(_, option) => setStatusFilter(option?.key as string || "All")}
                    styles={{ root: { width: 200 } }}
                  />
                </Stack>
                <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                  Showing {filteredItems.length} of {items.length} requests
                </Text>
              </div>

              <div className={contentCardClass} style={{ margin: '0 10px 20px 10px' }}>
                <Stack tokens={stackTokens}>
                  {error && (
                    <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
                      {error}
                    </MessageBar>
                  )}
                  <LeaveRequestList items={filteredItems.slice((currentPage - 1) * pageSize, currentPage * pageSize)} onSelectItem={onSelectItem} />
                  
                  {/* Pagination Controls */}
                  {filteredItems.length > pageSize && (
                    <Stack 
                      horizontal 
                      horizontalAlign="center" 
                      verticalAlign="center" 
                      tokens={{ childrenGap: 20 }}
                      style={{ marginTop: '20px', padding: '10px' }}
                    >
                      <DefaultButton
                        text="Previous"
                        iconProps={{ iconName: 'ChevronLeft' }}
                        onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                        disabled={currentPage === 1}
                      />
                      <Text variant="medium" style={{ fontWeight: 600 }}>
                        Page {currentPage} of {Math.ceil(filteredItems.length / pageSize)}
                      </Text>
                      <DefaultButton
                        text="Next"
                        iconProps={{ iconName: 'ChevronRight' }}
                        onClick={() => setCurrentPage(prev => Math.min(Math.ceil(filteredItems.length / pageSize), prev + 1))}
                        disabled={currentPage === Math.ceil(filteredItems.length / pageSize)}
                      />
                    </Stack>
                  )}
                </Stack>
              </div>
            </div>
          </PivotItem>
        )}

        {hasAccessToAdminDashboard && (
          <PivotItem headerText="Reports" itemIcon="BIDashboard">
            <div style={{ marginTop: '25px', paddingBottom: '30px' }}>
              <LeaveReports items={items} />
            </div>
          </PivotItem>
        )}
      </Pivot>

      <LeaveRequestDetail item={selectedItem} isOpen={isDetailOpen} onDismiss={onDismissDetail} onDismissed={onDismissedDetail} />
    </div>
  );
};

export default class LeaveRequestDashboard extends React.Component<ILeaveRequestDashboardProps, {}> {
  public render(): React.ReactElement<ILeaveRequestDashboardProps> {
    return <LeaveRequestDashboardFunctional {...this.props} />;
  }
}
