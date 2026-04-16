import * as React from 'react';
import {
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  getTheme,
  mergeStyles
} from '@fluentui/react';
import { ILeaveRequest } from '../../../../models/ILeaveRequest';

export interface ILeaveReportsProps {
  items: ILeaveRequest[];
}

const theme = getTheme();

const reportCardClass = mergeStyles({
  padding: '25px',
  backgroundColor: theme.palette.white,
  borderRadius: '12px',
  boxShadow: '0 4px 12px rgba(0, 0, 0, 0.05)',
  marginTop: '20px'
});

const kpiMiniCardClass = mergeStyles({
  padding: '15px 20px',
  backgroundColor: theme.palette.neutralLighterAlt,
  borderRadius: '8px',
  borderLeft: `4px solid ${theme.palette.themePrimary}`,
  flex: '1 1 150px'
});

export const LeaveReports: React.FC<ILeaveReportsProps> = (props) => {
  const { items } = props;
  
  const currentMonth = new Date().getMonth();
  const currentYear = new Date().getFullYear();
  
  const [selectedMonth, setSelectedMonth] = React.useState<number>(currentMonth);
  const [selectedYear, setSelectedYear] = React.useState<number>(currentYear);

  const monthOptions: IDropdownOption[] = [
    { key: 0, text: 'January' }, { key: 1, text: 'February' }, { key: 2, text: 'March' },
    { key: 3, text: 'April' }, { key: 4, text: 'May' }, { key: 5, text: 'June' },
    { key: 6, text: 'July' }, { key: 7, text: 'August' }, { key: 8, text: 'September' },
    { key: 9, text: 'October' }, { key: 10, text: 'November' }, { key: 11, text: 'December' },
  ];

  const yearOptions: IDropdownOption[] = [
    { key: 2024, text: '2024' },
    { key: 2025, text: '2025' },
    { key: 2026, text: '2026' },
  ];

  const aggregatedData = React.useMemo(() => {
    const reportData: { [key: string]: { name: string, totalDays: number, totalRequests: number } } = {};
    
    // Filter items first by month/year (based on StartDate)
    const filteredItems = items.filter(item => {
      const date = new Date(item.StartDate);
      return date.getMonth() === selectedMonth && date.getFullYear() === selectedYear && item.Status === 'Approved';
    });

    filteredItems.forEach(item => {
      if (!reportData[item.RequesterName]) {
        reportData[item.RequesterName] = { 
          name: item.RequesterName, 
          totalDays: 0, 
          totalRequests: 0 
        };
      }
      reportData[item.RequesterName].totalDays += item.NumberOfDays;
      reportData[item.RequesterName].totalRequests += 1;
    });

    return Object.values(reportData).sort((a, b) => b.totalDays - a.totalDays);
  }, [items, selectedMonth, selectedYear]);

  const totalMonthlyDays = aggregatedData.reduce((acc, curr) => acc + curr.totalDays, 0);

  const columns: IColumn[] = [
    {
      key: 'name',
      name: 'Employee Name',
      fieldName: 'name',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item) => <Text style={{ fontWeight: 600 }}>{item.name}</Text>
    },
    {
      key: 'totalDays',
      name: 'Total Leave Days',
      fieldName: 'totalDays',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item) => (
        <Stack horizontal verticalAlign="center">
          <Text variant="mediumPlus" style={{ color: theme.palette.themePrimary, fontWeight: 700 }}>
            {item.totalDays}
          </Text>
          <Text variant="small" style={{ marginLeft: 5, color: theme.palette.neutralSecondary }}>days</Text>
        </Stack>
      )
    },
    {
      key: 'totalRequests',
      name: 'Request Count',
      fieldName: 'totalRequests',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 20 }} style={{ padding: '0 10px' }}>
      {/* Report Summary Cards */}
      <Stack horizontal wrap tokens={{ childrenGap: 20 }}>
        <div className={kpiMiniCardClass}>
          <Text variant="small" style={{ color: theme.palette.neutralSecondary, textTransform: 'uppercase', fontWeight: 600 }}>
            Monthly Total (All Employees)
          </Text>
          <Text variant="xxLarge" style={{ display: 'block', fontWeight: 700, marginTop: 5 }}>
            {totalMonthlyDays} <Text variant="medium" style={{ fontWeight: 400 }}>days</Text>
          </Text>
        </div>
        <div className={kpiMiniCardClass}>
          <Text variant="small" style={{ color: theme.palette.neutralSecondary, textTransform: 'uppercase', fontWeight: 600 }}>
            Active Requesters
          </Text>
          <Text variant="xxLarge" style={{ display: 'block', fontWeight: 700, marginTop: 5 }}>
            {aggregatedData.length} <Text variant="medium" style={{ fontWeight: 400 }}>staff</Text>
          </Text>
        </div>
      </Stack>

      {/* Filter Bar */}
      <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 20 }} style={{ padding: '10px 0' }}>
        <Dropdown
          label="Select Year"
          options={yearOptions}
          selectedKey={selectedYear}
          onChange={(_, opt) => setSelectedYear(opt?.key as number)}
          styles={{ root: { width: 150 } }}
        />
        <Dropdown
          label="Select Month"
          options={monthOptions}
          selectedKey={selectedMonth}
          onChange={(_, opt) => setSelectedMonth(opt?.key as number)}
          styles={{ root: { width: 180 } }}
        />
      </Stack>

      <div className={reportCardClass}>
        <Text variant="large" style={{ fontWeight: 600, display: 'block', marginBottom: 20 }}>
          Monthly Leave Consumption Summary
        </Text>
        <DetailsList
          items={aggregatedData}
          columns={columns}
          selectionMode={SelectionMode.none}
          layoutMode={DetailsListLayoutMode.justified}
          styles={{
            root: {
              '.ms-DetailsHeader': {
                backgroundColor: theme.palette.neutralLighterAlt,
                paddingTop: 0,
              }
            }
          }}
        />
        {aggregatedData.length === 0 && (
          <Stack horizontalAlign="center" style={{ padding: '40px 0' }}>
            <Text style={{ color: theme.palette.neutralSecondary }}>No approved leave requests found for this period.</Text>
          </Stack>
        )}
      </div>
    </Stack>
  );
};
