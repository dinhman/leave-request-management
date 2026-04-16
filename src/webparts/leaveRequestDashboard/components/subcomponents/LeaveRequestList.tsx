import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Stack,
  Text,
  MessageBar,
  MessageBarType,
  getTheme,
  DetailsRow,
  IDetailsRowProps
} from '@fluentui/react';
import { ILeaveRequest } from '../../../../models/ILeaveRequest';
import { StatusBadge } from './StatusBadge';

export interface ILeaveRequestListProps {
  items: ILeaveRequest[];
  onSelectItem: (item: ILeaveRequest) => void;
}

const theme = getTheme();

const detailsListStyles = {
  root: {
    '.ms-DetailsHeader': {
      paddingTop: 0,
      backgroundColor: theme.palette.neutralLighterAlt,
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
    },
    '.ms-DetailsRow': {
      borderBottom: `1px solid ${theme.palette.neutralLighter}`,
      cursor: 'pointer',
      transition: 'background-color 0.2s ease',
      ':hover': {
        backgroundColor: theme.palette.neutralLighterAlt,
      }
    }
  }
};

// Removed local calculateDuration as per security/integrity review.
// We now rely purely on item.NumberOfDays from the source system.

export const LeaveRequestList: React.FC<ILeaveRequestListProps> = (props) => {
  const { items, onSelectItem } = props;

  const columns: IColumn[] = [
    {
      key: 'columnId',
      name: 'ID',
      fieldName: 'Id',
      minWidth: 40,
      maxWidth: 50,
      isResizable: true,
      onRender: (item: ILeaveRequest) => <span style={{ color: theme.palette.neutralSecondary }}>#{item.Id}</span>
    },
    {
      key: 'columnRequester',
      name: 'Requester',
      fieldName: 'RequesterName',
      minWidth: 100,
      maxWidth: 140,
      isResizable: true,
      onRender: (item: ILeaveRequest) => (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Text variant="mediumPlus" style={{ fontWeight: 600, color: theme.palette.themePrimary }}>{item.RequesterName}</Text>
        </Stack>
      )
    },
    {
      key: 'columnStatus',
      name: 'Status',
      fieldName: 'Status',
      minWidth: 110,
      maxWidth: 130,
      isResizable: true,
      onRender: (item: ILeaveRequest) => <StatusBadge status={item.Status} />
    },
    {
      key: 'columnStartDate',
      name: 'Start Date',
      fieldName: 'StartDate',
      minWidth: 90,
      maxWidth: 110,
      isResizable: true,
      onRender: (item: ILeaveRequest) => (
        <Text variant="medium">{item.StartDate ? item.StartDate.split('T')[0].split('-').reverse().join('/') : 'N/A'}</Text>
      )
    },
    {
      key: 'columnEndDate',
      name: 'End Date',
      fieldName: 'EndDate',
      minWidth: 90,
      maxWidth: 110,
      isResizable: true,
      onRender: (item: ILeaveRequest) => (
        <Text variant="medium">{item.EndDate ? item.EndDate.split('T')[0].split('-').reverse().join('/') : 'N/A'}</Text>
      )
    },
    {
      key: 'columnDuration',
      name: 'Duration',
      fieldName: 'Id',
      minWidth: 70,
      maxWidth: 90,
      isResizable: true,
      onRender: (item: ILeaveRequest) => (
        <Stack horizontal verticalAlign="center" horizontalAlign="center">
          <div style={{ padding: '2px 8px', backgroundColor: theme.palette.neutralLighter, borderRadius: '4px' }}>
            <Text variant="smallPlus" style={{ fontWeight: 600 }}>{item.NumberOfDays} Days</Text>
          </div>
        </Stack>
      )
    }
  ];

  const onRenderRow = (p: IDetailsRowProps | undefined) => {
    if (p) {
      return (
        <div 
          onClick={(e) => { 
            e.preventDefault();
            e.stopPropagation();
            onSelectItem(p.item); 
          }} 
          style={{ cursor: 'pointer' }}
        >
          <DetailsRow {...p} />
        </div>
      );
    }
    return null;
  };

  return (
    <div style={{ marginTop: '0px' }}>
      {items.length === 0 ? (
        <MessageBar messageBarType={MessageBarType.info} styles={{ root: { borderRadius: '8px' } }}>
          No leave requests found for this filter.
        </MessageBar>
      ) : (
        <DetailsList
          items={items}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          onRenderRow={onRenderRow}
          styles={detailsListStyles}
        />
      )}
    </div>
  );
};
