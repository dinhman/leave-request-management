import * as React from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  Label,
  Separator,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  getTheme,
  Icon,
  Link
} from '@fluentui/react';
import { ILeaveRequest } from '../../../../models/ILeaveRequest';
import { StatusBadge } from './StatusBadge';

export interface ILeaveRequestDetailProps {
  item: ILeaveRequest | undefined;
  isOpen: boolean;
  onDismiss: () => void;
  onDismissed?: () => void;
}

const theme = getTheme();

export const LeaveRequestDetail: React.FC<ILeaveRequestDetailProps> = (props) => {
  const { item, isOpen, onDismiss, onDismissed } = props;

  const isVN = item?.Region === 'VN';

  const onRenderFooterContent = React.useCallback(
    () => (
      <Stack horizontal tokens={{ childrenGap: 10 }} style={{ padding: '10px 0' }}>
        <PrimaryButton onClick={onDismiss}>{item?.Region === 'VN' ? 'Đóng chi tiết' : 'Close Detail'}</PrimaryButton>
      </Stack>
    ),
    [onDismiss, item],
  );

  return (
    <Panel
      headerText={item ? (item.Region === 'VN' ? `Chi tiết yêu cầu #${item.Id}` : `Request Detail #${item.Id}`) : '...'}
      isOpen={isOpen}
      onDismiss={onDismiss}
      onDismissed={onDismissed}
      type={PanelType.medium}
      onRenderFooterContent={onRenderFooterContent}
      isFooterAtBottom={true}
      headerClassName={undefined}
      isLightDismiss={true}
    >
      {item && (
      <Stack tokens={{ childrenGap: 20 }} style={{ marginTop: '20px' }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="xLarge" style={{ fontWeight: 600, color: theme.palette.themePrimary }}>
            {item.RequesterName}
          </Text>
          <StatusBadge status={item.Status} />
        </Stack>

        <Separator />

        <Stack verticalAlign="start">
          <Label style={{ color: theme.palette.neutralSecondary }}>{isVN ? 'Địa chỉ Email' : 'Email Address'}</Label>
          <Text variant="mediumPlus">{item.RequesterEmail || "N/A"}</Text>
        </Stack>

        {item.SelectedDates ? (
          <Stack>
            <Label style={{ color: theme.palette.neutralSecondary }}>{isVN ? 'Các ngày nghỉ' : 'Selected Dates'}</Label>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
              {item.SelectedDates.split(', ').map((dateStr, idx) => {
                const parts = dateStr.includes('T') ? dateStr.split('T')[0].split('-') : dateStr.split('-');
                const display = parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : dateStr;
                return (
                  <div key={idx} style={{ 
                    padding: '4px 10px', 
                    backgroundColor: theme.palette.themeLighter, 
                    border: `1px solid ${theme.palette.themeLight}`,
                    borderRadius: '16px'
                  }}>
                    <Text variant="smallPlus" style={{ fontWeight: 600, color: theme.palette.themePrimary }}>
                      {display}
                    </Text>
                  </div>
                );
              })}
            </Stack>
          </Stack>
        ) : (
          <Stack horizontal tokens={{ childrenGap: 40 }}>
            <Stack>
              <Label style={{ color: theme.palette.neutralSecondary }}>{isVN ? 'Từ ngày' : 'From Date'}</Label>
              <Text variant="mediumPlus">
                {item.StartDate ? item.StartDate.split('T')[0].split('-').reverse().join('/') : 'N/A'}
              </Text>
            </Stack>
            <Stack>
              <Label style={{ color: theme.palette.neutralSecondary }}>{isVN ? 'Đến ngày' : 'To Date'}</Label>
              <Text variant="mediumPlus">
                {item.EndDate ? item.EndDate.split('T')[0].split('-').reverse().join('/') : 'N/A'}
              </Text>
            </Stack>
          </Stack>
        )}

        <Stack>
          <Label style={{ color: theme.palette.neutralSecondary }}>{isVN ? 'Lý do nghỉ phép' : 'Reason for Leave'}</Label>
          <MessageBar
            messageBarType={MessageBarType.info}
            isMultiline={true}
            styles={{ root: { backgroundColor: theme.palette.neutralLighterAlt, borderRadius: '8px' } }}
          >
            {item.Reason || (isVN ? "Chưa có lý do." : "No reason provided.")}
          </MessageBar>
        </Stack>

        {/* Attachments Section */}
        {item.Attachments && item.Attachments.length > 0 && (
          <Stack>
            <Label style={{ color: theme.palette.neutralSecondary }}>{isVN ? 'Đính kèm' : 'Attachments'}</Label>
            <Stack tokens={{ childrenGap: 8 }}>
              {item.Attachments.map((file, idx) => (
                <Stack key={idx} horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}
                  style={{
                    padding: '8px 12px',
                    border: `1px solid ${theme.palette.neutralLighter}`,
                    borderRadius: '4px'
                  }}>
                  <Icon iconName="Page" style={{ color: theme.palette.themePrimary }} />
                  <Link href={file.ServerRelativeUrl} target="_blank" data-interception="off">
                    {file.FileName}
                  </Link>
                </Stack>
              ))}
            </Stack>
          </Stack>
        )}

        <Separator />

        <Stack verticalAlign="start">
          <Label style={{ color: theme.palette.neutralSecondary }}>{isVN ? 'Tiến trình phê duyệt' : 'Approval Progress'}</Label>
          <Stack tokens={{ childrenGap: 15 }}>
            <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 10 }}>
              <div style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: item.Status === 'Approved' ? theme.palette.green : theme.palette.themePrimary, marginTop: '8px' }} />
              <Stack>
                <Text variant="medium" style={{ fontWeight: 600 }}>{isVN ? 'Người phê duyệt thứ nhất' : 'Approver 1'}</Text>
                <Text variant="small">{isVN ? 'Người duyệt: ' : 'Approver: '} {item.Approver1Name || "Unassigned"}</Text>
                {item.Approver1Comment && (
                  <Text variant="small" style={{ fontStyle: 'italic', color: theme.palette.neutralSecondary }}>
                    {isVN ? 'Ghi chú: ' : 'Note: '} {item.Approver1Comment}
                  </Text>
                )}
              </Stack>
            </Stack>

            <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 10 }}>
              <div style={{ width: '8px', height: '8px', borderRadius: '50%', backgroundColor: item.Status === 'Approved' ? theme.palette.green : theme.palette.themePrimary, marginTop: '8px' }} />
              <Stack>
                <Text variant="medium" style={{ fontWeight: 600 }}>{isVN ? 'Người phê duyệt thứ hai' : 'Approver 2'}</Text>
                <Text variant="small">{isVN ? 'Người duyệt: ' : 'Approver: '} {item.Approver2Name || "Unassigned"}</Text>
                {item.Approver2Comment && (
                  <Text variant="small" style={{ fontStyle: 'italic', color: theme.palette.neutralSecondary }}>
                    {isVN ? 'Ghi chú: ' : 'Note: '} {item.Approver2Comment}
                  </Text>
                )}
              </Stack>
            </Stack>
          </Stack>
        </Stack>
      </Stack>
      )}
    </Panel>
  );
};
