import * as React from 'react';
import { Label, ILabelStyles } from '@fluentui/react';
import { LeaveStatus } from '../../../../models/ILeaveRequest';

export interface IStatusBadgeProps {
  status: LeaveStatus;
}

const statusStyles: { [key in LeaveStatus]: ILabelStyles } = {
  Approved: {
    root: {
      backgroundColor: '#dff6dd',
      color: '#107c10',
      padding: '2px 10px',
      borderRadius: '16px',
      fontSize: '12px',
      fontWeight: '600',
      border: '1px solid #107c10'
    }
  },
  'Awaiting approval': {
    root: {
      backgroundColor: '#fff4ce',
      color: '#797775',
      padding: '2px 10px',
      borderRadius: '16px',
      fontSize: '12px',
      fontWeight: '600',
      border: '1px solid #797775'
    }
  },
  Rejected: {
    root: {
      backgroundColor: '#fde7e9',
      color: '#a4262c',
      padding: '2px 10px',
      borderRadius: '16px',
      fontSize: '12px',
      fontWeight: '600',
      border: '1px solid #a4262c'
    }
  },
  Draft: {
    root: {
      backgroundColor: '#f3f2f1',
      color: '#605e5c',
      padding: '2px 10px',
      borderRadius: '16px',
      fontSize: '12px',
      fontWeight: '600',
      border: '1px solid #605e5c'
    }
  }
};

export const StatusBadge: React.FC<IStatusBadgeProps> = (props) => {
  const { status } = props;

  return (
    <Label styles={statusStyles[status]}>
      {status}
    </Label>
  );
};
