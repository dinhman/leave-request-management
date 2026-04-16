import * as React from 'react';
import {
  Stack,
  Text,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  getTheme,
  mergeStyles,
  IconButton,
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  TextField,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  SearchBox
} from '@fluentui/react';
import { ILeaveBalance } from '../../../../models/ILeaveBalance';
import { Region } from '../../../../models/ILeaveRequest';
import { ISPService, SPService } from '../../../../services/SPService';

export interface IBalanceManagementProps {
  region: Region;
}

const theme = getTheme();

const containerClass = mergeStyles({
  padding: '20px',
  backgroundColor: theme.palette.white,
  borderRadius: '12px',
  boxShadow: '0 4px 12px rgba(0, 0, 0, 0.05)',
});

export const BalanceManagement: React.FC<IBalanceManagementProps> = (props) => {
  const { region } = props;
  const [balances, setBalances] = React.useState<ILeaveBalance[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  
  const [editingItem, setEditingItem] = React.useState<ILeaveBalance | undefined>(undefined);
  const [editValue, setEditValue] = React.useState<string>("");
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  
  const [searchTerm, setSearchTerm] = React.useState<string>("");
  const [currentPage, setCurrentPage] = React.useState<number>(1);
  const pageSize = 10;

  const spService: ISPService = new SPService();

  const filteredBalances = React.useMemo(() => {
    if (!searchTerm) return balances;
    const lowerSearch = searchTerm.toLowerCase();
    return balances.filter(b => 
      (b.EmployeeName && b.EmployeeName.toLowerCase().indexOf(lowerSearch) > -1) || 
      (b.UserEmail && b.UserEmail.toLowerCase().indexOf(lowerSearch) > -1)
    );
  }, [balances, searchTerm]);

  const pagedBalances = React.useMemo(() => {
    return filteredBalances.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  }, [filteredBalances, currentPage]);

  const fetchData = async () => {
    try {
      setLoading(true);
      const data = await spService.getAllBalances(region);
      setBalances(data);
      setError(undefined);
    } catch (err) {
      console.error("Error fetching balances:", err);
      setError("Failed to load balances. Please ensure the 'LeaveBalance' list exists.");
    } finally {
      setLoading(false);
    }
  };

  React.useEffect(() => {
    fetchData().catch(() => {});
  }, [region]);

  const onEdit = (item: ILeaveBalance) => {
    setEditingItem(item);
    setEditValue(item.AnnualLeaveTotal.toString());
  };

  const onSave = async () => {
    if (!editingItem) return;
    try {
      setIsSaving(true);
      const newValue = parseFloat(editValue);
      if (isNaN(newValue)) throw new Error("Invalid number");

      await spService.updateUserBalance(editingItem.Id, { AnnualLeaveTotal: newValue });
      
      setEditingItem(undefined);
      await fetchData();
    } catch (err) {
      console.error("Error saving balance:", err);
      alert("Failed to save changes.");
    } finally {
      setIsSaving(false);
    }
  };

  const columns: IColumn[] = [
    {
      key: 'user',
      name: 'Email',
      fieldName: 'UserEmail',
      minWidth: 150,
      maxWidth: 250,
      isResizable: true,
    },
    {
      key: 'name',
      name: 'Employee Name',
      fieldName: 'EmployeeName',
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: 'total',
      name: 'Total Days',
      fieldName: 'AnnualLeaveTotal',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: ILeaveBalance) => (
        <Text style={{ fontWeight: 600 }}>{item.AnnualLeaveTotal}</Text>
      )
    },
    {
      key: 'used',
      name: 'Used',
      fieldName: 'AnnualLeaveUsed',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: ILeaveBalance) => (
        <Text style={{ color: theme.palette.neutralSecondary }}>{item.AnnualLeaveUsed}</Text>
      )
    },
    {
      key: 'remaining',
      name: 'Remaining',
      fieldName: 'AnnualLeaveRemaining',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: ILeaveBalance) => (
        <Text style={{ color: item.AnnualLeaveRemaining <= 0 ? theme.palette.red : theme.palette.green, fontWeight: 700 }}>
          {item.AnnualLeaveRemaining}
        </Text>
      )
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 50,
      maxWidth: 50,
      onRender: (item: ILeaveBalance) => (
        <IconButton
          iconProps={{ iconName: 'Edit' }}
          title="Edit Total Balance"
          onClick={() => onEdit(item)}
        />
      )
    }
  ];

  return (
    <div className={containerClass}>
      <Stack tokens={{ childrenGap: 20 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="xLarge" style={{ fontWeight: 600 }}>Leave Balance Management ({region})</Text>
          <DefaultButton iconProps={{ iconName: 'Refresh' }} onClick={fetchData} text="Reload" />
        </Stack>

        <SearchBox 
          placeholder="Search by name or email..." 
          onChange={(_, val) => {
            setSearchTerm(val || "");
            setCurrentPage(1);
          }}
          styles={{ root: { maxWidth: 400 } }}
        />
        
        {error && (
          <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
        )}

        {loading ? (
          <Spinner size={SpinnerSize.large} label="Loading balances..." />
        ) : (
          <DetailsList
            items={pagedBalances}
            columns={columns}
            selectionMode={SelectionMode.none}
            layoutMode={DetailsListLayoutMode.justified}
          />
        )}

        {/* Pagination Controls */}
        {!loading && filteredBalances.length > pageSize && (
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
              Page {currentPage} of {Math.ceil(filteredBalances.length / pageSize)}
            </Text>
            <DefaultButton
              text="Next"
              iconProps={{ iconName: 'ChevronRight' }}
              onClick={() => setCurrentPage(prev => Math.min(Math.ceil(filteredBalances.length / pageSize), prev + 1))}
              disabled={currentPage === Math.ceil(filteredBalances.length / pageSize)}
            />
          </Stack>
        )}
      </Stack>

      <Dialog
        hidden={!editingItem}
        onDismiss={() => setEditingItem(undefined)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Update Annual Leave Total',
          subText: `Updating total days for ${editingItem?.EmployeeName}`
        }}
        modalProps={{ isBlocking: true }}
      >
        <TextField
          label="Total Annual Leave Days"
          type="number"
          value={editValue}
          onChange={(_, val) => setEditValue(val || "")}
          autoFocus
        />
        <DialogFooter>
          {isSaving ? (
            <Spinner size={SpinnerSize.small} />
          ) : (
            <>
              <PrimaryButton onClick={onSave} text="Save" />
              <DefaultButton onClick={() => setEditingItem(undefined)} text="Cancel" />
            </>
          )}
        </DialogFooter>
      </Dialog>
    </div>
  );
};
