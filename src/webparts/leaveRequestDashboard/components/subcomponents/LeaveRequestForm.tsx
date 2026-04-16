import * as React from 'react';
import {
  Stack,
  Text,
  TextField,
  DatePicker,
  PrimaryButton,
  DefaultButton,
  IconButton,
  MessageBar,
  MessageBarType,
  getTheme,
  mergeStyles,
  Label,
  Dropdown,
  IDropdownOption,
  Icon,
  Dialog,
  DialogType,
  DialogFooter,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { Region } from '../../../../models/ILeaveRequest';
import { ILeaveBalance } from '../../../../models/ILeaveBalance';
import { SPService, ISPService } from '../../../../services/SPService';

interface ILeaveDate {
  date: Date;
  duration: 'Full' | 'Morning' | 'Afternoon';
}

export interface ILeaveRequestFormProps {
  userDisplayName: string;
  userEmail: string;
  automateUrl: string;
  httpClient: HttpClient;
  region: Region;
  onSuccess?: () => void;
}

const theme = getTheme();

const formContainerClass = mergeStyles({
  padding: '30px',
  backgroundColor: theme.palette.white,
  borderRadius: '12px',
  boxShadow: '0 8px 32px rgba(0, 0, 0, 0.08)',
  maxWidth: '600px',
  margin: '0 auto'
});

const formatDateToString = (date: Date | undefined): string => {
  if (!date) return "";
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}-${month}-${day}`;
};

export const LeaveRequestForm: React.FC<ILeaveRequestFormProps> = (props) => {
  const { userDisplayName, userEmail, automateUrl, httpClient, region } = props;
  const isVN = region === 'VN';

  const [selectedDates, setSelectedDates] = React.useState<ILeaveDate[]>([]);
  const [currentPickedDate, setCurrentPickedDate] = React.useState<Date | undefined>(undefined);
  const [leaveType, setLeaveType] = React.useState<string>("Annual Leave");
  const [numberOfDays, setNumberOfDays] = React.useState<number>(1);
  const [reason, setReason] = React.useState<string>("");
  // Optimizing memory: keep standard File objects in state instead of heavy Base64 strings.
  const [attachments, setAttachments] = React.useState<File[]>([]);
  const [balance, setBalance] = React.useState<ILeaveBalance | undefined>(undefined);
  const [loadingBalance, setLoadingBalance] = React.useState<boolean>(false);
  const [submitting, setSubmitting] = React.useState<boolean>(false);
  const [status, setStatus] = React.useState<{ type: MessageBarType, message: string } | undefined>(undefined);
  const [isSuccessDialogOpen, setIsSuccessDialogOpen] = React.useState<boolean>(false);
  const [isAttachmentWarningOpen, setIsAttachmentWarningOpen] = React.useState<boolean>(false);

  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const addDate = () => {
    if (!currentPickedDate) return;
    const dateStr = formatDateToString(currentPickedDate);
    const isDuplicate = selectedDates.some(d => formatDateToString(d.date) === dateStr);
    if (!isDuplicate) {
      const newDate: ILeaveDate = { date: currentPickedDate, duration: 'Full' };
      const newDates: ILeaveDate[] = [...selectedDates, newDate]
        .sort((a, b) => a.date.getTime() - b.date.getTime());
      setSelectedDates(newDates);
    }
  };

  const removeSelectedDate = (index: number) => {
    const newDates = [...selectedDates];
    newDates.splice(index, 1);
    setSelectedDates(newDates);
  };

  const updateDateDuration = (index: number, duration: 'Full' | 'Morning' | 'Afternoon'): void => {
    const newDates = [...selectedDates];
    newDates[index].duration = duration;
    setSelectedDates(newDates);
  };

  React.useEffect(() => {
    const total = selectedDates.reduce((acc, curr) => acc + (curr.duration === 'Full' ? 1 : 0.5), 0);
    setNumberOfDays(total);
  }, [selectedDates]);

  const spService: ISPService = new SPService();

  const fetchBalance = async (): Promise<void> => {
    try {
      setLoadingBalance(true);
      const userBalance = await spService.getUserBalance(userEmail, region);
      setBalance(userBalance);
    } catch (error) {
      console.error("Error fetching balance:", error);
    } finally {
      setLoadingBalance(false);
    }
  };

  React.useEffect(() => {
    fetchBalance().catch(() => {});
  }, [userEmail, region]);

  const isAnnualLeave = leaveType === 'Annual Leave';
  const remainingDays = balance ? balance.AnnualLeaveRemaining : 0;
  const isOverBalance = isAnnualLeave && numberOfDays > remainingDays;

  const leaveTypeOptionsIntl: IDropdownOption[] = [
    { key: 'Annual Leave', text: 'Annual Leave', disabled: balance !== undefined && balance.AnnualLeaveRemaining <= 0 },
    { key: 'Medical Leave', text: 'Medical Leave' },
    { key: 'Hospitalisation Leave', text: 'Hospitalisation Leave' },
    { key: 'Childcare Leave', text: 'Childcare Leave' },
    { key: 'Unpaid Leave', text: 'Unpaid Leave' },
    { key: 'Maternity Leave', text: 'Maternity Leave' },
    { key: 'Paternity Leave', text: 'Paternity Leave' },
    { key: 'Work from home', text: 'Work from home' }
  ];

  const leaveTypeOptionsVN: IDropdownOption[] = [
    { key: 'Business trip', text: 'Đi công tác/ Business trip' },
    { key: 'Unpaid Leave', text: 'Nghỉ không lương/ Unpaid Leave' },
    { key: 'Annual Leave', text: 'Nghỉ phép năm/ Annual Leave', disabled: balance !== undefined && balance.AnnualLeaveRemaining <= 0 },
    { key: 'Sick Leave', text: 'Nghỉ ốm, nghỉ con ốm/ Sick Leave, Sick Child Leave' },
    { key: 'Paternity Leave', text: 'Nghỉ vợ sinh/ Paternity Leave' },
    { key: 'Maternity Leave', text: 'Nghỉ thai sản/ Maternity Leave' },
    { key: 'Medical checkup', text: 'Nghỉ phép đi khám bệnh / Medical checkup' },
    { key: 'Marriage Leave', text: 'Nghỉ kết hôn / Marriage Leave' },
    { key: 'Child Marriage Leave', text: 'Nghỉ con kết hôn / Child\'s Marriage Leave' },
    { key: 'Funeral Leave', text: 'Nghỉ tang chế (Cha mẹ, Con, Vợ hoặc Chồng)/ Funeral Leave (Parents; Spouse\'s Parents; Child; Spouse)' }
  ];

  const leaveTypeOptions = isVN ? leaveTypeOptionsVN : leaveTypeOptionsIntl;

  const convertFileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        const result = reader.result as string;
        const base64 = result.split(',')[1];
        resolve(base64);
      };
      reader.onerror = (error) => reject(error);
    });
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    const newAttachments = [...attachments];
    for (let i = 0; i < files.length; i++) {
        newAttachments.push(files[i]);
    }

    setAttachments(newAttachments);
    if (fileInputRef.current) fileInputRef.current.value = "";
  };

  const removeAttachment = (index: number): void => {
    const newAttachments = [...attachments];
    newAttachments.splice(index, 1);
    setAttachments(newAttachments);
  };

  const typesRequiringAttachments = ['Paternity Leave', 'Sick Leave', 'Marriage Leave', 'Funeral Leave'];

  const performSubmit = async (uploadLater: boolean = false): Promise<void> => {
    try {
      setSubmitting(true);
      setStatus(undefined);

      // Perform conversion right before submission to save memory
      const validAttachments = [];
      for (const file of attachments) {
         try {
           const base64 = await convertFileToBase64(file);
           validAttachments.push({
             fileName: file.name,
             content: base64
           });
         } catch(e) {
           console.error("Error converting file:", e);
         }
      }

      const durationLabels: { [key: string]: string } = {
        'Full': isVN ? 'Cả ngày' : 'Full Day',
        'Morning': isVN ? 'Sáng' : 'Morning',
        'Afternoon': isVN ? 'Chiều' : 'Afternoon'
      };

      const selectedDatesStr = selectedDates.map(d => 
        `${d.date.getDate().toString().padStart(2, '0')}/${(d.date.getMonth() + 1).toString().padStart(2, '0')} (${durationLabels[d.duration]})`
      ).join(', ');

      const finalReason = uploadLater 
        ? `${reason}\n\n(Chưa có chứng từ đính kèm - Sẽ bổ sung sau / No attachments provided - Will provide later)` 
        : reason;

      const payload = {
        requester: userDisplayName,
        email: userEmail,
        region: region,
        leaveType: leaveType,
        numberOfDays: numberOfDays,
        attachments: validAttachments,
        startDate: formatDateToString(selectedDates[0].date),
        endDate: formatDateToString(selectedDates[selectedDates.length - 1].date),
        selectedDates: selectedDatesStr,
        reason: finalReason,
        submittedAt: new Date().toISOString(),
        annualLeaveRemaining: remainingDays,
        isAnnualLeave: isAnnualLeave
      };

      const requestOptions: IHttpClientOptions = {
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      };

      const response: HttpClientResponse = await httpClient.post(
        automateUrl,
        HttpClient.configurations.v1,
        requestOptions
      );

      if (response.ok) {
        setIsSuccessDialogOpen(true);
        setReason("");
        setAttachments([]);
        setSelectedDates([]);
        setCurrentPickedDate(undefined);
      } else {
        console.error("Submission failed:", response.statusText);
        setStatus({ type: MessageBarType.error, message: isVN ? `Gửi thất bại: ${response.statusText}` : `Failed to submit. Status: ${response.statusText}` });
      }
    } catch (err) {
      console.error("Error submitting form:", err);
      setStatus({ type: MessageBarType.error, message: isVN ? "Lỗi không xác định khi gửi form." : "An unexpected error occurred." });
    } finally {
      setSubmitting(false);
    }
  };

  const handleSubmit = async (): Promise<void> => {
    if (selectedDates.length === 0 || !reason || !leaveType || numberOfDays <= 0) {
      setStatus({ type: MessageBarType.error, message: isVN ? "Vui lòng điền đầy đủ thông tin bắt buộc." : "Please fill in all required fields." });
      return;
    }

    if (isOverBalance) {
      setStatus({ type: MessageBarType.error, message: isVN ? `Số ngày nghỉ phép năm vượt quá số dư hiện có (${remainingDays} ngày).` : `Requested days exceed your annual leave balance (${remainingDays} days).` });
      return;
    }

    if (!automateUrl) {
      setStatus({ type: MessageBarType.warning, message: isVN ? "Chưa cấu hình URL phê duyệt." : "Automate URL is not configured." });
      return;
    }

    // Check for required attachments
    if (typesRequiringAttachments.indexOf(leaveType) !== -1 && attachments.length === 0) {
      setIsAttachmentWarningOpen(true);
      return;
    }

    await performSubmit();
  };

  return (
    <div className={formContainerClass}>
      <Stack tokens={{ childrenGap: 25 }}>
        <Stack>
          <Text variant="xLarge" style={{ fontWeight: 700, color: theme.palette.themePrimary }}>
             {isVN ? 'Đăng ký nghỉ phép' : 'Submit Leave Request'}
          </Text>
          <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
            {isVN ? 'Hãy điền chi tiết bên dưới để thông báo cho Quản lý.' : 'Fill in the details below to notify your manager and HR.'}
          </Text>
        </Stack>

        {/* Display Balance Card */}
        <Stack 
          style={{ 
            padding: '12px 18px', 
            backgroundColor: isOverBalance ? '#fde7e9' : '#edf2f7', 
            borderRadius: '8px',
            borderLeft: `4px solid ${isOverBalance ? theme.palette.red : theme.palette.themePrimary}`
          }}
          tokens={{ childrenGap: 5 }}
        >
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="small" style={{ fontWeight: 600, color: theme.palette.neutralSecondary }}>
              {isVN ? 'SỐ DƯ PHÉP NĂM' : 'ANNUAL LEAVE BALANCE'}
            </Text>
            {loadingBalance && <Spinner size={SpinnerSize.xSmall} />}
          </Stack>
          <Text variant="large" style={{ fontWeight: 700, color: isOverBalance ? theme.palette.red : theme.palette.neutralPrimary }}>
            {balance ? `${balance.AnnualLeaveRemaining} ${isVN ? 'ngày' : 'days'}` : (loadingBalance ? '...' : (isVN ? 'Chưa khởi tạo' : 'Not initialized'))}
          </Text>
          {isOverBalance && (
            <Text variant="small" style={{ color: theme.palette.red, fontWeight: 600 }}>
              {isVN ? 'Cảnh báo: Vượt quá số dư!' : 'Warning: Exceeds balance!'}
            </Text>
          )}
        </Stack>

        {status && (
          <MessageBar messageBarType={status.type} onDismiss={() => setStatus(undefined)} isMultiline={false}>
            {status.message}
          </MessageBar>
        )}

        <Stack tokens={{ childrenGap: 15 }}>
          <Stack horizontal tokens={{ childrenGap: 20 }}>
            <Stack style={{ flex: 1 }}>
              <Label>{isVN ? 'Họ tên' : 'Requester Name'}</Label>
              <Text variant="mediumPlus" style={{ padding: '7px 0', borderBottom: `1px solid ${theme.palette.neutralLighter}` }}>
                {userDisplayName}
              </Text>
            </Stack>
            <Stack style={{ flex: 1 }}>
              <Label>Email</Label>
              <Text variant="mediumPlus" style={{ padding: '7px 0', borderBottom: `1px solid ${theme.palette.neutralLighter}` }}>
                {userEmail}
              </Text>
            </Stack>
          </Stack>

          <Stack horizontal tokens={{ childrenGap: 20 }} verticalAlign="end">
            <Dropdown
              label={isVN ? 'Loại nghỉ phép' : 'Type of Leave'}
              required={true}
              options={leaveTypeOptions}
              selectedKey={leaveType}
              onChange={(_, option) => setLeaveType(option?.key as string || "")}
              placeholder={isVN ? 'Chọn...' : 'Select...'}
              calloutProps={{ doNotLayer: true }}
              styles={{ root: { flex: 2 } }}
            />
            <TextField
              label={isVN ? 'Số ngày' : 'Number of Days'}
              type="number"
              required={true}
              readOnly={true}
              value={numberOfDays.toString()}
              styles={{ root: { flex: 1 } }}
            />
          </Stack>

          <Stack tokens={{ childrenGap: 10 }} style={{ padding: '15px', backgroundColor: theme.palette.neutralLighterAlt, borderRadius: '8px', border: `1px dashed ${theme.palette.neutralTertiaryAlt}` }}>
            <Label required>{isVN ? 'Chọn các ngày nghỉ' : 'Select Leave Dates'}</Label>
            <Stack horizontal tokens={{ childrenGap: 15 }} verticalAlign="end">
              <DatePicker
                value={currentPickedDate}
                onSelectDate={(date) => setCurrentPickedDate(date || undefined)}
                placeholder={isVN ? 'Chọn ngày...' : 'Select a date...'}
                styles={{ root: { flex: 1 } }}
              />
              <DefaultButton 
                text={isVN ? 'Thêm ngày' : 'Add Date'} 
                iconProps={{ iconName: 'Add' }} 
                onClick={addDate}
                disabled={!currentPickedDate}
                styles={{ root: { backgroundColor: theme.palette.themePrimary, color: 'white', border: 'none' }, rootHovered: { backgroundColor: theme.palette.themeDark, color: 'white' }, rootDisabled: { backgroundColor: theme.palette.neutralLighter } }}
              />
            </Stack>

            {selectedDates.length > 0 && (
              <Stack horizontal wrap tokens={{ childrenGap: 8 }} style={{ marginTop: '10px' }}>
                {selectedDates.map((item, idx) => (
                  <Stack 
                    key={idx} 
                    horizontal 
                    verticalAlign="center" 
                    tokens={{ childrenGap: 10 }}
                    style={{ 
                      padding: '8px 15px', 
                      backgroundColor: theme.palette.white, 
                      border: `1px solid ${theme.palette.neutralLight}`,
                      borderRadius: '8px',
                      width: '100%',
                      boxShadow: '0 2px 4px rgba(0,0,0,0.02)'
                    }}
                  >
                    <Text variant="medium" style={{ fontWeight: 600, width: '100px' }}>
                      {item.date.getDate().toString().padStart(2, '0')}/{((item.date.getMonth() + 1)).toString().padStart(2, '0')}
                    </Text>
                    
                    <Stack horizontal tokens={{ childrenGap: 5 }} style={{ flex: 1 }}>
                      {[
                        { key: 'Full', text: isVN ? 'Cả ngày' : 'Full' },
                        { key: 'Morning', text: isVN ? 'Sáng' : 'AM' },
                        { key: 'Afternoon', text: isVN ? 'Chiều' : 'PM' }
                      ].map(dur => (
                        <DefaultButton
                          key={dur.key}
                          text={dur.text}
                          onClick={() => updateDateDuration(idx, dur.key as 'Full' | 'Morning' | 'Afternoon')}
                          styles={{ 
                            root: { 
                              height: '28px', 
                              padding: '0 8px', 
                              fontSize: '11px',
                              borderRadius: '4px',
                              border: item.duration === dur.key ? `1px solid ${theme.palette.themePrimary}` : `1px solid ${theme.palette.neutralLighter}`,
                              backgroundColor: item.duration === dur.key ? theme.palette.themeLighter : theme.palette.white,
                              color: item.duration === dur.key ? theme.palette.themePrimary : theme.palette.neutralPrimary
                            } 
                          }}
                        />
                      ))}
                    </Stack>

                    <IconButton 
                      iconProps={{ iconName: "Cancel" }} 
                      styles={{ root: { color: theme.palette.neutralSecondary, height: '28px' } }} 
                      onClick={() => removeSelectedDate(idx)}
                    />
                  </Stack>
                ))}
              </Stack>
            )}
          </Stack>

          <TextField
            label={isVN ? 'Lý do' : 'Reason for Leave'}
            multiline
            rows={4}
            required={true}
            value={reason}
            onChange={(_, val) => setReason(val || "")}
            placeholder={isVN ? 'Cung cấp lý do ngắn gọn...' : 'Please provide a brief reason...'}
          />

          <Stack tokens={{ childrenGap: 10 }}>
            <Label>{isVN ? 'Đính kèm (Tùy chọn)' : 'Attachments (Optional)'}</Label>
            <input
              type="file"
              multiple
              ref={fileInputRef}
              style={{ display: 'none' }}
              onChange={handleFileChange}
            />
            <DefaultButton
              text={isVN ? 'Thêm file' : 'Attach File'}
              iconProps={{ iconName: 'Add' }}
              onClick={() => fileInputRef.current?.click()}
              styles={{ root: { width: 'fit-content' } }}
            />
            
            {attachments.length > 0 && (
              <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: '10px' }}>
                {attachments.map((file, index) => (
                  <Stack 
                    key={index} 
                    horizontal 
                    verticalAlign="center" 
                    horizontalAlign="space-between"
                    style={{ padding: '8px 12px', backgroundColor: theme.palette.neutralLighter, borderRadius: '6px' }}
                  >
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
                      <Icon iconName="Page" style={{ color: theme.palette.themePrimary }} />
                      <Text variant="smallPlus">{file.name}</Text>
                    </Stack>
                    <IconButton
                      iconProps={{ iconName: 'Cancel' }}
                      title={isVN ? 'Xóa' : 'Remove'}
                      onClick={() => removeAttachment(index)}
                    />
                  </Stack>
                ))}
              </Stack>
            )}
          </Stack>
        </Stack>

        <PrimaryButton
          text={submitting ? (isVN ? "Đang gửi..." : "Submitting...") : (isVN ? "Gửi yêu cầu" : "Send Request")}
          onClick={handleSubmit}
          disabled={submitting}
          iconProps={{ iconName: 'Send' }}
          styles={{ root: { height: '45px', borderRadius: '8px', marginTop: '15px' } }}
        />
      </Stack>
      <Dialog
        hidden={!isSuccessDialogOpen}
        onDismiss={() => setIsSuccessDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: isVN ? 'Đăng ký thành công' : 'Success',
          closeButtonAriaLabel: 'Close',
          subText: isVN ? 'Yêu cầu của bạn đã được gửi. Vui lòng chờ sự chấp thuận của Quản lý.' : 'Your request has been submitted. Please wait for Manager\'s approval.'
        }}
        modalProps={{ isBlocking: false, styles: { main: { maxWidth: 450 } } }}
      >
        <DialogFooter>
          <PrimaryButton 
            onClick={() => {
              setIsSuccessDialogOpen(false);
              if (props.onSuccess) props.onSuccess();
            }} 
            text={isVN ? 'Đồng ý' : 'OK'} 
            styles={{ root: { borderRadius: '6px' } }}
          />
        </DialogFooter>
      </Dialog>

      <Dialog
        hidden={!isAttachmentWarningOpen}
        onDismiss={() => setIsAttachmentWarningOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: isVN ? 'Thiếu chứng từ đính kèm' : 'Missing Attachments',
          subText: isVN 
            ? 'Loại nghỉ này yêu cầu phải có chứng từ đính kèm. Bạn muốn quay lại để tải lên hay sẽ bổ sung sau?' 
            : 'This leave type requires supporting documents. Would you like to go back and upload them or provide them later?'
        }}
        modalProps={{ isBlocking: true, styles: { main: { maxWidth: 450 } } }}
      >
        <DialogFooter>
          <PrimaryButton 
            onClick={() => setIsAttachmentWarningOpen(false)} 
            text={isVN ? 'Quay lại upload' : 'Back to upload'} 
          />
          <DefaultButton 
            onClick={() => {
              setIsAttachmentWarningOpen(false);
              performSubmit(true).catch(() => {});
            }} 
            text={isVN ? 'Bổ sung sau' : 'Upload later'} 
            styles={{ root: { color: theme.palette.orange } }}
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
