import * as React from 'react';
import { 
  Stack, 
  Text, 
  Icon, 
  ActionButton, 
  getTheme, 
  mergeStyles 
} from '@fluentui/react';

export interface IAppHeaderProps {
  currentRegion?: string;
  onResetRegion: () => void;
}

const theme = getTheme();

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

export function AppHeader(props: IAppHeaderProps): JSX.Element {
  const { currentRegion, onResetRegion } = props;

  return (
    <div className={headerContainerClass}>
      {/* Logo & Title */}
      <div className={logoSectionClass}>
        <div style={{ backgroundColor: theme.palette.themePrimary, padding: '8px', borderRadius: '8px' }}>
          <Icon iconName="SkypeMessage" style={{ color: 'white', fontSize: '20px' }} />
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

      {/* Region Info & Actions */}
      {currentRegion && (
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 20 }}>
          <div className={regionBadgeClass}>
            <Icon iconName="World" style={{ marginRight: '6px', verticalAlign: 'middle' }} />
            {currentRegion === 'VN' ? 'Vietnam' : 'Indonesia'}
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
}
