import * as React from 'react';
import { 
  Stack, 
  Text, 
  IStackTokens, 
  getTheme, 
  mergeStyles, 
  AnimationClassNames,
  Icon
} from '@fluentui/react';

export interface IRegionSelectorProps {
  onSelectRegion: (region: 'VN' | 'ID_SG') => void;
}

const theme = getTheme();
const stackTokens: IStackTokens = { childrenGap: 40 };

const cardClass = mergeStyles(
  {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    width: '320px', // slightly wider for better stability
    height: '320px',
    backgroundColor: theme.palette.white,
    borderRadius: '16px',
    boxShadow: '0 8px 32px rgba(0, 0, 0, 0.06)',
    cursor: 'pointer',
    transition: 'all 0.4s cubic-bezier(0.25, 1, 0.5, 1)',
    border: `2px solid transparent`,
    position: 'relative',
    overflow: 'hidden',
    selectors: {
      ':hover': {
        transform: 'translateY(-12px)',
        boxShadow: '0 20px 40px rgba(0, 0, 0, 0.12)',
        borderColor: theme.palette.themePrimary,
      },
      ':active': {
        transform: 'translateY(-4px) scale(0.98)',
      }
    }
  },
  AnimationClassNames.slideUpIn20
);

const iconContainerClass = mergeStyles({
  width: '100px',
  height: '100px',
  borderRadius: '50%',
  backgroundColor: theme.palette.themeLighterAlt,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  marginBottom: '24px',
  transition: 'background-color 0.3s ease',
  selectors: {
    '.ms-Icon': {
      fontSize: '48px',
      color: theme.palette.themePrimary,
    }
  }
});

const regionTitleClass = mergeStyles({
  fontWeight: 800,
  fontSize: '28px', // Replacement for 'mega' which was too big
  color: theme.palette.neutralPrimary,
  letterSpacing: '1px',
  marginBottom: '4px',
});

export const RegionSelector: React.FC<IRegionSelectorProps> = (props) => {
  const { onSelectRegion } = props;

  return (
    <Stack 
      horizontalAlign="center" 
      verticalAlign="center" 
      style={{ minHeight: '500px', width: '100%', paddingBottom: '60px' }}
      tokens={{ childrenGap: 10 }}
    >
      <Text variant="superLarge" style={{ fontWeight: 800, color: theme.palette.neutralPrimary, textAlign: 'center' }}>
        Welcome to Leave Dashboard
      </Text>
      <Text variant="large" style={{ color: theme.palette.neutralSecondary, marginBottom: '40px', textAlign: 'center' }}>
        Please select your region to access your dashboard
      </Text>

      <Stack horizontal tokens={stackTokens} horizontalAlign="center" wrap>
        {/* Vietnam Card */}
        <div className={cardClass} onClick={() => onSelectRegion('VN')}>
          <div className={iconContainerClass}>
            <Icon iconName="World" />
          </div>
          <div className={regionTitleClass}>VIETNAM</div>
          <Text variant="mediumPlus" style={{ color: theme.palette.neutralSecondary, fontWeight: 500 }}>
            Longan Group Vietnam
          </Text>
        </div>

        {/* Indonesia & Singapore Card */}
        <div className={cardClass} onClick={() => onSelectRegion('ID_SG')}>
          <div className={iconContainerClass}>
            <Icon iconName="CompassNW" />
          </div>
          <div className={regionTitleClass}>ID & SG</div>
          <div style={{ textAlign: 'center' }}>
            <Text block variant="mediumPlus" style={{ color: theme.palette.neutralSecondary, fontWeight: 500 }}>
              Longan Group International
            </Text>
            <Text block variant="small" style={{ color: theme.palette.themePrimary, fontWeight: 600, marginTop: '4px' }}>
              Indonesia & Singapore
            </Text>
          </div>
        </div>
      </Stack>
    </Stack>
  );
};
