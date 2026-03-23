import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './Header.module.scss';

export interface IHeaderTab {
  key: string;
  label: string;
}

export interface IHeaderProps {
  title: string;
  subtitle?: string;
  iconName?: string;
  metaText?: string;
  tabs?: IHeaderTab[];
  activeTab?: string;
  onTabChange?: (tab: string) => void;
  primaryActionLabel?: string;
  primaryActionIconName?: string;
  onPrimaryAction?: () => void;
  className?: string;
}

const Header: React.FC<IHeaderProps> = ({
  title,
  subtitle,
  iconName = 'Bullseye',
  tabs = [],
  activeTab,
  onTabChange,
  primaryActionLabel,
  primaryActionIconName = 'Add',
  onPrimaryAction,
  className
}) => {
  const initialTab = React.useMemo(() => {
    if (activeTab) {
      return activeTab;
    }

    return tabs.length > 0 ? tabs[0].key : undefined;
  }, [activeTab, tabs]);

  const [selectedTab, setSelectedTab] = React.useState<string | undefined>(initialTab);

  React.useEffect(() => {
    if (activeTab !== undefined) {
      setSelectedTab(activeTab);
    }
  }, [activeTab]);

  React.useEffect(() => {
    if (activeTab === undefined && tabs.length > 0 && selectedTab === undefined) {
      setSelectedTab(tabs[0].key);
    }
  }, [activeTab, tabs, selectedTab]);

  const handleTabClick = React.useCallback((tabKey: string) => {
    setSelectedTab(tabKey);
    if (onTabChange) {
      onTabChange(tabKey);
    }
  }, [onTabChange]);

  return (
    <header className={`${styles.header} ${className ?? ''}`.trim()} role="banner">
      <div className={styles.hero}>
        <div className={styles.branding}>
          <span className={styles.iconWrap} aria-hidden="true">
            <Icon iconName={iconName} />
          </span>
          <div className={styles.textGroup}>
            <span className={styles.title}>{title}</span>
            {subtitle && <span className={styles.subtitle}>{subtitle}</span>}
          </div>
        </div>
      </div>

      {tabs.length > 0 && (
        <div className={styles.tabRow} role="tablist">
          {tabs.map(tab => {
            const isActive = tab.key === selectedTab;
            return (
              <button
                key={tab.key}
                type="button"
                className={`${styles.tabButton} ${isActive ? styles.activeTab : ''}`.trim()}
                onClick={() => handleTabClick(tab.key)}
                role="tab"
                aria-selected={isActive}
              >
                {tab.label}
              </button>
            );
          })}
        </div>
      )}

      {primaryActionLabel && (
        <div className={styles.actionRow}>
          <button
            type="button"
            className={styles.primaryAction}
            onClick={() => { if (onPrimaryAction) { onPrimaryAction(); } }}
          >
            <Icon iconName={primaryActionIconName} className={styles.primaryActionIcon} />
            <span>{primaryActionLabel}</span>
          </button>
        </div>
      )}
    </header>
  );
};

export default Header;
