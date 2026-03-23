import * as React from 'react';
import styles from './CsiAndWorkingDays.module.scss';
import type { ICsiAndWorkingDaysProps } from './ICsiAndWorkingDaysProps';
import CSAT from './CSAT/CSAT';
import WorkingDays from './WorkingDays/WorkingDays';
import type { DashboardTabKey } from './types';

const CsiAndWorkingDays: React.FC<ICsiAndWorkingDaysProps> = ({ context }) => {
  const [activeTab, setActiveTab] = React.useState<DashboardTabKey>('csat');

  return (
    <div className={styles.csiAndWorkingDays}>
      {activeTab === 'csat' ? (
        <CSAT context={context} onSelectTab={setActiveTab} />
      ) : (
        <WorkingDays context={context} onSelectTab={setActiveTab} />
      )}
    </div>
  );
};

export default CsiAndWorkingDays;
