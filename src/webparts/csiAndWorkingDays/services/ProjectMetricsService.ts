import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { LIST_NAMES } from '../../../common/constants';

export interface IProjectMetricConfig {
  Goal?: string;
  USL?: number;
  LSL?: number;
}

/**
 * Service to fetch project metrics configuration from ProjectMetrics and ProjectMetricLogs lists
 */
export class ProjectMetricsService {
  private spService: SharePointService;

  constructor(context: WebPartContext) {
    this.spService = new SharePointService(context);
  }

  /**
   * Get the latest active approved version of Customer Satisfaction Index metrics
   * Fetches from ProjectMetricLogs (Status='Approved', IsActive=true)
   * then queries ProjectMetrics where VersionId lookup matches the ProjectMetricLogs ID
   * @returns Promise with metric configuration (Goal, USL, LSL)
   */
  public async getLatestCSATMetrics(): Promise<IProjectMetricConfig> {
    try {
      // Step 1: Query ProjectMetricLogs for the latest active approved version
      const logSelect = [
        'Id',
        'VersionId',
        'Status',
        'IsActive'
      ];

      // Query all Approved items (don't filter by IsActive in query as it's unreliable)
      const logFilter = "Status eq 'Approved'";

      let logItems = await this.spService.getListItems(
        LIST_NAMES.ProjectMetricLogs,
        logSelect,
        undefined,
        logFilter,
        50, // Get more items to ensure we find the right one
        'Id',
        false // Descending order
      );

      if (!logItems || logItems.length === 0) {
        console.warn('[ProjectMetricsService] No approved CSAT metrics found in ProjectMetricLogs');
        return {
          Goal: undefined,
          USL: undefined,
          LSL: undefined
        };
      }

      // Find the item with highest ID (latest) that has IsActive=true
      let maxIdActiveItem: any = null;
      for (const item of logItems) {
        if (item.IsActive === true || item.IsActive === 1 || item.IsActive === '1' || item.IsActive === 'Yes') {
          if (!maxIdActiveItem || item.Id > maxIdActiveItem.Id) {
            maxIdActiveItem = item;
          }
        }
      }
      
      if (!maxIdActiveItem) {
        console.warn('[ProjectMetricsService] No item with IsActive=true found');
        return {
          Goal: undefined,
          USL: undefined,
          LSL: undefined
        };
      }

      const approvedLogId = maxIdActiveItem.Id;

      // Step 2: Query ProjectMetrics where VersionId (lookup) equals the ProjectMetricLogs ID
      // Note: Not filtering by IsActive in ProjectMetrics as it may be a number field
      const metricsSelect = [
        'Id',
        'Title',
        'Metrics',
        'Goal',
        'USL',
        'LSL',
        'IsActive',
        'VersionId/Id'
      ];

      const metricsExpand = ['VersionId'];
      const metricsFilter = `VersionId/Id eq ${approvedLogId}`;

      const metricsItems = await this.spService.getListItems(
        LIST_NAMES.ProjectMetrics,
        metricsSelect,
        metricsExpand,
        metricsFilter,
        50,
        'Id',
        false
      );

      if (!metricsItems || metricsItems.length === 0) {
        console.warn('[ProjectMetricsService] No metrics found for approved ProjectMetricLogs ID:', approvedLogId);
        return {
          Goal: undefined,
          USL: undefined,
          LSL: undefined
        };
      }

      // Step 3: Find the Customer Satisfaction Index metric (case-insensitive)
      let csatMetric: any = null;
      for (const metric of metricsItems) {
        const metricName = (metric.Metrics || metric.Title || '').toString().toLowerCase();
        // Search for "customer satisfaction" in a case-insensitive way
        if (metricName.includes('customer satisfaction') || metricName.includes('customer-satisfaction') || metricName.includes('customersatisfaction')) {
          csatMetric = metric;
          break;
        }
      }

      if (csatMetric) {
        return {
          Goal: csatMetric.Goal,
          USL: csatMetric.USL !== null && csatMetric.USL !== undefined ? Number(csatMetric.USL) : undefined,
          LSL: csatMetric.LSL !== null && csatMetric.LSL !== undefined ? Number(csatMetric.LSL) : undefined
        };
      }

      // If no Customer Satisfaction Index found, return empty config
      console.warn('[ProjectMetricsService] No Customer Satisfaction Index metric found');
      return {
        Goal: undefined,
        USL: undefined,
        LSL: undefined
      };
    } catch (error) {
      console.error('[ProjectMetricsService] Error fetching CSAT metrics:', error);
      return {
        Goal: undefined,
        USL: undefined,
        LSL: undefined
      };
    }
  }

  /**
   * Get all project metrics from ProjectMetrics list
   * @returns Promise with array of project metrics
   */
  public async getAllProjectMetrics(): Promise<any[]> {
    try {
      const select = [
        'Id',
        'Title',
        'Metrics',
        'Goal',
        'USL',
        'LSL',
        'IsActive',
        'VersionId/Id',
        'Modified'
      ];

      const expand = ['VersionId'];

      const items = await this.spService.getListItems(
        LIST_NAMES.ProjectMetrics,
        select,
        expand,
        undefined,
        100,
        'Modified',
        false
      );
      
      return items || [];
    } catch (error) {
      console.error('[ProjectMetricsService] Error fetching project metrics:', error);
      return [];
    }
  }

  /**
   * Get all project metric logs from ProjectMetricLogs list
   * @returns Promise with array of project metric logs
   */
  public async getAllProjectMetricLogs(): Promise<any[]> {
    try {
      const select = [
        'Id',
        'VersionId',
        'Status',
        'IsActive',
        'PMComments',
        'ReviewerComments',
        'BUHComments',
        'CreatedVersion',
        'Modified'
      ];

      const items = await this.spService.getListItems(
        LIST_NAMES.ProjectMetricLogs,
        select,
        undefined,
        undefined,
        100,
        'Modified',
        false
      );
      
      return items || [];
    } catch (error) {
      console.error('[ProjectMetricsService] Error fetching project metric logs:', error);
      return [];
    }
  }

  /**
   * Check if there is an approved and active entry in ProjectMetricLogs
   * and retrieve the associated Customer Satisfaction Index metrics from ProjectMetrics
   * Note: ProjectMetrics.VersionId is a lookup to ProjectMetricLogs.ID
   * @returns Promise with validation result and metrics data
   */
  public async validateAndGetCSATMetrics(): Promise<{
    isValid: boolean;
    message?: string;
    metrics?: IProjectMetricConfig;
  }> {
    try {
      // Step 1: Query ProjectMetricLogs for approved entries
      const logSelect = [
        'Id',
        'VersionId',
        'Status',
        'IsActive'
      ];

      // Query all Approved items (don't filter by IsActive in query as it's unreliable)
      const logFilter = "Status eq 'Approved'";

      let logItems = await this.spService.getListItems(
        LIST_NAMES.ProjectMetricLogs,
        logSelect,
        undefined,
        logFilter,
        50, // Get more items to ensure we find the right one
        'Id',
        false // Descending order to get highest ID first
      );

      if (!logItems || logItems.length === 0) {
        return {
          isValid: false,
          message: 'PPO (Process Performance Objectives) are not submitted. Please submit PPO first then add CSAT entry.'
        };
      }

      // Find the item with highest ID (latest) that has IsActive=true
      let maxIdActiveItem: any = null;
      for (const item of logItems) {
        if (item.IsActive === true || item.IsActive === 1 || item.IsActive === '1' || item.IsActive === 'Yes') {
          if (!maxIdActiveItem || item.Id > maxIdActiveItem.Id) {
            maxIdActiveItem = item;
          }
        }
      }
      
      if (!maxIdActiveItem) {
        return {
          isValid: false,
          message: 'PPO (Process Performance Objectives) are not submitted. Please submit PPO first then add CSAT entry.'
        };
      }

      // Step 2: Get the approved active ProjectMetricLogs entry ID (highest ID)
      const approvedLogId = maxIdActiveItem.Id;

      // Step 3: Query ProjectMetrics where VersionId (lookup) equals the ProjectMetricLogs ID
      const metricsSelect = [
        'Id',
        'Title',
        'Metrics',
        'Goal',
        'USL',
        'LSL',
        'IsActive',
        'VersionId/Id'
      ];

      const metricsExpand = ['VersionId'];
      const metricsFilter = `VersionId/Id eq ${approvedLogId}`;

      const metricsItems = await this.spService.getListItems(
        LIST_NAMES.ProjectMetrics,
        metricsSelect,
        metricsExpand,
        metricsFilter,
        50,
        'Id',
        false
      );

      if (!metricsItems || metricsItems.length === 0) {
        return {
          isValid: false,
          message: 'Customer Satisfaction Index metric not found in approved PPO. Please configure PPO first.'
        };
      }

      // Step 4: Find the Customer Satisfaction Index metric (case-insensitive search)
      let csatMetric: any = null;
      for (const metric of metricsItems) {
        const metricName = (metric.Metrics || metric.Title || '').toString().toLowerCase();
        if (metricName.includes('customer satisfaction') || metricName.includes('customer-satisfaction') || metricName.includes('customersatisfaction')) {
          csatMetric = metric;
          break;
        }
      }

      if (!csatMetric) {
        return {
          isValid: false,
          message: 'Customer Satisfaction Index metric not found in approved PPO. Please configure PPO first.'
        };
      }

      return {
        isValid: true,
        metrics: {
          Goal: csatMetric.Goal,
          USL: csatMetric.USL !== null && csatMetric.USL !== undefined ? Number(csatMetric.USL) : undefined,
          LSL: csatMetric.LSL !== null && csatMetric.LSL !== undefined ? Number(csatMetric.LSL) : undefined
        }
      };

    } catch (error) {
      console.error('[ProjectMetricsService] Error validating CSAT metrics:', error);
      return {
        isValid: false,
        message: 'Error checking PPO status. Please try again or contact support.'
      };
    }
  }
}

export default ProjectMetricsService;
