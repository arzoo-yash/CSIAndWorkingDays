/* eslint-disable */
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/sites';
import { SharePointService } from './SharePointService';
import { PPOApprovers } from '../../../common/IPPOApprovers';
import { CurrentUserRole, PPO_Approvers_Columns } from '../../../common/PPOApproversConst';
import { LIST_NAMES } from '../../../common/constants';

export interface IFetchOptions {
	listTitle?: string;
	select?: string[];
	expand?: string[];
	filter?: string;
	orderBy?: string;
	ascending?: boolean;
	top?: number;
	context?: WebPartContext;
	forceSiteUrl?: string;
}

export interface IPPOApproversRepo {
    init(context: WebPartContext): void;
    fetchPPOApprovers(options: IFetchOptions): Promise<PPOApprovers[]>;
}

export interface IPPOApproverPermission {
    userRole: CurrentUserRole;
    canEdit: boolean;
    canDelete: boolean;
    currentUserEmail: string;
}

/**
 * PPO Approvers Repository implementation
 */
class PPOApproversRepository implements IPPOApproversRepo {
	private _spService: SharePointService | null = null;
	private _context: WebPartContext | null = null;

	constructor(context?: WebPartContext) {
		if (context) {
			this._context = context;
			this._spService = new SharePointService(context);
		}
	}

	public init(context: WebPartContext): void {
		this._context = context;
		this._spService = new SharePointService(context);
	}

	public async fetchPPOApprovers(options: IFetchOptions): Promise<PPOApprovers[]> {
		if (!this._spService) {
			throw new Error('PPOApproversRepository not initialized. Call init() or pass context to constructor.');
		}

		const defaultSelect = PPO_Approvers_Columns;
		const defaultExpand = ['Reviewer', 'ProjectManager', 'BUH'];

		const listTitle = options.listTitle || LIST_NAMES.PPO_APPROVERS;
		
		// Merge select fields with defaults, removing duplicates manually
		const selectSet: {[key: string]: boolean} = {};
		[...(options.select || []), ...defaultSelect].forEach(field => {
			selectSet[field] = true;
		});
		const select = Object.keys(selectSet);
		
		// Merge expand fields with defaults, removing duplicates manually
		const expandSet: {[key: string]: boolean} = {};
		[...(options.expand || []), ...defaultExpand].forEach(field => {
			expandSet[field] = true;
		});
		const expand = Object.keys(expandSet);
		const siteUrl = options.forceSiteUrl || this._context?.pageContext?.web?.absoluteUrl;

		if (!siteUrl) {
			throw new Error('Unable to determine site URL for fetching PPOApprovers.');
		}

		try {
			const items = await this._spService.getListItemsFromSite(
				siteUrl,
				listTitle,
				select,
				expand,
				options.filter,
				options.top,
				options.orderBy,
				options.ascending
			);

			return items as PPOApprovers[];
		} catch (error) {
			console.error('Error fetching PPOApprovers:', error);
			throw error;
		}
	}
}

/**
 * Service to manage PPOApprovers permissions
 * This service fetches PPOApprovers from the root site and determines user permissions
 */
export class PPOApproverService {
    private context: WebPartContext;
    private ppoApproversRepo: PPOApproversRepository;
    private rootSiteUrl: string = '';
    private currentUserEmail: string = '';

    constructor(context: WebPartContext) {
        this.context = context;
        this.ppoApproversRepo = new PPOApproversRepository(context);
        
        // Get current user email
        this.currentUserEmail = context.pageContext.user.email?.toLowerCase() || '';
    }

    /**
     * Get the root site URL from the current context
     * This extracts the root site from the absolute URL
     */
    private async getRootSiteUrl(): Promise<string> {
        if (this.rootSiteUrl) {
            return this.rootSiteUrl;
        }

        try {
            // Get the current site absolute URL
            const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
            
            // Parse to get the root site URL
            // For example: https://tenant.sharepoint.com/sites/MySite/SubSite
            // Root would be: https://tenant.sharepoint.com/sites/MySite
            const url = new URL(currentSiteUrl);
            const pathSegments = url.pathname.split('/').filter(s => s);
            
            // Typically, root site structure is /sites/[sitename]
            if (pathSegments.length >= 2 && pathSegments[0] === 'sites') {
                this.rootSiteUrl = `${url.origin}/${pathSegments[0]}/${pathSegments[1]}`;
            } else {
                // If not in /sites/, use the origin as root
                this.rootSiteUrl = url.origin;
            }

            return this.rootSiteUrl;
        } catch (error) {
            console.error('Error determining root site URL:', error);
            // Fallback to current site
            return this.context.pageContext.web.absoluteUrl;
        }
    }

    /**
     * Fetch all PPOApprovers from the root site
     * @param projectName - Optional filter by internal project name
     */
    public async fetchPPOApprovers(projectName?: string): Promise<PPOApprovers[]> {
        try {
            const rootSiteUrl = await this.getRootSiteUrl();
            
            const options: IFetchOptions = {
                listTitle: LIST_NAMES.PPO_APPROVERS,
                forceSiteUrl: rootSiteUrl,
                context: this.context,
                filter: projectName ? `InternalProjectName eq '${projectName}'` : undefined
            };

            const approvers = await this.ppoApproversRepo.fetchPPOApprovers(options);
            return approvers;
        } catch (error) {
            console.error('Error fetching PPOApprovers from root site:', error);
            return [];
        }
    }

    /**
     * Determine the current user's role based on PPOApprovers list
     * Checks if current user is a Reviewer (SQA), Project Manager, or BUH
     * @param projectName - Optional filter by project name
     */
    public async getCurrentUserRole(projectName?: string): Promise<CurrentUserRole> {
        try {
            // If no projectName was passed, attempt to use the current web (site) title
            // This ensures callers that pass an empty/blank projectName still get
            // the PPOApprovers filtered by the current site title.
            if (!projectName) {
                const siteTitle = this.context?.pageContext?.web?.title?.trim();
                if (siteTitle) {
                    projectName = siteTitle;
                }
            }

            const approvers = await this.fetchPPOApprovers(projectName);
            
            if (!approvers || approvers.length === 0) {
                console.warn('No PPOApprovers found');
                return CurrentUserRole.None;
            }

            // Helper to compare a people field (single object, array, or { results: [...] })
            const emailMatches = (field: any): boolean => {
                if (!field) return false;

                // If field is an array of users
                if (Array.isArray(field)) {
                    return field.some((u: any) => {
                        const em = (u && (u.EMail || u.Email || u.EmailAddress || u.email));
                        return typeof em === 'string' && em.toLowerCase() === this.currentUserEmail;
                    });
                }

                // If field has a results array (common for multi-value person fields)
                if (field.results && Array.isArray(field.results)) {
                    return field.results.some((u: any) => {
                        const em = (u && (u.EMail || u.Email || u.EmailAddress || u.email));
                        return typeof em === 'string' && em.toLowerCase() === this.currentUserEmail;
                    });
                }

                // If field is a single user object
                if (typeof field === 'object') {
                    const em = (field.EMail || field.Email || field.EmailAddress || field.email);
                    if (typeof em === 'string' && em.toLowerCase() === this.currentUserEmail) {
                        return true;
                    }

                    // Some PnP shapes may include nested user objects or arrays under other keys
                    // Try to find any property that looks like an email
                    for (const key of Object.keys(field)) {
                        const val = (field as any)[key];
                        if (typeof val === 'string' && val.indexOf('@') !== -1 && val.toLowerCase() === this.currentUserEmail) {
                            return true;
                        }
                    }
                }

                return false;
            };

            // Check if user is in any of the roles across all projects (or specific project)
            for (const approver of approvers) {
                if (emailMatches(approver.Reviewer)) {
                    return CurrentUserRole.Reviewer;
                }

                if (emailMatches(approver.ProjectManager)) {
                    return CurrentUserRole.ProjectManager;
                }

                if (emailMatches(approver.BUH)) {
                    return CurrentUserRole.BUH;
                }
            }

            // If not found in any role, return None
            return CurrentUserRole.None;
        } catch (error) {
            console.error('Error determining current user role:', error);
            return CurrentUserRole.None;
        }
    }

    /**
     * Get comprehensive permission information for the current user
     * @param projectName - Optional filter by project name
     */
    public async getUserPermissions(projectName?: string): Promise<IPPOApproverPermission> {
        const userRole = await this.getCurrentUserRole(projectName);
        
        // Define permission rules
        // Project Manager, BUH, and SQA (Reviewer) can edit and delete
        const canEdit = [
            CurrentUserRole.ProjectManager,
            CurrentUserRole.BUH,
            CurrentUserRole.Reviewer
        ].indexOf(userRole) !== -1;

        const canDelete = [
            CurrentUserRole.ProjectManager,
            CurrentUserRole.BUH,
            CurrentUserRole.Reviewer
        ].indexOf(userRole) !== -1;

        return {
            userRole,
            canEdit,
            canDelete,
            currentUserEmail: this.currentUserEmail
        };
    }

    /**
     * Check if current user has edit permission
     */
    public async canUserEdit(projectName?: string): Promise<boolean> {
        const permissions = await this.getUserPermissions(projectName);
        return permissions.canEdit;
    }

    /**
     * Check if current user has delete permission
     */
    public async canUserDelete(projectName?: string): Promise<boolean> {
        const permissions = await this.getUserPermissions(projectName);
        return permissions.canDelete;
    }
}
