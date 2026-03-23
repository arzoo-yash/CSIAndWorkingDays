/* eslint-disable */
export interface PPOApprovers {
    Id: number;
    ID: number;
    LinkTitle: string;
    InternalProjectName: string;
    
    // Reviewer fields
    ReviewerId?: number;
    Reviewer?: {
        Id: number;
        Title: string;
        Name: string;
        EMail: string;
    };
    
    // Project Manager fields
    ProjectManagerId?: number;
    ProjectManager?: {
        Id: number;
        Title: string;
        Name: string;
        EMail: string;
    };
    
    // BUH (Business Unit Head) fields
    BUHId?: number;
    BUH?: {
        Id: number;
        Title: string;
        Name: string;
        EMail: string;
    };
}
