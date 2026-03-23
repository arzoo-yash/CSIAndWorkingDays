export const PPO_Approvers_Columns = [
    'Id',
    'ID',
    'LinkTitle',
    'Reviewer',
    'ReviewerId',
    'Reviewer/Id',
    'Reviewer/Title',
    'Reviewer/Name',
    'Reviewer/EMail',
    'ProjectManager',
    'ProjectManagerId',
    'ProjectManager/Id',
    'ProjectManager/Title',
    'ProjectManager/Name',
    'ProjectManager/EMail',
    'BUH',
    'BUHId',
    'BUH/Id',
    'BUH/Title',
    'BUH/Name',
    'BUH/EMail',
    'InternalProjectName'
]

export enum CurrentUserRole{
    Reviewer = 'Reviewer',
    ProjectManager = 'ProjectManager',
    BUH = 'BUH',
    Parent_Site_Visitor = 'Parent_Site_Visitor',
    None = 'None'
}

export const Parent_Site_Visitory_Permission_Group = 'Process Performance Objectives uat Visitors';