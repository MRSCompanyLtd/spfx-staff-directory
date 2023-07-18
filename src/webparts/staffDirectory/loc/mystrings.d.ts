declare interface IStaffDirectoryWebPartStrings {
  PropertyPaneTitle: string;
  TitleFieldLabel: string;
  PageSizeFieldLabel: string;
  GroupSelectFieldLabel: string;
  DepartmentListFieldLabel: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'StaffDirectoryWebPartStrings' {
  const strings: IStaffDirectoryWebPartStrings;
  export = strings;
}
