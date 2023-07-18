import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react';

export interface IStaffDirectoryProps {
  title: string;
  group: string;
  departments: IDropdownOption[];
  showDepartmentFilter: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  pageSize: number;
  context: WebPartContext;
}
